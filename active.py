import datetime
import smtplib
from email.header import Header
from email.mime.text import MIMEText
from email.utils import parseaddr, formataddr
import time
import xlrd
from .SMS import sms_send
from .chinese_calendar.utils import is_workday, is_holiday
from .TimerTask import timer

# 全局数据库链接
from .sms_query import create_init, data_query
from duanxinstone import stoneobject, EmployeeInfo, Birthlist, Divisionlist, DivisionTable
from sqlalchemy import text, and_

global from_addr, password, to_addr, error_addr
from_addr = input("请输入发件人邮箱")
password = input("请输入发件人邮箱密码")
to_addr = input("请输入收件人邮箱")
error_addr = input("请输入异常收件人邮箱")
stone = stoneobject()
def init():
    # 清空所有数据
    stone.query(EmployeeInfo).delete()
    stone.query(Birthlist).delete()
    stone.query(Divisionlist).delete()
    stone.query(DivisionTable).delete()
    stone.commit()
    crete_emp_info()
    preprocessing()

# 单元格为空，则存入None
def one_or_none(obj):
    if obj == '':
        return None
    else:
        return obj

def crete_emp_info():
    workbook = xlrd.open_workbook('祝福短信人员.xlsx')
    cols = ['code', 'name', 'enterdate', 'Divisiondates', 'birthDate', 'Tel', 'leaveDate']
    for count, one in enumerate(workbook.sheet_names()):
        for i in range(1, workbook.sheet_by_name(one).nrows):
            if workbook.sheet_by_name(one).ncols > 7:
                print('错误')
                break
            empinfo = EmployeeInfo()
            setattr(empinfo, 'id', i)
            setattr(empinfo, 'Cover', None)
            print(i)
            for j, col in enumerate(cols):
                print(j)
                if col == 'enterdate' or col == 'Divisiondates' or col == 'birthDate' or col == 'leaveDate':
                    if one_or_none(workbook.sheet_by_name(one)._cell_values[i][j]):
                        setattr(empinfo, col, datetime.datetime.strptime(
                            workbook.sheet_by_name(one)._cell_values[i][j], '%Y-%m-%d'))
                    else:
                        setattr(empinfo, col, None)
                elif col == 'code' or col == 'Tel':
                    # 0开始的工号处理
                    if len(str(int(one_or_none(workbook.sheet_by_name(one)._cell_values[i][j])))) < 10 and col == 'code':
                        setattr(empinfo, col, '0' + str(int(one_or_none(workbook.sheet_by_name(one)._cell_values[i][j]))))
                    else:
                        setattr(empinfo, col, int(one_or_none(workbook.sheet_by_name(one)._cell_values[i][j])))
                elif col == 'name':
                    setattr(empinfo, col, one_or_none(workbook.sheet_by_name(one)._cell_values[i][j]))
            # print(empinfo.__dict__)
            # print(type(empinfo))
            stone.add(empinfo)
        stone.commit()

# 将人员信息放置Divisionlist
def preprocessing():
    result = stone.query(EmployeeInfo).filter(EmployeeInfo.leaveDate != None).all()
    for one in result:
        # print(one)
        one.Cover = (one.leaveDate-one.Divisiondates).days
    stone.commit()
    result = stone.query(EmployeeInfo).filter(EmployeeInfo.leaveDate == None).all()
    for one in result:
        # print(one)
        one.Cover = 0
    stone.commit()
    result = stone.query(EmployeeInfo).all()
    for one in result:
        division = Divisionlist()
        division.id = one.id
        division.name = one.name
        division.code = one.code
        division.realityenterdate = one.enterdate-datetime.timedelta(days=one.Cover)
        division.Tel = one.Tel
        stone.add(division)
    stone.commit()


# 如果明天是非工作日，继续，一直到工作日
def workexec(today):
    # march_first = datetime.date.today()
    # i=0
    global Days
    Days = 0
    march_first = today
    while True:
        print(is_workday(march_first))  # False
        print(is_holiday(march_first))  # True
        Days = Days+1
        march_first = today+datetime.timedelta(days=Days)
        if is_workday(march_first):
            break

# 生日祝福转存，司龄转存
# 装饰器待开发
def unloading(i, today):
    # (datetime.date.today()-EmployeeInfo.birthDate).year
    # print(datetime.datetime.strftime(datetime.date.today(),'%m-%d'))
    # print(EmployeeInfo.__dict__)
    # result=stone.query(EmployeeInfo).filter(EmployeeInfo.birthDate.like('%'+datetime.datetime.strftime(datetime.date.today(),'%m-%d')+'%'))
    # i=0
    # 2-29日已处理
    result = stone.query(EmployeeInfo).filter(
        text("strftime('%m%d',DATE (birthDate,'1 day'))=strftime('%m%d',date(:date,:value)) ")).params(
        value='{0} day'.format(i+1), date=today).all()
    # result=stone.query(EmployeeInfo).filter(text("id=(':value')")).params(value=224)
    # print(result.one())
    for one in result:
        # print(one)
        birth = Birthlist()
        # 处理后缀数值
        try:
            int(one.name[len(one.name)-1])
            birth.name = one.name[:len(one.name)-1]
        except ValueError:
            birth.name = one.name
        birth.code = one.code
        birth.birthDate = one.birthDate
        birth.Tel = one.Tel
        birth.flagnum=today.month
        birth.date = today + datetime.timedelta(days=i)
        birth.status = True
        # 可能会出现重复值
        stone.add(birth)
    stone.commit()
    # print('_________')
    result = stone.query(Divisionlist).filter(text(
        "strftime('%m%d',DATE (realityenterdate,'1 day'))=strftime('%m%d',date(:date,:value)) ")).params(
        value='{0} day'.format(i+1), date=today).all()
    for one in result:
        # print(one)
        division = DivisionTable()
        # 处理后缀数值
        try:
            int(one.name[len(one.name)-1])
            division.name = one.name[:len(one.name)-1]
        except ValueError:
            division.name = one.name
        division.code = one.code
        division.realityenterdate = one.realityenterdate
        division.Tel = one.Tel
        division.flagnum = today.year - one.realityenterdate.year
        division.date = today + datetime.timedelta(days=i)
        if division.flagnum > 1:
            division.status = True
        else:
            division.status = None
        # 可能会出现重复值
        stone.add(division)
    stone.commit()
    # print("____")
    # print(stone.query(EmployeeInfo).filter(EmployeeInfo.birthDate.like('09-05')))
    # Result=stone.execute("SELECT * FROM EmployeeInfo where strftime('%m%d',birthDate)"
    #                      "=strftime('%m%d',date('now','0 day')) ;")
    # for one in Result:
    #     print(one)

def _format_addr(s):
    name, addr = parseaddr(s)
    return formataddr((Header(name, 'utf-8').encode(), addr))

#     邮件发送
def send(header, body, to_address=to_addr):
    # smtp_server =r'smtp.qq.com'
    # 企业邮箱
    smtp_server = r'smtp.qiye.163.com'
    # smtp_server = r'smtp.163.com'

    # 正文
    msg = MIMEText(body, 'html', 'utf-8')
    # 主题，
    msg['Subject'] = Header(header, 'utf-8').encode()
    # 发件人别名
    msg['From'] = _format_addr('祝福管理站 <%s>' % from_addr)
    # 收件人别名
    msg['To'] = to_address

    # server = smtplib.SMTP(smtp_server, 25)
    # server.login(from_addr, password)
    server = smtplib.SMTP_SSL(smtp_server, 465)
    # QQ SSL 端口 587
    # 网易 SSL端口 994、465
    # server.starttls()
    server.login(from_addr, password)
    # server.set_debuglevel(1)
    server.sendmail(from_addr, to_address.split(','), msg.as_string())
    server.quit()
    print('发送成功')
# 邮件HTML发送
def draw(today):
    # HTML定义
    html = r"""<html>
<body>"""
    # 生日表格
    shengri = """<table width="580"  border="1" align="center">
    <caption>
    生日名单
    </caption>
    <tbody>
        <tr>
            <th width="20%" scope="col">工号</th>
            <th width="20%" scope="col">姓名</th>
            <th width="20%" scope="col">发送日期</th>
            <th width="20%" scope="col">生日日期</th>
            <th width="20%" scope="col">电话号码</th>
        </tr>"""
    brithstr = ''
    for i in range(Days):
        result = stone.query(Birthlist).filter(
            and_(Birthlist.date == today+datetime.timedelta(days=i), Birthlist.status == True)).all()
        for one in result:
            # print(one)
            str1 = r'<tr>      <td width="20%" align="center">{0}</td>      ' \
                   r'<td width="20%" align="center">{1}</td>  ' \
                   r'    <td width="20%" align="center">{2}</td>    ' \
                   r'<td width="20%" align="center">{3}</td>   ' \
                   r'<td width="20%" align="center">{4}</td> </tr>'\
                .format(one.code, one.name, one.date, one.birthDate, one.Tel)
            # print(str1)
            brithstr = brithstr + str1
    if brithstr == '':
        shengri = shengri + '<tr><td colspan=5 align="center">无人员</td></tr>'+'</tbody> </table>'
    else:
        shengri = shengri + brithstr + '  </tbody> </table>'
    # 司龄
    siling = """
    <table width="580" border="1" align="center"><caption>
      司龄名单
    </caption>
    <tbody>
      <tr>
        <th width="20%" scope="col">工号</th>
        <th width="20%" scope="col">姓名</th>
        <th width="20%" scope="col">发送日期</th>
        <th width="20%" scope="col">司龄</th>
        <th width="20%" scope="col">电话号码</th>
      </tr>"""
    slstr = ''
    for i in range(Days):
        result = stone.query(DivisionTable).filter(
            and_(DivisionTable.date == today + datetime.timedelta(days=i), DivisionTable.status == True)).all()
        for one in result:
            # print(one)
            str1 = r'<tr>      <td width="20%" align="center">{0}</td>      <td width="20%" align="center">{1}</td>  ' \
                 r'    <td width="20%" align="center">{2}</td>    <td width="20%" align="center">{3}</td>   ' \
                 r'<td width="20%" align="center">{4}</td> </tr>'.format(one.code, one.name, one.date, one.flagnum, one.Tel)
            # print(str1)
            slstr = slstr + str1
    if slstr == '':
        siling = siling + '<tr><td colspan=5 align="center">无人员</td></tr>' + '</tbody> </table>'
    else:
        siling = siling+slstr + '  </tbody></table>'
    wishbrithstr = html + shengri + siling + '</body></html>'
    print(wishbrithstr)
    # print('生日祝福短信 {0}'.format(datetime.date.today()))
    send("祝福短信{day}".format(day=datetime.date.today()), wishbrithstr)

def clear_stone(today):
    for i in range(Days):
        stone.query(Birthlist).filter(
            and_(Birthlist.date == today + datetime.timedelta(days=i), Birthlist.status == True)).delete()
        stone.query(DivisionTable).filter(
            and_(DivisionTable.date == today + datetime.timedelta(days=i), DivisionTable.status == True)).delete()

def smsdraw(today):
    brithstr=data_query(stone, brith, Birthlist,today)
    i=sms_send(brithstr[0],brithstr[1])
    silingstr=data_query(stone, siling, DivisionTable,today)
    j=sms_send(silingstr[0], silingstr[1])
    sendstr="共计有{sum}条短信没有发出，其中生日有{brith}没有发出，司龄有{siling}没有发出".format(
        sum=i + j , brith=i,siling=j)
    print(sendstr)
    if i + j != 0:
        send("祝福短信{day}".format(day=datetime.date.today()),sendstr,to_address=error_addr)
    pass

if __name__ == '__main__':
    init()
    Targettimestr = input('请输入定点时间，例如8:00')
    siling,brith=create_init()
    if Targettimestr == '':
        Targettimestr='8:00'
    Targettime = datetime.time(int(Targettimestr.split(':')[0]), int(Targettimestr.split(':')[1]))
    while True:
        if datetime.datetime.now().hour == Targettime.hour:
            date = datetime.date.today()+datetime.timedelta(days=0)
            workexec(date)
            for i in range(Days):
                unloading(i, date)
            if is_workday(date):
                draw(date)
            # smsdraw
            smsdraw(date)
            clear_stone(date)
            print(timer(Targettime))
            time.sleep(timer(Targettime))
        else:
            print(timer(Targettime))
            time.sleep(timer(Targettime))