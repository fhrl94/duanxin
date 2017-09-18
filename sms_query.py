import datetime
from sqlalchemy import and_

def create_init():
    siling = []
    brith = []
    file = open('司龄祝福', 'rb')
    for f in file.readlines():
        siling.append(f.decode('utf-8').replace('\r\n', '', -1))
    file = open('生日祝福', 'rb')
    for f in file.readlines():
        brith.append(f.decode('utf-8').replace('\r\n', '', -1))
    return siling,brith

def data_query(stone,strarrary, datatable,today):
    smsarray=[]
    smstel=[]
    result = stone.query(datatable).filter(
        and_(datatable.date == today, datatable.status == True)).all()
    for one in result:
        smsarray.append(strarrary[one.flagnum - 1].format(Name=one.name,
                                               Day=datetime.date.today().strftime('%Y{y}%m{m}%d{d}').format(y='年',
                                                   m='月', d='日')))
        smstel.append(one.Tel)
    return (smstel,smsarray)