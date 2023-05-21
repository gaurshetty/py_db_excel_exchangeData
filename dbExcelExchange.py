from py_db_excel_exchangeData.servisesImpl import ServisesImpl
from py_db_excel_exchangeData.empClass import Emp
x = ServisesImpl()


def db_to_excel(eid):
    e1 = x.db_get_one_employee(eid)
    e2 = Emp(e1[0], e1[1], e1[2], e1[3], e1[4])
    x.xl_add_update_employee(e2)
    print('Emp record transferred successfully!')


def excel_to_db(eid):
    e1 = x.xl_get_one_employee(eid)
    e2 = Emp(e1[0], e1[1], e1[2], e1[3], e1[4])
    x.db_add_update_employee(e2)
    print('Emp record transferred successfully!')


def db_to_excel_all():
    alldata = x.db_get_all_employee()
    if alldata:
        for e1 in alldata:
            e2 = Emp(e1[0], e1[1], e1[2], e1[3], e1[4])
            x.xl_add_update_employee(e2)
        print('All Emp record transferred successfully!')
    else:
        print('No records in Emp table...')


def excel_to_db_all():
    alldata = x.xl_get_all_employee()
    if alldata:
        for e1 in alldata:
            e2 = Emp(e1[0], e1[1], e1[2], e1[3], e1[4])
            x.db_add_update_employee(e2)
        print('All Emp record transferred successfully!')
    else:
        print('No records in Emp table...')


