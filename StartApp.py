from py_db_excel_exchangeData.servisesImpl import ServisesImpl
from py_db_excel_exchangeData.dbExcelExchange import *
x = ServisesImpl()

if __name__ == '__main__':

    while True:
        print("""Select Database or Excel:
        1: Database
        2: Excel
        3: Database to Excel
        4: Database to Excel all data
        5: Excel to Database
        6: Excel to Database all data
        7: Exit
        """)

        choice = int(input('Enter your choice: '))
        if choice == 1:
            while True:
                print("""Select Database operation:
                    1: In Database createTable,
                    2: In Database addEmployee,
                    3: In Database updateEmployee,
                    4: In Database deleteEmployee,
                    5: In Database getSingleEmployee,
                    6: In Database getAllEmployee,
                    7: In Database dropTable
                    8: Exit
                    """)
                choice = int(input('Enter your choice: '))
                if choice == 1:
                    x.db_create_table()
                elif choice == 2:
                    emp1 = x.empdata()
                    x.db_add_employee(emp1)
                elif choice == 3:
                    emp1 = x.empdata()
                    x.db_update_employee(emp1)
                elif choice == 4:
                    eid = int(input('Enter employee Id: '))
                    x.db_delete_employee(eid)
                elif choice == 5:
                    eid = int(input('Enter employee Id: '))
                    x.db_get_one_employee(eid)
                elif choice == 6:
                    x.db_get_all_employee()
                elif choice == 7:
                    x.db_drop_table()
                elif choice == 8:
                    break

                ch = input('Do you want to continue... Y | N : ')
                if ch == 'N' or ch == 'n':
                    break

        elif choice == 2:
            while True:
                print("""Select Excel operation:
                1: In Excel createTable,
                2: In Excel addEmployee,
                3: In Excel updateEmployee,
                4: In Excel deleteEmployee,
                5: In Excel getSingleEmployee,
                6: In Excel getAllEmployee,
                7: In Excel dropTable
                8: Exit
                """)
                choice = int(input('Enter your choice: '))
                if choice == 1:
                    x.xl_create_table()
                elif choice == 2:
                    e1 = x.empdata()
                    x.xl_add_employee(e1)
                elif choice == 3:
                    emp1 = x.empdata()
                    x.xl_update_employee(emp1)
                elif choice == 4:
                    eid = int(input('Enter employee Id: '))
                    x.xl_delete_employee(eid)
                elif choice == 5:
                    eid = int(input('Enter employee Id: '))
                    x.xl_get_one_employee(eid)
                elif choice == 6:
                    x.xl_get_all_employee()
                elif choice == 7:
                    x.xl_drop_table()
                elif choice == 8:
                    break

                ch = input('Do you want to continue... Y | N : ')
                if ch == 'N' or ch == 'n':
                    break

        elif choice == 3:
            eid = int(input('Enter employee Id: '))
            db_to_excel(eid)

        elif choice == 4:
            db_to_excel_all()

        elif choice == 5:
            eid = int(input('Enter employee Id: '))
            excel_to_db(eid)

        elif choice == 6:
            excel_to_db_all()

        elif choice == 7:
            break
