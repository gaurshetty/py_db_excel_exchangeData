from py_db_excel_exchangeData.servises import Services
from py_db_excel_exchangeData.AppQueries import *
from py_db_excel_exchangeData.dbExcelConnection import *
from py_db_excel_exchangeData.empClass import Emp


class ServisesImpl(Services):
    def db_create_table(self):
        dbExecution(query_createTable)
        print('Database table created successfully!')

    def db_get_one_employee1(self, EmpId):
        connect, channel = dbExecuteForFetch(query_getOneEmp.format(EmpId))
        record = channel.fetchone()
        return record

    def db_add_employee(self, emp):
        if type(emp) == Emp:
            if self.db_get_one_employee1(emp.EmpId):
                print("Employee record already Exist...")
                return
            dbExecution(query_addEmp.format(emp.EmpId, emp.EmpName, emp.EmpAge, emp.EmpSalary, emp.EmpAddress))
            print('Emp record added Successfully!')
        else:
            print('Invalid Emp type to add...')

    def db_update_employee(self, emp):
        if type(emp) == Emp:
            if self.db_get_one_employee1(emp.EmpId):
                dbExecution(query_updateEmp.format(emp.EmpName, emp.EmpAge, emp.EmpSalary, emp.EmpAddress, emp.EmpId))
                print('Emp record Updated Successfully!')
            else:
                print('No Emp record so can not update...')
        else:
            print('Invalid Emp type to update...')

    def db_delete_employee(self, EmpId):
        if self.db_get_one_employee1(EmpId):
            dbExecution(query_deleteEmp.format(EmpId))
            print('Emp record deleted successfully!')
        else:
            print('No Emp record so can not delete...')

    def db_get_one_employee(self, EmpId):
        connect, channel = dbExecuteForFetch(query_getOneEmp.format(EmpId))
        record = channel.fetchone()
        if record:
            print(record)
            return record
        else:
            print('No record avaiable...')

    def db_get_all_employee(self):
        connect, channel = dbExecuteForFetch(query_getAllEmp)
        record = channel.fetchall()
        if record:
            for i in record:
                print(i)
            return record
        else:
            print('No records present...')

    def db_drop_table(self):
        dbExecution(query_dropTable)
        print('Emp Table deleted successfully!')

    def empdata(self):
        eid1 = int(input('Enter employee Id: '))
        enm1 = (input('Enter employee Name: '))
        eage1 = int(input('Enter employee age: '))
        esal1 = float(input('Enter employee salary: '))
        eadd1 = (input('Enter employee address: '))
        empinfo = Emp(eid1, enm1, eage1, esal1, eadd1)
        print(empinfo)
        return empinfo

    def xl_create_table(self):
        wb, sheet = excelGetConnection()
        table = ('E_Id', 'E_Name', 'E_Age', 'E_Sal', 'E_Add')
        row, col = 1, 1
        for i in table:
            sheet.cell(row, col).value = i
            col = col + 1
        print('Table created successfully!')
        excelCloseConnection(wb)

    def xl_drop_table(self):
        wb, sheet = excelGetConnection()
        sheet.delete_rows(1, sheet.max_row+1)
        print('Employee Table deleted successfully!')
        excelCloseConnection(wb)

    def xl_get_one_employee1(self, eid):
        wb, sheet = excelGetConnection()
        row_count = sheet.max_row + 1
        for row in range(1, row_count):
            cell_val = sheet.cell(row, 1).value
            if cell_val == eid:
                return row

    def xl_add_employee(self, emp):
        wb, sheet = excelGetConnection()
        if self.xl_get_one_employee1(emp.EmpId) is None:
            cnt = sheet.max_row+1
            sheet['A'+str(cnt)] = emp.EmpId
            sheet['B'+str(cnt)] = emp.EmpName
            sheet['C'+str(cnt)] = emp.EmpAge
            sheet['D'+str(cnt)] = emp.EmpSalary
            sheet['E'+str(cnt)] = emp.EmpAddress
            print('Employee Added successfully!')
        else:
            print('Employee record already exist...')
        excelCloseConnection(wb)

    def xl_update_employee(self, emp):
        wb, sheet = excelGetConnection()
        if self.xl_get_one_employee1(emp.EmpId):
            row = self.xl_get_one_employee1(emp.EmpId)
            sheet['A'+str(row)] = emp.EmpId
            sheet['B'+str(row)] = emp.EmpName
            sheet['C'+str(row)] = emp.EmpAge
            sheet['D'+str(row)] = emp.EmpSalary
            sheet['E'+str(row)] = emp.EmpAddress
            print('Employee Added successfully!')
        else:
            print('Employee record not present...')
        excelCloseConnection(wb)

    def xl_delete_employee(self, eid):
        wb, sheet = excelGetConnection()
        if self.xl_get_one_employee1(eid):
            for row in range(1, sheet.max_row+1):
                cell_val = sheet.cell(row, 1).value
                if cell_val == eid:
                    sheet.delete_rows(row, 1)
                    print('Employee record deleted successfully!')
        else:
            print('Employee record not present...')
        excelCloseConnection(wb)

    def xl_get_one_employee(self, eid):
        e1 = []
        count = 1
        wb, sheet = excelGetConnection()
        row_count = sheet.max_row + 1
        for row in range(1, row_count):
            cell_val = sheet.cell(row, 1).value
            if cell_val == eid:
                for col in range(1, sheet.max_column + 1):
                    cell_values = sheet.cell(row, col).value
                    e1.append(cell_values)
                print(e1)
                return e1
            else:
                count = count + 1
                if count == row_count:
                    print('Employee record does not exist...')

    def xl_get_all_employee(self):
        elist = []
        e = []
        wb, sheet = excelGetConnection()
        for row in range(2, sheet.max_row + 1):
            for col in range(1, sheet.max_column + 1):
                cell_values = sheet.cell(row, col).value
                e.append(cell_values)
            e2 = e.copy()
            elist.append(e2)
            e.clear()
        for i in elist:
            print(i)
        return elist

    def xl_add_update_employee(self, emp):
        wb, sheet = excelGetConnection()
        if self.xl_get_one_employee1(emp.EmpId) is None:
            cnt = sheet.max_row+1
            sheet['A'+str(cnt)] = emp.EmpId
            sheet['B'+str(cnt)] = emp.EmpName
            sheet['C'+str(cnt)] = emp.EmpAge
            sheet['D'+str(cnt)] = emp.EmpSalary
            sheet['E'+str(cnt)] = emp.EmpAddress
        else:
            row = self.xl_get_one_employee1(emp.EmpId)
            sheet['A' + str(row)] = emp.EmpId
            sheet['B' + str(row)] = emp.EmpName
            sheet['C' + str(row)] = emp.EmpAge
            sheet['D' + str(row)] = emp.EmpSalary
            sheet['E' + str(row)] = emp.EmpAddress
        excelCloseConnection(wb)

    def db_add_update_employee(self, emp):
        if self.db_get_one_employee1(emp.EmpId) is None:
            dbExecution(query_addEmp.format(emp.EmpId, emp.EmpName, emp.EmpAge, emp.EmpSalary, emp.EmpAddress))
        else:
            dbExecution(query_updateEmp.format(emp.EmpName, emp.EmpAge, emp.EmpSalary, emp.EmpAddress, emp.EmpId))
