# all sql queries

query_createTable = '''create table emp(E_Id int, E_Name varchar(100), E_Age int, E_Sal float, E_Add varchar(200), PRIMARY KEY (E_Id))
'''

query_addEmp = '''insert into emp values({}, '{}', {}, {}, '{}') 
'''

query_updateEmp = '''update emp set E_Name='{}', E_Age={}, E_Sal={}, E_Add='{}' where E_Id={}
'''

query_deleteEmp = '''delete from emp where E_Id={}
'''

query_getOneEmp = '''select * from emp where E_Id={}
'''

query_getAllEmp = '''select * from emp
'''

query_dropTable = '''drop table emp
'''

query_desc = '''desc emp
'''