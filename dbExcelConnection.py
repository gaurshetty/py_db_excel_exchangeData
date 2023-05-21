import os
import pymysql
import openpyxl
os.chdir('D:\\mysqldb')


def excelGetConnection():
    wb = openpyxl.load_workbook('D:\\mysqldb\\empTable.xlsx')
    sheet = wb.active
    return wb, sheet


def excelCloseConnection(wb):
    wb.save('D:\\mysqldb\\empTable.xlsx')


def dbGetConnection():
    connect = pymysql.connect('localhost', 'root', 'root', 'pymydb')
    if connect:
        return connect


def dbExecution(query):
    connect = dbGetConnection()
    channel = connect.cursor()
    channel.execute(query)
    connect.commit()
    channel.close()
    connect.close()


def dbExecuteForFetch(query):
    connect = dbGetConnection()
    channel = connect.cursor()
    channel.execute(query)
    return connect, channel
