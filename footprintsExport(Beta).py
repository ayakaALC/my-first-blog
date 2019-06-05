#!/usr/bin/python
#C:\Users\has\AppData\Local\Programs\Python\Python37-32\Scripts
#CREATED BY
#SANG SOO HA
#On March 22nd 2019
import pyodbc
import os, time
import xlwt, csv, xlsxwriter
import datetime
import pandas as pd
import pandas.io.sql as psql
from pandas import ExcelWriter
from pandas import ExcelFile
from xlsxwriter.workbook import Workbook



#Instantiates a workbook and sheet to write the query results to.
#workbook = xlwt.Workbook(r'H:\DB.xlsx')
#sheet = workbook.add_sheet('Contact')

#Database Connection
conn = pyodbc.connect(driver='SQL Server', server='SQLGROUP1', database='fpscdb001',
                      user='vgn_readonly', password='Vgn_readonly1234!')

#It will let you execute all the queries you need 
cursor = conn.cursor()


#status_1 = 'Active'

departments = ['Transportation Services, Parks and Forestry Operations',
               'Strategic Planning', 'Legal Services',
               'Recreation Services','Financial Services',
               'Financial Planning and Development Finance',
               'Procurement Services','Corporate and Strategic Communications Department',
               'Economic and Cultural Development','Access Vaughan',
               'By-Law and Compliance, Licensing and permit Services',
               'Facility Services','Fire and Rescue Service',
               'Building Standards','Development Engineering',
               'Development Planning','Parks Development',
               'Policy Planning and Environmental Sustainability',
               'Vaughan Metropolitan Cetre Program',
               'Environmental Services','Fleet management Services',
               'Infrastructure Planning & Corporate Asset Management',
               'Infrastructure Delivery','Office of the Chief Human Resources Officer',
               'Legal Services','Real Estate','Office of the Chief Information Officer']

for i in departments:
    print(i)


query = """
                           SELECT *
                           FROM
                           (SELECT
                           device_idd AS "Device ID",
                           type_1 AS "Type",
                           serial_number AS "Serial Number",
                           model as "Model",
                           first_name AS "First Name",
                           last_name AS "Last Name",
                           title_1 AS "Title",
                           department AS "Department",
                           location_1 AS "Location",
                           invoice_date_1 AS "Invoice Date",
                           status_1 AS "Status",
                           sub_status AS "Sub Status",
                           comments AS "Comments"
                           FROM
                           [fpscdb001].[fpscdb001_cmdb_001].[notebooks]

                           UNION ALL
                           SELECT
                           device_idd as "Device ID",
                           type_1 as "Type",
                           serial_number_1 as "Serial Number",
                           model as "Model",
                           first_name_1 AS "First Name",
                           last_name_1 as "Last Name",
                           title as "Title",
                           department_1 as "Department",
                           location_2 as "Location",
                           invoice_date as "Invoice Date",
                           status_1 as "Status",
                           sub_status as "Sub Status",
                           comments as "Comments"
                           FROM
                           [fpscdb001].[fpscdb001_cmdb_001].[desktop]
                           
                           UNION ALL
                           SELECT
                           device_idd as "Device ID",
                           type_1 as "Type",
                           serial_number as "Serial Number",
                           model as "Model",
                           first_name as "First Name",
                           last_name as "Last Name",
                           title_1 as "Title",
                           department as "Department",
                           location_1 as "Location",
                           invoice_date as "Invoice Date",
                           status_1 as "Status",
                           sub_status as "Sub Status",
                           comments as "Comments"
                           FROM
                           [fpscdb001].[fpscdb001_cmdb_001].[tablets]
                           ) AS t 
                           WHERE
                           Department like ?
                           AND
                           Status like 'Active'
                           """
    

    #Admin
    #if input('Enter your id.') != 's':
    #location = input('Please enter the location of \
    #the file including its name + format.')

target = r'O:\OCIO\ESITAC\ASSET COORDINATOR & CONTRACTS FOLDER\Listings of Departments'

if not os.path.exists(target):
    os.mkdir(target)
    
today = target + os.sep + time.strftime('%Y%m%d') + '_XLSX'

if not os.path.exists(today):
    os.mkdir(today)


#for i in departments:
#    dept(i)
#    writer = pd.ExcelWriter(today + os.sep + i + ".xlsx")
#    dept.to_excel(writer, sheet_name = i)
#    writer.save
    


for i in departments:
    #workbook = Workbook(today + os.sep + i + 'xlsx')
    #worksheet = workbook.add_worksheet()
    #data = cursor.fetchall()
    a = today + os.sep + i + '.xlsx'
    writer = pd.ExcelWriter(a, engine='xlsxwriter')
    P_data = pd.read_sql_query(query, conn, params=(i,))
    P_data.to_excel(today + os.sep + i + '.xlsx')
    
    #workbook = xlsxwriter.Workbook(today + os.sep + i + '.xlsx')
    #worksheet = writer.sheets['Sheet1']
    #chart = workbook.add_chart({'type': 'column'})
    #chart.add_series({'values': '=Sheet1$B$2:$B$8'})
    #worksheet.insert_chart('D2', chart)
    #worksheet.add_table('B1:N99')

    #worksheet = writer.sheets[a]
    

    
    #file = today + os.sep + i + ".xlsx"
    #data = cursor.fetchall()
    rowHeaders = ["Device ID","Type","Serial Number",\
                     "Model","First Name","Last Name","Title",\
                     "Department","Location","Invoice Date",\
                     "Status","Sub Status","Comments"]

    #rowValues = [i]
    #for row in rowHeaders:
    #    worksheet.write_row(row, bold)
    #writer = pd.ExcelWriter(pd_data)    
    #pd_data.to_excel(writer, sheet_name = 'Sheet1')
    #writer.save()
    print(('Running...'))
    #time.sleep(0.8)
    #print('Done!')
    #time.sleep(0.2)
    print('Work In Progress...\n')
    print(i, 'has been exported successfully.')


    
    
os.startfile(today)

