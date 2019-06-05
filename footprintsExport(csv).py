#!/usr/bin/python
#C:\Users\has\AppData\Local\Programs\Python\Python37-32\Scripts
#CREATED BY
#SANG SOO HA
#On March 22nd 2019
import pyodbc
import os, time
import xlwt, csv
import datetime


#Instantiates a workbook and sheet to write the qurey results to.
workbook = xlwt.Workbook(r'H:\DB.xlsx')
sheet = workbook.add_sheet('Contact')

#Database Connection
conn = pyodbc.connect(driver='SQL Server', server='SQLGROUP1', database='fpscdb001',
                      user='vgn_readonly', password='Vgn_readonly1234!')

#It will let you execute all the queries you need 
cursor = conn.cursor()


status_1 = 'Active'

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
               'Legal Services','Real Estate','Office of the Chief Infomratin Officer']

def dept(department, status_1):
    
    argu = (department, status_1)
  
    cursor.execute("""
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
                   Department
                   like ?
                   AND
                   Status like ?
                   """, argu)

    #Admin
    #if input('Enter your id.') != 's':
    #location = input('Please enter the location of \
    #the file including its name + format.')

    
target = r'O:\OCIO\ESITAC\ASSET COORDINATOR & CONTRACTS FOLDER\Sang Soo Ha\Listings of Departments'
    
today = target + os.sep + time.strftime('%Y%m%d')

if not os.path.exists(today):
    os.mkdir(today)

for i in departments:
    dept(i, status_1)
    with open(today + os.sep + \
                i + ".csv","w", newline='') as csvFile:
        writer = csv.writer(csvFile, lineterminator='\n\n')
        data = cursor.fetchall()
        writer.writerow(["Device ID","Type","Serial Number",\
                         "Model","First Name","Last Name","Title",\
                         "Department","Location","Invoice Date",\
                         "Status","Sub Status","Comments"])
        
        for item in data:
                writer.writerow(item)

    print('\nWork In Progress...')
    print(i, 'has been exported successfully.\n')


def sharepoint():
    from sharepoint import SharePointSite, basic_auth_opener

    server_url = 'https://vol.vgn.cty/'
    site_url = + server_url + \
    'Sites/PMO/DesktopReplacementProgram/Desktop_Replacement/Shared%20Documents/Forms/AllItems.aspx'

    opener = basic_auth_opener(server_url, "has","Gkdlgkdl123!")

    site = SharePointSite(site_url, opener)

os.startfile(today)

