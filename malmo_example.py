# 0 - packages, initial settings, and user defined functions

#import packages for database connection
import mysql.connector
from mysql.connector import Error
#import selenium for browser automation
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
#import keyring for password management
import keyring
#import packages for xlsx management
import xlsxwriter
import openpyxl
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
#import plotly for data visualization
import plotly.graph_objects as go
#define today's date as a string 
from datetime import datetime
now = datetime.now() # current date and time
day = now.strftime("%d")
month = now.strftime("%m")
year = now.strftime("%y")
full_date = day + "-" + month + "-" + year
#other packages
import os
import time
import glob
#get passwords from windows credential manager
db_password = keyring.get_password("test_database", "sql7308056")
web_password = keyring.get_password("test_web", "test_web")
#user defined functions
def rename_last_file(path, filename):
    list_of_files = glob.glob(path + "/*")
    last_file = max(list_of_files, key=os.path.getctime)
    print(last_file)
    os.rename(last_file, path + "/" + filename)

# 1 - get data from the online database
try:
    #connection and fetching
    print("1- Connecting to database...")
    connection = mysql.connector.connect(host='sql7.freesqldatabase.com',
                                         database='sql7308056',
                                         user='sql7308056',
                                         password=db_password)
    
    sql_select_Query = "select * from ABBA2"
    cursor = connection.cursor()
    cursor.execute(sql_select_Query)
    records = cursor.fetchall()
    print("2- Connection successful. Saving data to file...")
    #get columns property
    col_number = len(records[0]) #column number
    col_names = cursor.column_names #column headers

    #create xlsx file to save data
    path = os.path.dirname(os.path.abspath(__file__))
    abba2_filename = "/ABBA2_" + full_date + ".xlsx"
    db_path = path + abba2_filename
    workbook = xlsxwriter.Workbook(db_path)
    worksheet = workbook.add_worksheet()
    #writing on xlsx file
    row_number = 0
    for rows in records:
        if row_number == 0: #first row, write headers
            for i in range(col_number-1):
                worksheet.write(row_number, i, col_names[i])
        else: #other rows, data
            for i in range(col_number-1):
                worksheet.write(row_number, i, rows[i])
        row_number += 1
    workbook.close() #close
    print("3- File created. Downloading data from web page...")
except Error as e:
    print("Error while connecting to MySQL ", e)
    exit()

# 2 - download file from internet
try:
    #set chromedriver options
    chromeOptions = webdriver.ChromeOptions()
    prefs = {"download.default_directory" : path} #default download path
    chromeOptions.add_experimental_option("prefs",prefs)
    driver = webdriver.Chrome(chrome_options=chromeOptions)
    #go to webpage
    driver.get("https://jbarbati.wordpress.com/2019/10/10/abba1/") #start browser
    driver.find_element_by_xpath('//*[@id="pwbox-369"]').send_keys(web_password) #input password
    driver.find_element_by_xpath('//*[@id="post-369"]/div[2]/div/form/p[2]/input').click() #submit
    time.sleep(5)
    driver.find_element_by_xpath('//*[@id="post-369"]/div[2]/div/p/a').click() #start download
    time.sleep(10)
    driver.quit() #close browser
    abba1_filename = "ABBA1_" + full_date + ".xlsx"
    rename_last_file(path, abba1_filename) #rename the file just downloaded
    print("4- File downloaded. Creating analysis file...")
except Exception as e:
    print("Error while downloading file ", e)
    exit()

# 3 - compare files
#read files with pandas
try:
    abba1 = pd.read_excel(path + "/" + abba1_filename)
    abba2 = pd.read_excel(path + "/" + abba2_filename)
    #create dictionaries for time and pm1 - ABBA1
    abba1_time = abba1['time'].tolist()
    abba1_pm1 = abba1['PM_1(ug/m3)'].tolist()
    abba1_dict = dict(zip(abba1_time, abba1_pm1))
    #create dictionaries for time and pm1 - ABBA2
    abba2_time = abba2['time'].tolist()
    abba2_pm1 = abba2['PM_1(ug/m3)'].tolist()
    abba2_dict = dict(zip(abba2_time, abba2_pm1))
    #set analysis workbook
    tot_path = path + "/analysis_" + full_date + ".xlsx"
    workbook_tot = xlsxwriter.Workbook(tot_path)
    worksheet_tot = workbook_tot.add_worksheet()
    #write analysis workbook
    #headers
    worksheet_tot.write(0, 0, "TIME")
    worksheet_tot.write(0, 1, "ABBA1_PM1")
    worksheet_tot.write(0, 2, "ABBA2_PM1")
    #data
    row = 1
    for time1, pm1 in abba1_dict.items():
        for time2, pm2 in abba2_dict.items():
            time1 = str(time1)
            time1 = time1[:5]
            time2 = str(time2)
            time2 = time2[:5]
            if time1 == time2: #only if time matches to the minute
                worksheet_tot.write(row, 0, time1)
                worksheet_tot.write(row, 1, pm1)
                worksheet_tot.write(row, 2, pm2)
                row += 1
    workbook_tot.close() #close
    print("5- Analysis file created. Visualizing data...")
except Exception as e:
    print("Error while creating analysis file ", e)
    exit()

# 4 - visualize data
try:
    df = pd.read_excel(tot_path)
    # Create figure
    fig = go.Figure()
    #add datasets
    fig.add_trace(go.Scatter(x=list(df.TIME), y=list(df.ABBA1_PM1), name = "ABBA1"))
    fig.add_trace(go.Scatter(x=list(df.TIME), y=list(df.ABBA2_PM1), name = "ABBA2"))
    # Set title and slider
    fig.update_layout(title_text='PM1 measures, ' + full_date, xaxis_rangeslider_visible=True)
    #show chart
    fig.show()
except Exception as e:
    print("Error while creating chart ", e)
    exit()