""" Sales Performance dashboard
    Python automatisation

    Steps:
    1_ Get data from kaggle CSV file 
    2_ Transform data
    3_ Create dashboard with data and Excel template
    4_ Share the report by outlook mail

"""

"""  SOURCE KAGGLE
    >> Supermarket sales

    https://www.kaggle.com/datasets/aungpyaeap/supermarket-sales

"""

""" Github repo
sales_performance_dashboard


"""

import pandas as pd
import datetime
import os
import openpyxl
import xlwings as xw

print ("'\nSTART - " + os.path.basename(__file__))
print( "Date : " + datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

#  _________________________________________________
#  VARIABLES
file_name = "supermarket_sales - Sheet1.csv"
today = datetime.date.today()
# fileExtension_name = today.strftime('%Y%m%d')
fileExtension_name = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
template_filename = "sales_performance"
data_sheet_name = "data"

#  def data_processing():

# _________________________________________________
# __________________________________________________
# DATA TRANSFORMATION

# retrieve data from csv into df
df_data_raw = pd.read_csv(file_name)
df_data_raw

#  Exploratory
df_data_raw.head()
df_data_raw.describe( include = "all")

# df_data_raw.columns
# Index(['Invoice ID', 'Branch', 'City', 'Customer type', 'Gender',
#        'Product line', 'Unit price', 'Quantity', 'Tax 5%', 'Total', 'Date',
#        'Time', 'Payment', 'cogs', 'gross margin percentage', 'gross income',
#        'Rating'],
#       dtype='object')

# Group by
df_data = df_data_raw.groupby(['Branch','City', 'Customer type', 'Gender', 'Product line', 'Payment']).agg({'Invoice ID': 'count', 'Quantity': 'sum', "Total": 'sum', "gross income": 'sum', "Rating": 'sum'}).reset_index()

#  rename columns
df_data = df_data.rename(columns={'Invoice ID' : 'Invoices'})
df_data = df_data.rename(columns={'Total' : 'Prices'})
df_data = df_data.rename(columns={'Rating' : 'Ratings score'})

# create new column

#  display distinct values of "Product line"
df_data["Product line"].drop_duplicates()

#  function to be used in order to create new colum with the category of Product 
def get_product_category(product_line):

    if product_line == "Food and beverages":
        return "consumables"
    
    elif product_line == "Health and beauty" or  product_line == "Electronic accessories":
        return "soft goods"
    
    elif product_line == "Fashion accessories" or  product_line == "Electronic accessories"  or  product_line == "Sports and travel":
        return "superior goods"
    
    else:
        return "others"

#  creation of new column depending of "Product line" columns value
df_data["Product category"] = df_data["Product line"].apply(get_product_category)


df_data
df_data.describe(include="all")

# _________________________________________________
# _________________________________________________
# ARCHIVING

# # Save old report in Archives folder 
# by moving all xlsx files
os.system('move /Y "'+ '*.xls*" "' + 'Archives"' )

# _________________________________________________
# _________________________________________________
# EXCEL DASHBOARD CREATION 

# with excel template
#  from def export_dfs_to_excel

# # just export data for the first time into excel sheet
# df_data.to_excel("output.xlsx")

#  export df to excel template with xlwings module 

#  open existing excel template
wb = xw.Book("Template\\" + template_filename + ".xlsx")

# app = xw.apps.active  
wb.sheets[data_sheet_name].activate()
ws = wb.sheets[data_sheet_name]
# paste the values of dataframe from B2 cell
ws.range("B2").value = df_data.values

report_file_name = template_filename +"_" +  fileExtension_name + ".xlsx"

# put the first sheet active before closing the report
wb.sheets[0].activate()

# save as the excel worbook: a nex file is created
wb.save(report_file_name)

# update pivot table -> "Refresh All" button
# Refresh all data connections.
wb.RefreshAll()

# close the excel workbook
wb.close()

# _________________________________________________
# _________________________________________________
# SHARING the file by email ( with outlook template )

# function to generate email with .oft template 
# ( for outlook user in windows OS)

import win32com.client
def displayEmail (email_template_path, attachmentPaths):

    print ( "\n>>> START- displayEmail ")

    # Create an email message from .oft template
    obj = win32com.client.Dispatch("Outlook.Application")
    report_mail = obj.CreateItemFromTemplate (email_template_path)
    report_mail.display()

    # Add attachments
    for attachment in attachmentPaths :
        print ("attachment = "+ attachment)
        report_mail.Attachments.Add(Source=attachment)
    
    # open the mail
    report_mail.display()

    # save it in outlook draft folder
    report_mail.save()

    print ( "\n>>> END - displayEmail ")

# the list of files to attached into email

attachment_files = list() 
attachment_files.append( project_folder + report_file_name ) 

# an oft template should be created before and save in "template" folder
emailTemplateName = "sales_performance.oft"
displayEmail (project_folder + "template\\" + emailTemplateName , attachmentPaths = attachment_files)   

##################################################################################
print ("'\nEND  - " + os.path.basename(__file__))
print( "Date : " + datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
