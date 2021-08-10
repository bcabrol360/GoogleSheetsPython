

import gspread          #lets you access Google Sheets API v4
from oauth2client.service_account import ServiceAccountCredentials  # oauth2client is a client library for accessing
                                                                    # resources protected by OAuth 2.0
                                                                    #this allows you to access the APIs with the key
import pandas as pd     #used for data manipulation

#Used to gain access to the APIS Used in this project
scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.file',
         'https://www.googleapis.com/auth/drive']

#Use the key to use the APIs
credentials = ServiceAccountCredentials.from_json_keyfile_name('googlesheetskey.json', scope)
client = gspread.authorize(credentials)
sheet = client.open("GoogleFinanceScrape").sheet1   #Uses the google sheets that was shared
data = sheet.get_all_records()  #access all data in sheets

#Get input for start date then update the cell
start_date = input("Enter the start date: ")
sheet.update_cell(3,3 , value=start_date)

#Get input for end date then update the cell
end_date = input("Enter the end date: ")
sheet.update_cell(4,3 , value=end_date)

#Get the ticker symbol and update the cell
ticker = input("Enter the ticker symbol: ")
sheet.update_cell(5,2 , value=ticker)

#in the next variable we need to enter the google finance formula as a string since it is a built in function that
#will execute and display the companys data in the range, with the attributes, we assigned previously in the code

#Get Google Finance formula -     ticker  attributes  start cell    end cell  interval
goofin_formula = '=GOOGLEFINANCE( $B$5,     "All",    $C$3,          $C$4,    "Daily")'

#update cell value with google finance formula
sheet.update_cell(6,2, value=goofin_formula)

#I will be updating this project with more code and comments as I add more features =)




