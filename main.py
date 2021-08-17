

import gspread          #lets you access Google Sheets API v4
from oauth2client.service_account import ServiceAccountCredentials  # oauth2client is a client library for accessing
                                                                    # resources protected by OAuth 2.0
                                                                    #this allows you to access the APIs with the key
import pandas as pd
import matplotlib
import matplotlib.pyplot as plt     #allows us to create charts
import scipy

#Used to gain access to the APIS Used in this project
scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.file',
         'https://www.googleapis.com/auth/drive']

#Use the key to use the APIs
credentials = ServiceAccountCredentials.from_json_keyfile_name('googlesheetskey.json', scope)
client = gspread.authorize(credentials)
sheet = client.open("GoogleFinanceScrape").sheet1   #Uses the google sheets that was shared
data = sheet.get_all_records()  #access all data in sheets

#Get input for start date then update the cell
#start_date = input("Enter the start date: ")
start_date= '1/1/2020'
sheet.update_cell(3,3 , value=start_date)

#Get input for end date then update the cell
#end_date = input("Enter the end date: ")
end_date = '3/1/2020'
sheet.update_cell(4,3 , value=end_date)

#Get the ticker symbol and update the cell
#ticker = input("Enter the ticker symbol: ")
ticker = 'NVDA'
sheet.update_cell(5,2 , value=ticker)

#in the next variable we need to enter the google finance formula as a string since it is a built in function that
#will execute and display the companys data in the range, with the attributes, we assigned previously in the code

#Get Google Finance formula -     ticker  attributes  start cell    end cell  interval
goofin_formula = '=GOOGLEFINANCE( $B$5,     "All",    $C$3,          $C$4,    "Daily")'

#update cell value with google finance formula
sheet.update_cell(6,2, value=goofin_formula)


#######Section for reading and graphing the data################

#Access all values in the google sheets
sheet_id = '1x6a3MZu8LFwvPVjQ300RfRZrvCknUEwn79VoEH6TcAw'   #The ID of the sheet found in the Google Sheets URL
open_sheetfile = client.open_by_key(sheet_id)               #Authorizes the program to open the file with the given credentials to access it
sheet = open_sheetfile.get_worksheet(0)                     #Open file and access the first worksheet
values = sheet.get_all_values()                             #Get all values in cells

#puts all values received previously, in a pandas dataframe
data = pd.DataFrame(values)

#Create a for loop to iterate over sheet values

#testy = close prices on chart
testy = []
#testx = date on chart
testx = []
#offset is used to get the for loop to see the column I want it to start at as the starting point for looping
#I need it to start at column 6 so 6-6 will make it the starting point
offset = 6

#My next plan of action is to figure out how to adjust the range and change interval of dates on the graph according to the
#users range of information

#in the position of 40 I would have to figure out how get difference between the start date and end date given by user
for n in range(offset, 40):
    testy += [data.iloc[n,5]]
    testx += [data.iloc[n,1]]
    print(testx[n-offset])
    print(testy[n-offset])

#plots the chart with the axis. I will add labels and a map key later on
plt.plot(testx, testy)
plt.show()

