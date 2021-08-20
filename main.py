

import gspread          #lets you access Google Sheets API v4
from oauth2client.service_account import ServiceAccountCredentials  # oauth2client is a client library for accessing
                                                                    # resources protected by OAuth 2.0
                                                                    #this allows you to access the APIs with the key
import pandas as pd
import matplotlib
import matplotlib.pyplot as plt     #allows us to create charts
import matplotlib.dates as mdates
import scipy
import datetime as dt
import numpy as np
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
start_date= '6/19/2021'
sheet.update_cell(3,3 , value=start_date)

#Get input for end date then update the cell
#end_date = input("Enter the end date: ")
end_date = '8/1/2021'
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


#Calculate difference between 2 dates
date_difference = '=NETWORKDAYS(C3,C4)'                     # =NETWORKDAYS() is a built in formula in google sheets to get
                                                            # the # of business days between 2 dates(exclude the weekend and holidays)
sheet.update_cell(3,4, value=date_difference)               #updates cell with value
data_range = data.values[2,3]
day_range = int(data_range)                                 #convert the days into an int
print(data_range)


#######Create a for loop to iterate over sheet values#######

testy = []                      #testy = close prices on chart
testx = []                      #testx = date on chart
offset = 6                      #offset is used to get the for loop to see the column I want it to start at as the starting point for looping
                                #I need it to start at column 6 so 6-6 will make it the starting point


for n in range(offset, day_range):
    testy += [data.iloc[n,5]]
    testx += [data.iloc[n,1]]
    y = testy[n-offset]
    x = testx[n-offset]
    print(x)
#convert testy values into an array with float values and store in variable testy1
n_arr = np.array(testy, dtype=float)
testy1 = n_arr
#print(testy1)

#plotting the data section

plt.title('Historical Stock Data')
#adding grid to plot
plt.grid(True)
#plots the values readd above
plt.plot(testx, n_arr)
#rotates the ticker symbols 90 degrees - formatting
plt.xticks(rotation = 90)
#plot and show graph
plt.show()

#My code is not displaying some dates that I ask of it at the moment but it is a minor fix. More additions to this code will come soon

