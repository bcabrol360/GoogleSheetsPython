import gspread                                                      #lets you access Google Sheets API v4

from oauth2client.service_account import ServiceAccountCredentials  # oauth2client is a client library for accessing
                                                                    # resources protected by OAuth 2.0
                                                                    #this allows you to access the APIs with the key

import pandas as pd                                                 #Data analysis library that provides a dataframe, which
                                                                    #is a 2-dimensional data structure containing rows and columns
                                                                    #similar to an excel sheet, that allows us to store and access data

import matplotlib.pyplot as plt                                     #matplotlib is a 2-Dimensional plotting library
                                                                    #that allows us to create plots and graphs

import numpy as np                                                  #provides us with a multidimensional array

#Scope is included so I am able to gain access to the APIs I used in this project
#Description on creating the Google Project on the Google Cloud Platform is included in the GoogleAPISetupDescription.txt file
scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.file',
         'https://www.googleapis.com/auth/drive']

#Uses the key to access and use the APIs iin the Google Cloud Platform
credentials = ServiceAccountCredentials.from_json_keyfile_name('googlesheetskey.json', scope)
client = gspread.authorize(credentials)
sheet = client.open("GoogleFinanceScrape").sheet1   #Uses the google sheets that was shared
data = sheet.get_all_records()  #access all data in sheets


#Get input for start date
start_date = input("Enter the start date in MM/DD/YYYY format: ")

#Update google sheets cell with start_date input
sheet.update_cell(3,3 , value=start_date)

#Get input for end date
end_date = input("Enter the end date in MM/DD/YYYY format: ")

#Update google sheets cell with end_date input
sheet.update_cell(4,3 , value=end_date)


#Get the ticker symbol
ticker = input("Enter the ticker symbol: ")

#Update google sheets cell with ticker input
sheet.update_cell(5,2 , value=ticker)

#in the goofin_formula variable we need to enter the google finance formula as a string since it is a built in formula in google sheets that
#will execute and display the companys data in the range, with the attributes, we acquired previously in the code

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


data = pd.DataFrame(values)                                 #puts all values received previously, in a pandas dataframe


date_difference = '=NETWORKDAYS(C3,C4)'                     # =NETWORKDAYS() is a built in formula in google sheets to get
                                                            # the number of business days between 2 dates(exclude the weekend and holidays)

sheet.update_cell(3,4, value=date_difference)               #updates cell with value
data_range = data.values[2,3]                               #Access value in specific data frame at
day_range = int(data_range) + 1                             #convert the days into an int. Im adding 1 so it includes the last day in the range as well


#######Create a for loop to iterate over sheet values#######

close_price = []                                            #Variable initialized to store all close prices of each day in a list
stock_date = []                                             #Variable initialized to store all dates for stock in a list
high_price = []                                             #Variable initialized to store all high prices for each day in a list
low_price = []                                              #Variable initialized to store all low prices for each day in a list


offset = 6                                                  #offset is used to get the for loop to see the row I want it to start at as the starting point for looping
                                                            #I need it to start at row 6 so 6-6 will make it the starting point

#for loop that loops through each column in data frame.
for n in range(offset,day_range):
    close_price += [data.iloc[n, 5]]                        #The loop will iterate over the 6th column in the dataframe and store it in close_price  .iloc [] is used for indexing and slicing in the dataframe
    stock_date += [data.iloc[n, 1]]                         #The loop will iterate over the 2nd column in the dataframe and store it in stock_date
    high_price += [data.iloc[n,3]]                          #The loop will iterate over the 4th column in the dataframe and store it in high_price
    low_price += [data.iloc[n,4]]                           #The loop will iterate over the 5th column in the dataframe and store it in low_price
    y = close_price[n - offset]                             #Stores close_price values
    x = stock_date[n - offset]                              #Stores stock_date values
    hp = high_price[n-offset]                               #Stores high_price values
    lp = low_price[n-offset]                                #Stores low_price values




#convert close_price values into an array with float values
n_arr = np.array(close_price, dtype=float)

#Converts high price from a list to an array with float values
highprice_arr = np.array(high_price, dtype=float)
highprice_plot = highprice_arr

#converts low price from a list to an array with float values
lowprice_arr = np.array(low_price, dtype=float)
lowprice_plot = lowprice_arr

#####################plotting the data section##############


plt.title('Historical Stock Data')                          #Gives plot a title
plt.grid(True)                                              #Adds grid to plot

#Plots the values read from above
#The first spot after the ( is always the X values. After that is the Y-values  which is then followed by the attributes you want to add or change.
#There are countless ways to change the charts appearance that can be found on the matplotlib website
plt.plot(stock_date, n_arr, color='red', label ="Close Price")
plt.plot(stock_date, highprice_plot, color ='blue', label ="High Price")
plt.plot(stock_date, lowprice_plot, color ='green', label ="Low Price")


plt.legend()                                                 #plots legend on chart with line labels
plt.xticks(rotation = 90)                                    #formatting that rotates the x-axis tick mark labels 90 degrees
plt.show()                                                   #plots and displays graph

#My code is not displaying some dates that I ask of it at the moment but I will be working to fix that.
#I narrowed the issue down to google sheets formula so I will try to fix this issue.

