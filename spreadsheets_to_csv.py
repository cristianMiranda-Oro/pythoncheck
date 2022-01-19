# -*- coding: utf-8 -*-
"""
Created on Wed Jan 19 15:36:29 2022

@author: CRISTIANMIRANDA
"""

"""
Inmputs:
    credentials: file .json that contains service_account
    sheetName: document sheet name
    numberSheet: number of the page
    routeOut: Route alocate file csv if not then it will be saved in the default path
Output:
    return: route where the CSV file was saved
"""
import os
import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime


def spreadsheetsToCSV(credentials = None , sheetName = None, 
                        numberSheet = None, routeOut = ""):
    
    outRoute = ""
    try:
        #**********************************************************************
        #************************input validation******************************
        #**********************************************************************
        txt = ""
        if credentials == None and type(credentials) != str:
            txt = txt + "Argument not found or invalid type"
        
        if sheetName == None and type(sheetName) != str:
            txt = txt + "Invalid second argument or type or"
            
        if numberSheet == None and type(numberSheet) != int:
            txt = txt + "Invalid third argument or type"
            
        if txt != "":
            raise NameError(txt)
        #**********************************************************************
        #**********************************************************************
        #**********************************************************************
        
        if(routeOut != ""):
            os.makedirs(routeOut, exist_ok=True)#create the path where the output file is saved
        
        scope = ["https://spreadsheets.google.com/feeds",
                 'https://www.googleapis.com/auth/spreadsheets',
                 "https://www.googleapis.com/auth/drive.file",
                 "https://www.googleapis.com/auth/drive"]
        
        # add credentials to the account
        creds = ServiceAccountCredentials.from_json_keyfile_name(credentials, scope)
    
        # authorize the clientsheet 
        client = gspread.authorize(creds)
        
        # get the instance of the Spreadsheet
        sheet = client.open(sheetName)
        
        # get the first sheet of the Spreadsheet
        sheet_instance = sheet.get_worksheet(numberSheet)
    
        # get all the records of the data
        records_data = sheet_instance.get_all_records()
        
        #**********************************************************************
        #*********************Convert and save file****************************
        #**********************************************************************
        
        # convert the json to dataframe
        records_df = pd.DataFrame.from_dict(records_data)
        
        # view the top records
        records_df.head()
        
        #file's name is datetime
        nameFile = datetime.today().strftime('%A_%B_%d_%Y_%H_%M_%S')
        
        #Final Route
        outRoute = routeOut+'/'+nameFile+".csv"
        
        records_df.to_csv(outRoute, index=False) 
        
        #**********************************************************************
        #**********************************************************************
        #**********************************************************************
        
        
    except NameError as exc:
        outRoute = "Error"
        raise RuntimeError('failed to convert sheet') from exc
        
        
    return outRoute


#Test1
outRoute = spreadsheetsToCSV("Keys.json", "MarkeBot", 1, "output")

