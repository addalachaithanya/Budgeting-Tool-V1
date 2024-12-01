import openpyxl
import os
import glob
import re
from collections import defaultdict
from datetime import datetime




# credit/debit statements directory
statement_reports_store = "C:\\Users\\saddala\\AppData\\Local\\Programs\\Python\\Python310\\Budget_App\\Statements"

# Global Dictionary to map all the transactions
Global_Transaction_Data = {}

'''
User_Name:
        [
            Bank : <Boolean>,
            Type : [Credit, Debit],
            Name :
            Category:
                    [
                        Category_type : Dining,
                        Amount : $xx,
                        Transaction_Date:
                        Merchant_Name : 
                        
                        
                    

                    ],
            
            Bank : <Boolean>,
            Type : [Credit, Debit],
            Name : 
            Category:
                    [
                        Category_type : Dining,
                        Amount : $xx,
                        Transaction_Date:
                        Merchant_Name : 
                        
                    ]
        ]
'''
Original_Data_Store = {}



Month_Id = {
    1 : 'January',
    2 : 'February',
    3 : 'March',
    4 : 'April',
    5 : 'May',
    6 : 'June',
    7 : 'July',
    8 : 'August',
    9 : 'September',
    10 : 'October',
    11 : 'November',
    12 : 'December',
}


# loop thru statements to load each statement and parse the trans
def exhibit_data_from_statements(statements):
    for i, xlsx_file in enumerate(statements):
        workbook = openpyxl.load_workbook(xlsx_file)
        sheet = workbook.active
        print(f"File: {xlsx_file}")
        trans_sum = 0.0
        dining_sum = 0.0
        month,day,year = 0,0,0
        monthly_charges = {}
        Transaction_Data = {}
        print(f" ------------------------------------ Start Statement {i} -----------------------------------")
        for row in sheet.iter_rows(values_only=True):
            
            if row is None:
                continue
            Transaction_Date, Posted_Date, Card_No, Description, Category, Debit, Credit, User_Id, Credit_Company = row
            Debit_Account = "Deb"
            Credit_Account = "Cred"
            major_year = 2024
            

            if isinstance(Transaction_Date,datetime):
                month = Transaction_Date.month
                day = Transaction_Date.day
                year = Transaction_Date.year
            
            month = int(month)
            year = int(year)
            
            if User_Id is None:
                continue

            if User_Id not in Global_Transaction_Data:
                Global_Transaction_Data[User_Id] = {}

            if Credit_Company not in Global_Transaction_Data[User_Id]:
                Global_Transaction_Data[User_Id][Credit_Company] = {'Dining': 0.0, 'Merchandise': 0.0, 'Miscellaneous': 0.0, 'Debit': 0.0, 'Credit': 0.0}

            if Category == 'Dining' and row[5] is not None:
                Global_Transaction_Data[User_Id][Credit_Company]['Dining'] += float(row[5])
                dining_sum += float(row[5])

            if Category == 'Merchandise' and row[5] is not None:
                Global_Transaction_Data[User_Id][Credit_Company]['Merchandise'] += float(row[5])
                
            if re.match(r'^(?!(Dining|Merchandise)$)', Category, re.IGNORECASE) and row[5] is not None:
                Global_Transaction_Data[User_Id][Credit_Company]['Miscellaneous'] += float(row[5])
                
            
            
            if row[5] is not None:
                Global_Transaction_Data[User_Id][Credit_Company]['Debit'] += float(row[5])
                if Month_Id[month] not in monthly_charges:
                    monthly_charges[Month_Id[month]] = 0.00
                monthly_charges[Month_Id[month]] += float(row[5])
                

            if row[6] is not None:
                Global_Transaction_Data[User_Id][Credit_Company]['Credit'] += abs(float(row[6]))
            
            trans_sum += Global_Transaction_Data[User_Id][Credit_Company]['Debit']

            #print(Global_Transaction_Data)

        print(Global_Transaction_Data)
            
        #print(trans_sum)
        #print(dining_sum)
        print(monthly_charges)
        print(f" ------------------------------------ End Statement {i} -----------------------------------")
        


def load_statement_reports(statement_reports_store):
    statement_files = [] 
    for file in glob.glob(os.path.join(statement_reports_store, "*.xlsx")):
        if "~" in file:
            continue
        statement_files.append(file)
    return statement_files


# Credit/debit statements directory
statement_reports_store = "C:\\Users\\saddala\\AppData\\Local\\Programs\\Python\\Python310\\Budget_App\\Statements"
statements = load_statement_reports(statement_reports_store)

# Statement files
print(statements)

# Print contents of xlsx files
exhibit_data_from_statements(statements)

