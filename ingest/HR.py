#################################################################################
# Python Formatting Script from HR to Excel               			    #
# Creates excel formatted output from excel files                               #
# Usage newHR.py 				        			                            #
# Last modification 09/April/2024                                               #
# Version 2 - added user_id and standarized some fields data      #
#################################################################################

import pandas as pd
import ingest.format_name

class HR_creation:
    def Creo_HR(self):
        # Read the CSV and Excel files into DataFrames
        # DT = pd.read_csv("./files/HR - DT Emails.csv")
        DT = pd.read_excel("./files/DT Active EE 7.3.2024.xlsx")
        AVERRO = pd.read_excel("./files/Averro Active EE 7.3.2024.xlsx")
        NUWEST = pd.read_excel("./files/NuWest Active EE 7.3.2024.xlsx")
        LATAM = pd.read_excel("./files/List of contractors and employees -  LATAM.xlsx")

        AVERRO.rename(columns={'Work_Email': 'email'}, inplace=True)
        AVERRO.rename(columns={'Employee_Name': 'name'.strip()}, inplace=True)
        NUWEST.rename(columns={'Employee_Name': 'name'.strip()}, inplace=True)
        NUWEST.rename(columns={'Work_Email': 'email'.strip()}, inplace=True)
        NUWEST.rename(columns={'Employee_Status': 'status'}, inplace=True)
        DT.rename(columns={'Work_Email': 'email'.strip()}, inplace=True)
        DT.rename(columns={'Employee_Name': 'name'.strip()}, inplace=True)
        AVERRO['source'] = 'AVERRO'
        AVERRO['status'] = 'Active'
        DT['source'] = 'DT'
        DT['status'] = 'Active'
        NUWEST['source'] = 'NUWEST'
        NUWEST['manager'] = ''
        DT['manager'] = ''
        AVERRO['manager'] = ''
        LATAM.rename(columns={'EMPLOYEES': 'name'.strip()}, inplace=True)
        LATAM.rename(columns={'EMAIL': 'email'.strip()}, inplace=True)
        LATAM.rename(columns={'MANAGER': 'manager'.strip()}, inplace=True)
        LATAM['source'] = 'LATAM'

        # # Apply formatear_nombre function to each element of Employee_Name column
        # DT['name'] = DT['name'].apply(lambda x: ' '.join((ingest.format_name.formatear_nombre(x))))
        # # Apply formatear_nombre function to each element of Employee_Name column
        # AVERRO['name'] = AVERRO['name'].apply(lambda x: ' '.join((ingest.format_name.formatear_nombre(x))))
        # # Apply formatear_nombre function to each element of Employee_Name column
        # NUWEST['name'] = NUWEST['name'].apply(lambda x: ' '.join((ingest.format_name.formatear_nombre(x))))
        # # Apply formatear_nombre function to each element of Employee_Name column
        # LATAM['name'] = LATAM['name'].apply(lambda x: ' '.join((ingest.format_name.formatear_nombre(x))))

        # Merge the DataFrames on the 'email' column
        HR = pd.concat([DT, AVERRO, NUWEST, LATAM], ignore_index=True)
        # HR = pd.concat([DT, AVERRO, NUWEST], ignore_index=True)

        # Assuming 'name' and 'status' are common columns in all DataFrames,
        HR = HR[['source','name', 'email', 'status', 'Employee_Code', 'manager']]

        # standarize email data and include a new field colled user_id to compare against the AD.
        HR['email'] = HR['email'].str.lower()
        HR['user_id'] = HR['email'].apply(lambda x: x.split('@')[0].lower() if isinstance(x, str) and '@' in x else x)

#       If you want to remove extra spaces after removing the patterns
        HR['user_id'] = HR['user_id'].str.strip()

        HR['Original_Name'] = HR['name']

        # Apply formatear_nombre function to each element of Employee_Name column
        HR['name'] = HR['name'].apply(lambda x: (ingest.format_name.formatear_nombre(x)))

        # Convert 'Employee_Code' to string
        HR['Employee_Code'] = HR['Employee_Code'].astype(str)

        return HR