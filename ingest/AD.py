#################################################################################
# Python Formatting Script from AD vs HR to Excel               			    #
# Creates excel formatted output from excel files                               #
# Last modification 3/April/2024                                                #
# Version 4                                                                     #
# Changes from v3: changed Generics to user_ids and added more tabs             #
#################################################################################

import pandas as pd
import ingest.format_name

class AD_creation:

    # Define a function to handle NaN values and lowercase conversion
    def process_email(email):
        if pd.isnull(email):
            return email
        else:
            return email.lower()
    
    # Function to check if a user is generic
    def is_generic(self, id_generic):
        return id_generic in self.generic_ids


    def Creo_AD(self):

        # Read generic users from generic_users.txt
        with open("./files/id_generic.txt", "r") as file:
            self.generic_ids = file.read().splitlines()

        # Read the Excel file with multiple tabs
        excel_file = pd.ExcelFile("./files/AccountReport_July2024.xlsx")

        # Initialize an empty list to store dataframes
        dfs = []

        # Iterate over each tab in the Excel file - you need to change the name of the TAB
        for sheet_name in excel_file.sheet_names:
            # Read each tab as a dataframe
            df = pd.read_excel(excel_file, sheet_name=sheet_name)
            # you need to change the name of the TAB everytime you receive the file ******************
            if sheet_name == "DirectTechnologyUsersReport_Gra": sheet_name = "DT"
            if sheet_name == "AverroUsersReport_Graph": sheet_name = "Averro"
            if sheet_name == "NuWestUsersReport_Graph": sheet_name = "NuWest"

            # Add a new column with tab names
            df['Tenant'] = sheet_name

            # Check if Displayname is a generic user
            df['Generic'] = df['Id'].apply(self.is_generic)

            # Convert UserPrincipalName to lowercase
            df['UserPrincipalName'] = df['UserPrincipalName'].str.lower()

            # Conditionally replace "_" with "@" based on specific pattern
            df['UserPrincipalName'] = df['UserPrincipalName'].str.replace('_', '@', regex=False).where(df['UserPrincipalName'].str.contains('_averro.com#ext#@'), other=df['UserPrincipalName'])

            # Conditionally replace "_" with "@" based on specific pattern
            df['UserPrincipalName'] = df['UserPrincipalName'].str.replace('_', '@', regex=False).where(df['UserPrincipalName'].str.contains('_directtechnology.com#ext#@'), other=df['UserPrincipalName'])

            # Conditionally replace "_" with "@" based on specific pattern
            df['UserPrincipalName'] = df['UserPrincipalName'].str.replace('_', '@', regex=False).where(df['UserPrincipalName'].str.contains('@tagroupholdings.com'), other=df['UserPrincipalName'])

            # Conditionally replace "_" with "@" based on specific pattern
            df['UserPrincipalName'] = df['UserPrincipalName'].str.replace('_', '@', regex=False).where(df['UserPrincipalName'].str.contains('_nuwestgroup.com#ext#@'), other=df['UserPrincipalName'])

            # Conditionally replace "_" with "@" based on specific pattern
            df['UserPrincipalName'] = df['UserPrincipalName'].str.replace('_', '@', regex=False).where(df['UserPrincipalName'].str.contains('_tagroupholdings.com#ext#@'), other=df['UserPrincipalName'])

            # Conditionally replace "_" with "@" based on UserType
            df['UserPrincipalName'] = df.apply(lambda row: row['UserPrincipalName'].replace('_', '@') if row['UserType'] == 'Guest' else row['UserPrincipalName'], axis=1)
            
            # Convert UserPrincipalName to lowercase
            df['UserPrincipalName'] = df['UserPrincipalName'].str.lower()
            
            # Create a new column called Work_Email
            df['Work_Email'] = df['UserPrincipalName'].apply(lambda x: x.split('#ext#')[0] if '#ext#' in x.lower() else x)

            # Create a new column called 'externo'
            df['externo'] = df['UserPrincipalName'].apply(lambda x: True if '#ext#' in x.lower() else False)

            # Create a new column called user_id
            df['user_id'] = df['Work_Email'].apply(lambda x: x.split('@')[0] if '@' in x.lower() else x)

            # Remove sa. from at the beginning of user IDs / for example sa.lblanco
            df['user_id'] = df['user_id'].str.replace('sa.', '')

            # Remove .UY at the end of user IDs / for example lblanco.uy
            df['user_id'] = df['user_id'].str.replace('.uy', '')

            # Remove numbers at the end of user IDs / for example fmichanie0720  / el CEO de la empresa
            df['user_id'] = df['user_id'].str.replace(r'\d+$', '', regex=True)
            
#            If you want to remove extra spaces after removing the patterns
            df['user_id'] = df['user_id'].str.strip()
            
            # normalize names by replacing  various "*admin*" and "company names" forms in DisplayName for future control against names in HR files
            df['Original_DisplayName'] = df['DisplayName']
            df['DisplayName'] = df['DisplayName'].str.replace(' Admin Account', '')
            df['DisplayName'] = df['DisplayName'].str.replace(' Admin account', '')
            df['DisplayName'] = df['DisplayName'].str.replace(' admin account', '')
            df['DisplayName'] = df['DisplayName'].str.replace(' admin Account', '')
            df['DisplayName'] = df['DisplayName'].str.replace( r'\'s ', '', regex=True)
            df['DisplayName'] = df['DisplayName'].str.replace(' Admin', '')
            df['DisplayName'] = df['DisplayName'].str.replace(' admin', '')
            df['DisplayName'] = df['DisplayName'].str.replace('SA.', '')
            df['DisplayName'] = df['DisplayName'].str.replace('sa.', '')
            df['DisplayName'] = df['DisplayName'].str.replace(r'\'s', '', regex=True)
            df['DisplayName'] = df['DisplayName'].str.replace(' administration account', '')
            df['DisplayName'] = df['DisplayName'].str.replace(' Administration Account', '')
            df['DisplayName'] = df['DisplayName'].str.replace(' Administration account', '')
            df['DisplayName'] = df['DisplayName'].str.replace(' administration Account', '')
            df['DisplayName'] = df['DisplayName'].str.replace(' - Averro', '')
            df['DisplayName'] = df['DisplayName'].str.replace(' - NuWest Group', '')
            df['DisplayName'] = df['DisplayName'].str.replace(' - NuWest', '')
            df['DisplayName'] = df['DisplayName'].str.replace('NuWest - ', '')
            df['DisplayName'] = df['DisplayName'].str.replace(' DT', '')
            df['DisplayName'] = df['DisplayName'].str.replace('DT ', '')
            df['DisplayName'] = df['DisplayName'].str.replace(' (TAG)', '')
            # If you want to remove extra spaces after removing the patterns
            df['DisplayName'] = df['DisplayName'].str.strip()
            
            # Concatenate the dataframe to the list
            dfs.append(df)

        # Concatenate all dataframes into one
        AD = pd.concat(dfs, ignore_index=True)

        # Reorder columns to place "Tenant" as the first column
        column_order = ['Tenant'] + [col for col in AD.columns if col != 'Tenant']
        AD = AD[column_order]

        AD['DisplayName'] = AD['DisplayName'].apply(lambda x: (ingest.format_name.formatear_nombre(x)))
       
        return AD