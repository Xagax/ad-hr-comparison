#################################################################################
# Python Formatting Script from AD vs HR to Excel               			    #
# Creates excel formatted output from excel files                               #
# Last modification 1/April/2024                                                #
# Version 3                                                                     #
# Changes from v2: changed File names and added more comments to the code       #
#################################################################################

import pandas as pd
import streamlit as st
from rapidfuzz import fuzz, process
from datetime import datetime, timedelta


# Define a function to handle NaN values and lowercase conversion
def process_email(email):
    if pd.isnull(email):
        return email
    else:
        return email.lower()
    

# Read the Excel file with multiple tabs
excel_file = pd.ExcelFile("files/AzureAccountsReport_v2.0.xlsx")

# Initialize an empty list to store dataframes
dfs = []

# Read generic users from generic_users.txt
with open("generic_users.txt", "r") as file:
    generic_users = file.read().splitlines()


def formatear_nombre(nombre_completo):
    # Check if the comma exists in nombre_completo
    if ", " in nombre_completo:
        # If the comma exists, split using ", "
        apellido, nombre = nombre_completo.split(", ", 1)
    else:
        # If the comma doesn't exist, split using any whitespace character
        apellido, nombre = nombre_completo.split(None, 1)

    # Now apellido will contain the last name and nombre will contain the first name

    # Convertir el nombre y el apellido a mayúsculas
    nombre = nombre.title()
    apellido = apellido.title()

    # Devolver el nombre y el apellido en el formato requerido
    nombre_formateado = f"{nombre} {apellido}"
    return nombre_formateado

# Function to check if a user is generic
def is_generic(display_name):
    return display_name in generic_users

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

    # Check if Displayname is a generic user
    df['Generic'] = df['DisplayName'].apply(is_generic)
    
    # Concatenate the dataframe to the list
    dfs.append(df)

# Concatenate all dataframes into one
AD = pd.concat(dfs, ignore_index=True)

# Reorder columns to place "Tenant" as the first column
column_order = ['Tenant'] + [col for col in AD.columns if col != 'Tenant']
AD = AD[column_order]

# 1. Merge azure_dt_df with HR - DT Emails CSV file
HR_DT = pd.read_csv("files/HR - DT Emails.csv")

# Convert email addresses to lowercase for comparison
HR_DT['Work_Email'] = HR_DT['Work_Email'].str.lower()

# Apply formatear_nombre function to each element of Employee_Name column
HR_DT['name_dt'] = HR_DT['Employee_Name'].apply(formatear_nombre)

# Create a new column called user_id
HR_DT['user_id'] = HR_DT['Work_Email'].apply(lambda x: x.split('@')[0] if '@' in x.lower() else x)

AD['fuzzy_ratio_dt'] = AD['DisplayName'].apply(lambda x: process.extractOne(x, HR_DT['name_dt'], scorer=fuzz.ratio))

# Extract the ratio value and store it as float
AD['fuzzy_ratio_dt'] = AD['fuzzy_ratio_dt'].apply(lambda x: x[1])

# Filter AD DataFrame based on fuzzy_ratio_dt > 80
AD_filtered = AD[AD['fuzzy_ratio_dt'] > 85]

# Merge AD_filtered and HR_DT based on user_id
merged_data = pd.merge(AD_filtered, HR_DT, how="left", left_on="user_id", right_on="user_id", suffixes=('_AD', '_DT'))

# Calculate ratio of names between AD DisplayName and HR_DT name_dt
merged_data['ratio_dt'] = merged_data.apply(lambda x: fuzz.ratio(x.DisplayName, x.name_dt), axis=1)


# Merge and check if email address contains a specific substring
AD = pd.merge(AD, HR_DT, how="left", left_on="user_id", right_on="user_id", suffixes=('_AD', '_DT'))

# ratio del nombre en el AD (displayname) y el nombre en HR.
AD['ratio_dt'] = AD.apply(lambda x: fuzz.ratio(x.DisplayName, x.name_dt), axis=1)

# 2. Merge averro_signin_df with HR- AVERRO Email Addresses Active XLSX file
HR_AVERRO = pd.read_excel("files/HR- Averro Email Addresses Active.xlsx")

# Identify non-integer values
non_integer_values_mask = HR_AVERRO['Employee_Code'].apply(lambda x: not str(x).isdigit())

# Get non-integer values
non_integer_values = HR_AVERRO.loc[non_integer_values_mask, 'Employee_Code']

HR_AVERRO['Rehire_Date'] = HR_AVERRO['Rehire_Date'].astype(str)

# Handle non-integer values (e.g., replace with NaN)
HR_AVERRO['Employee_Code'] = pd.to_numeric(HR_AVERRO['Employee_Code'], errors='coerce')

# Convert to integer type
HR_AVERRO['Employee_Code'] = HR_AVERRO['Employee_Code'].astype('Int64')

# Convert email addresses to lowercase for comparison
HR_AVERRO['Work_Email'] = HR_AVERRO['Work_Email'].str.lower()

# Apply formatear_nombre function to each element of Employee_Name column
HR_AVERRO['name_averro'] = HR_AVERRO['Employee_Name'].apply(formatear_nombre)

# Create a new column called user_id
HR_AVERRO['user_id'] = HR_AVERRO['Work_Email'].apply(lambda x: process_email(x).split('@')[0] if '@' in str(process_email(x)) else x)

AD['fuzzy_ratio_averro'] = AD['DisplayName'].apply(lambda x: process.extractOne(x, HR_AVERRO['name_averro'], scorer=fuzz.ratio))

# Extract the ratio value and store it as float
AD['fuzzy_ratio_averro'] = AD['fuzzy_ratio_averro'].apply(lambda x: x[1])

# Merge and check if email address contains a specific substring
AD = pd.merge(AD, HR_AVERRO, how="left", left_on="user_id", right_on="user_id", suffixes=('_DT', '_averro'))

# ratio del nombre en el AD (displayname) y el nombre en HR.
AD['ratio_averro'] = AD.apply(lambda x: fuzz.ratio(x.DisplayName, x.name_averro), axis=1)


# 3. Merge nw_signin_df with HR - Report for Hailey NuWest emails XLSX file
HR_NUWEST = pd.read_excel("files/HR - Report for Hailey NuWest emails.xlsx")

# Convert email addresses to lowercase for comparison
HR_NUWEST['Work_Email'] = HR_NUWEST['Work_Email'].str.lower()

# Apply formatear_nombre function to each element of Employee_Name column
HR_NUWEST['name_nuwest'] = HR_NUWEST['Employee Name'].apply(formatear_nombre)

# Create a new column called user_id
HR_NUWEST['user_id'] = HR_NUWEST['Work_Email'].apply(lambda x: process_email(x).split('@')[0] if '@' in str(process_email(x)) else x)

AD['fuzzy_ratio_nuwest'] = AD['DisplayName'].apply(lambda x: process.extractOne(x, HR_NUWEST['name_nuwest'], scorer=fuzz.ratio))

# Extract the ratio value and store it as float
AD['fuzzy_ratio_nuwest'] = AD['fuzzy_ratio_nuwest'].apply(lambda x: x[1])

# Merge and check if email address contains a specific substring
AD = pd.merge(AD, HR_NUWEST, how="left", left_on="user_id", right_on="user_id", suffixes=('_AVERRO', '_nuwest'))

# Check if a match was found in any of the columns and assign True/False accordingly
AD['match_found'] = ~AD[['Employee_Name_averro', 'Employee Name', 'Employee_Name_DT']].isnull().all(axis=1)

# ratio del nombre en el AD (displayname) y el nombre en HR.
AD['ratio_nuwest'] = AD.apply(lambda x: fuzz.ratio(x.DisplayName, x.name_nuwest), axis=1)

# Filter AD DataFrame based on ratio_nuwest == 100 and match_found == False
AD_nuwest_ratio_100_no_match = AD[(AD['ratio_nuwest'] > 85) & (AD['match_found'] == False)]

# Merge filtered AD DataFrame with HR_NUWEST data
AD_nuwest_with_hr_no_match = pd.merge(AD_nuwest_ratio_100_no_match, HR_NUWEST, how="left", left_on="name_nuwest", right_on="name_nuwest", suffixes=('_AD', '_HR'))

# Update match_found column to True for the merged records
AD_nuwest_with_hr_no_match['match_found'] = True

# Concatenate the merged DataFrame with the rest of the AD DataFrame
AD_with_hr_no_match = pd.concat([AD[~((AD['ratio_nuwest'] > 85) & (AD['match_found'] == False))], AD_nuwest_with_hr_no_match])

# Update match_found column to True for records where a match was found and match_found is False
AD.loc[(AD['match_found'] == False) & ((AD['fuzzy_ratio_nuwest'] > 85) | (AD['fuzzy_ratio_averro'] > 85) | (AD['fuzzy_ratio_dt'] >85)), 'match_found'] = True

# AD.to_excel("azure_dt_data.xlsx", index=False)

#++++++++++++++++++ fin ++++++++++++++++++++
# cambio nombres de campos para hacerlo mas amigable!!
AD = AD.rename(columns={'Employee Status': 'NuWest Employee Status'})
AD = AD.rename(columns={'Hire_Date': 'Hire_Date_Averro'})
AD = AD.rename(columns={'Rehire_Date': 'Rehire_Date_Averro'})
AD = AD.rename(columns={'Client_Desc': 'Client_Desc_Averro'})
AD = AD.rename(columns={'Employee Name': 'Employee Name NuWest'})
AD = AD.rename(columns={'Type Desc': 'Type Desc NuWest'})

# Calculate the date 90 days ago
current_date = datetime.now()
date_90_days_ago = current_date - timedelta(days=90)

# ***Genera tab con condiciones 
# Create a Pandas Excel writer

# Create a Pandas Excel writer
writer = pd.ExcelWriter('AD_vs_HR.xlsx', engine='openpyxl')

# Write your DataFrame to a new worksheet named 'FilteredData'
AD.to_excel(writer, sheet_name='ADvsHR', index=False)

# Filter the DataFrame based on the given condition
filtered_df = AD[(AD['AccountEnabled'] == True) & 
                 (AD['Generic'] == False) & 
                 (AD['match_found'] == True) & 
                 (AD['NuWest Employee Status'] == 'Terminated')]

# Select only the desired columns
filtered_df = filtered_df[['Tenant','DisplayName', 'UserPrincipalName', 'UserType', 'AccountEnabled','Department', 'JobTitle','Manager','CreatedDateTime', 'NuWest Employee Status' ]]


# Write the filtered DataFrame to a new worksheet named 'Enabled&Terminated'
filtered_df.to_excel(writer, sheet_name='Enabled&TerminatedALL', index=False)

# Filter the DataFrame based on the given condition
filtered_df = AD[(AD['Tenant'] == 'Averro') & 
                 (AD['AccountEnabled'] == True) & 
                 (AD['Generic'] == False) & 
                 (AD['match_found'] == True) & 
                 (AD['NuWest Employee Status'] == 'Terminated')]

# Select only the desired columns
filtered_df = filtered_df[['Tenant','DisplayName', 'UserPrincipalName', 'UserType', 'AccountEnabled', 'Department', 'JobTitle','Manager', 'CreatedDateTime', 'NuWest Employee Status', 'Employee_Code_averro', 'Employee_Name_averro', 'Hire_Date_Averro', 'Rehire_Date_Averro', 'Client_Desc_Averro' ]]

# Write the filtered DataFrame to a new worksheet named 'Enabled&Terminated'
filtered_df.to_excel(writer, sheet_name='Averro_Enabled&Terminated', index=False)

# Filter the DataFrame based on the given condition
filtered_df = AD[(AD['Tenant'] == 'DT') & 
                 (AD['AccountEnabled'] == True) & 
                 (AD['Generic'] == False) & 
                 (AD['match_found'] == True) & 
                 (AD['NuWest Employee Status'] == 'Terminated')]

# Select only the desired columns
filtered_df = filtered_df[['Tenant','DisplayName', 'UserPrincipalName', 'UserType', 'AccountEnabled', 'Department', 'JobTitle','Manager', 'CreatedDateTime', 'NuWest Employee Status', 'Employee_Code_DT', 'Employee_Name_DT', 'Work_Email_DT' ]]

# Write the filtered DataFrame to a new worksheet named 'Enabled&Terminated'
filtered_df.to_excel(writer, sheet_name='DT_Enabled&Terminated', index=False)

# Filter the DataFrame based on the given condition
filtered_df = AD[(AD['Tenant'] == 'NuWest') & 
                 (AD['AccountEnabled'] == True) & 
                 (AD['Generic'] == False) & 
                 (AD['match_found'] == True) & 
                 (AD['NuWest Employee Status'] == 'Terminated')]

# Select only the desired columns
filtered_df = filtered_df[['Tenant','DisplayName', 'UserPrincipalName', 'UserType', 'AccountEnabled', 'Department', 'JobTitle','Manager', 'CreatedDateTime','Employee Name NuWest', 'Type Desc NuWest', 'NuWest Employee Status', 'Work_Email_nuwest', 'name_nuwest']]

# Write the filtered DataFrame to a new worksheet named 'Enabled&Terminated'
filtered_df.to_excel(writer, sheet_name='NuWest_Enabled&Terminated', index=False)

# Filter the DataFrame based on the given condition
filtered_df = AD[(AD['AccountEnabled'] == True) & 
                 (AD['Generic'] == False) & 
                 (pd.isnull(AD['LastSignInDateTime']))]

# Select only the desired columns
filtered_df = filtered_df[['Tenant','DisplayName', 'UserPrincipalName', 'UserType', 'externo', 'AccountEnabled', 'Department', 'JobTitle','Manager', 'CreatedDateTime']]

# Write the filtered DataFrame to a new worksheet named 'FilteredResults'
filtered_df.to_excel(writer, sheet_name='Enabled&Never', index=False)

# Filter the DataFrame based on the given condition
filtered_df = AD[(AD['AccountEnabled'] == True) & 
                 (AD['Generic'] == False) & 
                 (AD['LastSignInDateTime'] < date_90_days_ago)]

# Select only the desired columns
filtered_df = filtered_df[['Tenant','DisplayName', 'UserPrincipalName', 'UserType', 'externo', 'CreatedDateTime', 'AccountEnabled', 'Department', 'JobTitle','Manager', 'LastSignInDateTime']]

# Write the filtered DataFrame to a new worksheet named 'Enabled&90days'
filtered_df.to_excel(writer, sheet_name='Enabled&90days', index=False)

# Save the changes
writer.close()


metrics_df = AD.groupby('Tenant').agg(
    Total_Users=('user_id', 'count'),
    Active_Users=('AccountEnabled', lambda x: x.sum()),
    Disabled_Users=('AccountEnabled', lambda x: (x == False).sum()),
    Normal_Users=('Generic', lambda x: (x == False).sum()),
    Generic_Users=('Generic', lambda x: x.sum()),
    Normal_Active_Users=('AccountEnabled', lambda x: ((x == True) & (AD['Generic'] == False)).sum()),
    Normal_Active_Users_Terminated = ('AccountEnabled', lambda x: ((x == True) & (AD['Generic'] == False) & (AD['match_found'] == True) & (AD['NuWest Employee Status'] == 'Terminated')).sum()),
    Normal_Active_Users_Never = ('AccountEnabled', lambda x: ((x == True) & (AD['Generic'] == False) & (pd.isnull(AD['LastSignInDateTime']))).sum()),
    # Modify the condition to check if LastSignInDateTime is greater than 90 days ago
    Normal_Active_Users_noaccess_90days = ('AccountEnabled', lambda x: ((x == True) & (AD['Generic'] == False) & (AD['LastSignInDateTime'] < date_90_days_ago)).sum())
).reset_index()


# Calculate total users, active users, and disabled users
total_users = metrics_df['Total_Users'].sum()
total_active_users = metrics_df['Active_Users'].sum()
total_disabled_users = metrics_df['Disabled_Users'].sum()
total_normal_users = metrics_df['Normal_Users'].sum()
total_generic_users = metrics_df['Generic_Users'].sum()
total_Normal_Active_Users = metrics_df['Normal_Active_Users'].sum()
total_Normal_Active_Users_Terminated = metrics_df['Normal_Active_Users_Terminated'].sum()
total_Normal_Active_Users_Never = metrics_df['Normal_Active_Users_Never'].sum()
total_Normal_Active_Users_noaccess_90days = metrics_df['Normal_Active_Users_noaccess_90days'].sum()

# Add a row with total counts
total_row = {'tenant': 'Total', 'Total_Users': total_users, 'Active_Users': total_active_users, 'Disabled_Users': total_disabled_users, 'Normal_Users': total_normal_users, 'Generic_Users': total_generic_users, 'Normal_Active_Users': total_Normal_Active_Users, 'Normal_Active_Users_Terminated': total_Normal_Active_Users_Terminated, 'Normal_Active_Users_Never': total_Normal_Active_Users_Never, 'Normal_Active_Users_noaccess_90days': total_Normal_Active_Users_noaccess_90days }
total_df = pd.DataFrame([total_row])

# Concatenate the total row with the metrics DataFrame
metrics_df = pd.concat([metrics_df, total_df], ignore_index=True)

# Save the metrics dataframe to a new Excel file
metrics_df.to_excel("Data4Metrics.xlsx", index=False)

# Create a Streamlit app
def main():
    st.set_page_config(layout='wide', initial_sidebar_state='expanded')
    with open('style.css') as f:
        st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)

    # st.title("Azure Data Analysis")

     # Display metrics
    st.markdown('### Metrics')
    col1, col2, col3, col4 = st.columns(4)
    
    # Access the "Total No. Users" value from the DataFrame
    total_users = metrics_df.loc[3, 'Total_Users']
    total_active_users = metrics_df.loc[3, 'Active_Users']
    total_active_users_terminated = metrics_df.loc[3, 'Normal_Active_Users_Terminated']
    total_active_users_never = metrics_df.loc[3, 'Normal_Active_Users_Never']
    total_Normal_Active_Users_noaccess_90days = metrics_df.loc[3, 'Normal_Active_Users_noaccess_90days']


    col1.metric("Total No. Users", total_users, f"{total_active_users} Active")
    col2.metric("Active Users Terminated", total_active_users_terminated, "%")
    col3.metric("Users Never Accessed", total_active_users_never, "%")
    col4.metric("Users no access for more 90 days", total_Normal_Active_Users_noaccess_90days, "%")    
    
    # Display the dataframe as a table
    st.subheader("Azure Data Table")
    st.write(AD)

    st.sidebar.header('Dashboard `version 1`')

    # # st.sidebar.subheader('Tenants')
    # time_hist_color = st.sidebar.selectbox('Tenants', ('DT', 'Averro', 'NueWest')) 

    # # st.sidebar.subheader('Estado de Usuarios')
    # time_hist_color = st.sidebar.selectbox('Accounts Status', ('Enabled', 'Disabled')) 

    # # # st.sidebar.subheader('Estado de Usuarios')
    # # time_hist_color = st.sidebar.selectbox('Accounts types', ('Member', 'Guest')) 


    # # Option to select data
    # st.sidebar.title("Select Data")
    # option = st.sidebar.selectbox(
    #     'Select a column:',
    #     AD.columns
    # )

    # # Display data based on selection
    # st.subheader("Selected Data")
    # st.write(AD[option])

    # # Plot a graph
    # st.sidebar.title("Graph")
    # graph_option = st.sidebar.selectbox(
    #     'Select a graph type:',
    #     ['Histogram', 'Bar Chart']
    # )

    st.sidebar.markdown('''
    ---
    Created by Luis Blanco.
    ''')


if __name__ == '__main__':
    main()