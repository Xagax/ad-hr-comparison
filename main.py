#################################################################################
# Python Formatting Script from AD vs HR to Excel               			    #
# Creates excel formatted output from excel files                               #
# Last modification 4/July/2024                                                #
# Version 6                                                                     #
# Changes from v5:  #
#################################################################################

import pandas as pd
import streamlit as st
from rapidfuzz import fuzz, process
from datetime import datetime, timedelta
from ingest.HR import HR_creation
from ingest.AD import AD_creation

# ----------------- AD File ------------------------
Azure_AD = AD_creation()
AD = Azure_AD.Creo_AD() # retorna el AD en un pandas!
AD.to_excel("AD_Original.xlsx", index=False)

# ----------------- HR Files ------------------------
panda_HR = HR_creation()
HR = panda_HR.Creo_HR()   # retorna el HR unificado en un pandas!
HR.to_excel("HR_Original.xlsx", index=False)


# # # Calculo el acercamiento de DisplayName y Nombre en HR.
AD['fuzzy_ratio'] = AD['DisplayName'].apply(lambda x: process.extractOne(x, HR['name'], scorer=fuzz.ratio))
AD['fuzzy_ratio'] = AD['fuzzy_ratio'].apply(lambda x: x[1])

def check_match(row, HR):
    if row['fuzzy_ratio'] == 100: # si el acercamiento es del 100 entonces el nombre es el mismo que esta en la base de datos de HR.
        return 'Yes'
    elif row['fuzzy_ratio'] > 84 and row['fuzzy_ratio'] < 100:
        hr_match = HR[HR['user_id'] == row['user_id']]
        if not hr_match.empty:
            return 'Yes'
        elif row['fuzzy_ratio'] > 84:
            return 'Yes'
    return 'No'

def update_match_found(AD, HR):
    AD['Match_found'] = AD.apply(lambda row: check_match(row, HR), axis=1)
    
    # Filter AD DataFrame where Match_found is 'Yes'
    matched_indices = AD.index[AD['Match_found'] == 'Yes']
    
    # Iterate over the matched indices and merge HR data into AD DataFrame
    for idx in matched_indices:
        nombre = AD.loc[idx, 'DisplayName']
        hr_match = HR[HR['name'] == nombre]
        if not hr_match.empty:
            for col in hr_match.columns:
                if col != 'name':  # Avoid duplicating name column
                    AD.loc[idx, col] = hr_match.iloc[0][col]
    
    # Filter AD DataFrame where Match_found is 'No'
    unmatched_indices = AD.index[AD['Match_found'] == 'No']
    
    # Iterate over the unmatched indices and include HR data in AD DataFrame
    for idx in unmatched_indices:
        user_id = AD.loc[idx, 'user_id']
        hr_match = HR[HR['user_id'] == user_id]
        if not hr_match.empty:
            for col in hr_match.columns:
                if col not in AD.columns:  # Add HR data if column doesn't exist in AD
                    AD[col] = None
                AD.loc[idx, col] = hr_match.iloc[0][col]
            AD.loc[idx, 'Match_found'] = 'Yes'  # Set Match_found to 'Yes' if match is found

    return AD

# Assuming you already have the AD DataFrame and HR DataFrame
AD = update_match_found(AD, HR)

# Calculate the date 90 days ago
current_date = datetime.now()
date_90_days_ago = current_date - timedelta(days=90)

# # Agrega a status "NotMatchHR" a todos los que no encontró al hacer el match con HR.
# AD['status'] = AD.apply(
#     lambda row: 'NotMatchHR' if (
#         row['AccountEnabled'] == True and   # Cuentas activas
#         row['Generic'] == False and         # Cuentas NO genericas
#         # pd.isnull(row['source']) and        # No lo encontro en HR
#         row['Match_found'] == "No"      # No lo encontro por cercanía de nombre
#         # row['LastSignInDateTime'] < date_90_days_ago    # No accedió por mas de 90 días.
#     ) else row['status'],
#     axis=1
# )

#**** comienza a colectar metricas ******

metrics_df = AD.groupby('Tenant').agg(
    Total_Users=('Id', 'count'),
    Active_Users=('AccountEnabled', lambda x: x.sum()),
    Disabled_Users=('AccountEnabled', lambda x: (x == False).sum()),
    Normal_Users=('Generic', lambda x: (x == False).sum()),
    Generic_Users=('Generic', lambda x: x.sum()),
    Generic_Active_Users=('Generic', lambda x: ((x == True) & (AD['AccountEnabled'] == True)).sum()),
    Normal_Active_Users=('AccountEnabled', lambda x: ((x == True) & (AD['Generic'] == False)).sum()),
    Normal_Active_Users_NotMatchHR = ('AccountEnabled', lambda x: ((x == True) & (AD['Generic'] == False) & (AD['Match_found'] == 'No')).sum()),
    Normal_Active_Users_Never = ('AccountEnabled', lambda x: ((x == True) & (AD['Generic'] == False) & (pd.isnull(AD['LastSignInDateTime']))).sum()),
    Normal_Active_Users_noaccess_90days = ('AccountEnabled', lambda x: ((x == True) & (AD['Generic'] == False) & (AD['LastSignInDateTime'] < date_90_days_ago)).sum()),
    Normal_Active_Users_NotMatch_HR = ('AccountEnabled', lambda x: ((x == True) & (AD['Generic'] == False) & (AD['UserType'] == 'Member') & (AD['Match_found'] == 'No')).sum())
).reset_index()


# Calculate total users, active users, and disabled users
total_users = metrics_df['Total_Users'].sum()
total_active_users = metrics_df['Active_Users'].sum()
total_disabled_users = metrics_df['Disabled_Users'].sum()
total_normal_users = metrics_df['Normal_Users'].sum()
total_generic_users = metrics_df['Generic_Users'].sum()
total_active_generic_users = metrics_df['Generic_Active_Users'].sum()
total_Normal_Active_Users = metrics_df['Normal_Active_Users'].sum()
total_Normal_Active_Users_NotMatchHR = metrics_df['Normal_Active_Users_NotMatchHR'].sum()
total_Normal_Active_Users_Never = metrics_df['Normal_Active_Users_Never'].sum()
total_Normal_Active_Users_noaccess_90days = metrics_df['Normal_Active_Users_noaccess_90days'].sum()
total_Normal_Active_Users_NotMatch_HR = metrics_df['Normal_Active_Users_NotMatch_HR'].sum()
# Add a row with total counts
total_row = {'tenant': 'Total', 
             'Total_Users': total_users, 
             'Active_Users': total_active_users, 
             'Disabled_Users': total_disabled_users, 
             'Normal_Users': total_normal_users, 
             'Generic_Users': total_generic_users,
             'Generic_Active_Users': total_active_generic_users,
             'Normal_Active_Users': total_Normal_Active_Users, 
             'Normal_Active_Users_NotMatchHR': total_Normal_Active_Users_NotMatchHR, 
             'Normal_Active_Users_Never': total_Normal_Active_Users_Never, 
             'Normal_Active_Users_noaccess_90days': total_Normal_Active_Users_noaccess_90days,
             'Normal_Actie_Users_NotMatch_HR': total_Normal_Active_Users_NotMatch_HR }
total_df = pd.DataFrame([total_row])

# Concatenate the total row with the metrics DataFrame
metrics_df = pd.concat([metrics_df, total_df], ignore_index=True)

#+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
# Create a Streamlit app
#+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def main():
    st.set_page_config(layout='wide', initial_sidebar_state='expanded')
    with open('style.css') as f:
        st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)


    # Access the "Total No. Users" value from the DataFrame
    total_users = metrics_df.loc[3, 'Total_Users']
    total_Normal_Active_Users = metrics_df.loc[3, 'Normal_Active_Users']
    total_active_users_NotMatchHR = metrics_df.loc[3, 'Normal_Active_Users_NotMatchHR']
    total_active_users_never = metrics_df.loc[3, 'Normal_Active_Users_Never']
    total_Normal_Active_Users_noaccess_90days = metrics_df.loc[3, 'Normal_Active_Users_noaccess_90days']
    total_Normal_Actie_Users_NotMatch_HR =  metrics_df.loc[3, 'Normal_Actie_Users_NotMatch_HR']
    total_active_generic_users = metrics_df.loc[3, 'Generic_Active_Users']

    st.title("Averro | NuWest | DT")
    st.header("Azure AD vs HR")
    # Display metrics in streamlit
    st.markdown('### Metrics')
    with st.container():
            col1, col2, col3 = st.columns(3)
            col1.metric("Total de Usuarios Activos (*)", total_Normal_Active_Users, f"{total_users} Total # in AD")
            col2.metric("Activos & NotMatchHR", total_active_users_NotMatchHR)
            col3.metric("Activos & Nunca Accedieron", total_active_users_never)
    
    st.caption('(*) - Número total de usuarios activos "NO" incluye a las cuentas genericas.')

    with st.container():
            col4, col5, col6 = st.columns(3)
            col4.metric("Activos & No accedieron por mas de 90 días", total_Normal_Active_Users_noaccess_90days)
            col5.metric("Activos & son Member & No_Match en HR", total_Normal_Actie_Users_NotMatch_HR)
            col6.metric("No. total de Genericos y Activos", total_active_generic_users, f"{total_generic_users} Total # in AD")


    # # Display the dataframe as a table
    # st.subheader("Azure AD Data Table")
    # st.write(AD)

    # st.sidebar.header('`version 6`')

    #+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    # Create a Pandas Excel writer
    writer = pd.ExcelWriter('AD_vs_HR.xlsx', engine='openpyxl')

    # Create a DataFrame for the summary
    summary_data = {
        'Summary - Tabs included': [
            "",
            "",
            "ADvsHR: AD completo con usuarios que encontró en los 3 archivos de HR.",
            "",
            "Enabled&NotMatchHRALL: Todos los usuarios Enabled en los 3 tenants y que encontró al compararlo contra el archivo de HR de NuWest y que además también tienen estado de Terminated.",
            "",
            "Averro_Enabled&NotMatchHR: Todos los usuarios en el Tenant de Averro y que tienen estado de Terminated en el archivo de HR de Nuwest.",
            "",
            "DT_Enabled&NotMatchHR: Todos los usuarios en el Tenant de DT y que tienen estado de Terminated en el archivo de HR de Nuwest.",
            "",
            "NuWest_Enabled&NotMatchHR: Todos los usuarios en el Tenant de Nuwest y que tienen estado de Terminated en el archivo de HR de Nuwest.",
            "",
            "Enabled&Never: Todos los usuarios en los 3 tenants que están Enabled y nunca han accedido (no incluye los genéricos).",
            "",
            "Enabled&90days: Todos los usuarios en los 3 tenants que están Enabled y nunca han por más de 90 días a la fecha de hoy (no incluye los genéricos).",
            "",
            "Enabled&NOTmatchHR: Todos los usuarios en los 3 tenants que están Enabled y no tienen un match en HR y además son del tipo de Usuario Member. (no incluye los genéricos)."
            "",
            "Generic: Todos los usuarios en los 3 tenants que son Genericos y están activos",
            "",
            "Guest: Todos los usuarios en los 3 tenants que son del tipo 'Guest' y están activos",
            "",
            "Externos: Todos los usuarios en los 3 tenants que son Externos y están activos",       
        ]
    }

    summary_df = pd.DataFrame(summary_data)

    # Write the summary DataFrame to a new worksheet named 'Summary'
    summary_df.to_excel(writer, sheet_name='Summary', index=False)

    # # Set the tab color
    wb = writer.book
    ws = wb['Summary']
    ws.sheet_properties.tabColor = '3B48F7'

    # Write the summary DataFrame to a new worksheet named 'Data4Metrics'
    metrics_df.to_excel(writer, sheet_name="Data4Metrics", index=False)

    # # Set the tab color
    wb = writer.book
    ws = wb['Summary']
    ws.sheet_properties.tabColor = 'FFFF00'

    # Write your DataFrame to a new worksheet named 'FilteredData'
    AD.to_excel(writer, sheet_name='ADvsHR', index=False)

    # # Set the tab color
    wb = writer.book
    ws = wb['ADvsHR']
    ws.sheet_properties.tabColor = '54DE8C'


    # Save HR to a new Excel file
    HR.to_excel(writer, sheet_name='HR_All', index=False)

    # # Set the tab color
    wb = writer.book
    ws = wb['HR_All']
    ws.sheet_properties.tabColor = 'C2DB57'

#------------------------------------------------------
    # Filter the DataFrame based on the given condition
    filtered_df = AD[(AD['AccountEnabled'] == True) & 
                    (AD['Generic'] == False) & 
                    (AD['Match_found'] == 'No')]

    # Select only the desired columns
    filtered_df = filtered_df[['Tenant','Id','DisplayName', 'UserPrincipalName', 'UserType', 'AccountEnabled','Department', 'JobTitle','Manager','CreatedDateTime']]

    # Write the filtered DataFrame to a new worksheet named 'Enabled&NotMatchHRALL'
    filtered_df.to_excel(writer, sheet_name='ALL-Enabled&NotMatchHR', index=False)

#------------------------------------------------------
    # Filter the DataFrame based on the given condition
    filtered_df = AD[
        (AD['Tenant'].str.startswith('Averro')) & 
        (AD['AccountEnabled'] == True) & 
        (AD['Generic'] == False) & 
        (AD['Match_found'] == 'No')
    ]

    # Select only the desired columns
    filtered_df = filtered_df[['Tenant','Id','DisplayName', 'UserPrincipalName', 'UserType', 'AccountEnabled','Department', 'JobTitle','Manager','CreatedDateTime']]

    # Write the filtered DataFrame to a new worksheet named 'Averro_Enabled&Terminated'
    filtered_df.to_excel(writer, sheet_name='Averro_Enabled&NotMatchHR', index=False)

#------------------------------------------------------
    # Filter the DataFrame based on the given condition
    filtered_df = AD[
        (AD['Tenant'].str.startswith('DT')) & 
        (AD['AccountEnabled'] == True) & 
        (AD['Generic'] == False) & 
        (AD['Match_found'] == 'No')
    ]

    # Select only the desired columns
    filtered_df = filtered_df[['Tenant','Id','DisplayName', 'UserPrincipalName', 'UserType', 'AccountEnabled','Department', 'JobTitle','Manager','CreatedDateTime']]

    # Write the filtered DataFrame to a new worksheet named 'DT_Enabled&NotMatchHR'
    filtered_df.to_excel(writer, sheet_name='DT_Enabled&NotMatchHR', index=False)

#------------------------------------------------------
    # Filter the DataFrame based on the given condition
    filtered_df = AD[
        (AD['Tenant'].str.startswith('NW')) & 
        (AD['AccountEnabled'] == True) & 
        (AD['Generic'] == False) & 
        (AD['Match_found'] == 'No')
    ]

    # Select only the desired columns
    filtered_df = filtered_df[['Tenant','Id','DisplayName', 'UserPrincipalName', 'UserType', 'AccountEnabled','Department', 'JobTitle','Manager','CreatedDateTime']]

    # Write the filtered DataFrame to a new worksheet named 'Enabled&NotMatchHR'
    filtered_df.to_excel(writer, sheet_name='NuWest_Enabled&NotMatchHR', index=False)

#------------------------------------------------------
    # Filter the DataFrame based on Enabled&NOTmatchHR ************
    filtered_df = AD[(AD['AccountEnabled'] == True) & 
                    (AD['Generic'] == False) & 
                    (AD['UserType'] == 'Member') &
                    (AD['Match_found'] == 'No')]

    # Select only the desired columns
    filtered_df = filtered_df[['Tenant','Id','DisplayName', 'UserPrincipalName', 'UserType', 'AccountEnabled','Department', 'JobTitle','Manager','CreatedDateTime']]

    # Write the filtered DataFrame to a new worksheet named 'Enabled&NOTmatchHR'
    filtered_df.to_excel(writer, sheet_name='ALL-MEMBER-Enabled&NOTmatchHR', index=False)

#------------------------------------------------------
    # Filter the DataFrame based on the given condition
    filtered_df = AD[(AD['AccountEnabled'] == True) & 
                    (AD['Generic'] == False) & 
                    (pd.isnull(AD['LastSignInDateTime']))]

    # Select only the desired columns
    filtered_df = filtered_df[['Tenant','Id','DisplayName', 'UserPrincipalName', 'UserType', 'externo', 'AccountEnabled', 'Department', 'JobTitle','Manager', 'CreatedDateTime']]

    # Write the filtered DataFrame to a new worksheet named 'FilteredResults'
    filtered_df.to_excel(writer, sheet_name='ALL-Enabled&Never', index=False)

#------------------------------------------------------
    # Filter the DataFrame based on Enabled&90days ************
    filtered_df = AD[(AD['AccountEnabled'] == True) & 
                    (AD['Generic'] == False) & 
                    (AD['LastSignInDateTime'] < date_90_days_ago)]

    # Select only the desired columns
    filtered_df = filtered_df[['Tenant','Id','DisplayName', 'UserPrincipalName', 'UserType', 'externo', 'CreatedDateTime', 'AccountEnabled', 'Department', 'JobTitle','Manager', 'LastSignInDateTime']]

    # Write the filtered DataFrame to a new worksheet named 'Enabled&90days'
    filtered_df.to_excel(writer, sheet_name='ALL-Enabled&>90days', index=False)

#------------------------------------------------------

    # Filter the DataFrame based on Generic ************
    filtered_df = AD[(AD['AccountEnabled'] == True) & 
                    (AD['Generic'] == True)]

    # Select only the desired columns
    filtered_df = filtered_df[['Tenant','Id','DisplayName', 'UserPrincipalName', 'UserType', 'externo', 'CreatedDateTime', 'AccountEnabled', 'Department', 'JobTitle', 'Manager', 'LastSignInDateTime']]

    # Write the filtered DataFrame to a new worksheet named 'Generic'
    filtered_df.to_excel(writer, sheet_name='ALL-GENERIC-Enabled', index=False)

#------------------------------------------------------
    # Filter the DataFrame based on Guest ************
    filtered_df = AD[(AD['AccountEnabled'] == True) & 
                    (AD['Generic'] == False) & 
                    (AD['UserType'] == 'Guest')]

    # Select only the desired columns
    filtered_df = filtered_df[['Tenant','Id','DisplayName', 'UserPrincipalName', 'UserType', 'externo', 'CreatedDateTime', 'AccountEnabled', 'Department', 'JobTitle', 'Manager', 'LastSignInDateTime']]

    # Write the filtered DataFrame to a new worksheet named 'Guest'
    filtered_df.to_excel(writer, sheet_name='ALL-Guest-Enabled', index=False)

#------------------------------------------------------
    # Filter the DataFrame based on Externos ************
    filtered_df = AD[(AD['AccountEnabled'] == True) & 
                    (AD['Generic'] == False) & 
                    (AD['externo'] == True)]

    # Select only the desired columns
    filtered_df = filtered_df[['Tenant','Id','DisplayName', 'UserPrincipalName', 'UserType', 'CreatedDateTime', 'AccountEnabled', 'Department', 'JobTitle', 'Manager', 'LastSignInDateTime']]

    # Write the filtered DataFrame to a new worksheet named 'Externos'
    filtered_df.to_excel(writer, sheet_name='ALL-EXTERNOS-Enabled', index=False)

    # Save the Pandas Excel writer
    writer.close()
    #*** Fin del proceso **********

if __name__ == '__main__':
    main()