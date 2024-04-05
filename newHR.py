#################################################################################
# Python Formatting Script from HR to Excel               			    #
# Creates excel formatted output from excel files                               #
# Usage newHR.py 				        			                            #
# Last modification 28/March/2024                                               #
# Version 1                                                                     #
#################################################################################

import pandas as pd

class HR_creation:
    def Creo_HR(self):
        # Read the CSV and Excel files into DataFrames
        DT = pd.read_csv("files/HR - DT Emails.csv")
        AVERRO = pd.read_excel("files/HR- Averro Email Addresses Active.xlsx")
        NUWEST = pd.read_excel("files/HR - Report for Hailey NuWest emails.xlsx")

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

        AVERRO.rename(columns={'Work_Email': 'email'}, inplace=True)
        AVERRO.rename(columns={'Employee_Name': 'name'}, inplace=True)
        NUWEST.rename(columns={'Employee Name': 'name'}, inplace=True)
        NUWEST.rename(columns={'Work_Email': 'email'}, inplace=True)
        NUWEST.rename(columns={'Employee Status': 'status'}, inplace=True)
        DT.rename(columns={'Work_Email': 'email'}, inplace=True)
        DT.rename(columns={'Employee_Name': 'name'}, inplace=True)
        AVERRO['source'] = 'AVERRO'
        AVERRO['status'] = 'Active'
        DT['source'] = 'DT'
        DT['status'] = 'Active'
        NUWEST['source'] = 'NUWEST'

        # Apply formatear_nombre function to each element of Employee_Name column
        DT['name'] = DT['name'].apply(formatear_nombre)
        # Apply formatear_nombre function to each element of Employee_Name column
        AVERRO['name'] = AVERRO['name'].apply(formatear_nombre)
        # Apply formatear_nombre function to each element of Employee_Name column
        NUWEST['name'] = NUWEST['name'].apply(formatear_nombre)

        # Merge the DataFrames on the 'email' column
        HR = pd.concat([DT, AVERRO, NUWEST], ignore_index=True)

        # Assuming 'name' and 'status' are common columns in all DataFrames,
        HR = HR[['source','name', 'email', 'status', 'Employee_Code']]

        # Save HR to a new Excel file
        HR.to_excel("HR.xlsx", index=False)

        return HR