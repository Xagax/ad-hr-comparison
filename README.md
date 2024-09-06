## Python Formatting Script to compare Azure Active Directory file vs HR files.

This Python script is designed to streamline the formatting process. Azure Active Directory file vs HR data file are formatted into Excel spreadsheets.

By executing this script, you can generate organized and structured Excel output from Active Directory and HR data files, making it easier to analyze and interpret users. Whether you're dealing with large datasets or frequent formatting tasks, this script provides a convenient solution to automate the process, saving you time and effort. Simply follow the instructions below to utilize this tool effectively and optimize your data analysis workflow.

A. To install the necessary Python packages, you can use pip, Python's package manager. Open your command line interface and execute the following command:

Optional: first create a virtualized environment using any path for environment you want. Then run source based on your virtual env path and shell.

```
python3 -m venv ~/Sites/pythonenv
source ~/Sites/pythonenv/bin/activate.fish

```

```
pip install pandas streamlit rapidfuzz
```

These libraries provide functionality for data manipulation (Pandas), creating interactive web applications (Streamlit), fuzzy string matching (RapidFuzz), and working with date and time data (datetime). With these packages installed and imported, you can proceed with your Python script execution.

Eensure that the "files" directory is created if it doesn't already exist. This directory will be used to store Active Directory and HR files .

To run the script you must run Streamlit app named mainv3.py, you need to execute the following command in your terminal or command prompt:

```
streamlit run mainv4.py
```

This command will start a local Streamlit server and launch your app in a web browser. You can then interact with the app in the browser window. One Excel file named "AD_vs_HR.xlsx" is automatically generated.
