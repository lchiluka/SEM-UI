
# STREAMLIT UI Installation Instructions

## Steps to be followed before running the script for UI in Remote Desktop

### Step 1
Before installing anaconda, run this **command** - pip install streamlit

if you are unable to install streamlit, download Anaconda 3 on your Remote Desktop.

### Step 2
Open Anaconda 3 terminal from Anaconda Navigator and install Streamlit.

- **Command:**
  
  pip install streamlit
  
### Step 3
If you encounter errors, update using the following command:

- **Command:**
  
  conda update streamlit
  
### Step 4
Install all required libraries using the PIP command 
For this project, Please install the below library

- **Install Office365-REST-Python-Client**
- **Command:**
  
  pip install Office365-REST-Python-Client
  
### Step 5
For this project, download the "UI_SEM_Final_Code.py" file and paste in `C:\Users\your_folder` 

In general, you can create a `.py` file in `C:\Users\your_folder` and paste your Streamlit code.

### Step 6
For this project, download Carlisle Logo and copy to `C:\Users\your_folder`.

- **Name of the logo:** `Carlisle_MasterLogo_RGB.jpg`

In general, you can use any image and replace it in the "UI_SEM_Final_Code.py" file.
  
### Step 7
Now run the script for the UI. 

- **Example script:** `app.py`
- **Command:**
  streamlit run app.py
  
  For this project -
  **command:**
  streamlit run UI_SEM_Final_Code.py

# Maintenance and Known Issues of Webapp
## In case of Outage/System Restart

1. Re-open the Anaconda 3 Terminal using Anaconda Navigator.
2. Run the command: 
   
   streamlit run UI_SEM_Final_Code.py (For this project)

3. Test the Network URL. The URL might change every time if you re-run the script. Double-check and provide the new URL to the users.

4. Bi-weekly login into RDP and check if the script is running
# Known Issues
1. The Anaconda terminal should remain running in the background continuously to ensure uninterrupted access to the web app
2. If you make changes to the script, please make sure to close the Anaconda terminal, reopen it and run the script again. This would help solve any screen-struck issues.

## URL

- For SEM: `http://10.4.204.140:8501`


# Script Maintenance

1. Changing Sharepoint folder location

   In the script UI_SEM_Final_Code.py, search for folder_in_sharepoint = '/teams/CCMRD7857/RDrive/Analytical/2024%20Projects/Insulation' and replace it with new URL.

