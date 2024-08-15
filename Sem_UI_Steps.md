
# STREAMLIT UI GENERAL STEPS

## Steps to be followed before running the script for UI in Remote Desktop

### Step 1
Create a `.py` file in `C:\Users\your_folder` and paste your combined Streamlit code.

### Step 2
Copy Carlisle Logo to `C:\Users\your_folder`.

- **Name of the logo:** `Carlisle_MasterLogo_RGB.jpg`

### Step 3
Download Anaconda 3 on your PC.

### Step 4
Open Anaconda 3 terminal from Anaconda Navigator and install Streamlit.

- **Command:**
  
  pip install streamlit
  

### Step 5
If you encounter errors, update using the following command:

- **Command:**
  
  conda update streamlit
  

### Step 6
Install all required libraries using the PIP command like below:

- **Install Office365-REST-Python-Client**
- **Command:**
  
  pip install Office365-REST-Python-Client
  

### Step 7
Now run the script for the UI.

- **Example script:** `app.py`
- **Command:**
  
  streamlit run app.py
  

## In case of Outage/System Restart

1. Re-open the Anaconda 3 Terminal using Anaconda Navigator.
2. Run the command: 
   
   streamlit run app.py
   
   (example)

3. Test the Network URL. The URL might change every time if you re-run the script. Double-check and provide the new URL to the users.

- **Note:** Bi-weekly login into RDP and check if the script is running.

## URLs

- For SEM: `http://10.4.204.140:8501`
- For RMQ: `http://10.4.204.140:8502`

## Tips

If you make changes to the script, instead of refreshing the URL in the browser, please make sure to restart the Anaconda terminal and run the script again. This would help solve any screen-struck issues.

