import streamlit as st
import pandas as pd
import msal
import requests
import io
import dotenv
import os
dotenv.load_dotenv()
CLIENT_ID = os.getenv('CLIENT_ID')
TENANT_ID = os.getenv('TENANT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
st.set_page_config(layout="wide")

# MSAL Configuration
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"]

def get_access_token():
    app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )
    result = app.acquire_token_for_client(scopes=SCOPE)
    if "access_token" in result:
        return result["access_token"]
    else:
        raise Exception("Could not obtain access token")

def get_site_id(access_token, hostname, site_relative_path):
    headers = {"Authorization": f"Bearer {access_token}"}
    url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:{site_relative_path}"
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        return response.json()["id"]
    else:
        st.error(f"Error fetching site ID: {response.status_code}")
        st.error(response.text)
        return None

def get_drive_id(access_token, site_id):
    headers = {"Authorization": f"Bearer {access_token}"}
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        drives = response.json()["value"]
        if drives:
            return drives[0]["id"]
    st.error(f"Error fetching drive ID: {response.status_code}")
    st.error(response.text)
    return None

def list_files(access_token, drive_id, folder_path):
    headers = {"Authorization": f"Bearer {access_token}"}
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{folder_path}:/children"
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        return response.json()
    else:
        st.error(f"Error fetching files: {response.status_code}")
        st.error(response.text)
        return {}

def download_file(access_token, drive_id, file_id):
    headers = {"Authorization": f"Bearer {access_token}"}
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}/content"
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        return response.content
    else:
        st.error(f"Error downloading file: {response.status_code}")
        st.error(response.text)
        return None

def extract_summary_from_file(file_content):
    try:
        df = pd.read_excel(io.BytesIO(file_content), sheet_name='Summary')
    except ValueError:
        return None

    # Remove any unnamed columns and set the first row as header
    df.columns = df.iloc[0]
    df = df[1:]

    # Reset the index to 'Month of Nivedan'
    df.set_index('Month of Nivedan', inplace=True)
    return df

def main():
    st.title("OneDrive Data Visualization")
    # Get access token using client credentials
    access_token = get_access_token()

    # Define SharePoint details
    hostname = "m365x50834976.sharepoint.com"  # Replace with your SharePoint domain
    site_relative_path = "/sites/100thYear"  # Replace with your site relative path

    # Get Site ID
    site_id = get_site_id(access_token, hostname, site_relative_path)
    if not site_id:
        st.stop()

    # Get Drive ID
    drive_id = get_drive_id(access_token, site_id)
    if not drive_id:
        st.stop()

    # User Role Selection
    user_role = st.sidebar.selectbox("Select Your Role", ["Admin", "User"])

    # OneDrive Folder Path
    folder_path = "data"  # Path within OneDrive
    
    # List Files
    files = list_files(access_token, drive_id, folder_path)

    summary_data = {}
    if 'value' in files:
        file_options = []
        for file in files['value']:
            if file['name'].endswith(".xlsx") and 'summary' not in file['name'].lower():
                file_options.append((file['name'], file['id']))

        for file_name, file_id in file_options:
            file_content = download_file(access_token, drive_id, file_id)
            if file_content:
                summary = extract_summary_from_file(file_content)
                if summary is not None:
                    summary_data[file_name.replace('.xlsx', '')] = summary

        if summary_data:
            combined_summary = pd.concat(summary_data, axis=1)
            combined_summary.columns = combined_summary.columns.map('_'.join)
            
            # Create a multi-index DataFrame for better visualization
            columns = pd.MultiIndex.from_tuples([(col.split('_')[0], col.split('_')[1]) for col in combined_summary.columns])
            combined_summary.columns = columns

            # Generate HTML table with merged headers
            html = '<table border="1" style="width:100%; border-collapse: collapse;">'
            html += '<tr><th rowspan="2">Month of Nivedan</th>'
            for file_name in summary_data.keys():
                col_span = len(summary_data[file_name].columns)
                html += f'<th colspan="{col_span}">{file_name}</th>'
            html += '</tr><tr>'
            for file_name in summary_data.keys():
                for col in summary_data[file_name].columns:
                    html += f'<th>{col}</th>'
            html += '</tr>'
            
            for idx, row in combined_summary.iterrows():
                html += f'<tr><td>{idx}</td>'
                for val in row:
                    html += f'<td>{val}</td>'
                html += '</tr>'
            html += '</table>'

            st.markdown(html, unsafe_allow_html=True)
        else:
            st.write("No summary data found.")
    else:
        st.error("No files found or error in fetching files.")

if __name__ == "__main__":
    main()
