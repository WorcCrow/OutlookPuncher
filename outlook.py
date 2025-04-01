import requests
import json
import re
import os
import time
import datetime
import webbrowser  # To open the URL in the browser
from tqdm import tqdm

# File paths
config_file = "config.json"
#{
#    "tenant_id": "********-9852-****-9757-4ac3063edb3f",
#    "client_id": "********-****-466a-8a75-50bef695c9a8",
#    "client_secret": ".****~********-nhPMiVVBVBueeIeooV_A6KdeT",
#}
file_path = "outlook.log"

data = []

def load_config():
    try:
        with open(config_file, 'r') as file:
            config = json.load(file)
        return config
    except Exception as e:
        print(f"Error loading config: {e}")
        return None

def save_config(config):
    try:
        with open(config_file, 'w') as file:
            json.dump(config, file, indent=4)
        print("Config updated successfully!")
    except Exception as e:
        print(f"save_config(): Error saving config: {e}")

def addRecord(link):
    try:
        data.append(link)
        writeRecord(link)
    except Exception as e:
        print(f"addRecord(): Error: addRecord - {e}")

def writeRecord(content, mode='a'):
    try:
        with open(file_path, mode) as file:
            file.write(content + '\n')
    except Exception as e:
        print(f"writeRecord(): Error: writeRecord - {e}")

def loadData():
    try:
        with open(file_path, 'r') as file:
            for line in file:
                data.append(line.strip())
    except FileNotFoundError:
        print("loadData(): No existing log file found. Creating a new one.")
    except Exception as e:
        print(f"loadData(): Error: loadData - {e}")

def printData():
    try:
        print("Current Record: ")
        for line in data:
            print(line)
    except Exception as e:
        print(f"printData(): Error: printData - {e}")

# Refresh OAuth2 token
def refreshToken():
    try:
        config = load_config()
        if not config:
            return ""

        url = f"https://login.microsoftonline.com/{config['tenant_id']}/oauth2/v2.0/token"
        
        payload = {
            'client_id': config['client_id'],
            'client_secret': config['client_secret'],
            'refresh_token': config['refresh_token'],
            'grant_type': 'refresh_token',
            'scope': 'offline_access Mail.ReadWrite Mail.Send'
        }

        headers = {
            'Content-Type': 'application/x-www-form-urlencoded'
        }

        response = requests.post(url, headers=headers, data=payload)

        if response.status_code == 200:
            token = response.json()
            return token
        else:
            print(f"refreshToken(): Error: {response.status_code}")
            print(response.text)
            get_refresh_token()
            return None

    except Exception as e:
        print(f"refreshToken(): Error: refreshToken - {e}")
        get_refresh_token()
        return ""

# Check for punchline in the latest email
def checkPunchline():
    token = ""
    try:
        token_data = refreshToken()
        if not token_data or "access_token" not in token_data:
            print("Failed to get access token.")
            return

        token = token_data["access_token"]

        url = "https://graph.microsoft.com/v1.0/me/messages?$search=\"from:attendancetracking@salary.com\"&top=1"
        headers = {
            'Authorization': f'Bearer {token}'
        }

        response = requests.get(url, headers=headers)

        if response.status_code != 200:
            print(f"Failed to retrieve messages: {response.status_code}")
            print(response.text)
            return

        responseJson = response.json()
        content = responseJson["value"][0]["body"]["content"]

        punchlink = re.findall('href="([^"]+)"', content)
        link = punchlink[0].replace("&amp;", "&")

        if link in data:
            print("Time Check:", datetime.datetime.now())
        else:
            print("\nGoing to Punch ", link)
            addRecord(link)
            os.system("start chrome \"" + link + "\"")

    except Exception as e:
        print("checkPunchline(): Time of Error:", datetime.datetime.now())
        print(f"checkPunchline(): Error: checkPunchline - {e}")

def open_authorization_url(tenant_id, client_id):
    auth_url = (
        f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/authorize"
        f"?client_id={client_id}"
        f"&response_type=code"
        f"&redirect_uri=https%3A%2F%2Flocalhost"
        f"&response_mode=query"
        f"&scope=offline_access%20Mail.ReadWrite%20Mail.Send"
        f"&state=12345644"
    )
    print(f"Opening authorization URL: {auth_url}")
    webbrowser.open(auth_url)
    print("\nPlease log in and authorize the app. Copy the `code` from the URL and update `config.json`.")

# Get new refresh token and update config.json
def get_refresh_token():
    config = load_config()

    if not config.get("code"):
        print("No authorization code found. Opening authorization URL...")
        open_authorization_url(config['tenant_id'], config['client_id'])
        return None

    url = f"https://login.microsoftonline.com/{config['tenant_id']}/oauth2/v2.0/token"
    
    payload = {
        "client_id": config['client_id'],
        "client_secret": config['client_secret'],
        "code": config['code'],
        "redirect_uri": "https://localhost",
        "grant_type": "authorization_code",
        "scope": "offline_access Mail.ReadWrite Mail.Send"
    }

    try:
      
        print("Getting new refresh token")
        response = requests.post(url, data=payload)

        if response.status_code == 200:
            token_data = response.json()
            config["refresh_token"] = token_data.get("refresh_token")
            save_config(config)
            print("New refresh token saved!")
            return token_data.get("refresh_token")

        else:
            error_response = response.json()
            
            # Handle "code already redeemed" error
            if error_response.get("error") == "invalid_grant":
                print("\nAuthorization code already redeemed. Opening new authorization URL...")
                open_authorization_url(config['tenant_id'], config['client_id'])
                return None

            print(f"get_refresh_token(): Error: {response.status_code} - {error_response}")
            return None

    except Exception as e:
        print(f"get_refresh_token(): Exception: {e}")
        return None
        
# Main Execution
loadData()
printData()

if len(data) != 0:
    writeRecord(data[-1], "w")
    data = [data[-1]]

while True:
    checkPunchline()
    
    # Wait for 5 minutes before checking again
    for i in tqdm(range(300), desc="Loading..."):
        time.sleep(1)
