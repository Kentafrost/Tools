import os, re
import logging, gspread
import pickle

from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.oauth2 import service_account
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.service_account import Credentials

gmail_cred = "XXXXX"
gsheet_cred = "XXXXX"

def authorize_gmail():
    try:
        scope = [
            "https://mail.google.com/",
            "https://www.googleapis.com/auth/gmail.modify",
            "https://www.googleapis.com/auth/gmail.readonly",
        ]
       
        creds = None
        token_file = "G:\\My Drive\\Tool\\Retrieve_gmail\\token.pickle"

        # Load existing credentials if available
        if os.path.exists(token_file):
            with open(token_file, "rb") as token:
                creds = pickle.load(token)

        # If no valid credentials, authenticate and save them
        if not creds or not creds.valid:
            # The file token.json stores the user's access and refresh tokens, and is        
            flow = InstalledAppFlow.from_client_secrets_file(gmail_cred, scope)
            creds = flow.run_local_server(port=0)
            
            # Save credentials for future use
            with open(token_file, "wb") as token:
                pickle.dump(creds, token)
        
        gmail_service = build('gmail', 'v1', credentials=creds)
        return gmail_service
    except Exception as e:
        logging.error('Googleの認証処理でエラーが発生しました。{}'.format(e))
        return None    
    
def authorize_gsheet():
    # Set up authentication
    try:
        scope = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        
        credentials = Credentials.from_service_account_file(
            gsheet_cred, scopes=scope
        )
        
        # Authorize the gspread client
        gspread_client = gspread.authorize(credentials)
        return gspread_client
    except Exception as e:
        logging.error('Googleの認証処理でエラーが発生しました。{}'.format(e))
        return None

def main():
    gmail_service = authorize_gmail()
    messages = []
    # Fetch messages
    # results = gmail_service.users().messages().list(userId='me', max_results=10000).execute()
    # messages = results.get('messages', [])

    page_token = None
    while len(messages) < 2000:  # Set your limit
        response = gmail_service.users().messages().list(
            userId='me',
            maxResults=2000,  # Adjust batch size
            pageToken=page_token
        ).execute()

        if 'messages' in response:
            messages.extend(response.get('messages', []))

        page_token = response.get('nextPageToken')
        if not page_token:
            break  # No more pages

    rakuten_cost_list = []
    rakuten_date_list = []
    coursera_list = []
    coursera_enrolled_list = []

    for message in messages:
        msg = gmail_service.users().messages().get(
            userId='me', 
            id=message['id']
        ).execute()
        
        # msg = gmail_service.users().messages().list(userId='me', id=message['id']).execute()
        if "XXカード" and "XXXXX" in msg['snippet']:
            
            # cost
            pattern = r'本人 *([\d,]+) 円'
            matches = re.findall(pattern, msg['snippet'])
            if matches:
                cost = int(matches[0].replace(',', ''))
                rakuten_cost_list.append(cost)
            
            # 利用金額～本人までの数字を抽出
            # 例: "ご利用金額 2025/05/04 本人"から、2025/05/04を抽出
            
            # date
            pattern = r'ご利用金額 (\d{4}/\d{2}/\d{2}) 本人'
            matches = re.findall(pattern, msg['snippet'])
            if matches:
                rakuten_date_list.append(matches[0])
            
        # Coursera certificate
        elif "Coursera" and "Congratulations! Your Certificate is Ready"in msg['snippet']:
            # print(msg['snippet'])
            coursera_list.append(msg['snippet'])
            
        elif "Coursera" and "Welcome to the Specialization!" or "We’re thrilled you enrolled" in msg['snippet']:
            pattern = r'you enrolled in (.+?) on Coursera'
            matches = re.findall(pattern, msg['snippet'])
            if matches:
                coursera_enrolled_list.append(matches[0])
                
    print( f"Total cost I spent in Rakuten Card: {sum(rakuten_cost_list)}円")
    print( f"Numbers of course completed in Coursera: {len(coursera_list)}")
    print( f"Numbers of course enrolled in Coursera: {len(coursera_enrolled_list)}")
    
    # top to down
    rakuten_cost_list.append(sum(rakuten_cost_list))
    rakuten_date_list.append("Total")
    
    rakuten_cost_list = sorted(rakuten_cost_list, reverse=True)
    rakuten_date_list = sorted(rakuten_date_list, reverse=True)
    
    gc = authorize_gsheet()
    workbook = gc.open("gmail_summary")
    sheet_name = "Summary"
    sheet = workbook.worksheet(sheet_name)
    
        # write all date, cost into gsheet
    if rakuten_cost_list:
        # Prepare data for batch update
        data = [['Date', 'Card Cost']]  # Header row
        data.extend(zip(rakuten_date_list, rakuten_cost_list))  # Add rows of data

        # Update the sheet in one call
        sheet.update(values=data, range_name='A1')

        # Print the data for confirmation
        for date, cost in zip(rakuten_date_list, rakuten_cost_list):
            print(date, cost)
    
    # get most recent data from rakuten_date_list
    if rakuten_date_list:
        most_recent_date = rakuten_date_list[1]
        min_recent_date = min(rakuten_date_list)
        print(f"Most recent date of Rakuten Card usage: {most_recent_date}")
        print(f"Most old date of Rakuten Card usage: {min_recent_date}")
    else:
        print("No Rakuten Card usage data found.")        
if __name__ == '__main__':
    main()
