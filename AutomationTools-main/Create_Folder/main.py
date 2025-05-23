import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pandas as pd, time
import upload_gdrive, upload_local
from common_tool import google_authorize, send_mail, check_sheet_exists, listup_all_files
import boto3
import logging


def main(sheet_name):
    sheet = workbook.worksheet(sheet_name)
    data = sheet.get_all_values()
    df = pd.DataFrame(data)
    extension = ".mp4"

    for index, row in df.iterrows():
        if not row[1] == "BasePath":       
            chara_name = row[0].replace(" ", "")
            base_path = row[1]
            folder_name = row[2]
            
            if not "XXX" in chara_name and not "XXX in chara_name:
                # case 1: local files
                destination_folder = upload_local.folder_path_create(sheet_name, chara_name, base_path) # define destination folder path
                
                upload_local.create_folder(destination_folder)
                result = upload_local.move_to_folder(destination_folder, chara_name, extension)
    if result == "Success":
        msg = f"Sheet name({sheet_name}): 処理完了"
    else:
        msg = f"Sheet name({sheet_name}): 処理失敗"
    return msg, base_path


if __name__ == "__main__":
    flg_filepath = input("Do you want to list file path? (y/n): ")
    # google Authorizations
    gc = google_authorize()
    workbook = gc.open("chara_name_list")
    
    game_sheet_name_list = []
    game_sheets = workbook.worksheets()
    
    # make list with all sheet names
    for game_sheet in game_sheets:
        game_sheet_name_list.append(game_sheet.title)
    print(game_sheet_name_list)
    
    msg_list = []
    folder_list = []

    for game_sheet in game_sheet_name_list:
        list = main(game_sheet)
        
        logging.info(f"Processing {game_sheet} is complete.")
        print(list[0])
        
        time.sleep(3)
        
        msg_list.append(list[0])
        folder_list.append(list[1]) # make the base paths list in google spreadsheet
            
    message = str(msg_list).encode('utf-8').decode('utf-8')
    print(message)
    
    # save files path in google spreadsheet
    file_workbook = gc.open("file_list")
    i = 0
    
    if flg_filepath == "y":
        for game_sheet_name in game_sheet_name_list:
            print(f"Game_sheet_name: {game_sheet_name}")
            
            try:
                check_sheet_exists(game_sheet_name, file_workbook)
                game_sheet = file_workbook.worksheet(game_sheet_name)
            except Exception as e:
                print(f"Error: {e}")
                logging.error(f"Error: {e}")
            
            print(folder_list[i])
            listup_all_files(folder_list[i], game_sheet)
            i = i + 1
                
    try:
        ssm_client = boto3.client('ssm', region_name='ap-southeast-2')
        #send_mail(ssm_client, msg_list) # if not necessary, comment out this line
    except Exception as e:
        print(f"Error sending email: {e}")
        
    print("--------------------------------")
    print("All operation completed")
    print("--------------------------------")
    
    logging.info("All processes are complete.")
    
    
