import requests
from bs4 import BeautifulSoup
from common_tool import google_authorize

def web_scrape(url):
    # ウェブサイトの内容を取得
    try:
        response = requests.get(url)
    except requests.exceptions.RequestException as e:
        print(f"Error: Unable to access the website. {e}")

    if response.status_code != 200:
        print(f"Error: Unable to access the website. Status code: {response.status_code}")
    else:
        # BeautifulSoupを使ってHTMLを解析
        soup = BeautifulSoup(response.text, 'html.parser')

        # テーブルのIDを指定し、キャラ名を取得する(テーブルIDをsheet_charaにまとめておく)
        table = soup.find('table', {'id': 'sortabletable1'})  
        name = table.find_all('a')
        data = []
        
        # retrieve only text between <a> tags from table
        for ele in name:
            text = ele.text
            if text != "編集" and text != "":
                data.append(text)
                
    return data

# 取得したいウェブサイトのURL
url_list = [
    'https://XXXXX/?XXX',
    'https://XXXXX/?XXX', 
    'https://XXXXX/?XXX'
    ]

data_list = []

for url in url_list:
    data = web_scrape(url)
    data_list.append(data)

unwanted_items = "追加"

if unwanted_items in data_list:
    data_list.remove(unwanted_items)
    print(f"Removed unwanted item: {unwanted_items}")

print(data_list)

# google authorization
gc = google_authorize()
sheet_chara = gc.open("chara_name_list")
sheet_url = gc.open("url_tbl_id_list")
sheet_list = ["XXXX"]
    
for required_sht_name in sheet_list:
    worksheet = sheet_chara.worksheet(required_sht_name)
    
    worksheet.clear()  # Clear the worksheet before writing new data
    
    count = 2
    worksheet.update("A1", [["キャラクター名"]])
    worksheet.update("B1", [["フォルダーパス"]])
    
    # several lists in a list
    formatted_data = [[item] for sublist in data_list for item in sublist]  
    existing_values = worksheet.col_values(1)  # Get all values from Column 
    
    # Filter out duplicates
    if (formatted_data) not in existing_values:
        worksheet.append_rows(formatted_data)
        print("New data added successfully!")
    else:
        print("No new data to add (all values already exist).")
    
    list_fold_path = [["D:\\XXXXX\\XXXXX\\XXXXX"] for _ in range(len(formatted_data))]
    
    worksheet.update(f"B2:B{len(formatted_data) + 1}", list_fold_path)

