import boto3
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import pandas as pd

# check all ssm parameters in SSM Parameter Store and list them

def check_ssm_parameter(ssm_client, parameter_name, parameter_value, before_dict, current_dict):
    try:
        # check all parameters in SSM Parameter Store and list them
        response = ssm_client.describe_parameters()
        parameters = response['Parameters']
        
        for parameter in parameters:
            if parameter['Name'] in parameter_name:
                print(f"Parameter {parameter_name} already exists.")
                response = ssm_client.get_parameter(
                    Name=parameter_name, WithDecryption=True
                )
                
                # list up parameters key, and value before updating
                before_dict[parameter_name] = response['Parameter']['Value']
                
                # retrieve parameter value for parameter_names
                for i in range(len(parameter_name)):
                    if parameter_name[i] == parameter['Name']:
                        parameter['Value'] = parameter_value[i]
                        value = parameter['Value']
                
                # update the parameter value
                ssm_client.put_parameter(
                    Name=parameter['Name'],
                    Value=value,
                    Type='String',
                    Overwrite=True
                )
                
                # list up parameters key, and value after updating
                current_dict[parameter['Name']] = value

            else:
                # if parameter not in before_dict:
                before_dict[parameter['Name']] = "None"
                
                ssm_client.put_parameter(
                    Name=parameter['Name'],
                    Value=parameter['Value'],
                    Type='String',
                    Overwrite=True
                )
                # list up parameters key, and value after updating
                current_dict[parameter['Name']] = parameter['Value']

        return before_dict, current_dict
    except Exception as e:
        print(f"Error checking SSM parameter: {e}")
        return before_dict, current_dict

# put data from csv data into SSM Parameter Store
def put_ssm_parameter(df, ssm_client):
    before_dict = {}
    current_dict = {}
    
    ssm_parameter_keys = df[0]
    ssm_parameter_values = df[1]
    #for index, row in df.iterrows():
    #    ssm_parameter_name = row[0]
    #    ssm_parameter_value = row[1]
        
    # check if each parameter already exists and if not store the value in dict
    before_dict, current_dict = check_ssm_parameter(ssm_client, ssm_parameter_keys, ssm_parameter_values, before_dict, current_dict)
        
    # update the csv with the current parameter value, and the before value
    # df.at[index, 'Before'] = before_dict.get(ssm_parameter_name, "None")
    # df.at[index, 'Current'] = current_dict.get(ssm_parameter_name, "None")
    # save the updated DataFrame to a new CSV file
    # df.to_csv("updated_ssm_parameters.csv", index=False)

if __name__ == "__main__":
    # Authorizations
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    credentials = Credentials.from_service_account_file(
        r"XXXXX", scopes=scope
    )
    # Authorize the gspread client
    gc = gspread.authorize(credentials)
    service = build('drive', 'v3', credentials=credentials)

    workbook = gc.open("XXXX")
    sheet_name = "SSM_parameter"
    sheet = workbook.worksheet(sheet_name)
    data = sheet.get_all_values()
    df = pd.DataFrame(data)

    ssm_client = boto3.client('ssm', region_name='ap-southeast-2')
    put_ssm_parameter(df, ssm_client)
