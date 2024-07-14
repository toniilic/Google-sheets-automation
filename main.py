import os
import pandas as pd
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def get_credentials():
    creds = None
    if os.path.exists('credentials/token.json'):
        creds = Credentials.from_authorized_user_file('credentials/token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists('credentials/credentials.json'):
                print("credentials.json not found. Please follow the setup instructions in the README.md file.")
                return None
            flow = InstalledAppFlow.from_client_secrets_file('credentials/credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('credentials/token.json', 'w') as token:
            token.write(creds.to_json())
    return creds

def read_sheet(service, spreadsheet_id, sheet_name, range_name):
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id, range=f'{sheet_name}!{range_name}'
        ).execute()
        return result.get('values', [])
    except HttpError as error:
        print(f"An error occurred: {error}")
        return None

def write_sheet(service, spreadsheet_id, sheet_name, range_name, values):
    try:
        body = {'values': values}
        result = service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id, range=f'{sheet_name}!{range_name}',
            valueInputOption='USER_ENTERED', body=body
        ).execute()
        print(f"{result.get('updatedCells')} cells updated.")
    except HttpError as error:
        print(f"An error occurred: {error}")

def remove_duplicates(df, subset):
    return df.drop_duplicates(subset=subset, keep='last')

def sort_dataframe(df, columns):
    return df.sort_values(by=columns)

def calculate_scores(df, weights):
    for column, weight in weights.items():
        df[f'{column}_score'] = df[column] * weight
    df['total_score'] = df[[col for col in df.columns if col.endswith('_score')]].sum(axis=1)
    return df

def create_backup(df, filename):
    df.to_csv(f'backups/{filename}.csv', index=False)

def get_user_input():
    spreadsheet_id = input("Enter your Google Sheet ID: ")
    sheet_name = input("Enter the sheet name: ")
    range_name = input("Enter the range (e.g., A1:Z): ")
    unique_identifier = input("Enter the column name for removing duplicates: ")
    sort_columns = input("Enter column names for sorting (comma-separated): ").split(',')
    
    weights = {}
    while True:
        column = input("Enter a column name for score calculation (or press Enter to finish): ")
        if not column:
            break
        weight = float(input(f"Enter the weight for {column}: "))
        weights[column] = weight
    
    return spreadsheet_id, sheet_name, range_name, unique_identifier, sort_columns, weights

def main():
    try:
        creds = get_credentials()
        if not creds:
            return

        service = build('sheets', 'v4', credentials=creds)
        
        spreadsheet_id, sheet_name, range_name, unique_identifier, sort_columns, weights = get_user_input()
        
        data = read_sheet(service, spreadsheet_id, sheet_name, range_name)
        
        if data:
            df = pd.DataFrame(data[1:], columns=data[0])  # Assuming first row is header
            
            df = remove_duplicates(df, subset=[unique_identifier])
            df = sort_dataframe(df, sort_columns)
            df = calculate_scores(df, weights)
            
            create_backup(df, 'data_backup')
            
            write_sheet(service, spreadsheet_id, sheet_name, range_name, [df.columns.tolist()] + df.values.tolist())
            
            print("Data processing complete.")
        else:
            print("Failed to retrieve data from the sheet.")
        
    except Exception as e:
        print(f"An error occurred: {e}")
        print("Please make sure you have set up your Google Sheets API credentials correctly.")
        print("Refer to the README.md file for setup instructions.")

if __name__ == '__main__':
    main()
