import os
import pandas as pd
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Set up credentials
creds = Credentials.from_authorized_user_file('credentials/token.json', ['https://www.googleapis.com/auth/spreadsheets'])

# Build the Sheets API service
service = build('sheets', 'v4', credentials=creds)

# Your Google Sheet ID
SPREADSHEET_ID = 'your_spreadsheet_id_here'

def read_sheet(sheet_name, range_name):
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID, range=f'{sheet_name}!{range_name}'
        ).execute()
        return result.get('values', [])
    except HttpError as error:
        print(f"An error occurred: {error}")
        return None

def write_sheet(sheet_name, range_name, values):
    try:
        body = {'values': values}
        result = service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID, range=f'{sheet_name}!{range_name}',
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

def main():
    # 1. Adding & Updating data
    sheet_name = 'Companies'
    range_name = 'A1:Z'  # Adjust based on your data
    data = read_sheet(sheet_name, range_name)
    
    if data:
        df = pd.DataFrame(data[1:], columns=data[0])  # Assuming first row is header
        
        # Remove duplicates
        df = remove_duplicates(df, subset=['Company Name'])  # Adjust based on your unique identifier
        
        # Sort the data
        df = sort_dataframe(df, ['Company Name'])  # Adjust sorting columns as needed
        
        # 2. Calculate scores
        weights = {
            'Revenue': 0.3,
            'Employees': 0.2,
            'Growth Rate': 0.5
        }  # Adjust weights as needed
        df = calculate_scores(df, weights)
        
        # Create a backup
        create_backup(df, 'companies_backup')
        
        # Write updated data back to sheet
        write_sheet(sheet_name, range_name, [df.columns.tolist()] + df.values.tolist())
        
        print("Data processing complete.")
    else:
        print("Failed to retrieve data from the sheet.")

if __name__ == '__main__':
    main()
