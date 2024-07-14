# Google Sheets Automation

This project automates various operations on Google Sheets, including data manipulation, score calculation, and backup creation.

## Setup

1. Clone this repository:
   ```
   git clone https://github.com/toniilic/Google-sheets-automation.git
   cd Google-sheets-automation
   ```

2. Install required packages:
   ```
   pip install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client pandas openpyxl
   ```

3. Set up Google Sheets API credentials:
   - Go to the [Google Cloud Console](https://console.cloud.google.com/)
   - Create a new project or select an existing one
   - Enable the Google Sheets API for your project
   - Create credentials (OAuth 2.0 Client ID) for a desktop application
   - Download the credentials and save them as 'credentials/credentials.json' in this project directory

4. Run the script:
   ```
   python3 main.py
   ```

5. Follow the authorization prompt in your web browser when you run the script for the first time.

## Usage

When you run the script, it will interactively prompt you for the following information:

- Google Sheet ID
- Sheet name
- Range (e.g., A1:Z)
- Column name for removing duplicates
- Column names for sorting (comma-separated)
- Columns and weights for score calculation

The script will then perform the following operations:

- Read data from the specified Google Sheet
- Remove duplicates based on the specified column
- Sort the data based on the specified columns
- Calculate scores based on the provided weights
- Create a backup of the data
- Write the updated data back to the Google Sheet

## Customization

You can modify the `main.py` script to add or customize operations such as:

- Additional data transformations
- More complex scoring algorithms
- Integration with other APIs or services

## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

## License

[MIT](https://choosealicense.com/licenses/mit/)
