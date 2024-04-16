import os
import gspread
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import openai

import keys

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


SPREADSHEET_ID = keys.SPREADSHEET_ID
OPENAI_API_KEY = keys.OPENAI_API_KEY
openai.api_key = OPENAI_API_KEY


def main():
    credentials = None
    if os.path.exists("token.json"):
        credentials = Credentials.from_authorized_user_file("token.json", SCOPES)
    
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            credentials = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(credentials.to_json())
    
    try:
        service = build("sheets", "v4", credentials=credentials)
        sheets = service.spreadsheets()


        product_a_data = []
        product_b_data = []

        # Reading data for Product A
        for row in range(3, 9):
            product_a_row_data = sheets.values().get(spreadsheetId=SPREADSHEET_ID, range=f"Sheet1!B{row}:F{row}").execute().get("values")
            if product_a_row_data:
                product_a_data.append(product_a_row_data[0])

        # Reading data for Product B
        for row in range(14, 20):
            product_b_row_data = sheets.values().get(spreadsheetId=SPREADSHEET_ID, range=f"Sheet1!B{row}:F{row}").execute().get("values")
            if product_b_row_data:
                product_b_data.append(product_b_row_data[0])
                
        # ----- PROMPT ONE ---------------------------------------------------------
        prompt = "For Product A:\n"
        for row in product_a_data:
            prompt += f"- {row[4]}: {row[2]} units sold, ${row[3]} revenue\n"
        prompt += "\nFor Product B:\n"
        for row in product_b_data:
            prompt += f"- {row[4]}: {row[2]} units sold, ${row[3]} revenue\n"
            
        question_1 = "\nComparison of Sales Performance between Product A and Product B:\n\n"
        prompt_1 = prompt + question_1
        
        response_1 = openai.ChatCompletion.create(
            model='gpt-3.5-turbo',
            messages=[{"role": "user", "content": prompt_1}],
            temperature=0.7,
            max_tokens=100
        )
        
        # Extracting the message content from the response
        response_1_content = response_1.choices[0].message.content.strip()

        # ------------------ PROMPT TWO ------------------------------------------------
        question_2 = "\nSeasonal Sales Trends:\n\n Which months experienced the highest sales quantity and revenue for both Product A and Product B?\n\n\nBased on the provided sales data for both products, please identify the months that experienced the highest sales quantity and revenue for each product."    
        
        prompt_two = prompt + question_2

        response_2 = openai.ChatCompletion.create(
            model='gpt-3.5-turbo',
            messages=[{"role": "user", "content": prompt_two}],
            temperature=0.7,
            max_tokens=100
        )

        # Extracting the message content from the response
        response_2_content = response_2.choices[0].message.content.strip()
        
        
        
        
        

        # Writing responses to the spreadsheet
        values = [
            [response_1_content],
            [response_2_content]
        ]
        body = {
            'values': values
        }
        result = sheets.values().update(spreadsheetId=SPREADSHEET_ID, range='Sheet1!B22', valueInputOption='RAW', body=body).execute()
        print(f'{result.get("updatedCells")} cells updated.')

    except HttpError as error:
        print(error)
        
if __name__ == "__main__":
    main()
