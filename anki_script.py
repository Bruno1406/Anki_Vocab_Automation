import requests
import json
import re
from bs4 import BeautifulSoup
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SHEET_ID = "your sheet id here"
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = "key.json"
ANKI_DECK_NAME = "German Vocabulary"
ANKI_MODEL_NAME = "German"

def check_if_is_one_word(word):
    return len(word.strip().split()) == 1

def fetch_new_sheet_entries(sheet, sheet_id):
    result = sheet.values().get(spreadsheetId=sheet_id,
                                range='A1:B').execute()
    rows = result.get('values', [])

    if not rows:
        print("No data found in the spreadsheet.")
        return

    new_entries = {}
    for index, row in enumerate(rows):
        row_num = index + 1 
        
        # Ensure the row actually has a word in Column A
        if len(row) > 0 and row[0].strip():
            word = row[0]
            
            # Safely check Column B. If the row is short, status is empty.
            status = row[1] if len(row) > 1 else ""

            if status == "" or status.lower() == "network error":
                new_entries[f'A{row_num}'] = word
    print(new_entries)
    return new_entries
    
def update_sheet_cells(sheet, sheet_id, updates):
    body = {
        'valueInputOption': 'RAW',
        'data': [{'range': cell, 'values': value} for cell, value in updates.items()]
    }
    result = sheet.values().batchUpdate(spreadsheetId=sheet_id, body=body).execute()
    return result

def delete_sheet_rows(sheet, sheet_id, rows_to_delete):
    requests_body = {
        "requests": [
            {
                "deleteDimension": {
                    "range": {
                        "sheetId": 0,
                        "dimension": "ROWS",
                        "startIndex": row - 1,
                        "endIndex": row
                    }
                }
            } for row in sorted(rows_to_delete, reverse=True)
        ]
    }
    response = sheet.batchUpdate(spreadsheetId=sheet_id, body=requests_body).execute()
    return response

def fetch_word_data(word):
    try:
        # Get the webpage
        url = f'https://www.verbformen.de/?w={word}'
        response = requests.get(url)
        if response.status_code != 200:
            print(f'Failed to fetch data for {word}: Status code {response.status_code}')
            return None, None, response.status_code
        soup = BeautifulSoup(response.content, 'html.parser')

        # Adverb word special case
        if soup.find('span', title= "Adverb"):
            base_word = soup.find('div', class_ = "rU6px rO0px")
            word_flexions = ""
            word_translation = soup.find('p', class_="r1Zeile rU6px rO0px")
            example_sentences = soup.find('p', class_ = "rNt rInf wKnFmt r1Zeile rU3px rO0px")

            if not base_word or not word_translation or not example_sentences:
                print(f'Incomplete data for adverb {word}')
                print(f'Base: {base_word}, Translation: {word_translation}, Examples: {example_sentences}')
                return None, None, response.status_code
            
            word_fields = {
                'Basis': base_word.text.strip(),
                'Flexionsformen': word_flexions,
                'Übersetzungen': word_translation.text.strip(),
                'Beispielsätze': re.sub(r'([.?!]).*', r'\1', example_sentences.text.replace('\n', ' ')).strip(),
                }
            word_tag = 'Adverb'
            return word_fields, word_tag, response.status_code

        base_word = soup.find('span', class_='vGrnd')
        if not base_word :
            print(f'Word {word} not found on the webpage.')
            return None, None, response.status_code
        
        word_flexions = soup.find('p', class_='vStm rCntr')
        word_translation = soup.find('p', class_='r1Zeile rU3px rO0px')
        example_sentences = soup.find('ul', class_='rLst rLstGt').find_all('li')
        word_info = soup.find('p', class_='rInf')

        if not base_word or not word_flexions or not word_translation or not example_sentences or not word_info: 
            print(f'Incomplete data for {word}')
            return None, None, response.status_code

        word_fields = {
            'Basis': base_word.text.strip(),
            'Flexionsformen': word_flexions.text.strip().replace('\n', '').replace('·', ' · '),
            'Übersetzungen': word_translation.text.strip(),
            'Beispielsätze': "<br>".join([re.sub(r'([.?!]).*', r'\1', sentence.text.replace('\n', ' ')).strip() for sentence in example_sentences]),
            }
        word_tag = ''
        info_lower = word_info.text.lower()
        if 'substantiv' in info_lower:
            word_tag = 'Substantiv'
        elif 'adjektiv' in info_lower:
            word_tag = 'Adjektiv'
        elif 'unregelmäßig' in info_lower:
            word_tag = 'Unregelmäßiges Verb'
        elif 'regelmäßig' in info_lower:
            word_tag = 'Regelmäßiges Verb'
        
        return word_fields, word_tag, response.status_code
    except Exception as e:
        print(f'Network error fetching data for {word}: {e}')
        return None, None, -1
    
def add_anki_note(deck_name, model_name, word_fields, word_tag, url='http://localhost:8765'):
    request_payload = {
        "action": "addNote",
        "version": 6,
        "params": {
            "note": {
                "deckName": deck_name,
                "modelName": model_name,
                "fields": word_fields,
                "tags": [
                    word_tag,
                ]
            }
        }
    }
    try:
        response = requests.post(url, json=request_payload)
        result = response.json()
        if 'error' in result and result['error'] is not None:
            print(f"AnkiConnect error: {result['error']} for word {word_fields['Basis']}")
        return result
    except Exception as e:
        print(e)
        return None


if __name__ == '__main__':
    credentials = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=credentials)
    sheet = service.spreadsheets()
    new_entries = fetch_new_sheet_entries(sheet, SHEET_ID)
    if new_entries:
        updates = {}
        for cell, word in new_entries.items():

            if not check_if_is_one_word(word):
                updates[cell.replace('A', 'B')] = [["expression"]]
                continue

            word_fields, word_tag, status_code = fetch_word_data(word)

            if status_code != 200:
                updates[cell.replace('A', 'B')] = [["network error"]]
                if status_code == 429:
                    print("Rate limited by the website. Stopping further requests. Maximum is 100 words at once. Try again later.")
                    break

            elif word_fields:
                anki_result = add_anki_note(ANKI_DECK_NAME, ANKI_MODEL_NAME, word_fields, word_tag)
                if anki_result:
                    if anki_result['error'] is None:
                        updates[cell] = [[word_fields['Basis'], "sync"]]
                    else:
                        updates[cell.replace('A', 'B')] = [["Anki Connect error: " + anki_result['error']]]
            else:
                updates[cell.replace('A', 'B')] = [["word not found"]]
        update_sheet_cells(sheet, SHEET_ID, updates)


