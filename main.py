import os
import pickle
import time
import base64
import json
import requests
from dotenv import load_dotenv
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
load_dotenv()


access_token = os.getenv('ZOHO_ACCESS_TOKEN')
organization_id = os.getenv('ORG_ID')
vendor_id = os.getenv('VENDOR_ID')
url = os.getenv('URL')

# If modifying these SCOPES, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def get_credentials():
    try:
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    
    return creds

def get_email_subject_and_body(service, message_id):
    message = service.users().messages().get(userId='me', id=message_id, format='full').execute()
    headers = message['payload']['headers']

    subject = None
    for header in headers:
        if header['name'] == 'subject' or header['name'] == 'Subject':
            subject = header['value']
            break

    body = None
    if 'parts' in message['payload']:
        parts = message['payload']['parts']
        for part in parts:
            if part['mimeType'] == 'text/plain':
                body = part['body']['data']
                break
    elif 'body' in message['payload']:
        body = message['payload']['body']['data']

    if body:
        body = body.replace('-', '+').replace('_', '/')
        body = base64.b64decode(body).decode('utf-8')

    return subject, body

def add_to_mongo(items, subject):
    data = {
    'items': items,
    'subject': subject
    }

    
    payload = json.dumps(data)
    headers = {
    'Content-Type': 'application/json'
    }

    response = requests.request("POST", url, headers=headers, data=payload)
    print(response.text)



def get_item_id(item_name):
    base_url = "https://books.zoho.com/api/v3/items"
    headers = {
        "Authorization": f"Zoho-oauthtoken {access_token}",
    }

    # Send a request to get a list of items
    response = requests.get(
        base_url,
        headers=headers,
        params={"organization_id": organization_id}
    )

    if response.status_code == 200:
        items = response.json()["items"]
        # Find the item with the given name
        for item in items:
            if item["name"] == item_name:
                return item["item_id"]

        print(f"Item not found: {item_name}")
        return None
    else:
        print(f"Failed to get items: {response.text}")
        return None

def create_item(item_name, item_description, item_rate):
    print(item_name, item_description, item_rate)
    base_url = "https://books.zoho.com/api/v3/items"
    headers = {
        "Authorization": f"Zoho-oauthtoken {access_token}",
        "Content-Type": "application/json;charset=UTF-8",
    }

    # Prepare the item data
    item_data = {
        "name": item_name,
        "description": item_description,
        "rate": item_rate,
        "account_id": "3450124000000000388",
        "tax_id": "",
        "tags": [],
        "custom_fields": [],
        "purchase_rate": item_rate,
        "purchase_account_id": "3450124000000034003",
        "item_type": "sales_and_purchases",
        "product_type": "goods"
    }

    # Send a request to create a new item
    response = requests.post(
        base_url,
        headers=headers,
        params={"organization_id": organization_id},
        data=json.dumps(item_data)
    )


    if response.status_code == 201:
        return response.json()
    else:
        raise Exception(f"Failed to create item: {response.text}")


def create_purchase_order(items):
    base_url = "https://books.zoho.com/api/v3/purchaseorders"
    headers = {
        "Authorization": f"Zoho-oauthtoken {access_token}",
        "Content-Type": "application/json;charset=UTF-8",
    }

    # Prepare the purchase order data
    purchase_order_data = {
        "vendor_id": vendor_id,
        "line_items": items
    }

    # Send a request to create a purchase order
    response = requests.post(
        base_url,
        headers=headers,
        params={"organization_id": organization_id},
        data=json.dumps(purchase_order_data)
    )

    # Check if the request was successful
    if response.status_code == 201:
        return response.json()
    else:
        raise Exception(f"Failed to create purchase order: {response.text}")




def get_items_from_the_email(text):

    items = []
    item_count = 0
    for line in text.split('\n'):
        if line.startswith('Item '):
            item_count += 1
            item = {}
            items.append(item)
            item['item_number'] = line.split(': ')[0]
            item['name'] = line.split(': ')[1]
        elif line.startswith('Quantity'):
            item['quantity'] = line.split(': ')[1]
        elif line.startswith('Description'):
            item['description'] = line.split(': ')[1]
        elif line.startswith('Price per item'):
            item['price'] = line.split(': ')[1]
        elif line.startswith('Total price') or line.startswith('Total purchase price'):
            item['total_price'] = line.split(': ')[1]

    return items

def watch_inbox():
    try:
        creds = get_credentials()
        service = build('gmail', 'v1', credentials=creds)
        last_email_id = None

        while True:
            results = service.users().messages().list(userId='me', labelIds=['INBOX'], maxResults=1).execute()
            messages = results.get('messages', [])

            if messages:
                msg_id = messages[0]['id']
                if msg_id != last_email_id:
                    last_email_id = msg_id
                    print(f"New email received. Email ID: {msg_id}")

                    subject, body = get_email_subject_and_body(service, msg_id)
                    if "purchase order" in subject.lower():
                        print(f"Subject: {subject}")
                        # print(f"Body: {body}")
                        items = get_items_from_the_email(body)

                        print(items)
                        add_to_mongo(items, subject)

                        
                        try:
                            processed_item = []
                            for item in items:
                                print(item)
                                item_name = item['name'].replace('\r', '')
                                item_description = item['description'].replace('\r', '')
                                item_rate = float(item['price'].replace('\r', '').replace('$',''))
                                quantity = int(item['quantity'].replace('bags\r', '').replace('\r', '').replace(',', ''))
                                item_id = get_item_id(item_name)
                                if item_id is None:
                                    print(f"Item not found. Creating a new item: {item_name}")
                                    new_item = create_item(item_name, item_description, item_rate)
                                    item_id = new_item["item"]["item_id"]

                                print(f"Item ID for '{item_name}': {item_id}")
                                processed_item.append({
                                    "item_id":item_id,
                                    "quantity":quantity,
                                    "rate": item_rate
                                })
                            purchase_order = create_purchase_order(processed_item)
                            print(f"Purchase order created successfully: {purchase_order}")
                        except Exception as e:
                            print(e)


                    else:
                        print("Email subject does not contain 'Purchase Order'.")

            time.sleep(5)
    except HttpError as error:
        print(f"An error occurred: {error}")

if __name__ == '__main__':
    watch_inbox()
