# Main modules
import os
import pickle

# Data modules
import pandas as pd

# Gmail API utils
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


# Basic Info for project
SCOPES = ['https://mail.google.com/']
our_email = 'BlackJoeNSK@gmail.com'

# Main Dataframe with all
email_dataframe = pd.DataFrame(columns=['Sender', 'From', 'To', 'Labels'])

# Function for e-mail Authentication
def gmail_authenticate():
    creds = None
    # the file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)
    # if there are no (valid) credentials availablle, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('configs/credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # save the credentials for the next run
        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)
    return build('gmail', 'v1', credentials=creds)


# Function for e-mail search
def search_messages(service, query):
    result = service.users().messages().list(userId='me', q=query, maxResults=500).execute()
    messages = []

    if 'messages' in result:
        messages.extend(result['messages'])

    counter = 0
    while 'nextPageToken' in result:
        counter += 1
        print('Now Parsing Page #', counter)
        page_token = result['nextPageToken']
        result = service.users().messages().list(userId='me', q=query, pageToken=page_token).execute()
        if 'messages' in result:
            messages.extend(result['messages'])
    return messages


# Function for e-mail parsing
def read_message(service, message):

    msg = service.users().messages().get(userId='me', id=message, format='full').execute()
    # parts can be the message body, or attachments
    payload = msg['payload']

    labels = msg['labelIds']
    headers = payload.get("headers")

    # Parsing only useful info
    inf_sender = ''
    inf_from = ''
    inf_to = ''

    if headers:
        for header in headers:

            name = header.get('name')
            value = header.get('value')

            if name.lower() == 'sender':
                inf_sender = value

            if name.lower() == 'from':
                inf_from = value

            if name.lower() == 'to':
                inf_to = value

    # Writing into dict
    writer = [inf_sender, inf_from, inf_to, labels]

    email_dataframe.loc[len(email_dataframe)] = writer

# get the Gmail API service
service = gmail_authenticate()

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print('Starting')

    found_messages = search_messages(service, 'From:')

    print('Got all messages', len(found_messages))

    counter = 1
    divider = round(len(found_messages) / 100, 0)

    for i in enumerate(found_messages):

        if i[0] == (divider * counter):
            print('Done with ', counter, '% of e-mails')
            counter += 1

        msg_id = i[1]['id']
        try:
            read_message(service, msg_id)
        except KeyError:
            print('End of List')

    print('Done processing')

    email_dataframe.to_excel('Email_List.xlsx', index=False)


    print('Done')

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
