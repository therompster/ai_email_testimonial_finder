
import os
import json
import subprocess
import base64
import requests

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from msal import PublicClientApplication

REFERENCE_LABEL = "Reference Emails"
KEYWORDS = ["thank you", "testimonial", "great work", "reference", "appreciate", "recommend", "endorse"]

def is_testimonial_llama(subject, body):
    prompt = f"""
You are filtering emails to detect **genuine client praise or testimonials** — not marketing spam or automated messages.

Only respond YES if the message:
- Is written by a real person (not a company bot or promotion)
- Contains meaningful appreciation, recommendation, or reference to working with the user
- Could reasonably be quoted or summarized as client praise

Subject: {subject}
Body:
{body[:1200]}

Answer with YES or NO and a short explanation.
"""
    result = subprocess.run(
        [r"C:\Users\thero\AppData\Local\Programs\Ollama\ollama.exe", "run", "llama3"],
        input=prompt.encode("utf-8"),
        stdout=subprocess.PIPE
    )
    output = result.stdout.decode("utf-8").strip()
    return "YES" in output.upper(), output


def get_gmail_service():
    SCOPES = ['https://www.googleapis.com/auth/gmail.modify']
    creds = None
    if os.path.exists('token_gmail.json'):
        creds = Credentials.from_authorized_user_file('token_gmail.json', SCOPES)
    else:
        flow = InstalledAppFlow.from_client_secrets_file('credentials_gmail.json', SCOPES)
        creds = flow.run_local_server(port=0)
        with open('token_gmail.json', 'w') as token:
            token.write(creds.to_json())
    return build('gmail', 'v1', credentials=creds)

def get_or_create_gmail_label(service, name=REFERENCE_LABEL):
    labels = service.users().labels().list(userId='me').execute().get('labels', [])
    for label in labels:
        if label['name'] == name:
            return label['id']
    label = service.users().labels().create(userId='me', body={'name': name}).execute()
    return label['id']

def process_gmail(service, only_with_replies=True, use_keywords=False):
    label_id = get_or_create_gmail_label(service)
    results = []

    # Gmail search query
    query = ""
    if only_with_replies:
        query += "has:userlabels"  # often works well as proxy for replied threads
    if use_keywords:
        keyword_query = " OR ".join(KEYWORDS)
        query += f" ({keyword_query})"

    next_page_token = None

    while True:
        response = service.users().messages().list(
            userId='me',
            q=query.strip(),
            pageToken=next_page_token,
            maxResults=100
        ).execute()

        messages = response.get('messages', [])
        if not messages:
            break

        for msg in messages:
            msg_id = msg['id']
            data = service.users().messages().get(userId='me', id=msg_id, format='full').execute()
            headers = {h['name']: h['value'] for h in data['payload']['headers']}
            subject = headers.get("Subject", "")
            sender = headers.get("From", "")
            date = headers.get("Date", "")
            snippet = data.get("snippet", "")

            # Optional basic spam filter
            if any(x in sender.lower() for x in ['no-reply', 'noreply', '@news', '@email.', '@mailer.']):
                continue

            # Extra reply check (if flag enabled)
            if only_with_replies:
                thread_id = data.get('threadId')
                thread = service.users().threads().get(userId='me', id=thread_id).execute()
                if len(thread.get('messages', [])) <= 1:
                    continue

            is_testimonial, reason = is_testimonial_llama(subject, snippet)
            if is_testimonial:
                service.users().messages().modify(
                    userId='me',
                    id=msg_id,
                    body={'addLabelIds': [label_id], 'removeLabelIds': ['INBOX']}
                ).execute()
                results.append({
                    "source": "Gmail",
                    "sender": sender,
                    "subject": subject,
                    "date": date,
                    "reason": reason
                })

        next_page_token = response.get('nextPageToken')
        if not next_page_token:
            break

    return results

def get_outlook_token():
    CLIENT_ID = "YOUR_CLIENT_ID"
    TENANT_ID = "YOUR_TENANT_ID"
    AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
    SCOPES = ['Mail.ReadWrite']
    app = PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
    accounts = app.get_accounts()
    result = app.acquire_token_silent(SCOPES, account=accounts[0]) if accounts else None
    if not result:
        result = app.acquire_token_interactive(SCOPES)
    return result['access_token']

def get_or_create_outlook_folder(token):
    headers = {'Authorization': f'Bearer {token}'}
    resp = requests.get("https://graph.microsoft.com/v1.0/me/mailFolders", headers=headers)
    folders = resp.json().get("value", [])
    for folder in folders:
        if folder['displayName'].lower() == REFERENCE_LABEL.lower():
            return folder['id']
    create_resp = requests.post(
        "https://graph.microsoft.com/v1.0/me/mailFolders",
        headers={**headers, "Content-Type": "application/json"},
        json={"displayName": REFERENCE_LABEL}
    )
    return create_resp.json()['id']

def process_outlook(token):
    folder_id = get_or_create_outlook_folder(token)
    headers = {'Authorization': f'Bearer {token}'}
    query = " OR ".join([f'"{k}"' for k in KEYWORDS])
    search_url = f"https://graph.microsoft.com/v1.0/me/messages?$search={query}"
    resp = requests.get(search_url, headers=headers)
    results = []
    for msg in resp.json().get("value", [])[:50]:
        subject = msg.get("subject", "")
        sender = msg.get("from", {}).get("emailAddress", {}).get("name", "")
        date = msg.get("receivedDateTime", "")
        snippet = msg.get("bodyPreview", "")
        msg_id = msg['id']
        is_testimonial, reason = is_testimonial_llama(subject, snippet)
        if is_testimonial:
            requests.post(
                f"https://graph.microsoft.com/v1.0/me/messages/{msg_id}/move",
                headers={**headers, "Content-Type": "application/json"},
                json={"destinationId": folder_id}
            )
            results.append({
                "source": "Outlook",
                "sender": sender,
                "subject": subject,
                "date": date,
                "reason": reason
            })
    return results

def main():
    all_results = []
    print("Processing Gmail...")
    gmail_service = get_gmail_service()
    gmail_results = process_gmail(gmail_service, only_with_replies=True, use_keywords=False)

    all_results.extend(gmail_results)

#    print("Processing Outlook...")
#    outlook_token = get_outlook_token()
#    outlook_results = process_outlook(outlook_token)
#    all_results.extend(outlook_results)

    with open("classified_testimonials.json", "w") as f:
        json.dump(all_results, f, indent=2)
    print(f"Saved {len(all_results)} testimonials.")

if __name__ == "__main__":
    main()
    