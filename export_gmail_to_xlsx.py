import pandas as pd
from gmail_utils import get_gmail_service

def export_labels_and_inbox_xlsx(service, xlsx_path="gmail_labels_inbox.xlsx", inbox_max=2000):
    # Export labels (Sheet: Labels)
    labels = service.users().labels().list(userId='me').execute().get('labels', [])
    df_labels = pd.DataFrame([{"Label": l["name"], "ID": l["id"]} for l in labels if l.get("type") == "user"])

    # Export inbox (Sheet: Inbox)
    messages = service.users().messages().list(userId='me', labelIds=['INBOX'], maxResults=inbox_max).execute().get('messages', [])
    emails = []
    for m in messages:
        msg = service.users().messages().get(userId='me', id=m['id'], format='metadata', metadataHeaders=['From','Subject','Date']).execute()
        headers = {h['name']: h['value'] for h in msg.get('payload', {}).get('headers', [])}
        emails.append({
            "ID": m['id'],
            "From": headers.get('From', ''),
            "Date": headers.get('Date', ''),
            "Subject": headers.get('Subject', ''),
            "Label": ""   # pentru completare manuală ulterior
        })
    df_inbox = pd.DataFrame(emails)

    # Scriere xlsx cu 2 sheet-uri
    with pd.ExcelWriter(xlsx_path, engine='xlsxwriter') as writer:
        df_labels.to_excel(writer, sheet_name="Labels", index=False)
        df_inbox.to_excel(writer, sheet_name="Inbox", index=False)
    print(f"Export complet: {xlsx_path}")
