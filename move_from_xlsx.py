import pandas as pd
from gmail_utils import get_gmail_service

def move_emails_from_xlsx(service, xlsx_path="gmail_labels_inbox.xlsx"):
    df = pd.read_excel(xlsx_path, sheet_name="Inbox")
    labels_gmail = service.users().labels().list(userId='me').execute().get('labels', [])
    labels_dict = {l['name']: l['id'] for l in labels_gmail}
    for _, row in df.iterrows():
        label = str(row.get('Label', '')).strip()
        msg_id = row['ID']
        if not label or not msg_id:
            continue
        # Creează label dacă nu există deja
        if label not in labels_dict:
            label_res = service.users().labels().create(userId='me', body={'name': label}).execute()
            label_id = label_res['id']
            labels_dict[label] = label_id
        else:
            label_id = labels_dict[label]
        # Mută mailul (adaugă label, scoate din INBOX)
        service.users().messages().modify(
            userId='me', id=msg_id,
            body={'addLabelIds': [label_id], 'removeLabelIds': ['INBOX']}
        ).execute()
    print("Toate emailurile cu Label completat au fost mutate.")
