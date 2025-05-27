import os
import pickle
import time
from threading import Thread
from dotenv import load_dotenv
from pyngrok import ngrok
from flask import Flask, request, redirect, session
from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request
import gradio as gr

from gmail_utils import get_gmail_service
from export_gmail_to_xlsx import export_labels_and_inbox_xlsx
from move_from_xlsx import move_emails_from_xlsx

os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'
load_dotenv()

FLASK_PORT = int(os.getenv("FLASK_PORT", 5099))
GRADIO_PORT = int(os.getenv("GRADIO_PORT", 7070))
NGROK_TOKEN = os.getenv("NGROK_TOKEN")
NGROK_HOSTNAME = os.getenv("NGROK_HOSTNAME")
CREDENTIALS_FILE = 'credentials.json'
TOKEN_FILE = 'token.pickle'
SCOPES = ['https://www.googleapis.com/auth/gmail.modify']
SECRET_KEY = os.getenv("SECRET_KEY", "abc123")

# --- Flask pentru OAuth2 ---
app = Flask(__name__)
app.secret_key = SECRET_KEY

@app.route("/")
def index():
    return """
    <div style='text-align:center;margin-top:32px;'>
      <h2 style='color:#1a73e8;'>MailManager - Gmail OAuth2</h2>
      <a href='/auth' style='font-size:1.3em;background:#1a73e8;color:white;padding:12px 36px;border-radius:8px;text-decoration:none;'>🔐 Autentificare Gmail</a>
    </div>"""

@app.route("/auth")
def auth():
    flow = Flow.from_client_secrets_file(
        CREDENTIALS_FILE,
        scopes=SCOPES,
        redirect_uri=f"https://{NGROK_HOSTNAME}/oauth2callback"
    )
    auth_url, state = flow.authorization_url(prompt='consent', access_type='offline')
    session['state'] = state
    return redirect(auth_url)

@app.route("/oauth2callback")
def oauth2callback():
    state = session.get('state')
    if not state:
        return "Missing OAuth state! Please restart the authentication.", 400
    flow = Flow.from_client_secrets_file(
        CREDENTIALS_FILE,
        scopes=SCOPES,
        redirect_uri=f"https://{NGROK_HOSTNAME}/oauth2callback",
        state=state
    )
    flow.fetch_token(authorization_response=request.url)
    creds = flow.credentials
    with open(TOKEN_FILE, "wb") as token:
        pickle.dump(creds, token)
    return "<h3 style='color:green;'>PAS /oauth2callback - Token salvat. Autentificare reușită!<br>Poți închide acest tab și reveni în Gradio.</h3>"

# --- Pornește Flask pe thread separat + ngrok stable ---
ngrok.set_auth_token(NGROK_TOKEN)
public_url = ngrok.connect(FLASK_PORT, "http", hostname=NGROK_HOSTNAME)
print("Ngrok stable link:", public_url)
print(f"Adaugă la Google Console: {public_url}/oauth2callback")

def run_flask():
    app.run(port=FLASK_PORT, host="0.0.0.0")

flask_thread = Thread(target=run_flask)
flask_thread.daemon = True
flask_thread.start()
time.sleep(5)

# --- Gradio UI ---
def check_auth():
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, "rb") as f:
            try:
                creds = pickle.load(f)
                if creds and creds.valid:
                    return "✅ Ești autentificat cu Gmail!"
            except Exception as e:
                return f"Eroare token: {e}"
    return ("❌ Nu ești autentificat. Click butonul de mai jos pentru autentificare și urmează pașii în browser.")

def open_auth_link():
    return f"Deschide <a href='https://{NGROK_HOSTNAME}/auth' target='_blank'>aici</a> pentru autentificare Gmail (OAuth2)."

def export_xlsx_ui(inbox_max):
    service = get_gmail_service()
    if not service:
        return None, "Eroare: nu ești autentificat Gmail!"
    path = "gmail_labels_inbox.xlsx"
    export_labels_and_inbox_xlsx(service, path, int(inbox_max))
    return path, f"Export XLSX gata: {path}"

def move_xlsx_ui(file):
    service = get_gmail_service()
    if not service:
        return "Eroare: nu ești autentificat Gmail!"
    move_emails_from_xlsx(service, file.name)
    return "Mutare finalizată! (vezi Gmail)"

with gr.Blocks(theme=gr.themes.Soft()) as demo:
    gr.Markdown("""
    <h2 style='color:#1a73e8'>MailManager – Gmail OAuth2 + Workflow Inbox/Labels</h2>
    <ol>
    <li>Click pe butonul de autentificare, urmează pașii din browser (vei genera token.pickle)</li>
    <li>Exportă inbox și labels în XLSX (download)</li>
    <li>Editează manual sheet-ul Inbox (coloana Label) în Excel</li>
    <li>Încarcă XLSX modificat și rulează mutarea automată a mailurilor în Labels</li>
    </ol>
    <hr>
    """)
    auth_status = gr.Markdown(check_auth())
    auth_btn = gr.Button("🔐 Deschide autentificarea Gmail în browser")
    auth_btn.click(open_auth_link, outputs=auth_status)

    gr.Markdown("### 1️⃣ Exportă Labels & Inbox în XLSX pentru editare")
    max_nr = gr.Number(label="Câte emailuri vrei să exporți din Inbox?", value=1000, precision=0)
    exp_btn = gr.Button("Export XLSX")
    exp_file = gr.File(label="Fișierul XLSX generat pentru download", interactive=False)
    exp_msg = gr.Textbox(label="Status export", lines=1)
    exp_btn.click(fn=export_xlsx_ui, inputs=max_nr, outputs=[exp_file, exp_msg])

    gr.Markdown("### 2️⃣ Încarcă XLSX editat pentru mutare automată în Labels")
    upl_file = gr.File(label="Încarcă XLSX editat (Inbox cu labeluri)", file_types=['.xlsx'])
    move_btn = gr.Button("Mută automat emailurile din fișier")
    move_msg = gr.Textbox(label="Status mutare", lines=2)
    move_btn.click(fn=move_xlsx_ui, inputs=upl_file, outputs=move_msg)

    gr.Markdown("""
    ---
    <details>
    <summary><b>Instrucțiuni complete</b></summary>
    <ol>
      <li>Click pe butonul de autentificare Gmail de mai sus și urmează pașii din browser (vei genera <code>token.pickle</code>).</li>
      <li>Folosește butonul de export pentru a salva un fișier XLSX cu două sheet-uri (<b>Labels</b> și <b>Inbox</b>).</li>
      <li>Deschide fișierul în Excel și completează/editează coloana <b>Label</b> din sheet Inbox.</li>
      <li>Încarcă fișierul modificat și apasă <b>Mută automat emailurile</b>.</li>
    </ol>
    </details>
    """)

if __name__ == "__main__":
    demo.launch(share=True, server_port=7070)
