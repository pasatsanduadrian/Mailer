import os
import pickle
from dotenv import load_dotenv
import gradio as gr

from gmail_utils import get_gmail_service
from export_gmail_to_xlsx import export_labels_and_inbox_xlsx
from move_from_xlsx import move_emails_from_xlsx

os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'
load_dotenv()
TOKEN_FILE = "token.pickle"

def check_auth():
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, "rb") as f:
            try:
                creds = pickle.load(f)
                if creds and creds.valid:
                    return "✅ Ești autentificat cu Gmail!"
            except Exception as e:
                return f"Eroare token: {e}"
    return "❌ Nu ești autentificat. Rulează întâi autentificarea Gmail din scriptul tău Colab sau local!"

def export_xlsx_ui(inbox_max):
    service = get_gmail_service()
    if not service:
        return None, "Eroare: nu ești autentificat Gmail!"
    path = "gmail_labels_inbox.xlsx"
    export_labels_and_inbox_xlsx(service, path, inbox_max)
    return path, f"Export XLSX gata: {path}"

def move_xlsx_ui(file):
    service = get_gmail_service()
    if not service:
        return "Eroare: nu ești autentificat Gmail!"
    move_emails_from_xlsx(service, file.name)
    return "Mutare finalizată! (vezi Gmail)"

with gr.Blocks(theme=gr.themes.Soft()) as demo:
    gr.Markdown("""
    <h2 style='color:#1a73e8'>MailManager – Flux complet organizare Gmail (Labels & Inbox)</h2>
    <ol>
    <li>Asigură-te că te-ai autentificat deja cu Gmail (token.pickle generat!)</li>
    <li>Exportă inbox și labels în XLSX (download)</li>
    <li>Editează manual sheet-ul Inbox (coloana Label) în Excel</li>
    <li>Încarcă XLSX modificat și rulează mutarea automată a mailurilor în Labels</li>
    </ol>
    <hr>
    """)
    auth_status = gr.Markdown(check_auth())
    gr.Button("Verifică autentificare Gmail").click(fn=check_auth, outputs=auth_status)

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
      <li>Rulează scriptul tău de autentificare Gmail pentru a genera <code>token.pickle</code>.</li>
      <li>Folosește butonul de export pentru a salva un fișier XLSX cu două sheet-uri (<b>Labels</b> și <b>Inbox</b>).</li>
      <li>Deschide fișierul în Excel și completează coloana <b>Label</b> din sheet Inbox.</li>
      <li>Încarcă fișierul modificat și apasă <b>Mută automat emailurile</b>.</li>
    </ol>
    </details>
    """)

if __name__ == "__main__":
    demo.launch(share=True)
