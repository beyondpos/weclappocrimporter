import os
import requests
import base64
import json
import time
import threading
from email import message_from_bytes
from io import BytesIO
from requests_toolbelt.multipart.encoder import MultipartEncoder
from dotenv import load_dotenv
from flask import Flask

# Umgebung laden
load_dotenv()

app = Flask(__name__)

# Konfiguration
CLIENT_ID = os.getenv('CLIENT_ID')
TENANT_ID = os.getenv('TENANT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
WECLAPP_API_KEY = os.getenv('WECLAPP_API_KEY')
WECLAPP_TENANT = os.getenv('WECLAPP_TENANT')
USER_EMAIL = os.getenv('USER_EMAIL')
FOLDER_NAME = os.getenv('FOLDER_NAME')
GRAPH_API_ENDPOINT = 'https://graph.microsoft.com/v1.0'
TOKEN_ENDPOINT = f'https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token'

# Token abrufen (Client Credentials Flow) mit Retry
def authenticate_graph():
    max_retries = 3
    for attempt in range(1, max_retries + 1):
        try:
            data = {
                'client_id': CLIENT_ID,
                'scope': 'https://graph.microsoft.com/.default',
                'client_secret': CLIENT_SECRET,
                'grant_type': 'client_credentials'
            }
            response = requests.post(TOKEN_ENDPOINT, data=data, timeout=10)
            response.raise_for_status()
            print("✅ Token erfolgreich abgerufen.", flush=True)
            return response.json()['access_token']
        except requests.exceptions.RequestException as e:
            print(f"❌ Fehler beim Authentifizieren (Versuch {attempt}): {e}", flush=True)
            if attempt < max_retries:
                print("🔄 Neuer Versuch...", flush=True)
                time.sleep(5)
            else:
                raise Exception(f"Verbindung zu Microsoft fehlgeschlagen nach {max_retries} Versuchen: {e}")

# E-Mails auslesen mit Retry
def fetch_emails(access_token):
    max_retries = 3
    for attempt in range(1, max_retries + 1):
        try:
            headers = {'Authorization': f'Bearer {access_token}'}
            response = requests.get(f'{GRAPH_API_ENDPOINT}/users/{USER_EMAIL}/mailFolders', headers=headers, timeout=10)
            response.raise_for_status()
            folders = response.json()
            folder_id = next((f['id'] for f in folders['value'] if f['displayName'] == FOLDER_NAME), None)
            if not folder_id:
                raise Exception(f"Ordner '{FOLDER_NAME}' nicht gefunden.")

            response = requests.get(f'{GRAPH_API_ENDPOINT}/users/{USER_EMAIL}/mailFolders/{folder_id}/messages', headers=headers, timeout=10)
            response.raise_for_status()
            messages = response.json()
            print(f"✅ {len(messages['value'])} E-Mails abgerufen.", flush=True)
            return messages['value']
        except requests.exceptions.RequestException as e:
            print(f"❌ Fehler beim Abrufen der E-Mails (Versuch {attempt}): {e}", flush=True)
            if attempt < max_retries:
                print("🔄 Neuer Versuch...", flush=True)
                time.sleep(5)
            else:
                raise Exception(f"E-Mails konnten nach {max_retries} Versuchen nicht abgerufen werden: {e}")

# E-Mail archivieren
def archive_email(access_token, message_id):
    headers = {'Authorization': f'Bearer {access_token}', 'Content-Type': 'application/json'}
    response = requests.get(f"{GRAPH_API_ENDPOINT}/users/{USER_EMAIL}/mailFolders", headers=headers, timeout=10)
    response.raise_for_status()
    folders = response.json()
    archive_folder_id = next((f['id'] for f in folders['value'] if f['displayName'] in ['Archiv', 'Archive']), None)
    if not archive_folder_id:
        raise Exception("Archiv-Ordner nicht gefunden.")

    move_url = f"{GRAPH_API_ENDPOINT}/users/{USER_EMAIL}/messages/{message_id}/move"
    data = {"destinationId": archive_folder_id}
    response = requests.post(move_url, headers=headers, json=data, timeout=10)
    response.raise_for_status()
    print(f"📥 E-Mail {message_id} archiviert.", flush=True)

# PDF Anhänge sammeln aus allen E-Mails und gemeinsam hochladen
def process_attachments(access_token, messages):
    headers = {'Authorization': f'Bearer {access_token}'}
    pdf_attachments = {}
    message_ids_to_archive = []
    attachment_counter = 1

    for msg in messages:
        response = requests.get(f"{GRAPH_API_ENDPOINT}/users/{USER_EMAIL}/messages/{msg['id']}/attachments", headers=headers, timeout=10)
        response.raise_for_status()
        attachments = response.json()
        has_pdf = False

        for attachment in attachments['value']:
            if attachment['@odata.type'] == '#microsoft.graph.fileAttachment' and attachment['contentType'].lower() == 'application/pdf':
                pdf_bytes = base64.b64decode(attachment['contentBytes'])
                pdf_attachments[f'file{attachment_counter}'] = (attachment['name'], BytesIO(pdf_bytes), 'application/pdf')
                print(f"📄 Gefundene PDF: {attachment['name']}", flush=True)
                attachment_counter += 1
                has_pdf = True

        if has_pdf:
            message_ids_to_archive.append(msg['id'])

    if pdf_attachments:
        upload_multiple_to_weclapp(pdf_attachments)
        for message_id in message_ids_to_archive:
            archive_email(access_token, message_id)
    else:
        print("ℹ️ Keine PDF-Anhänge gefunden.", flush=True)

# Upload mehrere PDFs zur weclapp OCR (mit MultipartEncoder) mit Retry
def upload_multiple_to_weclapp(pdf_attachments):
    url = f"https://{WECLAPP_TENANT}.weclapp.com/webapp/api/v1/purchaseInvoice/startInvoiceDocumentProcessing/multipartUpload"
    print(f"➡️ Upload zu Endpoint: {url}", flush=True)
    m = MultipartEncoder(fields=pdf_attachments)
    headers = {
        'AuthenticationToken': WECLAPP_API_KEY,
        'Accept': 'application/json',
        'Content-Type': m.content_type
    }

    max_retries = 3
    for attempt in range(1, max_retries + 1):
        try:
            response = requests.post(url, headers=headers, data=m, timeout=60)
            response.raise_for_status()
            uploaded_files = ', '.join(name for name, _, _ in pdf_attachments.values())
            print(f"✅ Upload erfolgreich: {uploaded_files}", flush=True)
            break
        except requests.exceptions.RequestException as e:
            print(f"❌ Fehler beim Upload (Versuch {attempt}): {e}", flush=True)

        if attempt < max_retries:
            print("🔄 Neuer Versuch...", flush=True)
            time.sleep(5)
        else:
            print("❌ Alle Upload-Versuche fehlgeschlagen.", flush=True)

# Hauptablauf
def main():
    try:
        access_token = authenticate_graph()
        messages = fetch_emails(access_token)
        if messages:
            process_attachments(access_token, messages)
        else:
            print("ℹ️ Keine neuen E-Mails im Ordner gefunden.", flush=True)
    except Exception as e:
        print(f"❗ Fehler im Hauptablauf: {e}", flush=True)

@app.route('/', methods=['GET'])
def index():
    print("✅ OCR Importer Service läuft! Nutze /run zum Ausführen.", flush=True)
    return "✅ OCR Importer Service läuft! Nutze /run zum Ausführen.", 200

@app.route('/run', methods=['GET'])
def run():
    print("▶️ Manueller Start über /run", flush=True)
    main()
    return "✅ Script ausgeführt", 200

def background_task():
    while True:
        print("⏳ Automatischer Lauf gestartet...", flush=True)
        main()
        print("✅ Automatischer Lauf beendet. Starte 60 Minuten Countdown...", flush=True)
        for remaining in range(60, 0, -1):
            print(f"⏳ Nächste Ausführung in {remaining} Minuten...", flush=True)
            time.sleep(60)

if __name__ == "__main__":
    threading.Thread(target=background_task, daemon=True).start()
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
