import os
import time
import base64
from io import BytesIO
import requests
from flask import Flask
from requests_toolbelt.multipart.encoder import MultipartEncoder
from dotenv import load_dotenv
from uuid import uuid4

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

print("✅ Weclapp OCR Importer Script startet...")


def request_with_retries(method, url, headers=None, data=None, json_data=None, retries=3, timeout=10):
    for attempt in range(1, retries + 1):
        try:
            response = requests.request(method, url, headers=headers, data=data, json=json_data, timeout=timeout)
            response.raise_for_status()
            return response
        except requests.RequestException as e:
            print(f"❌ Fehler bei {url} (Versuch {attempt}): {e}", flush=True)
            if attempt == retries:
                raise
            time.sleep(5)


def authenticate_graph():
    data = {
        'client_id': CLIENT_ID,
        'scope': 'https://graph.microsoft.com/.default',
        'client_secret': CLIENT_SECRET,
        'grant_type': 'client_credentials'
    }
    response = request_with_retries("POST", TOKEN_ENDPOINT, data=data)
    print("✅ Token erfolgreich abgerufen.", flush=True)
    return response.json()['access_token']


def get_folder_id(access_token, folder_name):
    headers = {'Authorization': f'Bearer {access_token}'}
    response = request_with_retries("GET", f'{GRAPH_API_ENDPOINT}/users/{USER_EMAIL}/mailFolders', headers=headers)
    folders = response.json().get('value', [])
    folder_id = next((f['id'] for f in folders if f['displayName'] == folder_name), None)
    if not folder_id:
        raise Exception(f"Ordner '{folder_name}' nicht gefunden.")
    return folder_id


def fetch_emails(access_token, folder_id):
    headers = {'Authorization': f'Bearer {access_token}'}
    response = request_with_retries("GET", f'{GRAPH_API_ENDPOINT}/users/{USER_EMAIL}/mailFolders/{folder_id}/messages', headers=headers)
    messages = response.json().get('value', [])
    print(f"✅ {len(messages)} E-Mails abgerufen.", flush=True)
    return messages


def archive_email(access_token, message_id, archive_folder_id):
    headers = {'Authorization': f'Bearer {access_token}', 'Content-Type': 'application/json'}
    move_url = f"{GRAPH_API_ENDPOINT}/users/{USER_EMAIL}/messages/{message_id}/move"
    data = {"destinationId": archive_folder_id}
    request_with_retries("POST", move_url, headers=headers, json_data=data)
    print(f"📥 E-Mail {message_id} archiviert.", flush=True)


def process_attachments(access_token, messages, archive_folder_id):
    headers = {'Authorization': f'Bearer {access_token}'}
    pdf_attachments = {}
    message_ids_to_archive = []

    for msg in messages:
        response = request_with_retries("GET", f"{GRAPH_API_ENDPOINT}/users/{USER_EMAIL}/messages/{msg['id']}/attachments", headers=headers)
        attachments = response.json().get('value', [])

        for attachment in attachments:
            if attachment['@odata.type'] == '#microsoft.graph.fileAttachment' and attachment['contentType'].lower() == 'application/pdf':
                pdf_bytes = base64.b64decode(attachment['contentBytes'])
                pdf_attachments[str(uuid4())] = (attachment['name'], BytesIO(pdf_bytes), 'application/pdf')
                print(f"📄 Gefundene PDF: {attachment['name']}", flush=True)
                message_ids_to_archive.append(msg['id'])

    if pdf_attachments:
        upload_multiple_to_weclapp(pdf_attachments)
        for message_id in message_ids_to_archive:
            archive_email(access_token, message_id, archive_folder_id)
    else:
        print("ℹ️ Keine PDF-Anhänge gefunden.", flush=True)


def upload_multiple_to_weclapp(pdf_attachments):
    url = f"https://{WECLAPP_TENANT}.weclapp.com/webapp/api/v1/purchaseInvoice/startInvoiceDocumentProcessing/multipartUpload"
    m = MultipartEncoder(fields=pdf_attachments)
    headers = {
        'AuthenticationToken': WECLAPP_API_KEY,
        'Accept': 'application/json',
        'Content-Type': m.content_type
    }
    request_with_retries("POST", url, headers=headers, data=m, timeout=60)
    uploaded_files = ', '.join(name for name, _, _ in pdf_attachments.values())
    print(f"✅ Upload erfolgreich: {uploaded_files}", flush=True)


def main():
    try:
        access_token = authenticate_graph()
        folder_id = get_folder_id(access_token, FOLDER_NAME)
        archive_folder_id = get_folder_id(access_token, 'Archiv')
        messages = fetch_emails(access_token, folder_id)
        if messages:
            process_attachments(access_token, messages, archive_folder_id)
        else:
            print("ℹ️ Keine PDFs zum Importieren gefunden.", flush=True)
    except Exception as e:
        print(f"❗ Fehler im Hauptablauf: {e}", flush=True)


@app.route('/', methods=['GET'])
def index():
    return "ℹ️ Nutze /run um das Skript manuell auszuführen.", 200


@app.route('/run', methods=['GET'])
def run():
    print("▶️ Manueller Start über /run", flush=True)
    main()
    return "✅ Script manuell ausgeführt", 200


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
