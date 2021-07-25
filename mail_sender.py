from __future__ import print_function
import os.path, base64, mimetypes, yaml
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart

with open("config.yaml", "r") as rf:
    config = (yaml.load(rf, Loader=yaml.Loader))

credentials_path = config['GMail']['credential_path']
SENDER = config['GMail']['sender']
SUBJECT = config['GMail']['subject']
with open(config['GMail']['html_path']) as f:
	CORPS = " ".join([l.rstrip() for l in f]) 
SCOPES = ['https://www.googleapis.com/auth/gmail.send']


def create_message_with_attachment(to, message_text, file):
    message = MIMEMultipart()
    message['to'] = to
    message['from'] = SENDER
    message['subject'] = SUBJECT

    msg = MIMEText(message_text, 'html')
    message.attach(msg)

    content_type, encoding = mimetypes.guess_type(file)

    if content_type is None or encoding is not None:
        content_type = 'application/octet-stream'
    main_type, sub_type = content_type.split('/', 1)
    if main_type == 'text':
        fp = open(file, 'rb')
        msg = MIMEText(fp.read(), _subtype=sub_type)
        fp.close()
    elif main_type == 'image':
        fp = open(file, 'rb')
        msg = MIMEImage(fp.read(), _subtype=sub_type)
        fp.close()
    else:
        fp = open(file, 'rb')
        msg = MIMEBase(main_type, sub_type)
        msg.set_payload(fp.read())
        fp.close()
    filename = os.path.basename(file)
    msg.add_header('Content-Disposition', 'attachment', filename=filename)
    message.attach(msg)

    b64_bytes = base64.urlsafe_b64encode(message.as_bytes())
    b64_string = b64_bytes.decode()
    body = {'raw': b64_string}
    return body


def send_message(service, user_id, message):
    try:
        message = (service.users().messages().send(userId=user_id, body=message)
                   .execute())
        return message
    except Exception as e:
        print(f"exception: {e}")


def send_email(row):
    # 0 Nom, 1 Prénom, 2 Email, 3 Cell, 4 PNG Path
    nom = row[0]
    prenom = row[1]
    email = row[2]
    path = row[4]
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('gmail', 'v1', credentials=creds)
    
    ms = create_message_with_attachment(
        email, ''.join(CORPS).format(prenom), path)
    m = send_message(service, 'me', ms)
    print(f' > Done: {prenom} {nom} ({m["id"]} sent)\n')
