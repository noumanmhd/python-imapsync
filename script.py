#!/usr/bin/env python3
import os
import sys
import json
import email
import imaplib
from pathlib import Path

config = json.loads(Path('config.json').read_text())


def mark_as_read(mail_id, mail):
    mail.store(mail_id, '+FLAGS', '\SEEN')


def get_attachments(mail_id, path, mail):
    print(f"Processing email with id {mail_id}...")
    _, data = mail.fetch(mail_id, '(RFC822)')
    raw_email_string = data[0][1].decode('utf-8')

    msg = email.message_from_string(raw_email_string)

    text = f"Subject: {msg['subject']}\n"
    text += f"From: {msg['from']}\n"
    text += f"To: {msg['to']}\n"

    for part in msg.walk():
        ctype = part.get_content_type()
        cdispo = str(part.get('Content-Disposition'))
        data = part.get_payload(decode=True)

        if 'attachment' in cdispo:
            filename = part.get_filename()
            if filename:
                file_path = os.path.join(path, filename)
                print(f"Downloading file [{filename}]...")
                with open(file_path, 'wb') as f:
                    f.write(data)

        elif 'text/plain' in ctype:
            text += "Message:\n"
            text += data.decode('utf-8')

        elif 'text/html' in ctype:
            file_path = os.path.join(path, 'email_body.html')
            with open(file_path, 'wb') as f:
                f.write(data)
    file_path = os.path.join(path, 'email_info.txt')
    with open(file_path, 'w') as f:
        f.write(text)
    mark_as_read(mail_id, mail)


def main():
    try:
        print(f"Connect to Server [{config['SERVER']}]")
        mail = imaplib.IMAP4_SSL(config['SERVER'], config['PORT'])
        mail.login(config['EMAIL'], config['PASSWORD'])
        mail.select()
    except:
        print("Unable to connect with the server!")
        sys.exit(1)
    try:
        print("Fecting ID's of UNSEEN emails...")
        _, data = mail.search(None, 'UNSEEN')
        mail_ids = data[0].split()
        mail_ids = [i.decode('utf-8') for i in mail_ids]
        if len(mail_ids) > 0:
            print(f"{len(mail_ids)} emails found!")
        else:
            print("No UNSEEN email found!")
            sys.exit(0)
    except:
        print("No UNSEEN email found!")
        sys.exit(0)

    for mail_id in mail_ids:
        try:
            path = os.path.join(config['OUTPUT_DIR'], mail_id)
            os.makedirs(path, exist_ok=True)
            get_attachments(mail_id, path, mail)
        except KeyboardInterrupt:
            sys.exit(0)
        except:
            pass


if __name__ == '__main__':
    main()
