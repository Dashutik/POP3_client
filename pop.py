import poplib
import email
from email.header import decode_header
import os
from datetime import datetime


def decode_mime_words(s):
    return ''.join(
        word.decode(encoding or 'utf8') if isinstance(word, bytes) else word
        for word, encoding in decode_header(s)
    )


def save_attachment(part, download_folder):
    filename = part.get_filename()
    if filename:
        filename = decode_mime_words(filename)
        filepath = os.path.join(download_folder, filename)
        with open(filepath, 'wb') as f:
            f.write(part.get_payload(decode=True))
        print(f"Сохранено вложение: {filename}")
        return filename
    return None


def fetch_email(server, port, username, password, download_folder="attachments"):
    if not os.path.exists(download_folder):
        os.makedirs(download_folder)

    mail_server = poplib.POP3_SSL(server, port)
    mail_server.user(username)
    mail_server.pass_(password)

    num_messages = len(mail_server.list()[1])

    if num_messages == 0:
        print("Нет писем на сервере")
        return

    response, lines, octets = mail_server.retr(num_messages)
    raw_email = b'\n'.join(lines).decode('utf-8', errors='ignore')
    msg = email.message_from_string(raw_email)

    print("\n=== Заголовки письма ===")
    print(f"От: {decode_mime_words(msg['From'])}")
    print(f"Кому: {decode_mime_words(msg['To'])}")
    print(f"Тема: {decode_mime_words(msg['Subject'])}")
    print(f"Дата: {decode_mime_words(msg['Date'])}")

    attachments = []
    for part in msg.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue

        attachment = save_attachment(part, download_folder)
        if attachment:
            attachments.append(attachment)

    if not attachments:
        print("Вложений не обнаружено")

    mail_server.quit()
    return msg, attachments


if __name__ == "__main__":
    POP3_SERVER = "pop.gmail.com"
    POP3_PORT = 995
    USERNAME = "name.com"
    PASSWORD = "password"

    print(f"Подключаемся к {POP3_SERVER}...")
    email_msg, attachments = fetch_email(POP3_SERVER, POP3_PORT, USERNAME, PASSWORD)