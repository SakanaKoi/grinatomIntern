import smtplib
import imaplib
import email
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd


def send_email_with_report(smtp_server, smtp_port, smtp_user, smtp_password, imap_server, imap_user, imap_password,
                           excel_file_path):
    try:
        # Подключаемся к почтовому ящику
        mail = imaplib.IMAP4_SSL(imap_server)
        mail.login(imap_user, imap_password)
        mail.select('inbox')

        # Ищем письмо с темой "Робот"
        subject_utf8 = '(SUBJECT "Тестовое задание (Стажер поддержки RPA)")'.encode('utf-8')
        status, data = mail.search(None, subject_utf8)
        mail_ids = data[0].split()  # Разделяем строку с идентификаторами на список

        if not mail_ids:
            raise Exception("Не найдено письмо с темой 'Тестовое задание (Стажер поддержки RPA)'")

        # Берем последнее письмо с темой "Робот"
        latest_email_id = mail_ids[-1]

        # Получаем полное содержимое сообщения
        status, data = mail.fetch(latest_email_id, '(RFC822)')
        raw_email = data[0][1]
        msg = email.message_from_bytes(raw_email)

        # Формируем текст письма
        int_lastRow = get_last_row(excel_file_path)
        rows = ""
        if int_lastRow == 1:
            rows = f"1 строка"
        elif (11 <= (int_lastRow) % 100 <= 19) or (int_lastRow % 10 in [0, 5, 6, 7, 8, 9]):
            rows = f"{int_lastRow} строк"
        else:
            rows = f"{int_lastRow} строки"
        body_for_mail = f"Здравствуйте. Это сообщение с отчетом отправлено программным роботом, написанным Броль Родионом на языке программирования Python\nВ отчете {rows}, не считая заголовков."

        # Формируем письмо
        reply = MIMEMultipart()
        reply['From'] = smtp_user
        reply['To'] = msg['Reply-To'] if msg['Reply-To'] else msg['From']
        reply['Cc'] = msg['Cc'] if msg['Cc'] else ""
        reply['Subject'] = f"Re: {msg['Subject']}"
        reply.attach(MIMEText(body_for_mail, 'plain', 'utf-8'))

        # Прикрепляем Excel файл
        with open(excel_file_path, 'rb') as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f"attachment; filename={excel_file_path.split('/')[-1]}")
            reply.attach(part)

        # Получаем список всех получателей
        to_addresses = [msg['Reply-To'] if msg['Reply-To'] else msg['From']]
        if msg['Cc']:
            to_addresses.extend(msg['Cc'].split(', '))

        # Отправляем письмо
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_user, smtp_password)
        server.sendmail(smtp_user, to_addresses, reply.as_string())
        server.quit()

    except Exception as e:
        raise Exception("Ошибка при отправке письма с отчетом") from e


def get_last_row(excel_file_path):
    df = pd.read_excel(excel_file_path)
    return len(df)
