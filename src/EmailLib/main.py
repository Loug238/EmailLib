import imaplib
import email
import datetime
import os.path
import email.mime.application
from email.header import decode_header
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Параметры для подключения к серверу почты (Яндекс Почта)
IMAP_SERVER = "imap.yandex.ru"
EMAIL_ACCOUNT = "TestEmailLib@yandex.ru"
APP_PASSWORD = "ehimznmjptswibcs"
SMTP_SERVER = "smtp.yandex.ru"
port = 465

# путь для сохранения писем
path = "C:/Users/Loug238/Desktop/EmailFolder"

def receive_emails(IMAP_SERVER, EMAIL_ACCOUNT, APP_PASSWORD, seen="ALL", count=None, start_date=None, end_date=None, read=0, rev=False):
    try:
        # Подключение к серверу и вход в аккаунт
        imap = imaplib.IMAP4_SSL(IMAP_SERVER)
        imap.login(EMAIL_ACCOUNT, APP_PASSWORD)

        # Выбор папки "INBOX"
        imap.select("INBOX")

        # Определение критерия поиска
        criteria = []
        if start_date:
            criteria.append(f"SINCE {start_date}")
        if end_date:
            criteria.append(f"BEFORE {end_date}")
        if seen == "UNSEEN":
            criteria.append("UNSEEN")
        if seen == "SEEN":
            criteria.append("SEEN")

        query = " ".join(criteria) if criteria else "ALL"

        # Поиск всех писем по критериям
        status, data = imap.search(None, query)

        # Преобразование списка ID сообщений
        email_ids = data[0].split()

        # Ограничение количества выгружаемых писем
        if count and len(email_ids) > count:
            email_ids = email_ids[:count]

        emails = []

        for email_id in email_ids:
            # Получение сообщения по ID
            if read == 0:
                status, msg_data = imap.fetch(email_id, "(BODY.PEEK[])")
            else:
                status, msg_data = imap.fetch(email_id, "(RFC822)")

            # Получение содержимого сообщения
            raw_msg = msg_data[0][1]
            emails.append(raw_msg)

        # Сортировка писем
        emails.sort(reverse=rev) # False - от нового к старому, True - наоборот

    finally:
        # Завершение сессии
        imap.logout()

        # возврат писем
        return emails


def print_emails(emails):
    msgs = []
    for raw_email in emails:
        data = email.message_from_bytes(raw_email)
        msgs.append(data)
    # Вывод информации о сервере и аккаунте
    print("Сервер:", IMAP_SERVER, "Аккаунт:", EMAIL_ACCOUNT)
    print()
    for msg in msgs:

        # Декодирование заголовка
        subject, encoding = decode_header(msg["Subject"])[0]

        # Проверка типа заголовков
        if isinstance(subject, bytes):
            subject = subject.decode(encoding if encoding else "utf-8")

        # Получение даты получения, ID письма, Email отправителя
        letter_date = datetime.datetime(*(email.utils.parsedate_tz(msg["Date"]))[0:6])
        letter_id = msg["Message-ID"]
        letter_from = msg["Return-path"]
        print("________________________________________________")
        print("Тема:", subject)
        print("Дата получения:", letter_date)
        print("ID письма:", letter_id)
        print("Email отправителя:", letter_from)

        # Получение текста сообщений
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    body = part.get_payload(decode=True).decode()
                    print("Текст сообщения:", body)
                    break
        else:
            body = msg.get_payload(decode=True).decode()
            print("Текст сообщения:", body)


def email_to_eml(email, path, email_name):
    email_name = os.path.join(path, email_name + ".eml")
    f = open(email_name, "wb")
    f.write(email)
    f.close()


def send_email(SMTP_SERVER, EMAIL_ACCOUNT, APP_PASSWORD, port, receiver_email, email_subject, email_body, attachment):
    try:
        # Подключение к серверу и вход в аккаунт
        smtp = smtplib.SMTP_SSL(SMTP_SERVER, port)
        smtp.login(EMAIL_ACCOUNT, APP_PASSWORD)

        # Создание письма
        msg = MIMEMultipart()
        msg["From"] = EMAIL_ACCOUNT
        msg["To"] = receiver_email
        msg["Subject"] = email_subject
        msg.attach(MIMEText(email_body, "plain"))

        # Добавление файла в письмо
        f = open(attachment, "rb")
        att = email.mime.application.MIMEApplication(f.read())
        f.close()
        att.add_header("Content-Disposition", "attachment", filename=attachment)
        msg.attach(att)


        # Отправка сообщения
        smtp.send_message(msg)
        print("Письмо отправлено")

    except Exception as exc:
        print(f"Не удалось отправить письмо: {exc}" )


    finally:
        smtp.quit()





#print_emails(receive_emails(IMAP_SERVER, EMAIL_ACCOUNT, APP_PASSWORD, start_date='01-Jan-2024', seen="UNSEEN"))
#emails = receive_emails(IMAP_SERVER, EMAIL_ACCOUNT, APP_PASSWORD, start_date='01-Jan-2024', seen="UNSEEN")
#email_to_eml(email=emails[0], path=path, email_name="test")
send_email(SMTP_SERVER, EMAIL_ACCOUNT, APP_PASSWORD, port, EMAIL_ACCOUNT, "Тест отправки 1", "vrnviervne", "testSend.docx")
#123
