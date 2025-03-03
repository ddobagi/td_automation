# 눈물 겹다.... 
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import gspread
from google.oauth2 import service_account
from googleapiclient.discovery import build

from google.auth.transport.requests import Request

import time
import json
import os
import imaplib
import email as email_lib
from email.utils import parsedate_to_datetime
from datetime import datetime, timedelta
import re
import pytz

# Google Sheets API 인증
SERVICE_ACCOUNT_JSON = os.getenv("GOOGLE_SHEETS_CREDENTIALS")
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

creds = service_account.Credentials.from_service_account_info(json.loads(SERVICE_ACCOUNT_JSON), scopes=SCOPES)
 # 인증 파일 경로

if creds and creds.expired and creds.refresh_token:
    creds.refresh(Request())
    print("new token refreshed")

client = gspread.authorize(creds)

# 스프레드시트 열기
spreadsheet = client.open("Identification Request")

# 워크시트 선택
worksheet = spreadsheet.get_worksheet(0)

# 이메일 처리 관련 변수
PROCESSED_COLUMN = 6
PROCESSED_COLUMN_PAYMENT = 16
PROCESSED_COLUMN_REQUEST = 9
last_requested_email = None
last_requested_timestamp = None
last_processed_email_id = None

# 이메일 관리 리스트 캐시
email_management_list_cache = None
email_management_list_last_fetched = None
email_management_list_cache_ttl = 10800  # 3시간 캐시 타임

# 새롭게 추가됨, is_email_in_management_list 내용

def fetch_email_management_list():
    """이메일 관리 리스트를 캐시하거나, 캐시에서 반환"""
    global email_management_list_cache, email_management_list_last_fetched
    current_time = time.time()
    
    if (email_management_list_cache is None or
        email_management_list_last_fetched is None or
        current_time - email_management_list_last_fetched > email_management_list_cache_ttl):
        
        SPREADSHEET_ID = '1JcG3bPyk7VLCevMJYWGWyXT5JcwM8Aex-KwoRiL8gFI'
        RANGE_NAME = 'pass_management!R:R'
        
        try:
            service = build('sheets', 'v4', credentials=creds)
            sheet = service.spreadsheets()
            result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
            email_management_list_cache = [row[0] for row in result.get('values', [])]
            email_management_list_last_fetched = current_time
        except gspread.exceptions.APIError as e:
            print(f"Google Sheets API 오류: {e}")
            send_email("[Error] fetch_email_management_list에서 google sheets api 오류 발생", 
                       "[Error] fetch_email_management_list에서 google sheets api 오류 발생", 
                       "xx11chotae@gmail.com")
            return []  # 오류 발생 시 빈 리스트 반환
        except Exception as e:
            print(f"예상치 못한 오류 발생 in fetch_email_management_list: {e}")
            send_email("[Error] fetch_email_management_list에서 예상치 못한 오류 발생", 
                       "[Error] fetch_email_management_list에서 예상치 못한 오류 발생", 
                       "xx11chotae@gmail.com")
            return []  # 오류 발생 시 빈 리스트 반환

    return email_management_list_cache



def is_email_in_management_list(email):
    """이메일 관리 리스트에서 이메일 존재 여부 확인"""
    email_list = fetch_email_management_list()
    print("현재 유효 이메일 목록:", email_list)
    return email in email_list

def send_email(subject, body, to_email):
    """이메일 전송 기능"""
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    sender_email = 'timelydrop.email@gmail.com'
    sender_password = 'nmtarxkezacupxcc'
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, to_email, msg.as_string())
        server.quit()
        print(f'이메일 전송 성공: {to_email}')
    except Exception as e:
        print(f'이메일 전송 실패: {e}')

def get_latest_email_from_sheet():
    """스프레드시트에서 최근 이메일과 해당 행 번호를 캐싱된 데이터를 이용하여 반환"""
    try:
        column_e = worksheet.col_values(5)  # 5번째 열 (이메일 열) 읽기
        processed_column = worksheet.col_values(PROCESSED_COLUMN)  # 처리 상태 열 읽기
    except gspread.exceptions.APIError as e:
        print(f"Google Sheets API 오류: {e}")
        send_email("[Error] get_latest_email_from_sheet에서 google sheets api 오류 발생", 
                   "[Error] get_latest_email_from_sheet에서 google sheets api 오류 발생", 
                   "xx11chotae@gmail.com")
        return None, None
    except Exception as e:
        print(f"예상치 못한 오류 발생 in get_latest_email_from_sheet: {e}")
        send_email("[Error] get_latest_email_from_sheet에서 예상치 못한 오류 발생", 
                   "[Error] get_latest_email_from_sheet에서 예상치 못한 오류 발생", 
                   "xx11chotae@gmail.com")
        return None, None

    for index in reversed(range(len(column_e))):
        email = column_e[index]
        if email and (len(processed_column) <= index or not processed_column[index]):
            return email, index + 1

    return None, None



def mark_email_as_processed(row):
    try:
        worksheet.update_cell(row, PROCESSED_COLUMN, "Processed")
    except gspread.exceptions.APIError as e:
        print(f"Google Sheets API 오류: {e}")
        send_email("[Error] mark_email_as_processed에서 google sheets api 오류 발생","[Error] mark_email_as_processed에서 google sheets api 오류 발생", "xx11chotae@gmail.com")
    except Exception as e:
        print(f"예상치 못한 오류 발생 in mark_email_as_processed: {e}")
        send_email("[Error] mark_email_as_processed에서 예상치 못한 오류 발생","[Error] mark_email_as_processed에서 예상치 못한 오류 발생", "xx11chotae@gmail.com")




def fetch_latest_sent_email(last_processed_email_id=None):
    """Gmail에서 마지막으로 전송된 이메일의 본문을 가져옴, 인증번호 문자가 메일로 전달될 후 3분 이내 시점이면 ok"""
    username = "timelydrop.email@gmail.com"
    password = "nmtarxkezacupxcc"
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(username, password)
    mail.select("inbox")
    result, data = mail.search(None, '(FROM "timelydrop.email@gmail.com" TO "timelydrop.email@gmail.com")')
    email_ids = data[0].split()
    if not email_ids:
        return None, None, last_processed_email_id

    latest_email_id = email_ids[-1]
    if latest_email_id == last_processed_email_id:
        return None, None, last_processed_email_id

    result, msg_data = mail.fetch(latest_email_id, "(RFC822)")
    for response_part in msg_data:
        if isinstance(response_part, tuple):
            msg = email_lib.message_from_bytes(response_part[1])
            timestamp = parsedate_to_datetime(msg["Date"])
            if datetime.now(pytz.utc) - timestamp > timedelta(minutes=3):
                return None, None, latest_email_id

            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() == "text/plain":
                        body = part.get_payload(decode=True).decode()
                        match = re.search(r'\[수신내용\](.*)', body, re.DOTALL)
                        if match:
                            return match.group(1).strip(), timestamp, latest_email_id
            else:
                body = msg.get_payload(decode=True).decode()
                match = re.search(r'\[수신내용\](.*)', body, re.DOTALL)
                if match:
                    return match.group(1).strip(), latest_email_id

    return None, None, last_processed_email_id

def mark_request_as_processed(row):
    # worksheet 객체가 함수 외부에서 접근 가능하도록 수정
    payment_sheet = client.open_by_key('1mzwPFpWZRyDblVmMwLaRtTvLLabV6p_sE6pHPROTsRY')
    worksheet = payment_sheet.worksheet('Sheet1')
    
    worksheet.update_cell(row, PROCESSED_COLUMN_REQUEST, "Processed")

def process_identification_request_email():
    """Identification Request 처리"""
    global last_requested_email, last_requested_timestamp
    latest_email, row = get_latest_email_from_sheet()
    if latest_email:
        if is_email_in_management_list(latest_email):
            subject = "[Timely Drop] Identification requested!"
            body = ("Hi! \n \n"
                    "We're delighted to be part of your online shopping journey. \n"
                    "To confirm your identity, kindly enter 01029901499 in the phone number field to get your verification code.\n\n"
                    "Thanks,\n"
                    "Timely Drop") 
        else:
            subject = "[Timely Drop] Identification request failed"
            body = ("We're sorry, \n"
                    "but Service Pass you are using is not valid. \n"
                    "Please purchase Timely Drop Service Pass first. \n\n"
                    "Service Pass purchase: https://tally.so/r/wA6Djz \n\n"
                    "If you've already purchased the Service Pass, confirming process for your payment is proceeding. \n"
                    "Please wait a bit, and you will get an email soon which tells you that Timely Drop service is available. \n\n"
                    "Sincerely,\n"
                    "Timely Drop")

        last_requested_email = latest_email
        last_requested_timestamp = datetime.now(pytz.utc)
        send_email(subject, body, latest_email)
        mark_email_as_processed(row)
        mark_request_as_processed(row)

def process_incoming_email():
    """수신된 이메일을 처리하여 재전송, identification requested 메일 도착 후 3분 이내 시점이면 ok"""
    global last_processed_email_id
    if last_requested_timestamp and (datetime.now(pytz.utc) - last_requested_timestamp <= timedelta(minutes=3)):
        received_body, received_timestamp, last_processed_email_id = fetch_latest_sent_email(last_processed_email_id)
        if received_body and received_timestamp:
            verification_subject = "[Timely Drop] Verification code arrived!"
            received_body_added = f"You're doing great! \n\n The message we received from the certification authority is as follows. \n Please read it carefully and make sure to input the correct verification code. \n\n Message: \n {received_body} \n\n Additionally, please read the manual carefully since the order may not proceed as expected if the manual is not followed precisely. \n\n If you have any questions or need further assistance, feel free to contact us via Chat in our Website. \n\n Sincerely, \n Timely Drop"
            send_email(verification_subject, received_body_added, last_requested_email)
    else:
        print("Process identification request email has not been executed in the last minute. Skipping the process.")

def mark_payment_as_processed(row):
    # worksheet 객체가 함수 외부에서 접근 가능하도록 수정
    try:
        payment_sheet = client.open_by_key('1jA52gS6N-I_8LrCmJnn2lXAZCE7NkGw95_HdLlEJJc8')
        worksheet = payment_sheet.worksheet('payment_list')
        worksheet.update_cell(row, PROCESSED_COLUMN_PAYMENT, "Processed")
    except gspread.exceptions.APIError as e:
        print(f"Google Sheets API 오류 in mark_payment_as_processed: {e}")
        send_email("[Error] mark_payment_as_processed에서 google sheets api 오류 발생","[Error] mark_payment_as_processed에서 google sheets api 오류 발생", "xx11chotae@gmail.com")
    except Exception as e:
        print(f"예상치 못한 오류 발생 in mark_email_as_processed: {e}")
        send_email("[Error] mark_payment_as_processed에서 예상치 못한 오류 발생","[Error] mark_payment_as_processed에서 예상치 못한 오류 발생", "xx11chotae@gmail.com")

def get_latest_payment_info():
    """스프레드시트 'payment_list'에서 최근 결제 정보를 가져옴"""
    try:
        payment_sheet = client.open_by_key('1jA52gS6N-I_8LrCmJnn2lXAZCE7NkGw95_HdLlEJJc8').worksheet('payment_list')
        column_f = payment_sheet.col_values(6)
        column_m = payment_sheet.col_values(9)
        column_d = payment_sheet.col_values(4)
        processed_column = payment_sheet.col_values(PROCESSED_COLUMN_PAYMENT)
    except gspread.exceptions.APIError as e:
        print(f"Google Sheets API 오류: {e}")
        send_email("[Error] get_latest_payment_info에서 google sheets api 오류 발생", 
                   "[Error] get_latest_payment_info에서 google sheets api 오류 발생", 
                   "xx11chotae@gmail.com")
        return None, None, None, None
    except Exception as e:
        print(f"예상치 못한 오류 발생 in get_latest_payment_info: {e}")
        send_email("[Error] get_latest_payment_info에서 예상치 못한 오류 발생", 
                   "[Error] get_latest_payment_info에서 예상치 못한 오류 발생", 
                   "xx11chotae@gmail.com")
        return None, None, None, None

    if not column_f or not column_m or not column_d:
        print("스프레드시트 열에서 데이터를 찾을 수 없습니다.")
        return None, None, None, None

    for index in reversed(range(len(column_d))):
        payment_method = column_f[index]
        total_amount_without_commission = column_m[index]
        payment_email = column_d[index]
        if payment_method and total_amount_without_commission and payment_email and \
           (len(processed_column) <= index or processed_column[index].strip() != "Processed"):
            if is_email_in_management_list(payment_email):
                return payment_method, total_amount_without_commission, payment_email, index + 1
            else:
                subject = "[Timely Drop] Inappropriate Email"
                body = ("We're sorry, but this email is not registered. \n"
                        "Please submit 'Payment Form' again using the email you had entered in 'Timely Drop Service Pass Purchase Form.' \n\n"
                        "Service Pass Purchase form: https://tally.so/r/w5LAyE \n\n"
                        "Sincerely,\n"
                        "Timely Drop")
                send_email(subject, body, payment_email)
                mark_payment_as_processed(index + 1)

    return None, None, None, None


def send_payment_email():
    """결제 요청 이메일 전송"""
    payment_method, total_amount_without_commission, payment_email, row = get_latest_payment_info()
    if payment_email:
        subject = "[Timely Drop] Complete your payment!"
        if payment_method == "Paypal":
            body = f"Thank you for using our service. \n\nWe would like to guide you on how to transfer the service fee to us via PayPal. Please follow the steps below to complete the transfer. \n\n 1. Please check the exact cost(KRW) of the items you purchased. \n 2. Please visit the following link and enter the KRW amount from step 1 accurately in ‘Receiving amount’ field. \n * Link : https://www.xoom.com/south-korea/send-money?locale=en-US \n 3. Once you enter the amount, the USD value in ‘Send amount’ field will automatically update. \n 4. Please transfer the updated USD amount to our PayPal account. \n * Timely Drop paypal account : timelydrop01@gmail.com \n\nIf you have any questions or need further assistance, feel free to contact us via Chat in our Website. We appreciate your business and look forward to serving you again! \n\nBest regards, \nTimely Drop"
        elif payment_method == "Wise":
            body = f"Thank you for using our service. \n\nWe would like to guide you on how to transfer the service fee to us via WISE. Please follow the steps below to complete the transfer. \n\n 1. Please check the exact cost(KRW) of the items you purchased. \n 2. Please visit the following link and enter the KRW amount from step 1 accurately in ‘Recipient gets’ field. \n * Link : https://wise.com/us/send-money/send-money-to-south-korea \n 3. Once you enter the amount, the USD value in ‘You send’ field will automatically update. \n 4. Please transfer the updated USD amount to our Wise account. \n <Wise account information> \n - Name: Timely Drop \n - Birthday: 2001.09.12 \n - Phone number: 010-8511-4979 \n - Bank account: Kakaobank 3333-31-4907631 \n - Email: paipaigks@naver.com \n\nIf you have any questions or need further assistance, feel free to contact us via Chat in our Website. We appreciate your business and look forward to serving you again! \n\nBest regards, \nTimely Drop"
        else:
            body = "결제 방법을 선택해주세요"
        send_email(subject, body, payment_email)
        send_email(subject, body, "timelydrop01@gmail.com")
        mark_payment_as_processed(row)

s = 1

def continuously_send_email():
    global s  # 함수 내에서 전역 변수를 사용
    if s > 29:  # s가 30 이상일 때 이메일 전송
        subject = "Server is working!!"
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        body = f"Server is working!!\nTimestamp: {current_time}"
        to_email = "xx11chotae@gmail.com"
        
        # 이메일 전송
        send_email(subject, body, to_email)
        
        # s를 다시 0으로 초기화
        s = 0
    else:
        # s 값 증가
        s = s + 1


# 메인 루프
while True:
    process_identification_request_email()
    process_incoming_email()
    send_payment_email()
    os.system("touch /tmp/keepalive")
    continuously_send_email()

    time.sleep(10)