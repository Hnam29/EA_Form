import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import re
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import logging

# Set up logging with a timestamp for each log entry
logging.basicConfig(
    filename='form.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

# Log helper function
def log_event(message):
    logging.info(message)
    
# Google Sheets connection setup
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('/Users/vuhainam/Documents/PROJECT_DA/EdtechAgency/Database/.streamlit/ea-database-form-fbf45c7169a9.json', scope)
client = gspread.authorize(creds)

# Create a new worksheet with the current date
def create_google_sheet():
    today_date = datetime.now().strftime("%Y-%m-%d")
    spreadsheet = client.open_by_url('https://docs.google.com/spreadsheets/d/1YdblYk8ovrtLmbkGBJAtdNoAXqYXKILGLJH9GTvbtpE')
    worksheet_title = f'Form_{today_date}'

    # Create the worksheet if it doesn't exist
    try:
        worksheet = spreadsheet.add_worksheet(title=worksheet_title, rows="100", cols="6")
        # Set header row
        worksheet.append_row(["Name", "Company", "Role", "PhoneNo", "Email", "Sentiment"])
    except gspread.exceptions.WorksheetExists:
        worksheet = spreadsheet.worksheet(worksheet_title)  # Get the existing worksheet

    return worksheet

# Insert user info into the Google Sheets
def add_info(data):
    worksheet = create_google_sheet()
    # Append user data
    worksheet.append_row([data['name'], data['company'], data['role'], data['phoneNo'], data['email'], data['sentiment']])

# Phone number cleaning function
def clean_phone_numbers(phone):
    if not isinstance(phone, str):
        return phone
    if re.match(r'(\d{4} \d{3} \d{3}|\d{4} \d{3} \d{4})( - (\d{4} \d{3} \d{3}|\d{4} \d{3} \d{4}))*', phone):
        return phone
    phone = phone.replace('|', '-')
    phone = re.sub(r'[.\s()]', '', phone)
    if phone.startswith('+84'):
        phone = '0' + phone[3:]
    elif phone.startswith('84'):
        phone = '0' + phone[2:]
    if '/' in phone:
        parts = phone.split('/')
        cleaned_parts = [clean_phone_numbers(part.strip().replace('-', '')) for part in parts]
        return ' - '.join(cleaned_parts)
    phone = phone.replace('-', '')
    match = re.match(r'(\d+)', phone)
    if match:
        digits = match.group(1)
        if not digits.startswith('0'):
            digits = '0' + digits
        if len(digits) == 10:
            return f"{digits[:4]} {digits[4:7]} {digits[7:]}"
        elif len(digits) == 11:
            return f"{digits[:4]} {digits[4:7]} {digits[7:]}"
    return phone

# Email typo correction dictionary and pattern
corrections = {
    '@domain': '.com', 'gmailcom': 'gmail.com', '.cm': '.com', 'gamil.com': 'gmail.com',
    'yahoocom': 'yahoo.com', 'gmai.com': 'gmail.com', 'yahoooo.com': 'yahoo.com',
    'gmal': 'gmail', 'hanoieduvn': 'hanoiedu.vn', 'tayho,edu,vn': 'tayho.edu.vn',
    'gmaill.com': 'gmail.com', 'gmil.com': 'gmail.com', 'yahô.com': 'yahoo.com',
    'yanhoo.com': 'yahoo.com', 'gmailk': 'gmail', 'gmail..com': 'gmail.com',
    'hanoiưdu': 'hanoiedu', 'gmaill.com': 'gmail.com', 'nocomment.con': 'nocomment.com',
    'gmaill@com': 'gmail.com', 'yahoocomvn': 'yahoo.com.vn', 'gmail. Com': 'gmail.com',
    'gmail,com': 'gmail.com', '@.gmail': '@gmail', '"gmail': '@gmail', 'Gmail': 'gmail',
    '.con': '.com', '.co': '.com'
}
email_pattern = re.compile(r'[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+')

def fix_common_typos(email):
    if email is None:
        return None
    email = email.strip().strip('-').replace(' ', '')  # Strip and remove spaces

    # Replace general patterns first
    for wrong, right in corrections.items():
        email = email.replace(wrong, right)

    # Use regex to validate email structure
    match = email_pattern.search(email)
    if match:
        email = match.group(0)

    # Handling multiple '@' cases
    if email.count('@') > 1:
        parts = email.split('@')
        email = parts[0] + '@' + ''.join(parts[1:])
        
    # Lowercase the email
    if email.isupper():
        email = email.lower()
        
    # Remove trailing dot if exists
    if email.endswith('.'):
        email = email[:-1]
        
    # Explicitly handle ending with 'gmail'
    if email.endswith('gmail'):
        email = email.replace('gmail', 'gmail.com')

    # Additional check for a final '.com' adjustment
    if email.endswith('.com'):
        return email
    elif email.endswith('.comm'):
        return email[:-1]  # Remove the extra 'm'

    return email

def validate_email(email):
    if email is None or not isinstance(email, str):
        return None, False
    email = fix_common_typos(email)
    if re.match(email_pattern, email):
        return email, True
    return email, False

def clean_emails(email_cell):
    if isinstance(email_cell, str):
        emails = [email.strip() for email in email_cell.split(' - ') if email.strip()]
        cleaned_emails = []
        is_valid = True
        for email in emails:
            cleaned_email, valid = validate_email(email)
            if not valid:
                is_valid = False
            cleaned_emails.append(cleaned_email)
        return ' - '.join(cleaned_emails), is_valid
    else:
        return None, False

# Main validation function
def validate_data(data):
    errors = []
    # Name validation
    if "name" in data and not re.match(r'^[\w\sÀÁẢÃẠÂẤẦẨẪẬĂẮẰẲẴẶÈÉẺẼẸÊẾỀỂỄỆÌÍỈĨỊÒÓỎÕỌÔỐỒỔỖỘƠỚỜỞỠỢÙÚỦŨỤƯỨỪỬỮỰÝỲÝỶỸỴ]+$', data["name"]):
        errors.append("Tên riêng không thể bao gồm số.")
    # Company validation
    if "company" in data and not data["company"].replace(" ", "").isalnum():
        errors.append("Tên công ty đang chứa ký hiệu đặc biệt.")
    # Phone number validation
    if "phoneNo" in data:
        cleaned_phone = clean_phone_numbers(data["phoneNo"])
        if cleaned_phone != data["phoneNo"]:
            data["phoneNo"] = cleaned_phone
        if len(re.sub(r'\D', '', cleaned_phone)) not in [10, 11]:
            errors.append("Số điện thoại chưa hợp lệ.")
    # Email validation
    if "email" in data:
        cleaned_email, is_valid = clean_emails(data["email"])
        data["email"] = cleaned_email
        if not is_valid:
            errors.append("Địa chỉ email chưa hợp lệ.")
    return errors


def send_confirmation_email(user_name, user_email):
    # Email configuration
    sender_email = "nam.vu@edtechagency.net"
    sender_password = "HnAm2002#@!" 
    subject = "[This is an auto email, no-reply] Confirmation of your submission"
    body = f"Hi {user_name}. EA wants to say a huge dupe cute thank you for your submission! We appreciate your interest. More information will be provided shortly."

    # Create message
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = user_email
    msg['Subject'] = subject

    # Attach body text
    msg.attach(MIMEText(body, 'plain'))

    try:
        # Set up the server
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()  # Upgrade to a secure connection
        server.login(sender_email, sender_password)

        # Send the email
        server.send_message(msg)
        log_event(f"Email sent successfully to {user_name}!")

    except Exception as e:
        log_event(f"Failed to send email to {user_name}: {e}")

    finally:
        server.quit()

# Form Creation
def form_creation():
    data = {}
    col1, col2 = st.columns([8, 2])
    col1.header('Mời Anh/Chị điền vào thông tin dưới đây để đăng ký tham gia sự kiện')
    col2.image("/Users/vuhainam/Documents/PROJECT_DA/EdtechAgency/Database/logo.png", caption="")
    
    with st.form(key='Registration Form'):
        name = st.text_input('Họ và tên của Anh/Chị: ')
        company = st.text_input('Tên doanh nghiệp của Anh/Chị: ')
        role = st.selectbox('Chức vụ của Anh/Chị: ', options=["C-level", "M-level", "E-level"], index=None)
        phoneNo = st.text_input('Số điện thoại của Anh/Chị: ')
        email = st.text_input('Địa chỉ email của Anh/Chị: ')
        sentiment = st.slider("Rate your experience:", 1, 5, 1, format="%d ⭐")
        submit = st.form_submit_button(label='Register')

        if submit:
            data['name'] = name
            data['company'] = company
            data['role'] = role
            data['phoneNo'] = phoneNo
            data['email'] = email
            data['sentiment'] = sentiment

            # Validate input data
            if not (name and company and role and phoneNo and email and sentiment):
                st.warning('Anh/Chị vui lòng nhập đầy đủ các trường thông tin. Xin cảm ơn!')
            else:
                # Validate the input data
                errors = validate_data(data)  
                if errors:
                    for error in errors:
                        st.error(error)
                else:
                    add_info(data)
                    st.success('Chúc mừng Anh/Chị đã đăng ký thành công.')
                    st.balloons()
                    st.markdown('Anh/Chị vui lòng kiểm tra email để nhận những thông tin cập nhật từ ban tổ chức.')

form_creation()