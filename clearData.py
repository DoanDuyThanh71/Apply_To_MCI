import pandas as pd 
import re

file_path = 'order.xlsx'
df = pd.read_excel(file_path)

missing_emails = df['cus_mail'].isnull().sum()

duplicate_emails = df['cus_mail'].duplicated().sum()

def is_valid_email(email):
    regex = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(regex, email) is not None

invalid_emails = df['cus_mail'].apply(lambda x: not is_valid_email(x) if pd.notnull(x) else False).sum()

missing_emails, duplicate_emails, invalid_emails
