import datetime as dt
import smtplib, ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

from gmail_credentials import *

# ___ Files and variables ___ #
EMAIL_NUMBER = 'email_number.txt'
LIST_EMAILS = 'TAP_Email_Addresses.txt'
# LIST_EMAILS = 'emails_test.txt'
LETTER_FILE_PATH = "TAP_Portugal_Explanation_Letter.pdf"


# ___ Get list of recipients ___ #
with open(LIST_EMAILS) as file:
    email_addresses = file.read().split("\n")

# ___ Calculate days between cancellation and now ___ #
today = dt.datetime.now()
cancellation_date = dt.datetime(2020, 3, 26)
total_days_difference = (today - cancellation_date).days
years = total_days_difference // 365
months = (total_days_difference - years * 365) // 30
days = (total_days_difference - years * 365 - months * 30)
duration = f"{total_days_difference} days (or {years} years, {months} months and {days} days)"
print(duration)

# ___ Get email number ___ #
def make_ordinal(n):
    """
    Convert an integer into its ordinal representation:
        make_ordinal(0)   => '0th'
        make_ordinal(3)   => '3rd'
        make_ordinal(122) => '122nd'
        make_ordinal(213) => '213th'
    """
    n = int(n)
    if 11 <= (n % 100) <= 13:
        suffix = 'th'
    else:
        suffix = ['th', 'st', 'nd', 'rd', 'th'][min(n % 10, 4)]
    return str(n) + suffix


with open(EMAIL_NUMBER) as file:
    integer_email_number = int(file.read())
    ordinal_email_number = make_ordinal(integer_email_number)
    print(ordinal_email_number)


# ___ Email settings and content ___ #
# port = 465  # For starttls
# context = ssl.create_default_context()
smtp_server = "smtp.gmail.com"

# ___ Loop over recipients and send email ___ #
# print(email_addresses)
for email in email_addresses:
    recipient = email
    print(recipient)

    # message_content = """\
    # Subject: this is a test from Python
    # This message is sent from Python.
    # """
    message_content = f"Hello,\n\n" \
                      f"I am reaching out to you today as I have been waiting for a refund from TAP for over 2 years now. I had booked flights with TAP from Toronto to Paris for April 2020 (reference XXXXXX). Due to Covid, the flights were cancelled on March 26, 2020, and I was offered a voucher as compensation. I have been trying for several months to get this voucher turned into a refund as I’m entitled to do. I am now seeking your help to accelerate this process. I’m awaiting a $XXXXX (around XXXXX€) refund. This is now my {ordinal_email_number} email regarding this matter.\n\n" \
                      f"As much as I understand how airlines have been affected by this pandemic, passengers have also very much been impacted and, in my case, $XXXXX represents a large amount of money. As of today, it has been {duration} since the flights were cancelled!\n\n" \
                      f"Attached is a document presenting the entire process I have gone through up until now. Could you please point me in the direction of the right person to help me with this matter?\n\n" \
                      f"Thank you,\n\n" \
                      f"Pierre Goaer\n" \
                      f"xxxxxxxxxxx@gmail.com\n"
    message = MIMEMultipart()
    message['Subject'] = "Refund status"
    message['From'] = f"Pierre Goaer <{GMAIL_SENDER}>"
    message['To'] = recipient
    message.attach(MIMEText(message_content))

    # ___ Add attachment ___ #
    part = MIMEBase('application', "octet-stream")
    with open(LETTER_FILE_PATH, 'rb') as file:
        part.set_payload(file.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition',
                    'attachment; filename={}'.format(Path(LETTER_FILE_PATH).name))
    message.attach(part)

    # ___ Create connection and send email ___ #
    try:
        with smtplib.SMTP(smtp_server) as server:
            server.starttls()
            server.login(GMAIL_EMAIL, GMAIL_PASSWORD)
            server.sendmail(from_addr=GMAIL_SENDER,
                            to_addrs=[recipient],
                            msg=message.as_string())
        print(f"Successfully sent to {recipient}.")
    except smtplib.SMTPRecipientsRefused:
        pass

# ___ Update email_number.text with new email count ___ #
with open(EMAIL_NUMBER, mode="w") as file:
    new_count = integer_email_number + 1
    file.write(f"{new_count}")
