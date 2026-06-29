import os
import json
import logging
import smtplib
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pandas as pd
from dotenv import load_dotenv

# Load environment variables from .env file if it exists
load_dotenv()

# Setup Logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

NOTIFICATION_LIST = [
    "230701026@rajalakshmi.edu.in",
    "sherind.official@gmail.com"
]

LOG_FILE = "email_sent_log.json"

def get_next_birthday(dob, today_date):
    """
    Returns the next birthday date object for the given DOB (datetime object) relative to today_date.
    Handles leap year birthdays (Feb 29). If it's not a leap year, it celebrates on Mar 1st.
    """
    try:
        # Try to replace year with current year
        bday_this_year = dob.replace(year=today_date.year)
    except ValueError:
        # Happens if dob is Feb 29 and current year is not a leap year
        # We'll celebrate on Mar 1st
        bday_this_year = dob.replace(year=today_date.year, month=3, day=1)
        
    if bday_this_year.date() < today_date.date():
        # Birthday already passed this year, next one is next year
        try:
            bday_next_year = dob.replace(year=today_date.year + 1)
        except ValueError:
            bday_next_year = dob.replace(year=today_date.year + 1, month=3, day=1)
        return bday_next_year.date()
        
    return bday_this_year.date()

def read_excel_data(filepath):
    """
    Reads the Excel file, cleans the data, handles duplicates and missing values.
    """
    try:
        df = pd.read_excel(filepath)
        
        # Standardize column names
        df.columns = [str(c).strip() for c in df.columns]
        if 'email' in df.columns:
            df.rename(columns={'email': 'Email'}, inplace=True)
            
        # Drop rows with missing Email or DOB
        df = df.dropna(subset=['Email', 'DOB'])
        
        # Clean emails: string type, lowercase, strip whitespace
        df['Email'] = df['Email'].astype(str).str.strip().str.lower()
        
        # Parse DOB (Handles mixed formats like MM/DD/YYYY and MM-DD-YYYY)
        df['DOB'] = pd.to_datetime(df['DOB'], errors='coerce', format='mixed')
        df = df.dropna(subset=['DOB'])
        
        logging.info(f"Successfully loaded {len(df)} valid employee records from {filepath}.")
        return df
    except Exception as e:
        logging.error(f"Error reading Excel file: {e}")
        return pd.DataFrame()

def send_email(subject, body, to_email=None, bcc_emails=None):
    """
    Sends an email using Gmail SMTP. 
    Can send to a single 'to_email' or multiple 'bcc_emails'.
    """
    sender_email = os.environ.get("EMAIL_ADDRESS")
    app_password = os.environ.get("APP_PASSWORD")

    if not sender_email or not app_password:
        logging.error("SMTP credentials (EMAIL_ADDRESS, APP_PASSWORD) not found in environment variables.")
        return False

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['Subject'] = subject

    if to_email:
        msg['To'] = to_email
        
    msg.attach(MIMEText(body, 'html'))

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, app_password)
        
        # Collect all recipients for the sendmail function
        recipients = []
        if to_email:
            recipients.append(to_email)
        if bcc_emails:
            recipients.extend(bcc_emails)
            
        if not recipients:
            logging.warning("No recipients specified. Skipping email.")
            return False

        server.sendmail(sender_email, recipients, msg.as_string())
        server.quit()
        logging.info(f"Email sent successfully: '{subject}' to {to_email} | bcc: {bcc_emails}")
        return True
    except Exception as e:
        logging.error(f"Failed to send email to {to_email} | bcc: {bcc_emails}. Error: {e}")
        return False

def load_sent_log():
    """
    Loads the state file to track which emails were sent today.
    Helps ensure idempotency.
    """
    if os.path.exists(LOG_FILE):
        try:
            with open(LOG_FILE, 'r') as f:
                return json.load(f)
        except Exception as e:
            logging.error(f"Error reading log file: {e}")
            return {}
    return {}

def save_sent_log(log_data):
    """
    Saves the state file.
    """
    try:
        with open(LOG_FILE, 'w') as f:
            json.dump(log_data, f, indent=4)
    except Exception as e:
        logging.error(f"Error writing to log file: {e}")

def process_wishes(df, today, sent_log):
    """
    Identifies whose birthday is today and sends them a wish email.
    """
    today_str = today.strftime("%Y-%m-%d")
    if today_str not in sent_log:
        sent_log[today_str] = []
        
    logging.info("--- Processing Birthday Wishes ---")
    for index, row in df.iterrows():
        bday = get_next_birthday(row['DOB'], today)
        days_until = (bday - today.date()).days
        logging.info(f"Evaluating {row['Name']} (DOB: {row['DOB'].date()}) -> Next Birthday: {bday} ({days_until} days away)")
        
        if bday == today.date():
            email = row['Email']
            name = row['Name']
            log_key = f"wish:{email}"
            
            if log_key not in sent_log[today_str]:
                subject = f"Happy Birthday, {name}!"
                body = f"<h3>Happy Birthday, {name}!</h3><p>Wishing you a fantastic day filled with joy and celebration.</p><p>Best regards,<br>The Team</p>"
                
                if send_email(subject, body, to_email=email):
                    sent_log[today_str].append(log_key)
            else:
                logging.info(f"Birthday wish already sent to {name} ({email}) today.")

def process_reminders(df, today, sent_log):
    """
    Sends reminders 1 day and 2 days before a birthday to the notification list.
    Excludes the birthday person from the reminder.
    """
    today_str = today.strftime("%Y-%m-%d")
    if today_str not in sent_log:
        sent_log[today_str] = []

    logging.info("--- Processing Birthday Reminders ---")
    for index, row in df.iterrows():
        bday = get_next_birthday(row['DOB'], today)
        days_until = (bday - today.date()).days
        
        if days_until in [1, 2]:
            email = row['Email']
            name = row['Name']
            log_key = f"reminder_{days_until}d:{email}"
            
            if log_key not in sent_log[today_str]:
                # Prepare notification list, excluding the birthday person
                recipients = [recip.lower().strip() for recip in NOTIFICATION_LIST if recip.lower().strip() != email]
                
                if not recipients:
                    logging.info(f"No reminder recipients for {name}'s birthday after exclusion.")
                    sent_log[today_str].append(log_key)
                    continue

                when_str = "Tomorrow" if days_until == 1 else "in 2 days"
                subject = f"Birthday Reminder: {name}'s birthday is {when_str.lower()}"
                body = f"<p>Just a quick reminder that <b>{name}'s</b> birthday is {when_str.lower()} ({bday.strftime('%B %d')}).</p>"
                
                # Send reminder using BCC
                if send_email(subject, body, to_email=None, bcc_emails=recipients):
                    sent_log[today_str].append(log_key)
            else:
                logging.info(f"Reminder ({days_until}d) already sent for {name}'s birthday today.")

def main():
    logging.info("Starting Birthday Email Automation System")
    
    # Load Excel data
    df = read_excel_data("bdays.xlsx")
    if df.empty:
        logging.warning("No data found in bdays.xlsx or file missing. Exiting.")
        return

    today = datetime.now(ZoneInfo("Asia/Kolkata"))
    logging.info(f"Current System Date/Time (IST): {today}")
    
    # Load state
    sent_log = load_sent_log()
    
    # Process
    process_wishes(df, today, sent_log)
    process_reminders(df, today, sent_log)
    
    # Save state
    save_sent_log(sent_log)
    logging.info("Completed execution")

if __name__ == "__main__":
    main()
