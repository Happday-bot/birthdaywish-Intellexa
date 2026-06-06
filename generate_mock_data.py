import pandas as pd
from datetime import datetime, timedelta

def generate_sample_data():
    # Set birthdays relative to today for testing
    today = datetime.now()
    bday_today = today.strftime("%m/%d/%Y")
    bday_tomorrow = (today + timedelta(days=1)).strftime("%m/%d/%Y")
    bday_day_after = (today + timedelta(days=2)).strftime("%m/%d/%Y")

    data = {
        "Name": ["John Doe", "Alice Smith", "Bob Jones", "Eve Adams"],
        "DOB": [bday_today, bday_tomorrow, bday_day_after, "02/29/2000"],
        "Email": ["john@example.com", "alice@example.com", "bob@example.com", "eve@example.com"]
    }
    
    df = pd.DataFrame(data)
    df.to_excel("bdays.xlsx", index=False)
    print("bdays.xlsx generated successfully.")

if __name__ == "__main__":
    generate_sample_data()
