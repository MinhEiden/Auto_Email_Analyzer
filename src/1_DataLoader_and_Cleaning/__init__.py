from data_loader import CsvLoader
import email
import pandas as pd
import re
from pathlib import Path

def clean_email_history(body):
    if not body:
        return ""
    split_patterns = (
        r"-----Original Message-----|"
        r"_{3,}|"                   
        r"-{3,}|"                    
        r"From:\s+|"                  
        r"Sent:\s+|"                 
        r"To:\s+|"                  
        r"Forwarded by\s+|"          
        r"On\s+.*wrote:"           
    )
    parts = re.split(split_patterns, body, flags=re.IGNORECASE)
    
    clean_text = parts[0].strip()
    return clean_text
    
def cleandata(raw_message):
    cleaned_data = []
    for message in raw_message:
        msg = email.message_from_string(message)
        
        email_data = {
            "Message-ID": msg.get("Message-ID"),
            "Date": msg.get("Date"),
            "From": msg.get("From"),
            "To": msg.get("To"),
            "Subject": msg.get("Subject")
        }
        body_text = ""
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    body_text = part.get_payload()
                    break 
        else:
            body_text = msg.get_payload()
    
        email_data["Body"] = clean_email_history(body_text)
        cleaned_data.append(email_data)

    df_cleaned = pd.DataFrame(cleaned_data)
    return df_cleaned

def save_cleaned_emails_to_excel(df_cleaned, filename="cleaned_emails.xlsx"):
    project_root = Path(__file__).resolve().parents[2]
    processed_dir = project_root / "data" / "processed"
    processed_dir.mkdir(parents=True, exist_ok=True)
    excel_path = processed_dir / filename
    df_cleaned.to_excel(excel_path, engine='openpyxl', index=False)
    print(f"đã clean data và log ra {excel_path}")
    return excel_path

if __name__ == "__main__":
    loader = CsvLoader(1500, 42)
    raw_series = loader.load_messages()
    df_final = cleandata(raw_series)
    save_cleaned_emails_to_excel(df_final)
        
        
        
    

