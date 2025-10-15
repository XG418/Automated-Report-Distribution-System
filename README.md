# 📨 Automated Report Distribution System

### Overview  
This is a small project I built to make my reporting workflow more efficient.  
It automatically updates scorecards with the latest metrics and then emails the finalized report to different departments using SendGrid.  

I originally built this at my company, where I often need to update and send reports to multiple teams.  
Before this, I had to open several Excel files, copy numbers, save new versions, attach them, and email everyone manually.  
Now, it all runs automatically in a few seconds.

---

### What It Does  
- **Updates Excel Scorecards** – Pulls data from source files (third-party website) and fills the correct cells automatically.  
- **Sends Updated Reports by Email** – Uses SendGrid’s SMTP to email reports to multiple recipients.    
- **Clear Status Messages** – Prints success or error messages in the terminal.  
- **Easy to Adjust** – You can change file paths, recipients, or email content easily.  

---

### How It Works  
1. **`Scorecard_Automation.py`**  
   - Reads data from different Excel reports  
   - Updates the main scorecard automatically using `openpyxl`  
   - Saves a new version (for example `Scorecard_Updated.xlsx`)  

2. **`Automated_Email_Report.py`**  
   - Sends the updated report to department leads  
   - Uses SendGrid SMTP and supports attachments of any type  

Together, these two scripts handle the full reporting workflow — from data update to delivery.

---

### How to Use  
1. Install required packages:
   
   pip install openpyxl pandas sendgrid python-dotenv
   
2. Update your file paths and email info inside the scripts:
   
  scorecard_path = "path/to/Scorecard.xlsx"
  source_file_path = "path/to/source_file.xlsx"
  sender_email = "your_verified_sendgrid_email"
  api_key = "your_sendgrid_api_key"

3. Run the scripts:
   
   Scorecard_Automation.py
   
   Automated_Email_Report.py
   

