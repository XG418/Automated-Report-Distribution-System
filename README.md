# Automated-Reporting-Email-Notification-System

## Overview
This is a small automation project I built to make my reporting process more efficient.
The script automatically sends reports (Excel, PDF, or Word) through SendGrid’s SMTP server, so I don’t need to manually attach or send emails anymore.

I built this because my position has a lot of reports, and I need to send to a bunch of people the lastest version, which cost lots of time

## What It Does

- Automatically sends emails through SendGrid SMTP

- Supports any attachment type (Excel, PDF, Word, image, etc.)

- Prints clear success/error messages in the terminal

- Easy to modify for different recipients or message content

- Can be integrated into a larger data pipeline (like Databricks or Airflow jobs)

## How to Use

- Install Python's libraries
  - smtplib
  - email.message

- Replace the following fields in the code with your own:

   - sender_email = "your_verified_sendgrid_email"
   - api_key = "your_sendgrid_api_key"
   - recipient_email = "recipient@example.com"


- Add your attachment path:

   - attachment_path = "/Users/guoxinning/Desktop/report.xlsx"


- Run:
   - You’ll see “✅ Email sent successfully!” if it worked.
