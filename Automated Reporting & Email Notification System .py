import smtplib
from email.message import EmailMessage



def send_email(sender_email, api_key, recipient_email, subject, body, attachment_path):
    """
    Sends an email using SendGrid's SMTP.


    Args:
        sender_email (str): Email address of the sender.
        api_key (str): SendGrid API Key.
        recipient_email (str): Email address of the recipient.
        subject (str): Subject of the email.
        body (str): The content of the email.
        attachment_path (str, optional): File path of attachment. Default is None.
    """
    # Create the email content
    msg = EmailMessage()
    msg["From"] = sender_email
    msg["To"] = recipient_email
    msg["Subject"] = subject
    msg.set_content(body)


    if attachment_path:
        with open(attachment_path, "rb") as attachment_file: # rb 可以读任何类型的文件，图片·PDF·Excel·视频等
            msg.add_attachment(
                attachment_file.read(),
                maintype="application",
                subtype="octet-stream",
                filename=attachment_path.split("\\")[-1]
            )

    
    try:
        with smtplib.SMTP("smtp.sendgrid.net", 587) as server:  
            server.starttls()
            server.login("apikey", api_key)  
            server.send_message(msg)
        print("Email sent successfully!")
    except Exception as e:
        print(f"Failed to send email: {e}")


# Example usage:
if __name__ == "__main__":
    
    sender_email = sender_email  # Your SendGrid-verified email
    api_key = api_key  # SendGrid API Key
    recipient_email = your_email  # Recipient email
    subject = "Automation Test Email"
    body = "This is an automated email sent from Python using SendGrid!"

    attachment_path = "/Users/guoxinning/Desktop/简历/Submit/Resume DA.docx"

   

    send_email(sender_email, api_key, recipient_email, subject, body, attachment_path)
