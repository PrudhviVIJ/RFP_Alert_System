from email.message import EmailMessage
import os
import smtplib

def test_email():
    msg = EmailMessage()
    msg["Subject"] = "Pharma RFP Monitor - Email Test"
    msg["From"] = os.environ["ALERT_FROM"]
    recipients = [
        x.strip()
        for x in os.environ["ALERT_TO"].replace(";", ",").split(",")
        if x.strip()
    ]
    msg["To"] = ", ".join(recipients)
    msg.set_content("SMTP email delivery is working.")

    with smtplib.SMTP(
        os.environ["SMTP_HOST"],
        int(os.getenv("SMTP_PORT", "587")),
        timeout=25,
    ) as server:
        server.starttls()
        server.login(
            os.environ["SMTP_USERNAME"],
            os.environ["SMTP_PASSWORD"],
        )
        server.send_message(msg, to_addrs=recipients)

    print("Test email sent successfully.")

if __name__ == "__main__":
    test_email()