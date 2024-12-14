import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Email account credentials
sender_email = ""
password = ""  # Consider using environment variables for better security

# Load the Excel file
file_path = "Yoga Studio.xlsx"  # Update the path to your file
df = pd.read_excel(file_path, engine="openpyxl")

# Extract the email addresses
recipients = df["Email"].dropna().tolist()

# Email content
subject = "Professional Accounting and Financial Services for Your Business"
body = """\
Subject: Professional Bookkeeping and Tax Services for Your Yoga Studio

Dear Yoga Owner,

I hope this email finds you well.

My name is Muhammad Asim Zubair, and I am a professional and qualified accountant with over 7 years of experience in handling finance, bookkeeping, and accounting work for various companies. I specialize in providing comprehensive bookkeeping and tax services tailored to meet the unique needs of businesses like yours.

My expertise includes:
- Bookkeeping using QuickBooks, Xero, Zoho, and Manager.io
- Preparation of Financial Statements
- Monthly Accruals & Prepayments Recording
- Monthly/Yearly Bank Reconciliation
- Preparation & Disbursement of Salaries
- Managing & Reconciling Vendor Ledgers and Commission Payable Accounts
- Audit of Inventory
- Corporate Filings (Annual Return, Sales Tax Return, etc.)
- USA Taxation (Form 1040, Form 1040 NR, Form 1040-X, Form 4868)
- LLC Registration in the USA and UK

I have had the privilege of working with a renowned chartered accounting firm, gaining invaluable experience in audits, tax planning, and financial analysis. I have also helped startups secure over 5 million USD in investments through comprehensive business plans and financials.

I am confident that my skills and experience can provide significant value to your yoga studio, ensuring your financial records are accurate and compliant with all tax regulations. If you are interested in learning more about how I can assist your business, please feel free to contact me via email or phone.

Contact Information:
Email: asimzubair454@gmail.com
Phone: +92 3216839454
LinkedIn: https://www.linkedin.com/in/muhammad-asim-zubair-22a39a27b?utm_source=share&utm_campaign=share_via&utm_content=profile&utm_medium=android_app
Upwork: https://www.upwork.com/freelancers/~01d7938842c243453b

Looking forward to the opportunity to work with you.

Best regards,

Muhammad Asim Zubair
Chartered Accountant
"""


# Send email to each recipient
try:
    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()  # Secure the connection
        server.login(sender_email, password)

        for recipient in recipients:
            # Create the email for each recipient
            msg = MIMEMultipart()
            msg["From"] = sender_email
            msg["To"] = recipient
            msg["Subject"] = subject
            msg.attach(MIMEText(body, "plain"))

            # Send the email
            server.sendmail(sender_email, recipient, msg.as_string())
            print(f"Email sent to {recipient}")

except Exception as e:
    print(f"Error: {e}")
