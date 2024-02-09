import os
import pandas as pd
import ssl
import smtplib
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from reportlab.lib.pagesizes import letter, landscape
from PyPDF2 import PdfReader, PdfWriter
from io import BytesIO
from reportlab.lib.pagesizes import landscape
from reportlab.pdfgen import canvas
import random


def send_email(receiver_email, subject, message, attachment_path=None):
    sender_email = "abc@gmail.com"
    sender_password = "1234"
    smtp_server = "smtp.gmail.com"
    smtp_port = 465

    msg = MIMEMultipart()
    msg["Subject"] = subject
    msg["From"] = sender_email
    msg["To"] = receiver_email
    msg.attach(MIMEText(message, "html"))

    if attachment_path:
        with open(attachment_path, "rb") as pdf_file:
            pdf_attachment = MIMEApplication(pdf_file.read(), _subtype="pdf")
            pdf_attachment.add_header(
                "Content-Disposition",
                "attachment",
                filename=os.path.basename(attachment_path),
            )
            msg.attach(pdf_attachment)

    context = ssl.create_default_context()
    server = smtplib.SMTP_SSL(smtp_server, smtp_port, context=context)
    server.login(sender_email, sender_password)
    server.sendmail(sender_email, receiver_email, msg.as_string())
    server.quit()
    
def update_log(log_file, data):
    with open(log_file, "a") as log:
        log.write(data + "\n")

def read_log(log_file):
    try:
        with open(log_file, "r") as log:
            data = log.read().splitlines()
        return data
    except FileNotFoundError:
        return []


excel_file_path = "./file.xlsx"
xls = pd.ExcelFile(excel_file_path)
sheet_names = xls.sheet_names
print(sheet_names)
certificate_template_path = (
    r"INTRA FAST CTF'24 Proposal.pdf"
)

output_directory = "./HACKWEEK_2024_CERTIFICATES"
i = 0
log_file = "sent_emails.txt"
error_file = "error_log.csv"
error_data = []
sent_emails_log = read_log(log_file)
os.makedirs(output_directory, exist_ok=True)
for sheet_name in sheet_names:
    df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
    for _, row in df.iterrows():
        member_name = row["OrganizationName"]
        # roll_number = row["ROLL NO"]
        email = row["Email"]
        if email in sent_emails_log:
            # print(f"Email already sent to {member_name}, {roll_number}, {email}. Skipping...")
            print(f"Email already sent to {member_name}, {email}. Skipping...")

            continue
        names = [f"{member_name}"]

        pdf_certificate_path = certificate_template_path

        subject = "Exclusive Sponsorship Invitation for INTRA FAST Capture the Flag Competition'24"
        html = f"""

    <html>
  <head>
    <title>Report</title>
    <meta name="viewport" content="width=device-width, initial-scale=1"> <!-- Added viewport meta tag -->

  </head>
  <body >
    <div >
      <div> <!-- Adjusted width -->

        <div >
          <p >
            <b>Dear {member_name} Team,</b>
          </p>
          <div class="container">
                <p>
I hope this message finds you well. I am writing to you as the Director of Marketing for the ACM Cyber Security-Chapter, the driving force behind the INTRA FAST Capture the Flag Competition, hosted by FAST NUCES Karachi. I am delighted to extend a unique sponsorship opportunity to <b>{member_name}</b> for this prestigious event, scheduled for 20th February 2024.<br><br>

FAST NUCES Karachi has a longstanding reputation for cultivating innovation and technological excellence. The INTRA FAST 'Capture the Flag Competition' exemplifies our dedication to pushing the boundaries of Cybersecurity. Recognizing <b>{member_name}</b> distinguished commitment to innovation and expertise in the tech industry, we are enthusiastic about the prospect of partnering with your esteemed company.<br><br>

As we gear up for what promises to be the most remarkable edition of the INTRA FAST 'Capture the Flag Competition', we believe that <b>{member_name}</b>, as a prominent industry player, is the ideal partner to contribute to the success of this landmark event. Aligning with us will not only provide your company exposure to a diverse audience but will also allow active participation in shaping the technological advancements showcased at this esteemed platform.<br><br>

For a comprehensive overview of the sponsorship opportunities available, please find attached our detailed sponsorship proposal. We are flexible and eager to discuss how we can tailor the sponsorship arrangement to align with <b>{member_name}'s</b> specific interests and objectives.<br><br>

Thank you for considering this exclusive partnership with the INTRA FAST Capture the Flag Competition' 24. We are excited about the prospect of <b>{member_name}</b> collaborating with us to make this edition a resounding success on 20th February 2024.<br><br>

Feel free to reach out to me directly for any additional information or to discuss this opportunity further. Thank you for considering our proposal.<br>

</p>

<br><br>
          </p></div>
          <p class="contain" style="font-size: 15px;">
            --<br>
            Best Regards,<br>
                Muhammad Osama Irfan, <br>
                Director Marketing, <br>
                ACM Cyber Security Chapter. <br>
          </p>
        </div>
      </div>
    </div>
  </body>
</html>"""
        try:
            send_email(email, subject, html, pdf_certificate_path)
            # message = member_name + ', ' + email + ', ' + roll_number
            message = member_name + ', ' + email

            print("Email sent for ", message)
            update_log(log_file, email)
            time.sleep(random.randint(1, 3))
        except Exception as e:
            error_data.append({
                "Member Name": member_name,
                # "Roll Number": roll_number,
                "Email": email,
            })
            error_df = pd.DataFrame(error_data)
            error_df.to_csv(error_file, mode='a', header=not os.path.exists(error_file), index=False)
            # print(f"Error sending email to {member_name}, {roll_number}, {email}: {e}")
            print(f"Error sending email to {member_name}, {email}: {e}")
