from flask import Flask, request, jsonify
import os
import smtplib
from email_validator import validate_email, EmailNotValidError
from email.message import EmailMessage
from dotenv import load_dotenv
from docx import Document

# Load environment variables
load_dotenv()

app = Flask(__name__)

# Email credentials and configuration
EMAIL_USER = os.getenv('EMAIL_USER')
EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD')
RESUME_PATH = 'resume.pdf'
SIWES_TEMPLATE_PATH = 'templates/siwes_template.docx'

# Function to send email
def send_email(to_email, subject, body, attachments):
    msg = EmailMessage()
    msg['From'] = EMAIL_USER
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.set_content(body)

    # Attach files
    for file_path in attachments:
        with open(file_path, 'rb') as f:
            file_data = f.read()
            file_name = os.path.basename(f.name)
            msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)
    
    # Send email
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_USER, EMAIL_PASSWORD)
        smtp.send_message(msg)

# Function to create customized SIWES letter
def create_custom_siwes_letter(company_address, output_path):
    doc = Document(SIWES_TEMPLATE_PATH)
    
    # Replace placeholder lines with the company address
    address_lines = company_address.split(', ')
    i = 0
    for paragraph in doc.paragraphs:
        if '_______________________' in paragraph.text and i < len(address_lines):
            paragraph.text = paragraph.text.replace('_______________________', address_lines[i], 1)
            i += 1
        if i == len(address_lines):
            break
    
    doc.save(output_path)

# Endpoint to send internship applications
@app.route('/send_applications', methods=['POST'])
def send_applications():
    data = request.json
    applications = data.get('applications', [])
    subject = "SIWES Internship Application"
    
    # Load email template
    with open('templates/email_template.txt', 'r') as file:
        email_body = file.read()

    for application in applications:
        email = application['email']
        company_address = application['address']
        custom_siwes_path = f'siwes_letter_{email}.docx'

        try:
            # Validate email
            validate_email(email)
            
            # Create custom SIWES letter
            create_custom_siwes_letter(company_address, custom_siwes_path)
            
            # Send email
            send_email(email, subject, email_body, [RESUME_PATH, custom_siwes_path])
            print(f"Email sent to {email}")

            # Clean up custom SIWES letter
            os.remove(custom_siwes_path)
        except EmailNotValidError as e:
            print(f"Invalid email {email}: {e}")
        except Exception as e:
            print(f"Failed to send email to {email}: {e}")

    return jsonify({"message": "Emails sent"}), 200

if __name__ == '__main__':
    app.run(debug=True)
