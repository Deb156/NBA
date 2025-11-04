import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

def send_email_with_logo(to_email, subject, html_content):
    # Create message container
    msg = MIMEMultipart('related')
    msg['Subject'] = subject
    msg['From'] = 'your_email@company.com'
    msg['To'] = to_email
    
    # Create the HTML part
    msg_html = MIMEText(html_content, 'html')
    msg.attach(msg_html)
    
    # Attach the logo image
    with open('CLEAR logo.png', 'rb') as f:
        img_data = f.read()
    
    image = MIMEImage(img_data)
    image.add_header('Content-ID', '<clear_logo>')
    image.add_header('Content-Disposition', 'inline', filename='clear_logo.png')
    msg.attach(image)
    
    # Send email
    server = smtplib.SMTP('your_smtp_server.com', 587)
    server.starttls()
    server.login('your_email@company.com', 'your_password')
    server.send_message(msg)
    server.quit()