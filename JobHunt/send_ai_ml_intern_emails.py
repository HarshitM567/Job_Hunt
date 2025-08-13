import os
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# ==== CONFIGURATION ====
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS")  # Set in .env file
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")  # Set in .env file

# Load HR data
df = pd.read_excel("hr_contacts.xlsx")  # Rename your file as needed

# Email template
subject = "Application for AI/ML Intern Role – Harshit Mittal"
body_template = """Dear Hiring Manager

I hope this message finds you well.

I am writing to express my keen interest in the AI/ML Intern position
at your organization. As a PG student specializing in Artificial
Intelligence and Machine Learning at IIT Guwahati, I have built and
deployed several impactful data-driven solutions. One of my key
projects, VoiceGPT, is a voice-based conversational system that
combines RAG pipelines and LLMs for intelligent human-computer
interaction.

Beyond that, I’ve worked on customer churn prediction systems,
resource allocation analytics for refugee camps, and NLP-powered
search engines—showcasing my ability to develop AI systems that solve
real-world problems.

Your company's focus on AI innovation to drive business impact strongly
aligns with my passion for applying machine learning to create
meaningful change. I bring hands-on expertise in Python, ML
frameworks, and cloud-based AI workflows, and I am excited about the
opportunity to learn and contribute to your team’s efforts.

Please find my resume attached for your consideration. I would be
thrilled to discuss how my background and skills could be of value to
the organization.

Thank you for your time and consideration.

Best regards,
Harshit Mittal
www.linkedin.com/in/harshit405
github.com/HarshitM567

"""

# Load resume
with open("Harshit_Mittal_Resume.pdf", "rb") as f:
    resume_data = f.read()

# Send emails
with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
    server.starttls()
    server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)

    for index, row in df.iterrows():
        name = row.get("Name", "HR")
        recipient = row.get("Email")

        if pd.isna(recipient):
            continue

        msg = MIMEMultipart()
        msg["From"] = EMAIL_ADDRESS
        msg["To"] = recipient
        msg["Subject"] = subject

        # Email body
        body = body_template.format(name=name)
        msg.attach(MIMEText(body, "plain"))

        # Attach resume
        part = MIMEApplication(resume_data, Name="Harshit_Mittal_Resume.pdf")
        part['Content-Disposition'] = 'attachment; filename="Harshit_Mittal_Resume.pdf"'
        msg.attach(part)

        try:
            server.send_message(msg)
            print(f"Email sent to {recipient}")
        except Exception as e:
            print(f"Failed to send to {recipient}: {e}")