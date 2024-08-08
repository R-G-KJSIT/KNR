import streamlit as st
import requests
from docx import Document
from io import BytesIO
from datetime import datetime
from groq import Groq

# Constants
GROQ_API_KEY = st.secrets["GROQ_API_KEY"]  # Replace with your actual API key

# Initialize the Groq client
client = Groq(api_key=GROQ_API_KEY)

# Azure AD details for Microsoft Graph API
client_id = st.secrets["client_id"]
client_secret = st.secrets['client_secret']
scopes = ['wl.signin', 'wl.offline_access', 'onedrive.readwrite']
tenant_id = st.secrets["tenant_id"]
scope = "https://graph.microsoft.com/.default"

def get_access_token():
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    response = requests.post(token_url, data={
        'grant_type': 'client_credentials',
        'client_id': client_id,
        'client_secret': client_secret,
        'scope': scope
    })
    response.raise_for_status()  # Raise an exception for HTTP errors
    return response.json().get('access_token')

# Placeholder for Microsoft Graph API
GRAPH_API_BASE_URL = "https://graph.microsoft.com/v1.0/me/drive/root:/Documents/College Documents/S4DS/Knowhow Reports/"

# Function to call Groq API using the Groq module
def call_groq_api(points):
    prompt = ("I am about to give you a few points that sum up an event that takes place, "
              "your task is to take those points and make a more detailed paragraph and only print the paragraph and absolutely nothing else."
              "Keep in mind you need to report in 3rd person (i.e 'The event conducted today comprised of the following things 'yada yada yada. You don't need to explain any concept, just elaborate what was done)")
    
    # Create the message for the chat completion
    messages = [
        {
            "role": "user",
            "content": f"{prompt} Points: {', '.join(points)}"
        }
    ]
    
    # Request completion from Groq
    try:
        chat_completion = client.chat.completions.create(
            messages=messages,
            model="llama3-8b-8192"
        )
        return chat_completion.choices[0].message.content
    except Exception as e:
        st.error(f"Error calling Groq API: {e}")
        return ""

# Function to create DOCX with title
def create_docx(title, paragraph):
    doc = Document()
    doc.add_heading(title, level=1)  # Add the title as a heading
    doc.add_paragraph(paragraph)  # Add the generated paragraph
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# Function to upload DOCX to OneDrive
def upload_to_onedrive(file_buffer, filename):
    access_token = get_access_token()
    upload_url = f"{GRAPH_API_BASE_URL}{filename}:/content"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/octet-stream"
    }
    response = requests.put(upload_url, headers=headers, data=file_buffer)
    if response.status_code == 201:
        st.success("File uploaded successfully!")
    else:
        st.error(f"Error uploading file to OneDrive: {response.status_code} - {response.text}")

# Streamlit UI
st.title("Event Report Generator")

# Use session state to persist points
if 'points' not in st.session_state:
    st.session_state.points = []

# Date picker
event_date = st.date_input("Date of Event")

# Domain Selector
domain = st.selectbox(
    "Domain",
    ["AI-ML", "DataScience and Analytics", "Cloud Native", "Web Development", "AR/VR", "IoT", "Blockchain", "Cybersecurity"]
)

# Format the date and domain for the file name
formatted_date = event_date.strftime("%Y-%m-%d")
domain_slug = domain.replace(" ", "_")  # Replace spaces with underscores
filename = f"report_{formatted_date}_{domain_slug}.docx"
title = f"Knowhow {domain} workshop {formatted_date}"  # Title format

# Brief Points Input
st.subheader("Brief Points")
point_input = st.text_input("Enter a point", "")

if st.button("Add Point"):
    if point_input:
        st.session_state.points.append(point_input)
        st.rerun()  # Refresh the app to display updated points
    else:
        st.warning("Please enter a point before adding.")

# Display points
if st.session_state.points:
    for i, point in enumerate(st.session_state.points):
        st.text(f"Point {i + 1}: {point}")

# Generate Report Button
if st.button("Generate Report"):
    if not st.session_state.points:
        st.warning("Please add at least one point.")
    else:
        # Call Groq API
        paragraph = call_groq_api(st.session_state.points)
        if paragraph:
            # Create DOCX with title
            docx_buffer = create_docx(title, paragraph)
            
            # Create a download button
            st.download_button(
                label="Download Report",
                data=docx_buffer,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
            # Upload to OneDrive with dynamic filename
            # Since upload_to_onedrive needs a file buffer, we need to reset buffer's position
            docx_buffer.seek(0)
            upload_to_onedrive(docx_buffer, filename)
