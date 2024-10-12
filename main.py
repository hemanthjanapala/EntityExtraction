import os
import streamlit as st
import openai
import fitz  # PyMuPDF for PDF
from PIL import Image
import base64
import json
import requests
import io
import re
import comtypes.client  # For PPTX to PDF conversion
import win32com.client as win32  # For XLSX to PDF conversion
import pythoncom  # For COM initialization
from pythoncom import com_error  # Correct import for COM error
import uuid  # For generating UUIDs

DEFAULT_PROMPT = """
Evaluate the image provided. Your response should be precise and strictly adhere to the instructions below:

1. Entity_Name: Extract every and each entity present in the provided diagram.

3. Relationships: Extract entities with parent-child relationships, and share holding between parent and child entity shown over relationship, if available in JSON format.

 

[ 
Example format:
{{
"Entity_ID: "A unique identifier for the extracted entity",
"Entity_Name": "A verbatim copy of the entity name as it appear on the diagram",
"Entity_Type": "For example, GP, or Fund, or LP, or Holding, etc.",
"Location": "A location specifying the jurisdiction of the entity formation if present"



"Relationships": [
{
  "parent": {
    "ID": "123456ABC",
    "name": "Elpam Asia Ltd.",
  }"child": "123478ABC",
    "name": "CPV 88 Ltd",
  "share_percent": {
    "Series C": 2733035"Ordinary": 2522296"Ordinary (Guaranteed)": 5255331
  }
}
]
}}

Relevancy Score: How confident are you of the diagram analysis as well as entity and relationships extraction.
]
"""


# Set your OpenAI API key here
openai.api_key = os.getenv(key)
MODEL_NAME = "key"
GPT4O_API_KEY = "key"

# Function to encode the image
def encode_image(image):
    buffer = io.BytesIO()
    image.save(buffer, format=image.format)  # Save image with its original format
    return base64.b64encode(buffer.getvalue()).decode('utf-8')

def analyze_image_with_gpt4o(image, prompt):
    # Convert image to binary format
    base64_image = encode_image(image)

    headers = {
        "Content-Type": "application/json",
        "api-key": GPT4O_API_KEY,
    }

    payload = {
        "messages": [
            {"role": "system", "content": "You are a business analyst working on CRM (Customer Relationship Management) systems to input and manage data related to corporate shareholding structures and have strong expertise in Corporate Structure Knowledge, Data Integrity and Validation, Reporting and Insights. Your goal is to analyse Corporate Shareholder Diagrams, extract entities and their respective relationships, as well as the shares that below to each one of the entities."},
            {"role": "user", "content": [
                {
                    "type": "text",
                    "text": prompt
                },
                {
                    "type": "image_url",
                    "image_url": {
                        "url": f"data:image/{image.format.lower()};base64,{base64_image}"
                    }
                }
            ]},
        ],
        "temperature": 0.2,
        "top_p": 0.95,
        "max_tokens": 4096,
        "response_format": {"type": "json_object"}
    }

    ENDPOINT = "https://relextraction.openai.azure.com/openai/deployments/gpt4o/chat/completions?api-version=2024-02-15-preview"
    
    try:
        response = requests.post(ENDPOINT, headers=headers, json=payload)
        response.raise_for_status()  # Raise an error for bad responses
        
        response_json = response.json()
        raw_content = response_json['choices'][0]['message']['content']
        analysis_result = json.loads(raw_content)
        
       
        return analysis_result
    
    except requests.RequestException as e:
        st.error(f"Request failed: {e}")
    except json.JSONDecodeError as e:
        st.error(f"Failed to parse JSON: {e}")
        st.write("Raw response content for debugging:", response.text)
    except KeyError as e:
        st.error(f"Unexpected response structure: {e}")
        st.write("Raw response content for debugging:", response.json())

    return None

def convert_pdf_to_images(pdf_stream):
    images = []
    pdf_document = fitz.open(stream=pdf_stream, filetype="pdf")
    for page_number in range(len(pdf_document)):
        page = pdf_document.load_page(page_number)
        pix = page.get_pixmap()
        img = Image.open(io.BytesIO(pix.tobytes()))
        images.append(img)
    pdf_document.close()
    return images

def extract_text_from_pdf(pdf_stream):
    text = ""
    pdf_document = fitz.open(stream=pdf_stream, filetype="pdf")
    for page_number in range(len(pdf_document)):
        page = pdf_document.load_page(page_number)
        text += page.get_text()
    pdf_document.close()
    return text

# Function to convert PPTX to PDF
def convert_pptx_to_pdf(pptx_stream, pdf_output_path):
    pythoncom.CoInitialize()  # Initialize COM
    
    with open("temp_presentation.pptx", "wb") as temp_pptx:
        temp_pptx.write(pptx_stream.getvalue())

    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    powerpoint.Visible = 1

    try:
        presentation = powerpoint.Presentations.Open(os.path.abspath("temp_presentation.pptx"))
        presentation.SaveAs(pdf_output_path, 32)  # 32 is the format type for PDF
    except com_error as e:
        st.error(f"COM Error: {e}")
    except Exception as e:
        st.error(f"Failed to convert PPTX to PDF: {e}")
    finally:
        presentation.Close()
        powerpoint.Quit()
        pythoncom.CoUninitialize()  # Uninitialize COM

# Function to convert XLSX to PDF
def convert_xlsx_to_pdf(xlsx_stream, pdf_output_path):
    pythoncom.CoInitialize()  # Initialize COM

    with open("temp_workbook.xlsx", "wb") as temp_xlsx:
        temp_xlsx.write(xlsx_stream.getvalue())

    excel = win32.Dispatch('Excel.Application')
    excel.Visible = 0

    try:
        workbook = excel.Workbooks.Open(os.path.abspath("temp_workbook.xlsx"))
        workbook.ExportAsFixedFormat(0, pdf_output_path)
    except com_error as e:
        st.error(f"COM Error: {e}")
    except Exception as e:
        st.error(f"Failed to convert XLSX to PDF: {e}")
    finally:
        workbook.Close(False)
        excel.Application.Quit()
        pythoncom.CoUninitialize()  # Uninitialize COM

def main():
    st.title("Document Analysis with GPT-4o")

    # Add a side panel for file upload and prompt input
    st.sidebar.header("File Upload and Prompt")

    # Sidebar for file upload (PDF, PPTX, XLSX, and Image)
    uploaded_file = st.sidebar.file_uploader("Choose a file", type=["pdf", "pptx", "xlsx", "jpeg", "jpg", "png"])

    # Sidebar for entering the custom prompt
    prompt = st.sidebar.text_area("Enter a prompt to analyze the document", value=DEFAULT_PROMPT)

    # Sidebar submit button
    if st.sidebar.button("Submit"):
        if uploaded_file is not None:
            file_extension = uploaded_file.name.split(".")[-1].lower()
            
            # Handle image files (JPEG, PNG)
            if file_extension in ["jpeg", "jpg", "png"]:
                # Process image files directly
                image = Image.open(uploaded_file)
                
                # Process the image and display results
                st.subheader(f"Processing Image")
                st.image(image, caption="Uploaded Image", use_column_width=True)
                
                with st.spinner(f"Analyzing Image with GPT-4o..."):
                    analysis_result = analyze_image_with_gpt4o(image, prompt)
                    if analysis_result:
                        st.subheader(f"Analysis Result:")
                        st.json(analysis_result)
                    else:
                        st.error(f"Failed to analyze the image")

            # Handle PPTX files
            elif file_extension == "pptx":
                # Convert PPTX to PDF
                pdf_output_path = os.path.abspath("converted_presentation.pdf")
                convert_pptx_to_pdf(uploaded_file, pdf_output_path)
                
                # Read the PDF for further processing
                with open(pdf_output_path, "rb") as pdf_file:
                    pdf_stream = io.BytesIO(pdf_file.read())
                
            # Handle XLSX files
            elif file_extension == "xlsx":
                # Convert XLSX to PDF
                pdf_output_path = os.path.abspath("converted_workbook.pdf")
                convert_xlsx_to_pdf(uploaded_file, pdf_output_path)
                
                # Read the PDF for further processing
                with open(pdf_output_path, "rb") as pdf_file:
                    pdf_stream = io.BytesIO(pdf_file.read())

            # Handle PDF files directly
            elif file_extension == "pdf":
                pdf_stream = io.BytesIO(uploaded_file.read())

            # Convert PDF to images
            if file_extension in ["pdf", "pptx", "xlsx"]:
                images = convert_pdf_to_images(pdf_stream)
                st.write(f"Total Pages: {len(images)}")

                all_results = []
                total_entities = 0

                # Reset the stream to the beginning for image conversion
                pdf_stream.seek(0)

                for index, image in enumerate(images):
                    st.subheader(f"Processing Page {index + 1}")

                    # Display the image
                    st.image(image, caption=f"Page {index + 1}", use_column_width=True)

                    # Process the image and display results
                    with st.spinner(f"Analyzing Page {index + 1} with GPT-4o..."):
                        analysis_result = analyze_image_with_gpt4o(image, prompt)
                        if analysis_result:
                            st.subheader(f"Analysis Result for Page {index + 1}:")
                            st.json(analysis_result)
                            all_results.append(analysis_result)
                            
                            # Count entities on each page
                            if 'entities' in analysis_result:
                                total_entities += len(analysis_result['entities'])

                # Display final analysis count
                st.subheader(f"Total Entities Extracted: {total_entities}")
                st.write(f"Analysis Complete")

if __name__ == "__main__":
    main()