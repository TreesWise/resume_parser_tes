import pytesseract
from pdf2image import convert_from_path
from PIL import Image
import os
import json
import openai
from dotenv import load_dotenv
from fastapi import HTTPException
from docx2pdf import convert
import tempfile
from openai import AzureOpenAI


load_dotenv()



AZURE_OPENAI_API_KEY = os.getenv("AZURE_OPENAI_API_KEY")
AZURE_OPENAI_ENDPOINT = os.getenv("AZURE_OPENAI_ENDPOINT")
AZURE_OPENAI_DEPLOYMENT_NAME = os.getenv("AZURE_OPENAI_DEPLOYMENT_NAME")
OPENAI_API_VERSION = os.getenv("OPENAI_API_VERSION")



# Set Tesseract path (Only needed for Windows)
pytesseract.pytesseract.tesseract_cmd = "/usr/bin/tesseract"  # Update this path
os.environ["TESSDATA_PREFIX"] = "/usr/share/tesseract-ocr/4.00/tessdata"
def extract_text_from_scanned_pdf(pdf_path):
    """Extracts text from a scanned PDF using Tesseract OCR."""
    images = convert_from_path(pdf_path, poppler_path="/usr/bin")  # Convert PDF pages to images
    text = ""
    for img in images:
        text += pytesseract.image_to_string(img, lang="eng") + "\n"  # OCR on each page
    return text.strip()

with open("output_json.json", "r", encoding="utf-8") as file:
        expected_structure = json.load(file)

# output_json_path = r"D:\OneDrive - MariApps Marine Solutions Pte.Ltd\liju_resume_parser\resume_parser_custom model\output_json.json"
# with open(output_json_path, "r") as f:
#     expected_structure = json.load(f)


# openai.api_key = os.getenv("openai_api_key")

# def get_openai_response(prompt, extracted_text):
#     response = openai.ChatCompletion.create(
#         model="gpt-4o", # "gpt-4",  # or 
#         messages=[
#             {"role": "system", "content": "You are an AI assistant."},
#             {"role": "user", "content": f"Prompt: {prompt}\nText: {extracted_text}"}
#         ],
#          response_format={"type": "json_object"}
#     )   
#     # print(response["choices"][0]["message"]["content"])
#     output_json = json.loads(response.choices[0].message.content)
#     return output_json



def get_openai_response(prompt, extracted_text):
    client = AzureOpenAI(
        azure_endpoint=AZURE_OPENAI_ENDPOINT,
        api_key=AZURE_OPENAI_API_KEY,
        api_version=OPENAI_API_VERSION,
    )

    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": "You are an AI assistant."},
            {"role": "user", "content": f"Prompt: {prompt}\nText: {extracted_text}"}
        ],
        response_format={"type": "json_object"},
    )
    output_json = json.loads(response.choices[0].message.content)
    return output_json

prompt = f"""You are an expert in data extraction and JSON formatting. Your task is to extract and format resume data **exactly** as per the provided JSON template `{expected_structure}`.  

### **Critical Rules – No Deviations Allowed:**  
#### 1. **Strict JSON Structure Compliance (No Changes Allowed)**  
- The output **MUST** exactly match the provided JSON **1:1, including structure, field names, order, nesting, and formatting**.  
- **Do NOT modify the JSON structure in any way.** The AI **must not** reorder, rename, remove, or nest fields differently than `{expected_structure}`.  
- **Every key must be present. If data is missing, return `null` instead of omitting the field.**  

#### 2. **Data Extraction Rules (No Data Should Be Missed)**  
##### **basic_details:**  
- Extract and correctly map `City`, `State`, `Country`, and `Zipcode`.  
- Split the address into `Address1`, `Address2`, `Address3`, and `Address4`.  

##### **Experience Table (Do NOT Miss Any Entries):**  
- **Every single experience entry must be included.** If an entry spans multiple lines, **merge those lines** into a single entry.  
- `TEU` must always be numerical; if missing, return `null`.  
- `IMO` must be a 7-digit number; if missing, return `null`.  
- `Flag` must be a valid country name (e.g., "Panama"); otherwise, return `null`.  
- **Output format must strictly follow the order in `{expected_structure}`. Do NOT rearrange the fields.**  

##### **Certificate Table (Do NOT Miss Any Documents):**  
- **Every certificate, visa, passport,Medical,Yellow Fever and flag document must be included—even if scattered across different sections. Don't omit any of these documents since these documnets are more important**
- Merge related certificates into a single entry (e.g., "GMDSS ENDORSEMENT").  
- If `NUMBER`, `ISSUING VALIDATION DATE`, or `ISSUING PLACE` are missing, return `null` but **do NOT remove the entry**.  
- If a certificate's NUMBER is missing, include the field as `null`.  
- **Do NOT modify the order or structure of certificates. Follow `{expected_structure}` exactly.**
- **Certificate Place & Country Extraction Rules:**
        - If the `PlaceOfIssue` contains both a **city and a country** (e.g., "Kochi, India"), split them correctly:
            - The city ("Kochi") should go under `"PlaceOfIssue"`.
            - The country ("India") should go under `"CountryOfIssue"`.
        - If **only one location is mentioned**, determine if it is a **city** or a **country**:
            - If it is a **recognized city**, it should go under `"PlaceOfIssue"`.
            - If it is a **recognized country**, it should go under `"CountryOfIssue"`.
        - If `"CountryOfIssue"` is missing but `"PlaceOfIssue"` contains a **comma-separated value**, assume the last part is the country.
        - Ensure that `"PlaceOfIssue"` and `"CountryOfIssue"` are never combined in a single field.

#### 4. **Date Format Standardization:**
- All date fields (`Dob`, `FromDt`, `ToDt`, `DateOfIssue`, `DateOfExpiry`, etc.) **must strictly follow the format in the sample JSON**: `DD-MM-YYYY`.
- Ensure no deviations (e.g., `YYYY-MM-DD`, `MM/DD/YYYY`, or `DD/MM/YY` are not alloweld).
- If a date is incomplete (e.g., missing the day or month), return `null` for that field.

#### 3. **Ensuring Accuracy & Completeness**  
- **EVERY experience and certificate entry must be included—NO omissions allowed.**  
- **The order of fields and tables must match `{expected_structure}` exactly.**  
- **No additional text, comments, or extra nesting—just the JSON output.**  

#### 4. **Strict JSON Output Formatting**  
- **Return ONLY a properly formatted, validated JSON response—no extra text, explanations, or code blocks.**  
- **Ensure the final JSON structure is a perfect match to `{expected_structure}`.**  
- **DO NOT modify field names, order, or nesting—output must be an exact match.**  

**Strictly follow these instructions to ensure 100% accuracy. Return only the structured JSON output.**  
"""


async def convert_docx_to_pdf(docx_path):
    """ Converts DOCX to PDF using LibreOffice (Linux) or Microsoft Word (Windows). """
    pdf_path = docx_path.replace(".docx", ".pdf")

    try:
        if platform.system() == "Windows":
            import win32com.client
            word = win32com.client.Dispatch("Word.Application")
            doc = word.Documents.Open(os.path.abspath(docx_path))
            doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17)  # PDF format
            doc.Close()
            word.Quit()
            print(f" Converted {docx_path} to {pdf_path} using Microsoft Word")
        else:
            libreoffice_path = "/usr/bin/libreoffice"
            if not os.path.exists(libreoffice_path):
                raise FileNotFoundError(f"LibreOffice not found at {libreoffice_path}")

            process = await asyncio.create_subprocess_exec(
                libreoffice_path, "--headless", "--convert-to", "pdf",
                "--outdir", os.path.dirname(docx_path), docx_path
            )
            await process.communicate()  # Ensure subprocess completes

            print(f" Converted {docx_path} to {pdf_path} using LibreOffice")

        return pdf_path
    except Exception as e:
        print(f" DOCX to PDF conversion failed: {e}")
        raise HTTPException(status_code=500, detail=f"DOCX to PDF conversion failed: {e}")
    
def replace_values(data, mapping):
    if isinstance(data, dict):
        return {mapping.get(key, key): replace_values(value, mapping) for key, value in data.items()}
    elif isinstance(data, list):
        return [replace_values(item, mapping) for item in data]
    elif isinstance(data, str):
        return mapping.get(data, data)  # Replace if found, else keep original
    return data

def replace_rank(json_data, rank_mapping):
    # Convert rank_mapping keys to lowercase for case-insensitive replacement
    rank_mapping = {key.lower(): value for key, value in rank_mapping.items()}

    if isinstance(json_data, dict):
        return {
            key: replace_rank(value, rank_mapping) if key != "2" else  # "2" corresponds to "Position"
            rank_mapping.get(value.lower(), value) if isinstance(value, str) else value
            for key, value in json_data.items()
        }
    elif isinstance(json_data, list):
        return [replace_rank(item, rank_mapping) for item in json_data]
    return json_data
