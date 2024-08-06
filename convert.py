import os
import pdf2image
import pytesseract
import pandas as pd
import openai
from dotenv import load_dotenv
from PIL import Image

# Load environment variables from .env file
load_dotenv()
openai.api_key = os.getenv('OPENAI_API_KEY')

# Path to the PDF file
pdf_path = 'file.pdf'

# Output paths for CSV and Excel
csv_output = 'output.csv'
excel_output = 'output.xlsx'

# Directory to save images
images_folder = 'images'
os.makedirs(images_folder, exist_ok=True)

# Convert PDF pages to images
print("Converting PDF pages to images...")
images = pdf2image.convert_from_path(pdf_path, dpi=300)

# Function to extract text from an image using OCR
def extract_text_from_image(image: Image) -> str:
    return pytesseract.image_to_string(image)

# Function to use OpenAI to parse and extract tabular data
def parse_table_with_openai(text: str) -> pd.DataFrame:
    try:
        # Send the text to OpenAI's API to extract table data
        response = openai.ChatCompletion.create(
            model="gpt-4",  # Use "gpt-4" if you have access
            messages=[
                {"role": "system", "content": "You are an expert in extracting and structuring tabular data from text."},
                {"role": "user", "content": f"Extract table data from the following text and format it as a CSV table:\n\n{text}"}
            ],
            max_tokens=2048,
            temperature=0.2
        )
        table_text = response.choices[0].message['content'].strip()
        
        # Print the extracted and corrected table data
        print("Extracted and Corrected Table Data:")
        print(table_text)
        
        # Convert the structured text into a DataFrame
        data = [line.split(',') for line in table_text.split('\n') if line]
        df = pd.DataFrame(data[1:], columns=data[0])
        return df
    except Exception as e:
        print(f"An error occurred during OpenAI processing: {e}")
        return pd.DataFrame()

# List to store all DataFrames
dataframes = []

# Process each image and extract tables
for i, image in enumerate(images):
    # Save each page image to the images folder
    image_path = os.path.join(images_folder, f'page_{i + 1}.jpg')
    image.save(image_path, 'JPEG')
    print(f"Saved image for page {i + 1} as {image_path}")

    # Extract text from image
    text = extract_text_from_image(image)
    print(f"Extracted text from page {i + 1}:\n{text}\n")

    # Parse table data using OpenAI
    df = parse_table_with_openai(text)
    if not df.empty:
        # Add page number information
        df['Source Page'] = i + 1
        dataframes.append(df)

# Concatenate all DataFrames into a single DataFrame
try:
    full_table = pd.concat(dataframes, ignore_index=True)
    # Save the full table to CSV
    full_table.to_csv(csv_output, index=False)
    print(f'Table data saved to {csv_output}')

    # Save the full table to Excel
    full_table.to_excel(excel_output, index=False)
    print(f'Table data saved to {excel_output}')
except ValueError as e:
    print(f"An error occurred while concatenating tables: {e}")
    print("Please check the structure of the tables in the PDF and ensure they have compatible formats.")
