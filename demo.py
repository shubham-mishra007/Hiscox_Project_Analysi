
import os
#from unstructured.partition.pdf import partition_pdf
from unstructured.partition.docx import partition_docx
from unstructured.staging.base import elements_to_json
import pathlib
from pathlib import Path
import json

filename = "Hiscox-Policy-Wording.docx" # For this notebook I uploaded Nvidia's earnings into the Files directory called "/content/"
output_dir = "/data/"

# Define parameters for Unstructured's library
strategy = "hi_res" # Strategy for analyzing PDFs and extracting table structure
model_name = "yolox" # Best model for table extraction. Other options are detectron2_onnx and chipper depending on file layout

# Extracts the elements from the PDF
elements = partition_docx(
filename=filename
)

# Store results in json
elements_to_json(elements, filename=f"{filename}.json") 

def process_json_file(input_filename):
    # Read the JSON file
    print(input_filename)
    with open(input_filename, 'r') as file:
        data = json.load(file)

    # Iterate over the JSON data and extract required table elements
    extracted_elements = []
    text_prev = ""
    for i,entry in enumerate(data):
        if entry["type"] == "Title":
            text = "" + entry["text"] + ""
        elif entry["type"] == "Table":
            try:
                text = entry["metadata"]["text_as_html"]
            except Exception as ex:
                try:
                    text = "" + entry["text"] + ""
                except Exception:
                    text = "Not Found"
        else:
            text = "" + entry["text"] + ""

        if text != text_prev: extracted_elements.append(text)
        text_prev = text

    # Write the extracted elements to the output file
    html_start = """

    
    
    Document Information
    
    
    
    """

    html_end = """
    
    
    """

    output_file_html = Path(input_filename).name.replace(".json", "") + "_" + model_name + ".html"
    print(output_file_html)
    with open(output_file_html, 'w') as output_file:
        output_file.write(html_start + "\n")
        for element in extracted_elements:
            output_file.write(element + "\n")
        output_file.write(html_end + "\n")

    return str(output_file_html)

process_json_file(f"{filename}.json")