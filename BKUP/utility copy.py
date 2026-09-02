import os
import re
import json
import pandas as pd


API_KEY_FILE = "./utility_files/AzureOpeapiKeys.txt"



def save_txt_to_file(text, filepath):
    """Save the prompt to a text file."""
    with open(filepath, "w", encoding="utf-8") as f:
        f.write(text)

def load_llm_credentials(config_path=API_KEY_FILE):
    """
    Parse endpoint / api-key / api-version / deployment-name out of the
    '${VAR:default}' style config file (AzureOpeapiKeys.txt).
    """
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"Config file '{config_path}' not found.")

    with open(config_path, "r", encoding="utf-8") as f:
        content = f.read()

    def extract(field):
        match = re.search(rf"{field}:\s*\$\{{[^:]+:([^}}]+)\}}", content)
        if not match:
            raise ValueError(f"Could not find '{field}' in '{config_path}'.")
        return match.group(1).strip()

    endpoint = extract("endpoint")
    api_key = extract("api-key")
    api_version = extract("api-version")
    deployment_name = extract("deployment-name")

    return endpoint, api_key, api_version, deployment_name


def get_model_version():
    model = "gpt-5.6-sol"  
    model = "gpt-5.6-terra"  
    model = "gpt-5.6-luna" 
    model = "gpt-5.4-mini"

    model = "gpt-5.4" 
    return model

 
def read_json_file_and_convert_to_string(file_path):
    """Read a JSON file and return its content."""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"JSON file '{file_path}' not found.")

    with open(file_path, "r", encoding="utf-8") as f:
        return json.dumps(json.load(f), indent=2)


def read_csv_file_and_convert_to_string(file_path):
    """Read a CSV file and return its content as a string."""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"CSV file '{file_path}' not found.")

    df = pd.read_csv(file_path)
    csv_json = df.to_dict(orient="records")
    return json.dumps(csv_json, indent=2)


def read_excel_file_and_convert_to_string(file_path, sheet_name=0):
    """Read an Excel sheet and return row-wise JSON string for LLM grounding."""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Excel file '{file_path}' not found.")

    df = pd.read_excel(file_path, sheet_name=sheet_name, dtype=str)
    df = df.fillna("")

    # Preserve full rule table so the model can reason over row relationships.
    excel_json = df.to_dict(orient="records")
    return json.dumps(excel_json, indent=2, ensure_ascii=False)


def save_json(data, output_path):
    """Save the LLM's JSON response to a file."""

    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)
    print(f"Saved LLM response to '{output_path}'") 
