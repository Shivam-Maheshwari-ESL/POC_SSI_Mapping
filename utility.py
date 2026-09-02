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


def slim_doc_intelligence_json(doc_json):
    """Strip Azure Document Intelligence geometry/offset noise the LLM never uses.

    Raw prebuilt-layout output carries per-word/line polygons, spans, and
    boundingRegions for every element, which can inflate a single document to
    hundreds of thousands of tokens and blow through TPM rate limits. The
    model only needs the plain text (`content`) plus table structure/content,
    so drop `pages` (words/lines with polygons) and strip geometry from
    `tables` and `paragraphs` while keeping their text and structure.
    """
    if not isinstance(doc_json, dict):
        return doc_json

    slim = {
        key: value
        for key, value in doc_json.items()
        if key not in ("pages", "tables", "paragraphs", "styles")
    }

    tables = doc_json.get("tables")
    if isinstance(tables, list):
        slim["tables"] = [
            {
                "rowCount": table.get("rowCount"),
                "columnCount": table.get("columnCount"),
                "cells": [
                    {
                        k: cell.get(k)
                        for k in ("kind", "rowIndex", "columnIndex", "rowSpan", "colSpan", "content")
                        if k in cell
                    }
                    for cell in table.get("cells", [])
                    if isinstance(cell, dict)
                ],
            }
            for table in tables
            if isinstance(table, dict)
        ]

    paragraphs = doc_json.get("paragraphs")
    if isinstance(paragraphs, list):
        slim["paragraphs"] = [
            {k: paragraph.get(k) for k in ("role", "content") if k in paragraph}
            for paragraph in paragraphs
            if isinstance(paragraph, dict)
        ]

    return slim


def read_json_file_and_convert_to_slim_string(file_path):
    """Read a Document Intelligence JSON file, strip geometry noise, and return it as a string."""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"JSON file '{file_path}' not found.")

    with open(file_path, "r", encoding="utf-8") as f:
        doc_json = json.load(f)

    return json.dumps(slim_doc_intelligence_json(doc_json), indent=2, ensure_ascii=False)


def read_csv_file_and_convert_to_string(file_path):
    """Read a CSV file and return its content as a string."""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"CSV file '{file_path}' not found.")

    df = pd.read_csv(file_path)
    csv_json = df.to_dict(orient="records")
    return json.dumps(csv_json, indent=2)


def read_excel_file_and_convert_to_string(file_path, sheet_name=0):
    """Read the mapping-rule sheet and return a compact CSV string for LLM grounding.

    The workbook is laid out with attributes (Section, Field, Mapping rule, ...) as rows
    and one column per rule entry, with blank spacer rows/columns mixed in. Reading it with
    the default header made pandas treat it row-wise, producing one JSON record per
    attribute with ~87 meaningless "Unnamed: N" keys each (~170K chars of repeated,
    unlabeled keys). Read raw (no header), transpose so each rule entry becomes one row
    with the real attribute names as columns, drop blank spacer rows/columns, and emit
    CSV instead of indented JSON to remove the repeated-key overhead entirely.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Excel file '{file_path}' not found.")

    raw = pd.read_excel(file_path, sheet_name=sheet_name, dtype=str, header=None, engine="openpyxl")
    raw = raw.fillna("")

    labels = raw.iloc[:, 0]
    entries = raw.iloc[:, 1:].transpose()
    entries.columns = labels.values

    entries = entries.loc[:, [str(c).strip() != "" for c in entries.columns]]
    entries = entries.loc[~(entries == "").all(axis=1)]

    return entries.to_csv(index=False)


def save_json(data, output_path):
    """Save the LLM's JSON response to a file."""

    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)
    print(f"Saved LLM response to '{output_path}'") 

