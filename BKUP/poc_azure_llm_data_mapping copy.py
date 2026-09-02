import json
import logging
import os
import time
from datetime import datetime

from openai import AzureOpenAI
from utility import *

from prompt_file import get_mapping_prompt

API_KEY_FILE = "./utility_files/AzureOpeapiKeys.txt"

JSON_FILE_NAME = "output_1_prebuilt-layout"
JSON_FILE_NAME = "output_4_prebuilt-layout"

JSON_FILE_PATH = f"./input_ADI_json/{JSON_FILE_NAME}.json"

# JSON_FILE_NAME = "3"
# JSON_FILE_PATH = f"./input_final_json_extracted/{JSON_FILE_NAME}.json"


OUTPUT_FILE = f"./output_llm_on_ADI/output_{JSON_FILE_NAME}_{get_model_version()}.json"
OUTPUT_FILE_intermediate_llm = f"./output_llm_on_ADI/output_{JSON_FILE_NAME}_{get_model_version()}_intermediate.json"

example_json_path = "./utility_files/EXAMPLE_inversis.json"
country_code_mapping_path = "./utility_files/country_code_mapping.csv"
mapping_rule_excel_path = "./utility_files/FMSB_Securities_SSI_Create_Update_v.2.0.3.xlsx"

LOG_LEVEL = os.getenv("LOG_LEVEL", "DEBUG").upper()
LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO").upper()
LOG_FILE = os.getenv("LOG_FILE", "")
LOG_DIR = os.getenv("LOG_DIR", "./logs/ssi_mapper")


def resolve_log_file_path():
    """Create one unique log file per run unless LOG_FILE is explicitly provided."""
    if LOG_FILE:
        return LOG_FILE

    run_stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    return os.path.join(LOG_DIR, f"ssi_mapper_{JSON_FILE_NAME}_{run_stamp}.log")


ACTIVE_LOG_FILE = resolve_log_file_path()


def setup_logging():
    log_parent_dir = os.path.dirname(ACTIVE_LOG_FILE)
    if log_parent_dir:
        os.makedirs(log_parent_dir, exist_ok=True)

    handlers = [logging.StreamHandler()]
    handlers.append(logging.FileHandler(ACTIVE_LOG_FILE, encoding="utf-8"))

    logging.basicConfig(
        level=getattr(logging, LOG_LEVEL, logging.INFO),
        format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        handlers=handlers,
    )


setup_logging()
logger = logging.getLogger("ssi_mapper")

APPROVED_FIELD_ORDER = [
    "action",
    "instructingParty.identifier",
    "instructingParty.lei",
    "instructingParty.bic",
    "instructingParty.proprietaryIdentification",
    "instructingParty.proprietarySchemeName",
    "beneficiary.entityIdentification",
    "effectiveDate.parameter",
    "effectiveDate.start",
    "effectiveDate.end",
    "market.settlementPurpose",
    "market.settlementCountry",
    "market.classificationIdentifier",
    "market.isin",
    "market.settlementCurrency",
    "market.psafeBic",
    "depository.psetBic",
    "party1.identifier",
    "party1.bic",
    "party1.accountIdentification",
    "party1.accountName",
    "party2.identifier",
    "party2.bic",
    "party2.accountIdentification",
    "party2.accountName",
    "party3.identifier",
    "party3.bic",
    "party3.accountIdentification",
    "party3.accountName",
    "party4.identifier",
    "party4.bic",
    "party4.accountIdentification",
    "party4.accountName",
]
 

def call_llm(prompt, endpoint, api_key, api_version, deployment_name, max_retries=5, base_delay_seconds=2):
    """Send the PDF file itself (no local text/image extraction) + SSI prompt to Azure OpenAI."""
    client = AzureOpenAI(
        azure_endpoint=endpoint,
        api_key=api_key,
        api_version=api_version,
    )
    logger.info("Starting LLM call for deployment '%s'", deployment_name)
    logger.debug("Prompt size (chars): %s", len(prompt))

    for attempt in range(1, max_retries + 1):
        try:
            logger.debug("LLM attempt %s/%s on deployment '%s'", attempt, max_retries, deployment_name)
            response = client.chat.completions.create(
                model=deployment_name,
                temperature=0,
                messages=[
                    {"role": "user", "content": prompt},
                ],
                response_format={"type": "json_object"},
            )

            content = response.choices[0].message.content
            logger.info("LLM response received from deployment '%s'", deployment_name)
            return json.loads(content)
        except Exception as exc:
            error_text = str(exc).lower()
            is_rate_limit = "429" in error_text or "rate_limit" in error_text
            if not is_rate_limit or attempt == max_retries:
                logger.exception("LLM call failed for deployment '%s' on attempt %s", deployment_name, attempt)
                raise

            delay = base_delay_seconds * (2 ** (attempt - 1))
            logger.warning(
                "Rate limit hit on deployment '%s' attempt %s/%s. Retrying in %ss",
                deployment_name,
                attempt,
                max_retries,
                delay,
            )
            time.sleep(delay)


def call_llm_with_fallback(prompt, endpoint, api_key, api_version, deployment_candidates):
    """Try deployments in order; fall back when a deployment remains rate-limited."""
    last_error = None
    logger.info("Deployment fallback order: %s", deployment_candidates)
    for deployment_name in deployment_candidates:
        logger.info("Trying deployment: %s", deployment_name)
        try:
            return call_llm(prompt, endpoint, api_key, api_version, deployment_name)
        except Exception as exc:
            last_error = exc
            error_text = str(exc).lower()
            if "429" in error_text or "rate_limit" in error_text:
                logger.warning("Deployment '%s' is rate-limited. Trying next fallback.", deployment_name)
                continue
            if "deploymentnotfound" in error_text:
                logger.warning("Deployment '%s' not found. Trying next fallback.", deployment_name)
                continue
            logger.exception("Deployment '%s' failed with non-retryable error", deployment_name)
            raise

    if last_error:
        logger.error("All deployment candidates failed")
        raise last_error
    raise RuntimeError("No deployment candidates were provided.")


def _empty_field(field_id, field_name):
    return {
        "fieldId": field_id,
        "fieldName": field_name,
        "mappedValue": None,
        "components": [],
        "mappingStatus": "NOT_AVAILABLE",
        "mappingBasis": [],
        "sourceGroupIds": [],
        "sourceValues": [],
        "evidence": [],
        "reviewRequired": False,
        "reviewReasonCodes": [],
    }


def enforce_output_contract(result):
    """Guarantee top-level structure and strict 33-field mapping contract."""
    if not isinstance(result, dict):
        raise ValueError("LLM output must be a JSON object.")

    result.setdefault("promptVersion", "P2_SANTANDER_FMSB_MAPPER_V1.0")
    result.setdefault("rulesVersion", "Santander-FMSB-v2.0.3-Custom-Final")
    result.setdefault("documentName", "")
    result.setdefault("records", [])
    result.setdefault("unmappedAndReview", [])

    if not isinstance(result["records"], list):
        result["records"] = []

    for index, record in enumerate(result["records"], start=1):
        record.setdefault("recordId", f"SSI-{index:03d}")
        record.setdefault("sourceBlockId", "")
        record.setdefault("recordScopeStatus", "ACTION_NOT_STATED")
        record.setdefault("recordReviewRequired", False)
        record.setdefault("recordReviewReasonCodes", [])
        fields = record.get("fields", [])
        if not isinstance(fields, list):
            fields = []

        existing_by_name = {
            field.get("fieldName"): field
            for field in fields
            if isinstance(field, dict) and field.get("fieldName")
        }

        extra_fields = [
            field for field_name, field in existing_by_name.items()
            if field_name not in APPROVED_FIELD_ORDER
        ]
        if extra_fields:
            logger.info(
                "Record %s returned %s non-approved fields; moved to unmappedAndReview",
                record.get("recordId", f"SSI-{index:03d}"),
                len(extra_fields),
            )
            result["unmappedAndReview"].append(
                {
                    "recordId": record["recordId"],
                    "reason": "NON_APPROVED_FIELDS_RETURNED",
                    "fields": extra_fields,
                }
            )

        reordered = []
        for field_pos, field_name in enumerate(APPROVED_FIELD_ORDER, start=1):
            candidate = existing_by_name.get(field_name)
            field_id = f"F-{field_pos:02d}"
            if not isinstance(candidate, dict):
                reordered.append(_empty_field(field_id, field_name))
                continue

            candidate["fieldId"] = candidate.get("fieldId") or field_id
            candidate["fieldName"] = field_name
            candidate.setdefault("mappedValue", None)
            candidate.setdefault("components", [])
            candidate.setdefault("mappingStatus", "NOT_AVAILABLE")
            candidate.setdefault("mappingBasis", [])
            candidate.setdefault("sourceGroupIds", [])
            candidate.setdefault("sourceValues", [])
            candidate.setdefault("evidence", [])
            candidate.setdefault("reviewRequired", False)
            candidate.setdefault("reviewReasonCodes", [])
            reordered.append(candidate)

        record["fields"] = reordered

    return result

 
  

def get_prompt(input_json_path=JSON_FILE_PATH):
     # Read JSON output from Azure Document Intelligence
    with open(input_json_path, "r", encoding="utf-8") as f:
        doc_json = json.load(f)

    # Convert JSON to string for the LLM
    json_content = json.dumps(doc_json, ensure_ascii=False, indent=2)
    return json_content



def main():

    try:
        logger.info("SSI mapping run started")
        logger.info("Log file for this run: %s", ACTIVE_LOG_FILE)

        endpoint, api_key, api_version, deployment_name = load_llm_credentials()
        preferred_deployment = os.getenv("AZURE_OPENAI_DEPLOYMENT", get_model_version())
        deployment_candidates = []
        for model_name in [preferred_deployment, deployment_name, "gpt-5.4-mini", "gpt-5.6-luna", "gpt-5.6-terra"]:
            if model_name and model_name not in deployment_candidates:
                deployment_candidates.append(model_name)

        ## to take only one model gpt 5.4
        deployment_candidates = [deployment_candidates[0]]


        logger.info("Using input JSON file: %s", JSON_FILE_PATH)
        logger.info("Using mapping rules Excel: %s", mapping_rule_excel_path)

        example_json = read_json_file_and_convert_to_string(example_json_path)
        country_code_mapping = read_csv_file_and_convert_to_string(country_code_mapping_path)
        mapping_rule_excel_data = read_excel_file_and_convert_to_string(mapping_rule_excel_path)

        data_json = read_json_file_and_convert_to_string(JSON_FILE_PATH)
        logger.debug("Loaded example JSON chars: %s", len(example_json))
        logger.debug("Loaded country mapping chars: %s", len(country_code_mapping))
        logger.debug("Loaded mapping Excel JSON chars: %s", len(mapping_rule_excel_data))
        logger.debug("Loaded input data JSON chars: %s", len(data_json))


        
        try:

            prompt = get_mapping_prompt(mapping_rule_excel_data, country_code_mapping, example_json, data_json)

            prompt_save_file_path = f"./output_llm_on_ADI/prompt_{JSON_FILE_NAME}_{get_model_version()}.txt"
            logger.info("Saving prompt to file : %s", prompt_save_file_path)
            save_txt_to_file(prompt, prompt_save_file_path)

            logger.debug("Final prompt chars: %s", len(prompt))

            result = call_llm_with_fallback(
                prompt,
                endpoint,
                api_key,
                api_version,
                deployment_candidates,
            )

            save_json(result, OUTPUT_FILE_intermediate_llm)
            
            logger.info("Saved intermediate LLM output: %s", OUTPUT_FILE_intermediate_llm)

            result = enforce_output_contract(result)
        except Exception as e:
            logger.exception("Error during LLM call: %s", e)
            raise

        save_json(result, OUTPUT_FILE)
        logger.info("Saved final mapped output: %s", OUTPUT_FILE)
        logger.info("SSI mapping run completed")

        logger.info("Extracted data preview: %s", json.dumps(result, ensure_ascii=False)[:500])
    except Exception as e:
        logger.exception("Pipeline failed: %s", e)


if __name__ == "__main__":
    main()

 