import argparse
import json
import os
import re
from dataclasses import dataclass
from typing import Dict, List, Set, Tuple

import pandas as pd


@dataclass
class PairScore:
	excel_row_idx: int
	excel_row_number: int
	record_id: str
	exact_match_count: int
	compared_field_count: int
	mismatch_count: int
	match_ratio: float


PRIMARY_CANONICAL_FIELDS = [
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

PRIMARY_FIELDS_SET = set(PRIMARY_CANONICAL_FIELDS)

DISPLAY_TO_CANONICAL = {
	"action": "action",
	"instructing party identifier": "instructingParty.identifier",
	"instructing entity lei": "instructingParty.lei",
	"instructing party bic": "instructingParty.bic",
	"proprietary identification": "instructingParty.proprietaryIdentification",
	"proprietary scheme name": "instructingParty.proprietarySchemeName",
	"beneficiary entity identification": "beneficiary.entityIdentification",
	"effective date parameter": "effectiveDate.parameter",
	"effective start date": "effectiveDate.start",
	"effective end date": "effectiveDate.end",
	"settlement purpose": "market.settlementPurpose",
	"settlement country": "market.settlementCountry",
	"classification identifier": "market.classificationIdentifier",
	"isin": "market.isin",
	"settlement currency": "market.settlementCurrency",
	"psafe bic": "market.psafeBic",
	"pset party identifier bic": "depository.psetBic",
}


def normalize_header_tokens(text: str) -> Tuple[str, ...]:
	tokens = re.findall(r"[a-z0-9]+", text.lower())
	return tuple(tokens)


def normalize_header_key(text: str) -> str:
	return " ".join(normalize_header_tokens(text))


def split_pandas_dedup_suffix(column_name: str) -> Tuple[str, int]:
	match = re.match(r"^(.*?)(?:\.(\d+))?$", column_name)
	if not match:
		return column_name, 1
	base = match.group(1)
	suffix = match.group(2)
	occurrence = (int(suffix) + 1) if suffix is not None else 1
	return base, occurrence


def title_from_json_field(field_name: str) -> str:
	temp = re.sub(r"([a-z])([A-Z])", r"\1 \2", field_name)
	temp = temp.replace(".", " ")
	temp = re.sub(r"(party)(\d+)", r"\1 \2", temp)
	temp = re.sub(r"\s+", " ", temp).strip()
	return temp.title()


def value_to_str(value) -> str:
	if pd.isna(value):
		return ""
	if value is None:
		return ""
	return str(value)


def load_excel(file_path: str, sheet_name: str, header_row_number: int) -> pd.DataFrame:
	if not os.path.exists(file_path):
		raise FileNotFoundError(f"Excel file not found: {file_path}")

	header_index = header_row_number - 1
	df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_index, dtype=str)
	df = df.dropna(how="all")
	df = df.reset_index(drop=True)

	# Convert all values to string while preserving exact characters and case.
	for col in df.columns:
		df[col] = df[col].map(value_to_str)

	return df


def load_excel_group_row(file_path: str, sheet_name: str, header_row_number: int, column_count: int) -> List[str]:
	"""Read row above header to identify subsection/group labels (e.g., Party 1..5)."""
	raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None, dtype=str)
	group_row_index = header_row_number - 2
	if group_row_index < 0 or group_row_index >= len(raw):
		return ["BASE"] * column_count

	group_values = raw.iloc[group_row_index].tolist()
	if len(group_values) < column_count:
		group_values = group_values + [""] * (column_count - len(group_values))
	group_values = group_values[:column_count]

	normalized_groups: List[str] = []
	current_group = "BASE"
	for value in group_values:
		text = value_to_str(value).strip()
		if text:
			current_group = text
		normalized_groups.append(current_group)

	return normalized_groups


def load_json_records(file_path: str) -> Tuple[Dict[str, Dict[str, str]], Dict[str, str], Set[str]]:
	if not os.path.exists(file_path):
		raise FileNotFoundError(f"JSON file not found: {file_path}")

	with open(file_path, "r", encoding="utf-8") as handle:
		payload = json.load(handle)

	records = payload.get("records", [])
	if not isinstance(records, list):
		raise ValueError("JSON payload has invalid 'records' format. Expected a list.")

	flattened: Dict[str, Dict[str, str]] = {}
	primary_field_names_available: Set[str] = set()
	for idx, record in enumerate(records, start=1):
		if not isinstance(record, dict):
			continue

		record_id = record.get("recordId") or f"ROW-{idx:03d}"
		fields = record.get("fields", [])
		if not isinstance(fields, list):
			fields = []

		by_field: Dict[str, str] = {}
		for field in fields:
			if not isinstance(field, dict):
				continue
			field_name = value_to_str(field.get("fieldName")).strip()
			field_id = value_to_str(field.get("fieldId")).strip()
			if not field_name and not field_id:
				continue

			primary_key = ""
			if field_id in PRIMARY_FIELDS_SET:
				primary_key = field_id
			elif field_name:
				display_key = normalize_header_key(field_name)
				if display_key in DISPLAY_TO_CANONICAL:
					primary_key = DISPLAY_TO_CANONICAL[display_key]
				elif field_name in PRIMARY_FIELDS_SET:
					primary_key = field_name

			value_str = value_to_str(field.get("mappedValue"))
			if primary_key:
				by_field[primary_key] = value_str
				primary_field_names_available.add(primary_key)

			if field_name:
				by_field[field_name] = value_str

		flattened[str(record_id)] = by_field

	json_field_names: Set[str] = set()
	for field_map in flattened.values():
		json_field_names.update(field_map.keys())

	alias_by_header_key: Dict[str, str] = {}

	# Build aliases from fields present in records. Prefer canonical primary keys when collisions occur.
	ordered_fields = sorted(
		json_field_names,
		key=lambda name: (0 if name in PRIMARY_FIELDS_SET else 1, name.lower()),
	)
	for field_name in ordered_fields:
		for alias_key in (
			normalize_header_key(field_name),
			normalize_header_key(title_from_json_field(field_name)),
		):
			existing = alias_by_header_key.get(alias_key)
			if existing is None:
				alias_by_header_key[alias_key] = field_name
				continue
			if existing not in PRIMARY_FIELDS_SET and field_name in PRIMARY_FIELDS_SET:
				alias_by_header_key[alias_key] = field_name

	# Build aliases from unmappedAndReview non-approved fields where fieldId often carries canonical keys.
	extra = payload.get("unmappedAndReview", [])
	if isinstance(extra, list):
		for item in extra:
			if not isinstance(item, dict):
				continue
			if item.get("reason") != "NON_APPROVED_FIELDS_RETURNED":
				continue
			fields = item.get("fields", [])
			if not isinstance(fields, list):
				continue
			for field in fields:
				if not isinstance(field, dict):
					continue
				canonical = field.get("fieldId")
				display_name = field.get("fieldName")
				if not canonical or not display_name:
					continue
				canonical = str(canonical)
				display_name = str(display_name)
				if canonical in json_field_names:
					alias_by_header_key[normalize_header_key(display_name)] = canonical

	if not flattened:
		raise ValueError("No comparable records found in JSON 'records'.")

	if not primary_field_names_available:
		# Fallback so execution can continue even for unexpected JSON variants.
		primary_field_names_available = set(PRIMARY_CANONICAL_FIELDS)

	return flattened, alias_by_header_key, primary_field_names_available


def build_column_mapping(
	excel_columns: List[str],
	group_labels: List[str],
	json_field_names: Set[str],
	primary_field_names: Set[str],
	alias_by_header_key: Dict[str, str],
) -> Tuple[Dict[str, str], List[str], Dict[str, str]]:
	mapping: Dict[str, str] = {}
	extra_fields: List[str] = []
	column_category: Dict[str, str] = {}
	used_primary_targets: Set[str] = set()

	token_to_json_fields: Dict[Tuple[str, ...], List[str]] = {}
	for field_name in sorted(json_field_names):
		token_key = normalize_header_tokens(field_name)
		token_to_json_fields.setdefault(token_key, []).append(field_name)

	ordered_duplicate_templates: Dict[str, List[str]] = {}

	if "party1.identifier" in json_field_names:
		ordered_duplicate_templates["party identifier"] = [
			"party1.identifier",
			"party2.identifier",
			"party3.identifier",
			"party4.identifier",
		]
	if "party1.bic" in json_field_names:
		ordered_duplicate_templates["bic"] = [
			"party1.bic",
			"party2.bic",
			"party3.bic",
			"party4.bic",
		]
	if "party1.accountIdentification" in json_field_names:
		ordered_duplicate_templates["account identification"] = [
			"party1.accountIdentification",
			"party2.accountIdentification",
			"party3.accountIdentification",
			"party4.accountIdentification",
		]
	if "party1.accountName" in json_field_names:
		ordered_duplicate_templates["account name"] = [
			"party1.accountName",
			"party2.accountName",
			"party3.accountName",
			"party4.accountName",
		]

	# Alternate duplicate mapping for display-name style JSON fields.
	if "Party 1 Party Identifier" in json_field_names:
		ordered_duplicate_templates["party identifier"] = [
			"Party 1 Party Identifier",
			"Party 2 Party Identifier",
			"Party 3 Party Identifier",
			"Party 4 Party Identifier",
		]
	if "Party 1 BIC" in json_field_names:
		ordered_duplicate_templates["bic"] = [
			"Party 1 BIC",
			"Party 2 BIC",
			"Party 3 BIC",
			"Party 4 BIC",
		]
	if "Party 1 Account Identification" in json_field_names:
		ordered_duplicate_templates["account identification"] = [
			"Party 1 Account Identification",
			"Party 2 Account Identification",
			"Party 3 Account Identification",
			"Party 4 Account Identification",
		]
	if "Party 1 Account Name" in json_field_names:
		ordered_duplicate_templates["account name"] = [
			"Party 1 Account Name",
			"Party 2 Account Name",
			"Party 3 Account Name",
			"Party 4 Account Name",
		]

	explicit_alias_overrides: Dict[str, str] = {
		"action": "action",
		"instructing party identifier": "instructingParty.identifier",
		"instructing entity lei": "instructingParty.lei",
		"instructing party bic": "instructingParty.bic",
		"proprietary identification": "instructingParty.proprietaryIdentification",
		"proprietary scheme name": "instructingParty.proprietarySchemeName",
		"beneficiary entity identification": "beneficiary.entityIdentification",
		"effective date parameter": "effectiveDate.parameter",
		"effective start date": "effectiveDate.start",
		"effective end date": "effectiveDate.end",
		"settlement purpose": "market.settlementPurpose",
		"settlement country": "market.settlementCountry",
		"classification identifier": "market.classificationIdentifier",
		"isin": "market.isin",
		"settlement currency": "market.settlementCurrency",
		"psafe bic": "market.psafeBic",
		"pset party identifier bic": "depository.psetBic",
	}

	# Keep only aliases that point to actual fields in records.
	explicit_alias_overrides = {
		k: v for k, v in explicit_alias_overrides.items() if v in json_field_names
	}

	repeated_party_headers = {
		"party identifier": "identifier",
		"bic": "bic",
		"proprietary identification": "proprietaryIdentification",
		"proprietary scheme name": "proprietarySchemeName",
		"proprietary issuer": "proprietaryIssuer",
		"account identification": "accountIdentification",
		"account name": "accountName",
	}

	for idx, column in enumerate(excel_columns):
		base_name, occurrence = split_pandas_dedup_suffix(column)
		group_label = group_labels[idx] if idx < len(group_labels) else "BASE"
		group_key = normalize_header_key(group_label)

		key_base = normalize_header_key(base_name)
		key_full = normalize_header_key(column)

		# Party-aware mapping using row-2 subsection context.
		party_match = re.match(r"party\s+(\d+)$", group_key)
		if party_match and key_base in repeated_party_headers:
			party_no = int(party_match.group(1))
			candidate = f"party{party_no}.{repeated_party_headers[key_base]}"
			if candidate in primary_field_names:
				if candidate in used_primary_targets:
					extra_fields.append(column)
					column_category[column] = "EXTRA_FIELD"
				else:
					mapping[column] = candidate
					column_category[column] = "PRIMARY_FIELD"
					used_primary_targets.add(candidate)
				continue
			# Party repeated column that does not map to the 33 primary fields is extra by definition.
			extra_fields.append(column)
			column_category[column] = "EXTRA_FIELD"
			continue

		if column in json_field_names and column in primary_field_names:
			if column in used_primary_targets:
				extra_fields.append(column)
				column_category[column] = "EXTRA_FIELD"
			else:
				mapping[column] = column
				column_category[column] = "PRIMARY_FIELD"
				used_primary_targets.add(column)
			continue

		if base_name in json_field_names and base_name in primary_field_names:
			if base_name in used_primary_targets:
				extra_fields.append(column)
				column_category[column] = "EXTRA_FIELD"
			else:
				mapping[column] = base_name
				column_category[column] = "PRIMARY_FIELD"
				used_primary_targets.add(base_name)
				continue

		if key_full in alias_by_header_key and alias_by_header_key[key_full] in json_field_names:
			candidate = alias_by_header_key[key_full]
			if candidate in primary_field_names:
				if candidate in used_primary_targets:
					extra_fields.append(column)
					column_category[column] = "EXTRA_FIELD"
				else:
					mapping[column] = candidate
					column_category[column] = "PRIMARY_FIELD"
					used_primary_targets.add(candidate)
				continue

		if key_base in alias_by_header_key and alias_by_header_key[key_base] in json_field_names:
			candidate = alias_by_header_key[key_base]
			if candidate in primary_field_names:
				if candidate in used_primary_targets:
					extra_fields.append(column)
					column_category[column] = "EXTRA_FIELD"
				else:
					mapping[column] = candidate
					column_category[column] = "PRIMARY_FIELD"
					used_primary_targets.add(candidate)
				continue

		if key_base in explicit_alias_overrides:
			candidate = explicit_alias_overrides[key_base]
			if candidate in primary_field_names:
				if candidate in used_primary_targets:
					extra_fields.append(column)
					column_category[column] = "EXTRA_FIELD"
				else:
					mapping[column] = candidate
					column_category[column] = "PRIMARY_FIELD"
					used_primary_targets.add(candidate)
				continue

		if key_base in ordered_duplicate_templates:
			candidates = ordered_duplicate_templates[key_base]
			if 1 <= occurrence <= len(candidates):
				candidate = candidates[occurrence - 1]
				if candidate in json_field_names and candidate in primary_field_names:
					if candidate in used_primary_targets:
						extra_fields.append(column)
						column_category[column] = "EXTRA_FIELD"
					else:
						mapping[column] = candidate
						column_category[column] = "PRIMARY_FIELD"
						used_primary_targets.add(candidate)
					continue

		token_key = normalize_header_tokens(column)
		candidates = token_to_json_fields.get(token_key, [])

		if len(candidates) == 1:
			candidate = candidates[0]
			if candidate in primary_field_names:
				if candidate in used_primary_targets:
					extra_fields.append(column)
					column_category[column] = "EXTRA_FIELD"
				else:
					mapping[column] = candidate
					column_category[column] = "PRIMARY_FIELD"
					used_primary_targets.add(candidate)
				continue

		extra_fields.append(column)
		column_category[column] = "EXTRA_FIELD"

	for column in mapping:
		if column not in column_category:
			column_category[column] = "PRIMARY_FIELD"

	return mapping, extra_fields, column_category


def build_pair_scores(
	df: pd.DataFrame,
	mapping: Dict[str, str],
	json_records: Dict[str, Dict[str, str]],
	header_row_number: int,
) -> List[PairScore]:
	scores: List[PairScore] = []
	comparable_columns = list(mapping.keys())

	for row_idx in range(len(df)):
		row = df.iloc[row_idx]
		excel_row_number = header_row_number + 1 + row_idx

		for record_id, field_map in json_records.items():
			compared_count = 0
			match_count = 0

			for column in comparable_columns:
				field_name = mapping[column]
				excel_value = value_to_str(row[column])
				json_value = value_to_str(field_map.get(field_name, ""))

				compared_count += 1
				if excel_value == json_value:
					match_count += 1

			mismatch_count = compared_count - match_count
			ratio = (match_count / compared_count) if compared_count else 0.0
			scores.append(
				PairScore(
					excel_row_idx=row_idx,
					excel_row_number=excel_row_number,
					record_id=record_id,
					exact_match_count=match_count,
					compared_field_count=compared_count,
					mismatch_count=mismatch_count,
					match_ratio=ratio,
				)
			)

	return scores


def assign_best_unique_pairs(pair_scores: List[PairScore]) -> Dict[int, PairScore]:
	sorted_scores = sorted(
		pair_scores,
		key=lambda s: (
			-s.exact_match_count,
			s.mismatch_count,
			-s.match_ratio,
			s.excel_row_number,
			s.record_id,
		),
	)

	assigned_rows: Set[int] = set()
	assigned_records: Set[str] = set()
	assignments: Dict[int, PairScore] = {}

	for score in sorted_scores:
		if score.excel_row_idx in assigned_rows:
			continue
		if score.record_id in assigned_records:
			continue
		assignments[score.excel_row_idx] = score
		assigned_rows.add(score.excel_row_idx)
		assigned_records.add(score.record_id)

	return assignments


def compare_assigned_pairs(
	df: pd.DataFrame,
	mapping: Dict[str, str],
	json_records: Dict[str, Dict[str, str]],
	assignments: Dict[int, PairScore],
	extra_fields: List[str],
	group_labels_by_column: Dict[str, str],
	column_category: Dict[str, str],
) -> Tuple[List[Dict[str, object]], Dict[str, float]]:
	rows: List[Dict[str, object]] = []
	all_columns_for_comparison = list(df.columns)

	# Metrics including blank-vs-blank matches.
	total_all = 0
	matches_all = 0
	tp_all = 0
	fp_all = 0
	fn_all = 0

	# Metrics excluding pairs where both sides are blank.
	total_non_blank = 0
	matches_non_blank = 0
	tp_non_blank = 0
	fp_non_blank = 0
	fn_non_blank = 0

	for row_idx in sorted(assignments.keys()):
		score = assignments[row_idx]
		row = df.iloc[row_idx]
		field_map = json_records[score.record_id]

		for column in all_columns_for_comparison:
			field_name = mapping.get(column, "")
			excel_value = value_to_str(row[column])
			json_value = value_to_str(field_map.get(field_name, "")) if field_name else ""
			is_blank_pair = excel_value == "" and json_value == ""
			decision = "MATCH" if excel_value == json_value else "MISMATCH"
			is_primary = column_category.get(column, "EXTRA_FIELD") == "PRIMARY_FIELD"
			extra_risk_flag = "RISK" if (not is_primary and excel_value != "") else "NO_RISK"

			# Primary 33 fields drive metrics; extras are informational/risk only.
			if is_primary:
				total_all += 1
				if decision == "MATCH":
					matches_all += 1

				# Including-blank metrics.
				if excel_value == json_value:
					tp_all += 1
				elif excel_value == "" and json_value != "":
					fp_all += 1
				elif excel_value != "" and json_value == "":
					fn_all += 1
				elif excel_value != "" and json_value != "" and excel_value != json_value:
					fp_all += 1
					fn_all += 1

				# Non-blank-only metrics.
				if not is_blank_pair:
					total_non_blank += 1
					if decision == "MATCH":
						matches_non_blank += 1

					if excel_value == json_value:
						tp_non_blank += 1
					elif excel_value == "" and json_value != "":
						fp_non_blank += 1
					elif excel_value != "" and json_value == "":
						fn_non_blank += 1
					elif excel_value != "" and json_value != "" and excel_value != json_value:
						fp_non_blank += 1
						fn_non_blank += 1

			rows.append(
				{
					"rowType": "DETAIL",
					"excel_row_number": score.excel_row_number,
					"assigned_recordId": score.record_id,
					"excel_group_row2": group_labels_by_column.get(column, "BASE"),
					"excel_column": column,
					"json_fieldName": field_name,
					"mapping_state": "PRIMARY_FIELD" if is_primary else "EXTRA_FIELD",
					"excel_value": excel_value,
					"json_mappedValue": json_value,
					"decision": decision,
					"blank_match": "YES" if is_blank_pair else "NO",
					"extra_field_risk": extra_risk_flag,
					"row_record_match_ratio": round(score.match_ratio, 6),
					"pair_exact_match_count": score.exact_match_count,
					"pair_compared_field_count": score.compared_field_count,
				}
			)

	accuracy_all = (matches_all / total_all) if total_all else 0.0
	precision_all = (tp_all / (tp_all + fp_all)) if (tp_all + fp_all) else 0.0
	recall_all = (tp_all / (tp_all + fn_all)) if (tp_all + fn_all) else 0.0

	accuracy_non_blank = (matches_non_blank / total_non_blank) if total_non_blank else 0.0
	precision_non_blank = (tp_non_blank / (tp_non_blank + fp_non_blank)) if (tp_non_blank + fp_non_blank) else 0.0
	recall_non_blank = (tp_non_blank / (tp_non_blank + fn_non_blank)) if (tp_non_blank + fn_non_blank) else 0.0

	metrics = {
		"primary_fields_expected_count": float(len(PRIMARY_CANONICAL_FIELDS)),
		"primary_fields_mapped_count": float(len(mapping)),
		"extra_fields_count": float(len(extra_fields)),
		"extra_fields_with_value_risk_count": float(
			sum(1 for r in rows if r.get("mapping_state") == "EXTRA_FIELD" and r.get("extra_field_risk") == "RISK")
		),
		"total_compared_fields_all": float(total_all),
		"total_matches_all": float(matches_all),
		"total_mismatches_all": float(total_all - matches_all),
		"blank_pair_matches": float(tp_all - tp_non_blank),
		"accuracy_including_blanks": accuracy_all,
		"precision_including_blanks": precision_all,
		"recall_including_blanks": recall_all,
		"tp_including_blanks": float(tp_all),
		"fp_including_blanks": float(fp_all),
		"fn_including_blanks": float(fn_all),
		"total_compared_fields_non_blank": float(total_non_blank),
		"total_matches_non_blank": float(matches_non_blank),
		"total_mismatches_non_blank": float(total_non_blank - matches_non_blank),
		"accuracy_non_blank_only": accuracy_non_blank,
		"precision_non_blank_only": precision_non_blank,
		"recall_non_blank_only": recall_non_blank,
		"tp_non_blank_only": float(tp_non_blank),
		"fp_non_blank_only": float(fp_non_blank),
		"fn_non_blank_only": float(fn_non_blank),
	}

	return rows, metrics


def append_summary_rows(
	rows: List[Dict[str, object]],
	metrics: Dict[str, float],
	extra_fields: List[str],
	unassigned_excel_rows: List[int],
	unassigned_record_ids: List[str],
) -> List[Dict[str, object]]:
	out = list(rows)

	out.append(
		{
			"rowType": "SUMMARY",
			"excel_row_number": "",
			"assigned_recordId": "",
			"excel_column": "",
			"json_fieldName": "",
			"excel_value": "",
			"json_mappedValue": "",
			"decision": "",
			"row_record_match_ratio": "",
			"pair_exact_match_count": "",
			"pair_compared_field_count": "",
			"metric": "accuracy_including_blanks",
			"metric_value": metrics["accuracy_including_blanks"],
			"notes": "Exact case-sensitive equality across all comparisons, including blank-vs-blank",
		}
	)
	out.append(
		{
			"rowType": "SUMMARY",
			"excel_row_number": "",
			"assigned_recordId": "",
			"excel_column": "",
			"json_fieldName": "",
			"excel_value": "",
			"json_mappedValue": "",
			"decision": "",
			"row_record_match_ratio": "",
			"pair_exact_match_count": "",
			"pair_compared_field_count": "",
			"metric": "precision_including_blanks",
			"metric_value": metrics["precision_including_blanks"],
			"notes": "TP / (TP + FP), counting blank-vs-blank as TP",
		}
	)
	out.append(
		{
			"rowType": "SUMMARY",
			"excel_row_number": "",
			"assigned_recordId": "",
			"excel_column": "",
			"json_fieldName": "",
			"excel_value": "",
			"json_mappedValue": "",
			"decision": "",
			"row_record_match_ratio": "",
			"pair_exact_match_count": "",
			"pair_compared_field_count": "",
			"metric": "recall_including_blanks",
			"metric_value": metrics["recall_including_blanks"],
			"notes": "TP / (TP + FN), counting blank-vs-blank as TP",
		}
	)
	out.append(
		{
			"rowType": "SUMMARY",
			"excel_row_number": "",
			"assigned_recordId": "",
			"excel_column": "",
			"json_fieldName": "",
			"excel_value": "",
			"json_mappedValue": "",
			"decision": "",
			"row_record_match_ratio": "",
			"pair_exact_match_count": "",
			"pair_compared_field_count": "",
			"metric": "accuracy_non_blank_only",
			"metric_value": metrics["accuracy_non_blank_only"],
			"notes": "Exact case-sensitive equality excluding blank-vs-blank pairs",
		}
	)
	out.append(
		{
			"rowType": "SUMMARY",
			"excel_row_number": "",
			"assigned_recordId": "",
			"excel_column": "",
			"json_fieldName": "",
			"excel_value": "",
			"json_mappedValue": "",
			"decision": "",
			"row_record_match_ratio": "",
			"pair_exact_match_count": "",
			"pair_compared_field_count": "",
			"metric": "precision_non_blank_only",
			"metric_value": metrics["precision_non_blank_only"],
			"notes": "TP / (TP + FP) excluding blank-vs-blank pairs",
		}
	)
	out.append(
		{
			"rowType": "SUMMARY",
			"excel_row_number": "",
			"assigned_recordId": "",
			"excel_column": "",
			"json_fieldName": "",
			"excel_value": "",
			"json_mappedValue": "",
			"decision": "",
			"row_record_match_ratio": "",
			"pair_exact_match_count": "",
			"pair_compared_field_count": "",
			"metric": "recall_non_blank_only",
			"metric_value": metrics["recall_non_blank_only"],
			"notes": "TP / (TP + FN) excluding blank-vs-blank pairs",
		}
	)
	out.append(
		{
			"rowType": "SUMMARY",
			"excel_row_number": "",
			"assigned_recordId": "",
			"excel_column": "",
			"json_fieldName": "",
			"excel_value": "",
			"json_mappedValue": "",
			"decision": "",
			"row_record_match_ratio": "",
			"pair_exact_match_count": "",
			"pair_compared_field_count": "",
			"metric": "blank_pair_matches",
			"metric_value": metrics["blank_pair_matches"],
			"notes": "Count of comparisons where both Excel and JSON were blank",
		}
	)
	out.append(
		{
			"rowType": "SUMMARY",
			"excel_row_number": "",
			"assigned_recordId": "",
			"excel_column": "",
			"json_fieldName": "",
			"excel_value": "",
			"json_mappedValue": "",
			"decision": "",
			"row_record_match_ratio": "",
			"pair_exact_match_count": "",
			"pair_compared_field_count": "",
			"metric": "primary_fields_expected_count",
			"metric_value": metrics["primary_fields_expected_count"],
			"notes": "Primary JSON fields expected for scoring",
		}
	)
	out.append(
		{
			"rowType": "SUMMARY",
			"excel_row_number": "",
			"assigned_recordId": "",
			"excel_column": "",
			"json_fieldName": "",
			"excel_value": "",
			"json_mappedValue": "",
			"decision": "",
			"row_record_match_ratio": "",
			"pair_exact_match_count": "",
			"pair_compared_field_count": "",
			"metric": "primary_fields_mapped_count",
			"metric_value": metrics["primary_fields_mapped_count"],
			"notes": "Primary fields mapped from Excel headers",
		}
	)
	out.append(
		{
			"rowType": "SUMMARY",
			"excel_row_number": "",
			"assigned_recordId": "",
			"excel_column": "",
			"json_fieldName": "",
			"excel_value": "",
			"json_mappedValue": "",
			"decision": "",
			"row_record_match_ratio": "",
			"pair_exact_match_count": "",
			"pair_compared_field_count": "",
			"metric": "extra_fields_count",
			"metric_value": metrics["extra_fields_count"],
			"notes": " | ".join(extra_fields),
		}
	)
	out.append(
		{
			"rowType": "SUMMARY",
			"excel_row_number": "",
			"assigned_recordId": "",
			"excel_column": "",
			"json_fieldName": "",
			"excel_value": "",
			"json_mappedValue": "",
			"decision": "",
			"row_record_match_ratio": "",
			"pair_exact_match_count": "",
			"pair_compared_field_count": "",
			"metric": "extra_fields_with_value_risk_count",
			"metric_value": metrics["extra_fields_with_value_risk_count"],
			"notes": "Count of extra-field cells that contain value and are flagged RISK",
		}
	)
	out.append(
		{
			"rowType": "SUMMARY",
			"excel_row_number": "",
			"assigned_recordId": "",
			"excel_column": "",
			"json_fieldName": "",
			"excel_value": "",
			"json_mappedValue": "",
			"decision": "",
			"row_record_match_ratio": "",
			"pair_exact_match_count": "",
			"pair_compared_field_count": "",
			"metric": "unassigned_excel_rows_count",
			"metric_value": len(unassigned_excel_rows),
			"notes": " | ".join(str(v) for v in unassigned_excel_rows),
		}
	)
	out.append(
		{
			"rowType": "SUMMARY",
			"excel_row_number": "",
			"assigned_recordId": "",
			"excel_column": "",
			"json_fieldName": "",
			"excel_value": "",
			"json_mappedValue": "",
			"decision": "",
			"row_record_match_ratio": "",
			"pair_exact_match_count": "",
			"pair_compared_field_count": "",
			"metric": "unassigned_json_records_count",
			"metric_value": len(unassigned_record_ids),
			"notes": " | ".join(unassigned_record_ids),
		}
	)

	return out


def append_compact_summary_rows(rows: List[Dict[str, object]], metrics: Dict[str, float]) -> List[Dict[str, object]]:
	"""Append a compact roll-up block at the end of the same CSV file."""
	out = list(rows)
	detail_rows = [r for r in out if r.get("rowType") == "DETAIL"]
	primary_rows = [r for r in detail_rows if r.get("mapping_state") == "PRIMARY_FIELD"]
	extra_rows = [r for r in detail_rows if r.get("mapping_state") == "EXTRA_FIELD"]
	risk_rows = [r for r in extra_rows if r.get("extra_field_risk") == "RISK"]

	unique_assignments = sorted({str(r.get("assigned_recordId", "")) for r in detail_rows if r.get("assigned_recordId")})
	primary_matches = sum(1 for r in primary_rows if r.get("decision") == "MATCH")

	risk_field_set = sorted({str(r.get("excel_column", "")) for r in risk_rows if r.get("excel_column")})
	risk_preview = " | ".join(risk_field_set[:10])
	if len(risk_field_set) > 10:
		risk_preview += f" | ... (+{len(risk_field_set) - 10} more)"

	compact_rows = [
		{
			"rowType": "COMPACT_SUMMARY",
			"metric": "assigned_records",
			"metric_value": len(unique_assignments),
			"notes": " | ".join(unique_assignments),
		},
		{
			"rowType": "COMPACT_SUMMARY",
			"metric": "primary_fields_mapped_count",
			"metric_value": metrics["primary_fields_mapped_count"],
			"notes": "Primary comparison scope (target: 33)",
		},
		{
			"rowType": "COMPACT_SUMMARY",
			"metric": "primary_matches",
			"metric_value": primary_matches,
			"notes": f"Out of {len(primary_rows)} primary comparisons",
		},
		{
			"rowType": "COMPACT_SUMMARY",
			"metric": "accuracy_including_blanks",
			"metric_value": metrics["accuracy_including_blanks"],
			"notes": "Primary fields only",
		},
		{
			"rowType": "COMPACT_SUMMARY",
			"metric": "accuracy_non_blank_only",
			"metric_value": metrics["accuracy_non_blank_only"],
			"notes": "Primary fields only",
		},
		{
			"rowType": "COMPACT_SUMMARY",
			"metric": "extra_fields_count",
			"metric_value": len(extra_rows),
			"notes": "All extra-field comparisons across assigned rows",
		},
		{
			"rowType": "COMPACT_SUMMARY",
			"metric": "extra_fields_with_value_risk_count",
			"metric_value": len(risk_rows),
			"notes": risk_preview,
		},
	]

	for row in compact_rows:
		row.setdefault("excel_row_number", "")
		row.setdefault("assigned_recordId", "")
		row.setdefault("excel_group_row2", "")
		row.setdefault("excel_column", "")
		row.setdefault("json_fieldName", "")
		row.setdefault("mapping_state", "")
		row.setdefault("excel_value", "")
		row.setdefault("json_mappedValue", "")
		row.setdefault("decision", "")
		row.setdefault("blank_match", "")
		row.setdefault("extra_field_risk", "")
		row.setdefault("row_record_match_ratio", "")
		row.setdefault("pair_exact_match_count", "")
		row.setdefault("pair_compared_field_count", "")

	out.extend(compact_rows)
	return out


def run_comparison(
	excel_path: str,
	json_path: str,
	output_csv_path: str,
	sheet_name: str,
	header_row_number: int,
):
	df = load_excel(excel_path, sheet_name=sheet_name, header_row_number=header_row_number)
	group_labels = load_excel_group_row(excel_path, sheet_name, header_row_number, len(df.columns))
	group_labels_by_column = {column: group_labels[idx] for idx, column in enumerate(df.columns)}
	json_records, alias_by_header_key, primary_field_names = load_json_records(json_path)

	json_field_names: Set[str] = set()
	for field_map in json_records.values():
		json_field_names.update(field_map.keys())

	mapping, extra_fields, column_category = build_column_mapping(
		list(df.columns),
		group_labels,
		json_field_names,
		primary_field_names,
		alias_by_header_key,
	)
	if not mapping:
		raise ValueError(
			"No Excel columns could be mapped to primary JSON fields. "
			"Please verify sheet selection and header row."
		)

	pair_scores = build_pair_scores(df, mapping, json_records, header_row_number)
	assignments = assign_best_unique_pairs(pair_scores)

	assigned_record_ids = {score.record_id for score in assignments.values()}
	unassigned_record_ids = sorted(set(json_records.keys()) - assigned_record_ids)
	unassigned_excel_rows = sorted(
		header_row_number + 1 + idx
		for idx in range(len(df))
		if idx not in assignments
	)

	detail_rows, metrics = compare_assigned_pairs(
		df,
		mapping,
		json_records,
		assignments,
		extra_fields,
		group_labels_by_column,
		column_category,
	)
	final_rows = append_summary_rows(
		detail_rows,
		metrics,
		extra_fields,
		unassigned_excel_rows,
		unassigned_record_ids,
	)
	final_rows = append_compact_summary_rows(final_rows, metrics)

	output_dir = os.path.dirname(output_csv_path)
	if output_dir:
		os.makedirs(output_dir, exist_ok=True)
	pd.DataFrame(final_rows).to_csv(output_csv_path, index=False, encoding="utf-8-sig")

	print("Comparison completed.")
	print(f"Excel rows loaded: {len(df)}")
	print(f"JSON records loaded: {len(json_records)}")
	print(f"Primary fields mapped (of 33): {len(mapping)}")
	print(f"Extra Excel headers: {len(extra_fields)}")
	print(f"Total headers compared per row: {len(df.columns)}")
	if extra_fields:
		print("Extra header list:")
		for header in extra_fields:
			print(f"  - {header}")
	print(f"Assigned row-record pairs: {len(assignments)}")
	print(f"Unassigned Excel rows: {len(unassigned_excel_rows)}")
	print(f"Unassigned JSON records: {len(unassigned_record_ids)}")
	print(f"Extra-field risk cells: {int(metrics['extra_fields_with_value_risk_count'])}")

	print("\nOverall metrics")
	print("  Including blank-vs-blank matches")
	print(f"    Accuracy : {metrics['accuracy_including_blanks']:.6f}")
	print(f"    Precision: {metrics['precision_including_blanks']:.6f}")
	print(f"    Recall   : {metrics['recall_including_blanks']:.6f}")
	print(
		f"    TP={int(metrics['tp_including_blanks'])}, FP={int(metrics['fp_including_blanks'])}, FN={int(metrics['fn_including_blanks'])}"
	)
	print(f"    Blank pair matches: {int(metrics['blank_pair_matches'])}")

	print("  Excluding blank-vs-blank matches")
	print(f"    Accuracy : {metrics['accuracy_non_blank_only']:.6f}")
	print(f"    Precision: {metrics['precision_non_blank_only']:.6f}")
	print(f"    Recall   : {metrics['recall_non_blank_only']:.6f}")
	print(
		f"    TP={int(metrics['tp_non_blank_only'])}, FP={int(metrics['fp_non_blank_only'])}, FN={int(metrics['fn_non_blank_only'])}"
	)
	print(f"\nCSV report: {output_csv_path}")


def normalize_sheet_or_file_key(text: str) -> str:
	"""Concatenate lowercase alnum tokens, dropping 'pdf' so sheet/file naming variants line up."""
	tokens = re.findall(r"[a-z0-9]+", text.lower())
	return "".join(token for token in tokens if token != "pdf")


def extract_json_document_key(json_stem: str) -> str:
	"""Derive the source-document key from an LLM output stem, e.g.
	'output_output_ADI_bracebridge_selectable_text_gpt-5.4_intermediate' -> 'bracebridgeselectabletext'.
	"""
	match = re.match(r"^output_(?:output_)?(?:ADI_)?(.+?)_gpt[-.\w]*?(?:_intermediate)?$", json_stem)
	core = match.group(1) if match else json_stem
	return normalize_sheet_or_file_key(core)


def match_sheet_for_json(json_key: str, sheet_names: List[str]) -> str:
	"""Find the workbook sheet whose normalized key best matches the JSON document key."""
	sheet_keys = {sheet: normalize_sheet_or_file_key(sheet) for sheet in sheet_names}

	for sheet, key in sheet_keys.items():
		if key == json_key:
			return sheet

	for sheet, key in sheet_keys.items():
		if key and json_key and (key.startswith(json_key) or json_key.startswith(key)):
			return sheet

	return ""


def run_batch_comparison(
	excel_path: str,
	sheet_names: List[str],
	header_row_number: int,
	input_dir: str,
	output_dir: str,
):
	if not os.path.isdir(input_dir):
		print(f"[SKIP] Input directory not found: {input_dir}")
		return

	os.makedirs(output_dir, exist_ok=True)
	json_files = sorted(f for f in os.listdir(input_dir) if f.lower().endswith(".json"))

	for json_file in json_files:
		json_stem = os.path.splitext(json_file)[0]
		json_key = extract_json_document_key(json_stem)
		sheet_name = match_sheet_for_json(json_key, sheet_names)

		if not sheet_name:
			print(f"[SKIP] No matching Excel sheet found for '{json_file}' (key='{json_key}')")
			continue

		json_path = os.path.join(input_dir, json_file)
		output_csv_path = os.path.join(output_dir, f"comparison_{json_stem}.csv")

		print(f"\n=== Comparing '{json_file}' against sheet '{sheet_name}' ===")
		try:
			run_comparison(excel_path, json_path, output_csv_path, sheet_name, header_row_number)
		except Exception as exc:
			print(f"[ERROR] Failed comparing '{json_file}': {exc}")


def parse_args() -> argparse.Namespace:
	parser = argparse.ArgumentParser(
		description=(
			"Compare Excel expected values with JSON mappedValue using exact case-sensitive matching, "
			"across every intermediate and final LLM output JSON."
		)
	)
	parser.add_argument(
		"--excel",
		default="./utility_files/SSI_Golden_Set_Expected_Extraction_Output.xlsx",
		help="Path to expected Excel file (golden set with one sheet per source document).",
	)
	parser.add_argument(
		"--header-row",
		type=int,
		default=3,
		help="Header row number in Excel (1-based).",
	)
	parser.add_argument(
		"--intermediate-dir",
		default="./output_llm_on_ADI/intermediate_llm_json",
		help="Directory containing intermediate (pre-contract) LLM output JSON files.",
	)
	parser.add_argument(
		"--final-dir",
		default="./output_llm_on_ADI/final_processed_json",
		help="Directory containing final (post-contract) LLM output JSON files.",
	)
	parser.add_argument(
		"--intermediate-out-dir",
		default="./comparision_results/intermediate_jsons",
		help="Directory to save comparison CSVs for intermediate JSON files.",
	)
	parser.add_argument(
		"--final-out-dir",
		default="./comparision_results/final_jsons",
		help="Directory to save comparison CSVs for final JSON files.",
	)
	return parser.parse_args()


def main():
	args = parse_args()
	sheet_names = pd.ExcelFile(args.excel).sheet_names
	print(f"Discovered {len(sheet_names)} sheets in '{args.excel}': {sheet_names}")

	run_batch_comparison(
		args.excel, sheet_names, args.header_row, args.intermediate_dir, args.intermediate_out_dir
	)
	run_batch_comparison(
		args.excel, sheet_names, args.header_row, args.final_dir, args.final_out_dir
	)


if __name__ == "__main__":
	main()
