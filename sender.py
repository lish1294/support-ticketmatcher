import json
import pandas as pd
import re
from pathlib import Path

def read_excel_and_generate_json(excel_file_path, sheet_name=0, output_file=None):
    """
    Reads an Excel file row by row and generates JSON for each row.
    
    Args:
        excel_file_path (str): Path to the Excel file
        sheet_name (str or int): Sheet name or index (default: 0)
        output_file (str): Optional file to write JSON output. If None, prints to console.
    """
    try:
        # Read the Excel file
        df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
        
        json_list = []
        
        # Iterate through rows
        for index, row in df.iterrows():
            # Convert row to dictionary
            row_dict = row.to_dict()
            
            # Handle NaN values (convert to None for JSON compatibility)
            row_dict = {k: (None if pd.isna(v) else v) for k, v in row_dict.items()}
            
            # Convert to JSON string
            json_str = json.dumps(row_dict, indent=2, default=str)
            
            json_list.append(json_str)
            
            # Print each row's JSON
            print(f"Row {index + 1}:")
            print(json_str)
            print()
        
        # Optionally write to file
        if output_file:
            with open(output_file, 'w') as f:
                # Write as JSON array
                f.write(json.dumps([json.loads(j) for j in json_list], indent=2, default=str))
            print(f"JSON output saved to {output_file}")
        
        return json_list
        
    except FileNotFoundError:
        print(f"Error: Excel file '{excel_file_path}' not found.")
    except Exception as e:
        print(f"Error reading Excel file: {e}")


def _cast_value(value):
    """Try to cast a string value to int/float/bool if applicable."""
    if value is None:
        return None
    if isinstance(value, (int, float, bool)):
        return value
    s = str(value).strip()
    if s == "":
        return None
    # booleans
    low = s.lower()
    if low in ("true", "yes", "y"):
        return True
    if low in ("false", "no", "n"):
        return False
    # int
    try:
        return int(s)
    except Exception:
        pass
    # float
    try:
        return float(s)
    except Exception:
        pass
    # fallback
    return s


def parse_user_input(user_input):
    """
    Parse free-form user input into a dict of criteria.

    Supported formats:
      - JSON object: '{"Col": "value", "Age": 10}'
      - Comma-separated key=value pairs: 'Col=val, Age=10'
      - Colon-separated pairs: 'Col: val; Age:10'

    Returns a dict where values are cast to int/float/bool when possible.
    """
    if not user_input:
        return {}
    # Try JSON first
    try:
        data = json.loads(user_input)
        if isinstance(data, dict):
            return {str(k).strip(): _cast_value(v) for k, v in data.items()}
    except Exception:
        pass

    # Normalize separators
    # Split on commas or semicolons
    parts = [p for sep in [",", ";"] for p in user_input.split(sep)]
    # If no separators found, try splitting by newline
    if len(parts) == 1:
        parts = user_input.replace("\n", ";").split(";")

    criteria = {}
    for part in parts:
        if not part or not part.strip():
            continue
        if "=" in part:
            k, v = part.split("=", 1)
        elif ":" in part:
            k, v = part.split(":", 1)
        else:
            # single token, place under a generic key 'query'
            criteria["query"] = _cast_value(part.strip())
            continue
        criteria[str(k).strip()] = _cast_value(v.strip())

    return criteria


def _is_numeric_only_search(user_input):
    """Return True if the user input is only a number, with no descriptive text."""
    if not user_input:
        return False
    value = str(user_input).strip()
    if value == "":
        return False
    return bool(re.fullmatch(r"[+-]?\d+(?:\.\d+)?", value))


def find_matching_rows(excel_file_path, criteria, sheet_name=0, match_all=True):
    """
    Read the Excel file and return rows that match the given criteria.

    Args:
      excel_file_path: path to excel file
      criteria: dict of column->value to match (case-insensitive column matching)
      sheet_name: sheet name or index
      match_all: if True, require all criteria to match (AND); if False, match any (OR)

    Returns a list of dicts (one per matching row) with NaN converted to None.
    """
    if not Path(excel_file_path).exists():
        raise FileNotFoundError(f"Excel file '{excel_file_path}' not found")

    df = pd.read_excel(excel_file_path, sheet_name=sheet_name)

    if df.empty:
        return []

    # Build a mapping of lowercase column names to actual column names
    col_map = {c.lower(): c for c in df.columns}

    masks = []
    for key, val in criteria.items():
        if key == "query":
            # free-text search across all columns
            q = "" if val is None else str(val).strip().lower()
            if q == "":
                continue
            col_masks = []
            for c in df.columns:
                col_masks.append(df[c].astype(str).fillna("").str.lower().str.contains(q, na=False))
            masks.append(pd.concat(col_masks, axis=1).any(axis=1))
            continue

        col = col_map.get(key.lower())
        if col is None:
            # no such column, create a mask that's all False
            masks.append(pd.Series([False] * len(df)))
            continue

        if val is None:
            masks.append(df[col].isna() | (df[col].astype(str).str.strip() == ""))
        else:
            # match stringified values case-insensitively
            masks.append(df[col].astype(str).fillna("").str.strip().str.lower() == str(val).lower())

    if not masks:
        return []

    if match_all:
        combined = masks[0]
        for m in masks[1:]:
            combined = combined & m
    else:
        combined = masks[0]
        for m in masks[1:]:
            combined = combined | m

    matched = df[combined]

    results = []
    for _, row in matched.iterrows():
        row_dict = {k: (None if pd.isna(v) else v) for k, v in row.to_dict().items()}
        results.append(row_dict)

    return results


def check_excel_and_match(excel_file_path, user_input, sheet_name=0, match_all=True):
    """
    Convenience function: parse `user_input`, then search the Excel and return matches as JSON strings.
    """
    if _is_numeric_only_search(user_input):
        raise ValueError(
            "Please enter a description of the issue. If you are searching by a number, "
            "specify whether it is an error code, ticket ID, case number, or status code "
            "(example: 'error: 401')."
        )
    criteria = parse_user_input(user_input)
    rows = find_matching_rows(excel_file_path, criteria, sheet_name=sheet_name, match_all=match_all)
    return [json.dumps(r, default=str, indent=2) for r in rows]


if __name__ == "__main__":
    # Example usage
    excel_path = "tickets.xlsx"  # Replace with your Excel file path
    output_path = "output.json"  # Optional: file to save JSON output
    
    read_excel_and_generate_json(excel_path, output_file=output_path)
