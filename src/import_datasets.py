import pandas as pd
import requests
from datetime import datetime
import json
import sys
from config import *
from codelist_utils import map_theme_to_code, map_license_to_code, map_access_rights_to_code

def create_language_object(text, lang="de"):
    """Create a language object with only German"""
    return {lang: text}

def create_uri_label_object(uri, label=None):
    """Create an object with URI and optional label"""
    obj = {"uri": uri}
    if label:
        obj["label"] = create_language_object(label)
    return obj

def safe_get(row, key, default=None):
    """Safely get a value from a row, whether it's a Series, dict, or other object"""
    if isinstance(row, dict):
        return row.get(key, default)
    elif isinstance(row, pd.Series):
        return row[key] if key in row and pd.notna(row[key]) else default
    else:
        try:
            val = getattr(row, key, default)
            return val if pd.notna(val) else default
        except:
            return default

def process_keywords(row):
    """Process keywords from the Excel row"""
    keywords = []
    for i in range(1, 4):
        key = f'keywords_{i}'
        value = safe_get(row, key)
        if pd.notna(value):
            keywords.append(create_language_object(value))
    return keywords

def process_distribution(row, index):
    """Process a single distribution"""
    access_key = f'distribution_accessUrl_{index}'
    download_key = f'distribution_downloadUrl_{index}'
    license_key = f'distribution_license_label_{index}'
    
    access_url = safe_get(row, access_key)
    download_url = safe_get(row, download_key)
    
    if pd.isna(access_url) and pd.isna(download_url):
        return None

    distribution = {}
    
    # According to API validation, download URL must be the same as access URL
    if pd.notna(access_url) and pd.notna(download_url):
        # Use the access URL for both to satisfy API requirement
        url_to_use = access_url
        distribution["accessUrl"] = create_uri_label_object(url_to_use)
        distribution["downloadUrl"] = create_uri_label_object(url_to_use)
    elif pd.notna(access_url):
        # If only access URL is provided, use it for both
        distribution["accessUrl"] = create_uri_label_object(access_url)
        distribution["downloadUrl"] = create_uri_label_object(access_url)
    elif pd.notna(download_url):
        # If only download URL is provided, use it for both
        distribution["accessUrl"] = create_uri_label_object(download_url)
        distribution["downloadUrl"] = create_uri_label_object(download_url)
    
    license_value = safe_get(row, license_key)
    if pd.notna(license_value):
        license_code = map_license_to_code(license_value)
        if license_code:
            distribution["license"] = {"code": license_code}

    distribution['title'] = {'de': 'Datenexport'}
    distribution['description'] = {'de': 'Export der Daten'}

    return distribution

def create_dataset_payload(row):
    """Transform Excel row into I14Y API payload"""
    # Get mapped codes for required fields
    access_rights_value = safe_get(row, 'accessRights')
    access_rights_code = map_access_rights_to_code(access_rights_value)
    
    payload = {
        "data": {
            "title": create_language_object(row['title']),
            "description": create_language_object(row['description']),
            "identifiers": [row['identificator']],
            "publisher": DEFAULT_PUBLISHER,
            "accessRights": {"code": access_rights_code or "PUBLIC"},
            "issued": row['issued'].isoformat() if pd.notna(row['issued']) else None,
            "modified": row['modified'].isoformat() if pd.notna(row['modified']) else None,
        }
    }

    # Add keywords if present
    keywords = process_keywords(row)
    if keywords:
        payload["data"]["keywords"] = keywords

    # Add contact points if present
    contact_fn = safe_get(row, 'contactPoints_fn')
    contact_email = safe_get(row, 'contactPoints_hasEmail')
    contact_phone = safe_get(row, 'contactPoints_hasTelephone')
    
    if pd.notna(contact_fn) or pd.notna(contact_email):
        contact_point = {
            "kind": "Organization"
        }
        if pd.notna(contact_fn):
            contact_point["fn"] = create_language_object(contact_fn)
        if pd.notna(contact_email):
            contact_point["hasEmail"] = contact_email
        if pd.notna(contact_phone):
            contact_point["hasTelephone"] = contact_phone
        payload["data"]["contactPoints"] = [contact_point]

    # Add themes if present
    theme_value = safe_get(row, 'themes_label')
    if pd.notna(theme_value):
        theme_code = map_theme_to_code(theme_value)
        if theme_code:
            payload["data"]["themes"] = [{"code": theme_code}]

    # Add spatial if present
    spatial_value = safe_get(row, 'spatial')
    if pd.notna(spatial_value):
        payload["data"]["spatial"] = [spatial_value]

    # Add temporal coverage if present
    temp_start = safe_get(row, 'temporalCoverage_start')
    temp_end = safe_get(row, 'temporalCoverage_end')
    
    if pd.notna(temp_start) or pd.notna(temp_end):
        coverage = {}
        if pd.notna(temp_start):
            coverage["start"] = temp_start.isoformat()
        if pd.notna(temp_end):
            coverage["end"] = temp_end.isoformat()
        payload["data"]["temporalCoverage"] = [coverage]

    # Process distributions
    distributions = []
    for i in range(1, 4):
        dist = process_distribution(row, i)
        if dist:
            distributions.append(dist)
    
    if distributions:
        payload["data"]["distributions"] = distributions

    return payload

def submit_to_api(payload):
    """Submit dataset to I14Y API"""
    headers = {
        "Authorization": API_TOKEN,
        "Content-Type": "application/json"
    }
    
    print(payload)
    response = requests.post(
        f"{API_BASE_URL}/datasets",
        headers=headers,
        json=payload
    )
    
    if response.status_code not in (200, 201):
        raise Exception(f"API submission failed: {response.status_code} - {response.text}")
    
    return response.json()

def main():
    # Read Excel file
    try:
        df = pd.read_excel(TEMPLATE_PATH)
        print(f"Successfully loaded {len(df)} entries from {TEMPLATE_PATH}")
    except Exception as e:
        print(f"Error loading Excel file: {e}")
        sys.exit(1)
    
    # Process each row
    success_count = 0
    error_count = 0
    
    print("\nStarting dataset import...\n")
    
    for idx, row in df.iterrows():
        title_value = safe_get(row, 'title')
        if pd.isna(title_value):  # Skip empty rows
            continue
            
        identificator = safe_get(row, 'identificator', 'No ID')
        print(f"Processing dataset {idx + 1}: {identificator}")
        
        try:
            payload = create_dataset_payload(row)
            response = submit_to_api(payload)
            success_count += 1
            print(f"✓ Success - Dataset ID: {response.get('id', 'N/A')}\n")
            
        except Exception as e:
            error_count += 1
            print(f"✗ Error: {str(e)}\n")
    
    # Print summary
    print("=== Import Summary ===")
    print(f"Total processed: {success_count + error_count}")
    print(f"Successful: {success_count}")
    print(f"Failed: {error_count}")

if __name__ == "__main__":
    main()
