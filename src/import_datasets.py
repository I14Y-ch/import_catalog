import pandas as pd
import requests
from datetime import datetime
import json
from config import *

def create_language_object(text, lang="de"):
    """Create a language object with only German"""
    return {lang: text}

def create_uri_label_object(uri, label=None):
    """Create an object with URI and optional label"""
    obj = {"uri": uri}
    if label:
        obj["label"] = create_language_object(label)
    return obj

def process_keywords(row):
    """Process keywords from the Excel row"""
    keywords = []
    for i in range(1, 4):
        if pd.notna(row.get(f'keywords_{i}')):
            keywords.append(create_language_object(row[f'keywords_{i}']))
    return keywords

def process_distribution(row, index):
    """Process a single distribution"""
    if pd.isna(row.get(f'distribution_accessUrl_{index}')) and pd.isna(row.get(f'distribution_downloadUrl_{index}')):
        return None

    distribution = {}
    
    if pd.notna(row.get(f'distribution_accessUrl_{index}')):
        distribution["accessUrl"] = create_uri_label_object(row[f'distribution_accessUrl_{index}'])
    
    if pd.notna(row.get(f'distribution_downloadUrl_{index}')):
        distribution["downloadUrl"] = create_uri_label_object(row[f'distribution_downloadUrl_{index}'])
    
    if pd.notna(row.get(f'distribution_license_label_{index}')):
        distribution["license"] = {"code": row[f'distribution_license_label_{index}']}

    return distribution

def create_dataset_payload(row):
    """Transform Excel row into I14Y API payload"""
    payload = {
        "data": {
            "title": create_language_object(row['title']),
            "description": create_language_object(row['description']),
            "identifiers": [row['identificator']],
            "publisher": DEFAULT_PUBLISHER,
            "accessRights": {"code": row['accessRights']},
            "issued": row['issued'].isoformat() if pd.notna(row['issued']) else None,
            "modified": row['modified'].isoformat() if pd.notna(row['modified']) else None,
        }
    }

    # Add keywords if present
    keywords = process_keywords(row)
    if keywords:
        payload["data"]["keywords"] = keywords

    # Add contact points if present
    if pd.notna(row.get('contactPoints_fn')) or pd.notna(row.get('contactPoints_hasEmail')):
        contact_point = {
            "kind": "Organization"
        }
        if pd.notna(row.get('contactPoints_fn')):
            contact_point["fn"] = create_language_object(row['contactPoints_fn'])
        if pd.notna(row.get('contactPoints_hasEmail')):
            contact_point["hasEmail"] = row['contactPoints_hasEmail']
        if pd.notna(row.get('contactPoints_hasTelephone')):
            contact_point["hasTelephone"] = row['contactPoints_hasTelephone']
        payload["data"]["contactPoints"] = [contact_point]

    # Add themes if present
    if pd.notna(row.get('themes_label')):
        payload["data"]["themes"] = [{"code": row['themes_label']}]

    # Add spatial if present
    if pd.notna(row.get('spatial')):
        payload["data"]["spatial"] = [row['spatial']]

    # Add temporal coverage if present
    if pd.notna(row.get('temporalCoverage_start')) or pd.notna(row.get('temporalCoverage_end')):
        coverage = {}
        if pd.notna(row.get('temporalCoverage_start')):
            coverage["start"] = row['temporalCoverage_start'].isoformat()
        if pd.notna(row.get('temporalCoverage_end')):
            coverage["end"] = row['temporalCoverage_end'].isoformat()
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
    df = pd.read_excel(TEMPLATE_PATH)
    
    # Process each row
    success_count = 0
    error_count = 0
    
    print("\nStarting dataset import...\n")
    
    for idx, row in df.iterrows():
        if pd.isna(row['title']):  # Skip empty rows
            continue
            
        print(f"Processing dataset {idx + 1}: {row['identificator']}")
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
