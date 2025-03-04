import requests
import pandas as pd

# Cache for the codelists to avoid multiple API calls
_themes_cache = None
_license_cache = None
_access_rights_cache = None

def get_themes_codelist():
    """Fetch themes codelist from I14Y API"""
    global _themes_cache
    if _themes_cache is not None:
        return _themes_cache
    
    url = "https://api.i14y.admin.ch/api/public/v1/concepts/08da58dc-4dc8-f9cb-b6f2-7d16b3fa0cde/codelist-entries/exports/json"
    try:
        response = requests.get(url)
        response.raise_for_status()
        data = response.json()
        
        result = {}
        for item in data['data']:
            code = item.get('code')
            label = item.get('name', {}).get('de', '')
            if code and label:
                result[label] = code
                # Also store the code as key for direct mapping
                result[code] = code
        
        _themes_cache = result
        return result
    except Exception as e:
        print(f"Error fetching themes codelist: {e}")
        return {}

def get_license_codelist():
    """Fetch license codelist from I14Y API"""
    global _license_cache
    if _license_cache is not None:
        return _license_cache
    
    url = "https://api.i14y.admin.ch/api/public/v1/concepts/08db7eb9-8d92-b301-982e-5f7cbd44e45f/codelist-entries/exports/json"
    try:
        response = requests.get(url)
        response.raise_for_status()
        data = response.json()
        
        result = {'Unknown': 'UNKNOWN'}  # Default mapping
        for item in data['data']:
            code = item.get('code')
            label = item.get('name', {}).get('de', '')
            if code and label:
                result[label] = code
                # Also store the code as key for direct mapping
                result[code] = code
        
        _license_cache = result
        return result
    except Exception as e:
        print(f"Error fetching license codelist: {e}")
        return {'Unknown': 'UNKNOWN'}

def get_access_rights_codelist():
    """Create access rights codelist"""
    global _access_rights_cache
    if _access_rights_cache is not None:
        return _access_rights_cache
    
    mapping = {
        'Nicht-öffentlich': 'NON_PUBLIC',
        'Öffentlich': 'PUBLIC',
        'Eingeschränkt': 'RESTRICTED',
        'Vertraulich': 'CONFIDENTIAL',
        # Also include direct code mappings
        'NON_PUBLIC': 'NON_PUBLIC',
        'PUBLIC': 'PUBLIC',
        'RESTRICTED': 'RESTRICTED',
        'CONFIDENTIAL': 'CONFIDENTIAL'
    }
    
    _access_rights_cache = mapping
    return mapping

def map_theme_to_code(theme_value):
    """Map a theme label or code to its code"""
    if pd.isna(theme_value) or not theme_value:
        return None
        
    themes_map = get_themes_codelist()
    return themes_map.get(theme_value, theme_value)

def map_license_to_code(license_value):
    """Map a license label or code to its code"""
    if pd.isna(license_value) or not license_value:
        return None
        
    license_map = get_license_codelist()
    return license_map.get(license_value, license_value)

def map_access_rights_to_code(access_rights_value):
    """Map an access rights label or code to its code"""
    if pd.isna(access_rights_value) or not access_rights_value:
        return None
        
    access_rights_map = get_access_rights_codelist()
    return access_rights_map.get(access_rights_value, access_rights_value)
