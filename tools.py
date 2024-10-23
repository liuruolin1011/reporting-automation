import re
import os
from typing import Union, Optional

def get_file(file_name: str, folder_path: str) -> str:
    """
    Combines folder path and file name to create a full file path.
    
    Args:
        file_name (str): The name of the file.
        folder_path (str): The path to the folder.
    
    Returns:
        str: Full file path.
    """
    return os.path.join(folder_path, file_name)


def extract_restriction_code(s: Union[str, None]) -> Optional[int]:
    """
    Extracts the posting restriction code from a given string.
    
    Args:
        s (str): String containing the restriction code.
    
    Returns:
        int or None: The extracted restriction code or None if not found.
    """
    if isinstance(s, str):
        match = re.search(r'\b\d+\b', s)
        if match:
            return int(match.group())
    return None


def extract_sector_code(s: Union[str, None]) -> Optional[int]:
    """
    Extracts the sector code (digits) from a string.
    
    Args:
        s (str): String containing the sector code.
    
    Returns:
        int or None: The extracted sector code or None if not found.
    """
    if isinstance(s, str):
        match = re.search(r'\d+', s)
        if match:
            return int(match.group())
    return None


def extract_sector_description(cell_value: Union[str, None]) -> str:
    """
    Extracts the sector description by splitting the string at ' - '.
    
    Args:
        cell_value (str): The full string containing sector description.
    
    Returns:
        str: The sector description before the ' - ' separator.
    """
    cell_value = str(cell_value)  # Ensure it is a string
    return cell_value.split(" - ")[0] if " - " in cell_value else cell_value


def extract_branch(s: Union[str, None]) -> Optional[str]:
    """
    Extracts the branch or business unit enclosed in parentheses.
    
    Args:
        s (str): String containing the branch in parentheses.
    
    Returns:
        str or None: The extracted branch or None if not found.
    """
    if isinstance(s, str):
        match = re.search(r'\((.*?)\)', s)
        if match:
            return match.group(1)
    return None


# Map legal entity types to segment codes
le_type_map = {
    'Non-Natural Person': 'NNP',
    'Natural Person': 'NP',
    'Financial Institution': 'FI'
}

def extract_segment(le_type: str) -> str:
    """
    Maps the legal entity type to a segment.
    
    Args:
        le_type (str): Legal entity type.
    
    Returns:
        str: Mapped segment or original type if not found.
    """
    return le_type_map.get(le_type, le_type)


def pep_type(row: dict) -> str:
    """
    Determines the PEP (Politically Exposed Person) type based on multiple columns.
    
    Args:
        row (dict): A row of data containing 'Is PEP', 'PEP Associated', and 'LE Type'.
    
    Returns:
        str: The PEP type (e.g., 'NP PEP', 'NNP PEP Associated') or an empty string.
    """
    if row.get('Is PEP') == 'Yes':
        return 'NP PEP' if row.get('LE Type') == 'NP' else 'NNP PEP'
    elif row.get('Is PEP') == 'No' and row.get('PEP Associated') == 'Yes':
        return 'NP PEP Associated' if row.get('LE Type') == 'NP' else 'NNP PEP Associated'
    return ''


def segment_type(row: dict) -> str:
    """
    Determines the segment type based on the 'LE Type' and 'Business Unit'.
    
    Args:
        row (dict): A row of data containing 'LE Type' and 'Business Unit'.
    
    Returns:
        str: The segment type ('NP', 'FI', or 'NNP').
    """
    if row.get('LE Type') == 'NP':
        return 'NP'
    elif row.get('Business Unit') == 'FID':
        return 'FI'
    return 'NNP'


def offboarding_reason(s: str) -> Optional[str]:
    """
    Extracts the offboarding reason from a string enclosed in parentheses.
    
    Args:
        s (str): String containing the offboarding reason.
    
    Returns:
        str or None: The extracted offboarding reason or None if not found.
    """
    match = re.search(r'\((.*?)\)', s)
    return match.group(1) if match else None
