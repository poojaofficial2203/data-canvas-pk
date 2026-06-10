"""Utility functions for data processing and file handling"""

import os
from pathlib import Path
import pandas as pd
import json
import csv
import io
from typing import Optional


def parse_raw_data(raw_text, delimiter=None):
    """
    Parse raw text data into DataFrame with auto-detection.
    Supports CSV, TSV, pipe-delimited, space-separated, and other formats.
    
    Args:
        raw_text (str): Raw text data
        delimiter (str): Optional specific delimiter to use
        
    Returns:
        pd.DataFrame: Parsed data
    """
    raw_text = raw_text.strip()
    
    if not raw_text:
        raise ValueError("Empty data provided")
    
    if delimiter:
        try:
            return pd.read_csv(io.StringIO(raw_text), delimiter=delimiter)
        except Exception as e:
            raise ValueError(f"Failed to parse with delimiter '{delimiter}': {e}")
    
    delimiters_to_try = [',', '\t', '|', ';', ' ']
    
    for delim in delimiters_to_try:
        try:
            df = pd.read_csv(io.StringIO(raw_text), delimiter=delim)
            
            if len(df) > 0 and len(df.columns) > 1:
                return df
            elif delim == ' ':
                continue
            else:
                return df
                
        except Exception:
            continue
    
    try:
        data = json.loads(raw_text)
        if isinstance(data, list):
            return pd.DataFrame(data)
        elif isinstance(data, dict):
            return pd.DataFrame([data])
    except:
        pass
    
    try:
        return pd.read_csv(io.StringIO(raw_text), sep='\s+')
    except:
        pass
    
    raise ValueError("Could not parse the provided data with any supported format")


def load_data(file_path):
    """
    Load data from various file formats.
    
    Args:
        file_path (str): Path to the data file
        
    Returns:
        pd.DataFrame: Loaded data
    """
    file_path = Path(file_path)
    
    if not file_path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")
    
    extension = file_path.suffix.lower()
    
    if extension == ".csv":
        return pd.read_csv(file_path)
    elif extension in [".xls", ".xlsx"]:
        return pd.read_excel(file_path)
    elif extension == ".json":
        return pd.read_json(file_path)
    elif extension in [".txt", ".tsv", ".dat"]:
        with open(file_path, 'r', encoding='utf-8') as f:
            raw_text = f.read()
        return parse_raw_data(raw_text)
    else:
        with open(file_path, 'r', encoding='utf-8') as f:
            raw_text = f.read()
        return parse_raw_data(raw_text)


def ensure_directory(directory):
    """Create directory if it doesn't exist"""
    Path(directory).mkdir(parents=True, exist_ok=True)


def get_data_summary(df):
    """
    Generate a summary of the dataset.
    
    Args:
        df (pd.DataFrame): Input dataframe
        
    Returns:
        dict: Summary statistics
    """
    return {
        "rows": len(df),
        "columns": len(df.columns),
        "column_names": df.columns.tolist(),
        "data_types": df.dtypes.to_dict(),
        "missing_values": df.isnull().sum().to_dict(),
        "numeric_summary": df.describe().to_dict(),
    }


def is_numeric(series):
    """Check if a pandas series contains numeric data"""
    return pd.api.types.is_numeric_dtype(series)


def get_numeric_columns(df):
    """Get list of numeric columns in dataframe"""
    return df.select_dtypes(include=[float, int]).columns.tolist()


def get_categorical_columns(df):
    """Get list of categorical columns in dataframe"""
    return df.select_dtypes(include=["object"]).columns.tolist()


def detect_delimiter(raw_text, sample_size=1000):
    """
    Detect the most likely delimiter in raw text data.
    
    Args:
        raw_text (str): Sample of raw text data
        sample_size (int): Number of characters to analyze
        
    Returns:
        str: Detected delimiter
    """
    sample = raw_text[:sample_size]
    lines = sample.split('\n')[:5]
    
    delimiter_counts = {}
    delimiters = [',', '\t', '|', ';', ' ']
    
    for delim in delimiters:
        total_count = sum(line.count(delim) for line in lines)
        if total_count > 0:
            delimiter_counts[delim] = total_count
    
    if delimiter_counts:
        return max(delimiter_counts, key=delimiter_counts.get)
    
    return ','
