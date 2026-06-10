"""Data Canvas - Universal Data to PowerPoint Converter Package"""

from .universal_converter import UniversalDataToPPT
from .performance import PerformanceOptimizer, ChunkProcessor, CacheManager
from .summarizer import DataSummarizer, StructuredReportGenerator, PerformanceSummary, DataSummary
from .analyzers import DataAnalyzer
from .generators import ChartGenerator
from .utils import (
    parse_raw_data,
    load_data,
    get_numeric_columns,
    get_categorical_columns,
    detect_delimiter
)

__version__ = "2.0.0"
__author__ = "Data Canvas Team"
__all__ = [
    "UniversalDataToPPT",
    "PerformanceOptimizer",
    "ChunkProcessor",
    "CacheManager",
    "DataSummarizer",
    "StructuredReportGenerator",
    "PerformanceSummary",
    "DataSummary",
    "DataAnalyzer",
    "ChartGenerator",
    "parse_raw_data",
    "load_data",
    "get_numeric_columns",
    "get_categorical_columns",
    "detect_delimiter",
]
