"""Structured output and data summarization module for comprehensive reporting"""

import pandas as pd
import numpy as np
from typing import Dict, List, Any, Optional
from dataclasses import dataclass, asdict
from datetime import datetime
import json


@dataclass
class DataSummary:
    """Data structure for comprehensive data summary"""
    total_rows: int
    total_columns: int
    column_names: List[str]
    data_types: Dict[str, str]
    missing_values: Dict[str, int]
    numeric_columns: List[str]
    categorical_columns: List[str]
    date_columns: List[str]
    generated_at: str
    
    def to_dict(self) -> Dict:
        """Convert to dictionary"""
        return asdict(self)
    
    def to_json(self) -> str:
        """Convert to JSON string"""
        return json.dumps(self.to_dict(), indent=2, default=str)


class DataSummarizer:
    """Generate comprehensive structured summaries of datasets"""
    
    def __init__(self):
        """Initialize data summarizer"""
        self.summary = None
        
    def generate_summary(self, df: pd.DataFrame) -> DataSummary:
        """
        Generate comprehensive summary of the dataset.
        
        Args:
            df (pd.DataFrame): Input dataframe
            
        Returns:
            DataSummary: Structured summary object
        """
        numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
        categorical_cols = df.select_dtypes(include=['object']).columns.tolist()
        date_cols = df.select_dtypes(include=['datetime64']).columns.tolist()
        
        self.summary = DataSummary(
            total_rows=len(df),
            total_columns=len(df.columns),
            column_names=df.columns.tolist(),
            data_types={col: str(df[col].dtype) for col in df.columns},
            missing_values=df.isnull().sum().to_dict(),
            numeric_columns=numeric_cols,
            categorical_columns=categorical_cols,
            date_columns=date_cols,
            generated_at=datetime.now().isoformat()
        )
        
        return self.summary
    
    def get_column_insights(self, df: pd.DataFrame, column: str) -> Dict[str, Any]:
        """
        Get detailed insights for a specific column.
        
        Args:
            df (pd.DataFrame): Input dataframe
            column (str): Column name
            
        Returns:
            Dict: Detailed insights
        """
        if column not in df.columns:
            raise ValueError(f"Column '{column}' not found in dataframe")
        
        series = df[column]
        insights = {
            'column_name': column,
            'data_type': str(series.dtype),
            'non_null_count': series.notna().sum(),
            'null_count': series.isnull().sum(),
            'null_percentage': round(series.isnull().sum() / len(df) * 100, 2),
            'unique_values': series.nunique(),
            'unique_percentage': round(series.nunique() / len(df) * 100, 2),
        }
        
        if pd.api.types.is_numeric_dtype(series):
            insights.update({
                'min': float(series.min()),
                'max': float(series.max()),
                'mean': float(series.mean()),
                'median': float(series.median()),
                'std_dev': float(series.std()),
                'q25': float(series.quantile(0.25)),
                'q75': float(series.quantile(0.75)),
                'skewness': float(series.skew()),
                'kurtosis': float(series.kurtosis()),
            })
        else:
            top_values = series.value_counts().head(5)
            insights.update({
                'top_values': top_values.to_dict(),
                'top_value_frequency': float(top_values.iloc[0]) if len(top_values) > 0 else 0,
            })
        
        return insights
    
    def get_all_column_insights(self, df: pd.DataFrame) -> Dict[str, Dict[str, Any]]:
        """
        Get insights for all columns.
        
        Args:
            df (pd.DataFrame): Input dataframe
            
        Returns:
            Dict: Insights for all columns
        """
        return {col: self.get_column_insights(df, col) for col in df.columns}


class StructuredReportGenerator:
    """Generate structured reports with insights and recommendations"""
    
    def __init__(self):
        """Initialize report generator"""
        self.report = {}
        
    def generate_report(self, df: pd.DataFrame, title: str = "Data Analysis Report") -> Dict:
        """
        Generate comprehensive structured report.
        
        Args:
            df (pd.DataFrame): Input dataframe
            title (str): Report title
            
        Returns:
            Dict: Structured report
        """
        summarizer = DataSummarizer()
        summary = summarizer.generate_summary(df)
        
        self.report = {
            'metadata': {
                'title': title,
                'generated_at': datetime.now().isoformat(),
                'data_summary': summary.to_dict(),
            },
            'statistics': {
                'overview': self._get_overview_stats(df),
                'column_details': summarizer.get_all_column_insights(df),
                'data_quality': self._analyze_data_quality(df),
            },
            'insights': self._generate_insights(df),
            'recommendations': self._generate_recommendations(df),
        }
        
        return self.report
    
    def _get_overview_stats(self, df: pd.DataFrame) -> Dict:
        """Get overview statistics"""
        return {
            'shape': {'rows': len(df), 'columns': len(df.columns)},
            'memory_usage_mb': round(df.memory_usage(deep=True).sum() / 1024 / 1024, 2),
            'completeness': round((df.notna().sum().sum() / (len(df) * len(df.columns)) * 100), 2),
            'duplicate_rows': len(df) - len(df.drop_duplicates()),
        }
    
    def _analyze_data_quality(self, df: pd.DataFrame) -> Dict:
        """Analyze data quality"""
        quality = {
            'missing_data': {
                'columns_with_missing': len(df.columns[df.isnull().any()]),
                'missing_percentage': round(df.isnull().sum().sum() / (len(df) * len(df.columns)) * 100, 2),
            },
            'duplicates': {
                'total_duplicate_rows': len(df) - len(df.drop_duplicates()),
                'duplicate_percentage': round((len(df) - len(df.drop_duplicates())) / len(df) * 100, 2),
            },
            'outliers': self._detect_outliers(df),
        }
        return quality
    
    def _detect_outliers(self, df: pd.DataFrame) -> Dict:
        """Detect outliers in numeric columns"""
        outliers = {}
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        
        for col in numeric_cols:
            Q1 = df[col].quantile(0.25)
            Q3 = df[col].quantile(0.75)
            IQR = Q3 - Q1
            lower_bound = Q1 - 1.5 * IQR
            upper_bound = Q3 + 1.5 * IQR
            
            outlier_count = ((df[col] < lower_bound) | (df[col] > upper_bound)).sum()
            
            if outlier_count > 0:
                outliers[col] = {
                    'outlier_count': int(outlier_count),
                    'percentage': round(outlier_count / len(df) * 100, 2),
                }
        
        return outliers
    
    def _generate_insights(self, df: pd.DataFrame) -> List[str]:
        """Generate data insights"""
        insights = []
        
        # Missing data insight
        missing_pct = df.isnull().sum().sum() / (len(df) * len(df.columns)) * 100
        if missing_pct > 0:
            insights.append(f"Dataset has {missing_pct:.1f}% missing values across all columns")
        
        # Duplicate insight
        dup_pct = (len(df) - len(df.drop_duplicates())) / len(df) * 100
        if dup_pct > 0:
            insights.append(f"Found {dup_pct:.1f}% duplicate rows in the dataset")
        
        # Numeric columns insight
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        if len(numeric_cols) > 0:
            insights.append(f"Dataset contains {len(numeric_cols)} numeric columns for statistical analysis")
        
        # Categorical insight
        cat_cols = df.select_dtypes(include=['object']).columns
        if len(cat_cols) > 0:
            insights.append(f"Dataset contains {len(cat_cols)} categorical columns")
        
        return insights
    
    def _generate_recommendations(self, df: pd.DataFrame) -> List[str]:
        """Generate actionable recommendations"""
        recommendations = []
        
        # Missing data recommendations
        missing_cols = df.columns[df.isnull().any()].tolist()
        if missing_cols:
            recommendations.append(f"Consider handling missing values in columns: {', '.join(missing_cols[:3])}")
        
        # Duplicate recommendations
        if len(df) > len(df.drop_duplicates()):
            recommendations.append("Remove or investigate duplicate rows for data consistency")
        
        # Outlier recommendations
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        for col in numeric_cols:
            Q1 = df[col].quantile(0.25)
            Q3 = df[col].quantile(0.75)
            IQR = Q3 - Q1
            outlier_count = ((df[col] < Q1 - 1.5 * IQR) | (df[col] > Q3 + 1.5 * IQR)).sum()
            
            if outlier_count > 0:
                recommendations.append(f"Column '{col}' contains {outlier_count} potential outliers - review for data quality")
                break
        
        return recommendations
    
    def export_report(self, filepath: str = None) -> str:
        """
        Export report to JSON file.
        
        Args:
            filepath (str): Output filepath (optional)
            
        Returns:
            str: JSON report string
        """
        if not self.report:
            raise ValueError("No report generated yet. Call generate_report() first.")
        
        json_report = json.dumps(self.report, indent=2, default=str)
        
        if filepath:
            with open(filepath, 'w') as f:
                f.write(json_report)
        
        return json_report


class PerformanceSummary:
    """Generate performance and efficiency summaries"""
    
    @staticmethod
    def summarize_conversion(
        input_size_mb: float,
        num_rows: int,
        num_columns: int,
        processing_time_sec: float,
        charts_generated: int
    ) -> Dict:
        """
        Summarize conversion performance.
        
        Args:
            input_size_mb (float): Input file size in MB
            num_rows (int): Number of data rows
            num_columns (int): Number of columns
            processing_time_sec (float): Processing time in seconds
            charts_generated (int): Number of charts generated
            
        Returns:
            Dict: Performance summary
        """
        rows_per_sec = num_rows / max(processing_time_sec, 0.01)
        
        return {
            'input_metrics': {
                'file_size_mb': round(input_size_mb, 2),
                'rows': num_rows,
                'columns': num_columns,
            },
            'processing_metrics': {
                'processing_time_seconds': round(processing_time_sec, 2),
                'rows_processed_per_second': round(rows_per_sec, 0),
                'charts_generated': charts_generated,
            },
            'efficiency': {
                'average_time_per_row_ms': round((processing_time_sec / max(num_rows, 1)) * 1000, 3),
                'time_per_chart_sec': round(processing_time_sec / max(charts_generated, 1), 2),
            }
        }
