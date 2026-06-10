"""Performance optimization module for handling large datasets efficiently"""

import pandas as pd
import numpy as np
from datetime import datetime
import gc
from typing import Dict, List, Optional, Tuple


class PerformanceOptimizer:
    """Optimize data processing and PowerPoint generation for large datasets"""
    
    def __init__(self, max_rows_per_chart: int = 10000, memory_limit_mb: int = 500):
        """
        Initialize performance optimizer.
        
        Args:
            max_rows_per_chart (int): Maximum rows to process per chart
            memory_limit_mb (int): Memory limit for operations in MB
        """
        self.max_rows_per_chart = max_rows_per_chart
        self.memory_limit_mb = memory_limit_mb
        self.metrics = {}
        self.start_time = None
        
    def start_tracking(self):
        """Start performance tracking"""
        self.start_time = datetime.now()
        gc.collect()
        
    def end_tracking(self, operation: str = "operation"):
        """End performance tracking and log metrics"""
        if self.start_time:
            duration = (datetime.now() - self.start_time).total_seconds()
            self.metrics[operation] = {
                'duration_seconds': round(duration, 2),
                'timestamp': datetime.now().isoformat()
            }
            gc.collect()
            
    def optimize_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Optimize dataframe memory usage through type conversion.
        
        Args:
            df (pd.DataFrame): Input dataframe
            
        Returns:
            pd.DataFrame: Memory-optimized dataframe
        """
        self.start_tracking()
        
        # Optimize numeric types
        for col in df.select_dtypes(include=['int64']).columns:
            col_min = df[col].min()
            col_max = df[col].max()
            
            if col_min > np.iinfo(np.int8).min and col_max < np.iinfo(np.int8).max:
                df[col] = df[col].astype(np.int8)
            elif col_min > np.iinfo(np.int16).min and col_max < np.iinfo(np.int16).max:
                df[col] = df[col].astype(np.int16)
            elif col_min > np.iinfo(np.int32).min and col_max < np.iinfo(np.int32).max:
                df[col] = df[col].astype(np.int32)
        
        # Optimize float types
        for col in df.select_dtypes(include=['float64']).columns:
            df[col] = df[col].astype(np.float32)
        
        # Optimize categorical columns
        for col in df.select_dtypes(include=['object']).columns:
            if df[col].nunique() / len(df) < 0.5:  # If < 50% unique values
                df[col] = df[col].astype('category')
        
        self.end_tracking('dataframe_optimization')
        return df
    
    def downsample_for_visualization(self, df: pd.DataFrame, max_rows: Optional[int] = None) -> pd.DataFrame:
        """
        Downsample data for visualization if it exceeds max_rows.
        
        Args:
            df (pd.DataFrame): Input dataframe
            max_rows (int): Maximum rows to keep (default: self.max_rows_per_chart)
            
        Returns:
            pd.DataFrame: Downsampled dataframe
        """
        self.start_tracking()
        
        max_rows = max_rows or self.max_rows_per_chart
        
        if len(df) <= max_rows:
            self.end_tracking('downsampling_skipped')
            return df
        
        # Use stratified sampling for categorical columns
        numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
        
        if numeric_cols and len(df) > max_rows:
            # Sample while preserving distribution
            sample = df.sample(n=max_rows, replace=False, random_state=42)
        else:
            sample = df.head(max_rows)
        
        self.end_tracking('downsampling')
        return sample
    
    def get_performance_summary(self) -> Dict:
        """
        Get performance metrics summary.
        
        Returns:
            Dict: Performance metrics
        """
        return {
            'operations': self.metrics,
            'total_operations': len(self.metrics),
            'total_time': sum(m['duration_seconds'] for m in self.metrics.values())
        }


class ChunkProcessor:
    """Process large dataframes in chunks to manage memory"""
    
    def __init__(self, chunk_size: int = 5000):
        """
        Initialize chunk processor.
        
        Args:
            chunk_size (int): Number of rows per chunk
        """
        self.chunk_size = chunk_size
        
    def process_in_chunks(self, df: pd.DataFrame, func, *args, **kwargs):
        """
        Process dataframe in chunks.
        
        Args:
            df (pd.DataFrame): Input dataframe
            func: Function to apply to each chunk
            *args: Additional positional arguments
            **kwargs: Additional keyword arguments
            
        Returns:
            List: Results from processing each chunk
        """
        results = []
        
        for i in range(0, len(df), self.chunk_size):
            chunk = df.iloc[i:i + self.chunk_size]
            result = func(chunk, *args, **kwargs)
            results.append(result)
            gc.collect()
        
        return results
    
    def aggregate_chunks(self, results: List[Dict]) -> Dict:
        """
        Aggregate results from chunk processing.
        
        Args:
            results (List[Dict]): Results from each chunk
            
        Returns:
            Dict: Aggregated results
        """
        if not results:
            return {}
        
        # Simple aggregation - can be customized
        aggregated = {}
        
        for key in results[0].keys():
            values = [r[key] for r in results if key in r]
            
            if isinstance(values[0], (int, float)):
                aggregated[key] = {
                    'mean': np.mean(values),
                    'sum': np.sum(values),
                    'min': np.min(values),
                    'max': np.max(values)
                }
            else:
                aggregated[key] = values
        
        return aggregated


class CacheManager:
    """Manage caching of expensive operations"""
    
    def __init__(self):
        """Initialize cache manager"""
        self.cache = {}
        self.access_count = {}
        
    def get(self, key: str):
        """Get cached value"""
        if key in self.cache:
            self.access_count[key] = self.access_count.get(key, 0) + 1
            return self.cache[key]
        return None
    
    def set(self, key: str, value):
        """Set cached value"""
        self.cache[key] = value
        self.access_count[key] = 0
        
    def clear(self):
        """Clear cache"""
        self.cache.clear()
        self.access_count.clear()
    
    def get_stats(self) -> Dict:
        """Get cache statistics"""
        return {
            'cached_items': len(self.cache),
            'total_accesses': sum(self.access_count.values()),
            'cache_keys': list(self.cache.keys())
        }
