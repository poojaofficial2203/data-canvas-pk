"""Unit tests for data-canvas-pk"""

import unittest
import pandas as pd
from pathlib import Path
from data_canvas import DataToPPT
from data_canvas.utils import load_data
from data_canvas.analyzers import DataAnalyzer


class TestDataToPPT(unittest.TestCase):
    """Test DataToPPT class"""
    
    @classmethod
    def setUpClass(cls):
        """Set up test fixtures"""
        cls.test_csv = "examples/sample_data.csv"
    
    def test_load_data(self):
        """Test data loading"""
        df = load_data(self.test_csv)
        self.assertIsInstance(df, pd.DataFrame)
        self.assertGreater(len(df), 0)
    
    def test_analyzer_overview(self):
        """Test data analyzer overview"""
        df = load_data(self.test_csv)
        analyzer = DataAnalyzer(df)
        overview = analyzer.get_overview()
        
        self.assertEqual(overview["total_rows"], len(df))
        self.assertEqual(overview["total_columns"], len(df.columns))
        self.assertGreater(overview["numeric_columns"], 0)
    
    def test_converter_initialization(self):
        """Test converter initialization"""
        converter = DataToPPT(title="Test Report")
        self.assertEqual(converter.title, "Test Report")
    
    def test_data_summary(self):
        """Test data summary generation"""
        df = load_data(self.test_csv)
        analyzer = DataAnalyzer(df)
        
        # Test for a numeric column
        stats = analyzer.get_descriptive_stats("age")
        self.assertIn("mean", stats)
        self.assertIn("median", stats)
        self.assertIn("std_dev", stats)


if __name__ == "__main__":
    unittest.main()
