import pandas as pd

class DataAnalyzer:

    def __init__(self, df):
        self.df = df

    def get_overview(self):

        return {
            "total_rows": len(self.df),
            "total_columns": len(self.df.columns),
            "numeric_columns": len(
                self.df.select_dtypes(include="number").columns
            ),
            "categorical_columns": len(
                self.df.select_dtypes(exclude="number").columns
            )
        }

    def get_descriptive_stats(self, column):

        s = self.df[column]

        return {
            "mean": float(s.mean()),
            "median": float(s.median()),
            "std_dev": float(s.std()),
            "min": float(s.min()),
            "max": float(s.max())
        }
