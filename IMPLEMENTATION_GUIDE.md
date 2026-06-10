# Data Canvas - Complete Implementation Guide

## Overview

**Data Canvas** is a Python tool that automatically converts raw data (CSV, JSON, Excel) into professional PowerPoint presentations with charts, statistics, and insights.

---

## 📁 Project Structure

```
data-canvas-pk/
├── data_canvas/                 # Main package
│   ├── __init__.py             # Package initialization
│   ├── converter.py            # Main conversion engine
│   ├── analyzers.py            # Data analysis module
│   ├── generators.py           # Chart generation
│   └── utils.py                # Utility functions
├── tests/
│   └── test_converter.py       # Unit tests
├── examples/
│   └── sample_data.csv         # Sample employee data
├── main.py                      # Entry point script
├── requirements.txt             # Dependencies
├── README.md                    # Project documentation
└── .gitignore                   # Git ignore rules
```

---

## 🚀 Quick Start

### Installation

```bash
# Clone the repository
git clone https://github.com/poojaofficial2203/data-canvas-pk.git
cd data-canvas-pk

# Install dependencies
pip install -r requirements.txt
```

### Basic Usage

```python
from data_canvas import DataToPPT

# Create converter
converter = DataToPPT(title="My Report")

# Convert data to PowerPoint
converter.convert("data.csv", "output.pptx")
```

### Run Example

```bash
python main.py
```

This generates `output.pptx` with analysis of sample employee data.

---

## 📊 Generated Presentation Slides

The tool automatically creates:

1. **Title Slide** - Report title with generation date
2. **Data Overview** - Dataset dimensions (rows, columns, data types)
3. **Data Summary Table** - Column-by-column breakdown
4. **Statistics Slides** - Descriptive stats for numeric columns:
   - Mean, Median, Std Dev
   - Min, Max, Quartiles
5. **Visualization Slides** - Automatic charts:
   - Histogram for numeric distributions
   - Bar chart for categorical data
   - Scatter plots for relationships

---

## 💻 Code Examples

### Example 1: Convert CSV File

```python
from data_canvas import DataToPPT

converter = DataToPPT(title="Sales Analysis 2024")
converter.convert("sales_data.csv", "sales_report.pptx")
```

### Example 2: Convert Excel File

```python
converter = DataToPPT(title="Q4 Performance Report")
converter.convert("performance.xlsx", "q4_report.pptx")
```

### Example 3: Manual Data Analysis

```python
import pandas as pd
from data_canvas.analyzers import DataAnalyzer

df = pd.read_csv("data.csv")
analyzer = DataAnalyzer(df)

# Get overview
overview = analyzer.get_overview()
print(f"Rows: {overview['total_rows']}, Columns: {overview['total_columns']}")

# Get statistics for a column
stats = analyzer.get_descriptive_stats("age")
print(f"Mean: {stats['mean']}, Median: {stats['median']}")

# Get correlation matrix
corr = analyzer.get_correlation_matrix()
print(corr)
```

### Example 4: Custom Chart Generation

```python
from data_canvas.generators import ChartGenerator

# Create chart generator
cg = ChartGenerator()

# Generate histogram
histogram_img = cg.histogram(df['salary'], title="Salary Distribution")

# Generate bar chart
bar_img = cg.bar_chart(df['department'].value_counts(), title="Employees by Department")

# Generate scatter plot
scatter_img = cg.scatter_plot(df, 'age', 'salary', title="Age vs Salary")
```

---

## ��� Module Details

### `converter.py` - DataToPPT Class

**Main Methods:**
- `__init__(title)` - Initialize converter with presentation title
- `convert(input_file, output_file)` - Convert data file to PowerPoint

**Private Methods:**
- `_create_presentation(df)` - Create presentation structure
- `_add_title_slide()` - Add title slide
- `_add_overview_slide(df)` - Add data overview
- `_add_data_summary_slide(df)` - Add detailed summary table
- `_add_statistics_slides(df)` - Add statistics for numeric columns
- `_add_visualization_slides(df)` - Add charts and graphs
- `_add_image_to_slide(slide, image)` - Add image to slide

### `analyzers.py` - DataAnalyzer Class

**Key Methods:**
- `get_overview()` - Dataset dimensions
- `get_descriptive_stats(column)` - Stats for any column
- `get_correlation_matrix()` - Correlation for numeric columns
- `get_top_values(column, n)` - Top N values in a column
- `get_missing_data_summary()` - Missing data analysis

### `generators.py` - ChartGenerator Class

**Chart Methods:**
- `bar_chart(data, title, figsize)` - Generate bar chart
- `pie_chart(data, title, figsize)` - Generate pie chart
- `line_chart(df, x, y, title, figsize)` - Generate line chart
- `histogram(data, title, bins, figsize)` - Generate histogram
- `scatter_plot(df, x, y, title, figsize)` - Generate scatter plot

### `utils.py` - Utility Functions

- `load_data(file_path)` - Load CSV, Excel, or JSON
- `ensure_directory(directory)` - Create directory
- `get_data_summary(df)` - Get data summary
- `is_numeric(series)` - Check if column is numeric
- `get_numeric_columns(df)` - Get numeric column names
- `get_categorical_columns(df)` - Get categorical column names

---

## 🧪 Testing

Run unit tests:

```bash
python -m unittest tests.test_converter
```

Test Coverage:
- Data loading (CSV, Excel, JSON)
- Data analysis and statistics
- Converter initialization
- Chart generation (basic tests)

---

## 📋 Supported Data Formats

| Format | Extension | Support |
|--------|-----------|---------|
| CSV | `.csv` | ✅ Full |
| Excel | `.xlsx`, `.xls` | ✅ Full |
| JSON | `.json` | ✅ Full |

---

## 📊 Sample Data

The `examples/sample_data.csv` includes employee data with:
- `name` - Employee name
- `age` - Employee age
- `salary` - Annual salary
- `department` - Department (Sales, Engineering, Marketing, Management)
- `experience` - Years of experience
- `performance_score` - Performance rating (1-10)

---

## 🔧 Dependencies

All dependencies are in `requirements.txt`:

```
python-pptx==0.6.23    # PowerPoint generation
pandas==2.0.3          # Data manipulation
openpyxl==3.1.2        # Excel file handling
matplotlib==3.7.2      # Charting
seaborn==0.12.2        # Statistical visualization
numpy==1.24.3          # Numerical computing
Pillow==10.0.0         # Image processing
```

---

## 🎨 Customization

### Change Presentation Title

```python
converter = DataToPPT(title="Custom Report Title")
```

### Change Output Filename

```python
converter.convert("input.csv", "my_custom_report.pptx")
```

### Modify Chart Styles

Edit the `ChartGenerator.__init__()` style parameter to change matplotlib styles.

---

## 🐛 Troubleshooting

### Issue: Module not found error
**Solution:** Ensure you're in the project directory and have run `pip install -r requirements.txt`

### Issue: File not found error
**Solution:** Verify the input file path is correct and relative to your working directory

### Issue: No charts appearing
**Solution:** Check that your data has numeric columns for histograms and numeric columns for scatter plots

---

## 🤝 Contributing

To add features or fix issues:

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests
5. Submit a pull request

---

## 📄 License

This project is open source and available on GitHub.

---

## 📞 Support

For issues or questions, open an issue on GitHub or contact the maintainer.

---

## Changelog

### v1.0.0 (Current)
- Initial release
- Support for CSV, Excel, JSON
- Automatic chart generation
- Statistical analysis
- Professional PowerPoint generation
