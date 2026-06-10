# Data Canvas - Universal Data to PowerPoint Converter

Convert **ANY** raw data format to professional PowerPoint presentations instantly!

## 🚀 New Feature: Universal Data Converter

Now supports:
- ✅ **Raw Text Data** (CSV, TSV, pipe-delimited, space-separated, etc.)
- ✅ **JSON Strings**
- ✅ **Python Dictionaries**
- ✅ **CSV Files**
- ✅ **Excel Files** (.xls, .xlsx)
- ✅ **Text Files** (.txt, .dat, .tsv)
- ✅ **Auto-Format Detection**

## 📋 Installation

```bash
git clone https://github.com/poojaofficial2203/data-canvas-pk.git
cd data-canvas-pk
pip install -r requirements.txt
```

## 💡 Quick Examples

### Example 1: Raw CSV Text
```python
from data_canvas import UniversalDataToPPT

raw_csv = """name,age,salary
Alice,28,75000
Bob,35,85000
Carol,32,90000"""

converter = UniversalDataToPPT(title="My Report")
converter.convert_text(raw_csv, "report.pptx")
```

### Example 2: Tab-Separated Data
```python
raw_tsv = """Product\tQ1\tQ2\tQ3
Widget A\t100\t120\t150
Widget B\t200\t210\t205"""

converter.convert_text(raw_tsv, "report.pptx", delimiter='\t')
```

### Example 3: JSON String
```python
json_data = '[{"name": "John", "age": 30}, {"name": "Jane", "age": 28}]'

converter.convert_json_string(json_data, "report.pptx")
```

### Example 4: Python Dictionary
```python
data = {
    'Month': ['Jan', 'Feb', 'Mar'],
    'Sales': [1000, 1200, 1500]
}

converter.convert_dict(data, "report.pptx")
```

### Example 5: File Conversion
```python
# CSV File
converter.convert_file("data.csv", "report.pptx")

# Excel File
converter.convert_file("data.xlsx", "report.pptx")

# Text File (auto-detects format)
converter.convert_file("data.txt", "report.pptx")
```

## 🎯 Supported Delimiters (Auto-Detected)

| Format | Delimiter | Example |
|--------|-----------|---------|
| CSV | `,` | `name,age,salary` |
| TSV | `\t` | `name	age	salary` |
| Pipe | `\|` | `name\|age\|salary` |
| Semicolon | `;` | `name;age;salary` |
| Space | ` ` | `name age salary` |
| JSON | N/A | `[{...}, {...}]` |

## 📊 Generated Presentation Includes

1. ✅ **Title Slide** - Report title with data format info
2. ✅ **Data Overview** - Rows, columns, data types
3. ✅ **Data Preview** - First 5 rows
4. ✅ **Column Summary** - Detailed column info
5. ✅ **Statistics** - Mean, median, std dev, min, max
6. ✅ **Visualizations** - Histograms, bar charts, scatter plots

## 🏃 Run All Examples

```bash
python universal_examples.py
```

This creates 9 different PPT files showcasing all conversion types!

## 📝 Use Cases

- 📊 **Business Reports** - Convert sales data to presentations
- 📈 **Data Analysis** - Turn raw data into insights
- 📋 **Reports** - Automated slide generation
- 🎓 **Education** - Data visualization for students
- 📱 **Data Import** - From any source format

## 🔄 Conversion Flow

```
Raw Data (Any Format)
        ↓
    Auto-Detect Format
        ↓
    Parse to DataFrame
        ↓
    Analyze Data
        ↓
    Generate Charts
        ↓
    Create PowerPoint
        ↓
Professional Presentation
```

## 🛠️ Advanced Options

### Custom Delimiter
```python
converter.convert_text(data, "output.pptx", delimiter='|')
```

### Custom Title
```python
converter = UniversalDataToPPT(title="My Custom Title")
```

### DataFrame Input
```python
import pandas as pd

df = pd.read_csv("data.csv")
converter.convert_dataframe(df, "report.pptx")
```

## 📋 File Format Support

| Format | Read | Write | Notes |
|--------|------|-------|-------|
| CSV | ✅ | ✅ | Auto-detected |
| Excel | ✅ | ✅ | .xls, .xlsx |
| JSON | ✅ | ✅ | Lists of dicts |
| Text | ✅ | ✅ | TSV, pipe, space delimited |
| TXT | ✅ | ✅ | Auto-format detection |

## 🎓 Learning Resources

- `IMPLEMENTATION_GUIDE.md` - Complete API documentation
- `universal_examples.py` - 10 working examples
- `tests/test_converter.py` - Unit tests

## 🐛 Error Handling

```python
converter = UniversalDataToPPT()

# Returns True/False for success
success = converter.convert_text(data, "output.pptx")

if not success:
    print("Conversion failed - check data format")
```

## 📞 Support

For issues or questions:
1. Check the examples in `universal_examples.py`
2. Review `IMPLEMENTATION_GUIDE.md`
3. Open an issue on GitHub

## 📄 License

Open source - Feel free to use and modify!

---

**Made with ❤️ for data enthusiasts**
