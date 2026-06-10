# Data Canvas - Complete Testing Guide

## 🧪 Quick Test Steps

### Step 1: Clone the Repository
```bash
git clone https://github.com/poojaofficial2203/data-canvas-pk.git
cd data-canvas-pk
```

### Step 2: Install Dependencies
```bash
pip install -r requirements.txt
```

Required packages:
- python-pptx==0.6.23
- pandas==2.0.3
- openpyxl==3.1.2
- matplotlib==3.7.2
- seaborn==0.12.2
- numpy==1.24.3
- Pillow==10.0.0

### Step 3: Run All Examples
```bash
python universal_examples.py
```

This will generate 10 PowerPoint files:
1. ✅ example1_raw_csv.pptx
2. ✅ example2_tsv.pptx
3. ✅ example3_pipe.pptx
4. ✅ example4_json.pptx
5. ✅ example5_dict.pptx
6. ✅ example6_csv_file.pptx
7. ✅ example8_auto_detect.pptx
8. ✅ example9_space.pptx
9. ✅ example10_dataframe.pptx
10. ✅ example11_semicolon.pptx

---

## 📝 Individual Test Cases

### Test 1: Raw CSV Text Conversion
```python
from data_canvas import UniversalDataToPPT

raw_csv = """name,age,salary
Alice,28,75000
Bob,35,85000
Carol,32,90000"""

converter = UniversalDataToPPT(title="Employee Report")
converter.convert_text(raw_csv, "test1_csv.pptx")
print("✓ Test 1 passed: CSV text converted")
```

### Test 2: Tab-Separated Data
```python
raw_tsv = """Product\tPrice\tStock
Laptop\t1200\t45
Mouse\t25\t150
Keyboard\t75\t100"""

converter = UniversalDataToPPT(title="Inventory")
converter.convert_text(raw_tsv, "test2_tsv.pptx", delimiter='\t')
print("✓ Test 2 passed: TSV converted")
```

### Test 3: Pipe-Delimited Data
```python
raw_pipe = """ID|Name|Status
1001|John|Active
1002|Jane|Inactive
1003|Bob|Active"""

converter = UniversalDataToPPT(title="User Status")
converter.convert_text(raw_pipe, "test3_pipe.pptx", delimiter='|')
print("✓ Test 3 passed: Pipe-delimited converted")
```

### Test 4: JSON String
```python
import json

json_data = [
    {"name": "John", "age": 30, "city": "NYC"},
    {"name": "Jane", "age": 28, "city": "LA"},
    {"name": "Bob", "age": 35, "city": "Chicago"}
]

converter = UniversalDataToPPT(title="JSON Report")
converter.convert_json_string(json.dumps(json_data), "test4_json.pptx")
print("✓ Test 4 passed: JSON string converted")
```

### Test 5: Python Dictionary
```python
data = {
    'Month': ['Jan', 'Feb', 'Mar', 'Apr'],
    'Sales': [1000, 1200, 1500, 1800],
    'Profit': [200, 300, 400, 500]
}

converter = UniversalDataToPPT(title="Monthly Report")
converter.convert_dict(data, "test5_dict.pptx")
print("✓ Test 5 passed: Dictionary converted")
```

### Test 6: CSV File
```python
converter = UniversalDataToPPT(title="Employee Analysis")
success = converter.convert_file("examples/sample_data.csv", "test6_csv_file.pptx")
if success:
    print("✓ Test 6 passed: CSV file converted")
```

### Test 7: Pandas DataFrame
```python
import pandas as pd

df = pd.DataFrame({
    'Product': ['A', 'B', 'C', 'D'],
    'Q1': [100, 200, 150, 180],
    'Q2': [120, 210, 160, 200],
    'Q3': [150, 250, 180, 220]
})

converter = UniversalDataToPPT(title="Sales Analysis")
converter.convert_dataframe(df, "test7_dataframe.pptx")
print("✓ Test 7 passed: DataFrame converted")
```

### Test 8: Semicolon-Separated Data
```python
raw_semi = """Name;Role;Salary
Alice;Engineer;95000
Bob;Manager;105000
Carol;Analyst;75000"""

converter = UniversalDataToPPT(title="Payroll")
converter.convert_text(raw_semi, "test8_semicolon.pptx", delimiter=';')
print("✓ Test 8 passed: Semicolon-delimited converted")
```

### Test 9: Space-Separated Data
```python
raw_space = """Year Revenue Profit
2022 1000000 200000
2023 1200000 300000
2024 1500000 400000"""

converter = UniversalDataToPPT(title="Financial Report")
converter.convert_text(raw_space, "test9_space.pptx")
print("✓ Test 9 passed: Space-separated converted")
```

### Test 10: Auto-Delimiter Detection
```python
raw_data = """Name,Department,Salary
John,Engineering,95000
Jane,Marketing,75000
Bob,Sales,70000"""

# No delimiter specified - should auto-detect comma
converter = UniversalDataToPPT(title="Auto-Detect Test")
converter.convert_text(raw_data, "test10_auto.pptx")
print("✓ Test 10 passed: Auto-detection worked")
```

---

## 🔍 Verification Steps

After running tests, verify:

1. **Check Output Files:**
   ```bash
   ls -lah example*.pptx  # Should show all 10 files
   ```

2. **Verify File Sizes:**
   - Each PPT should be 100KB-500KB
   - Larger files indicate more content/slides

3. **Open PowerPoint Files:**
   - Open any `.pptx` file in Microsoft PowerPoint or LibreOffice
   - Check that slides contain:
     - Title slide with data format info
     - Data overview with row/column counts
     - Data preview showing first 5 rows
     - Column summary table
     - Statistics for numeric columns
     - Visualization charts

4. **Run Unit Tests:**
   ```bash
   python -m unittest tests.test_converter -v
   ```

---

## ✅ Expected Results

### Each Generated PPT Should Have:

1. **Slide 1: Title Slide**
   - Report title
   - Data format indicator
   - Generation date/time
   - Professional blue background

2. **Slide 2: Data Overview**
   - Total rows (formatted with commas)
   - Total columns
   - Numeric columns count
   - Categorical columns count
   - Data source format

3. **Slide 3: Data Preview**
   - Table with first 5 rows
   - All columns displayed
   - Formatted headers

4. **Slide 4: Column Summary**
   - Column names
   - Data types
   - Non-null count
   - Unique values

5. **Slide 5+: Statistics & Charts**
   - Mean, median, std dev for numeric columns
   - Histograms for distributions
   - Bar charts for categories
   - Scatter plots for relationships

---

## 🐛 Troubleshooting

### Issue: "No module named 'data_canvas'"
**Solution:**
```bash
pip install -e .
# or
export PYTHONPATH=$PYTHONPATH:$(pwd)
```

### Issue: Missing dependencies
**Solution:**
```bash
pip install --upgrade -r requirements.txt
```

### Issue: File not found error
**Solution:**
- Ensure you're in the project directory
- Check that `examples/sample_data.csv` exists
- Use absolute paths if needed

### Issue: Permission denied
**Solution:**
```bash
chmod +x universal_examples.py
python universal_examples.py
```

### Issue: Matplotlib style error
**Solution:**
- The code handles this automatically
- If still failing, reinstall matplotlib:
```bash
pip install --upgrade matplotlib
```

---

## 📊 Sample Test Output

When running `python universal_examples.py`, you should see:

```
======================================================================
     Data Canvas - Universal Data to PowerPoint Converter
======================================================================

This tool converts ANY raw data format to professional presentations!

Supported Formats:
  ✓ Raw CSV, TSV, Pipe-delimited, Semicolon-separated, Space-separated
  ✓ JSON Strings & Files
  ✓ Python Dictionaries & DataFrames
  ✓ Excel Files (.xlsx, .xls)
  ✓ Text Files (.txt, .dat, .tsv)
======================================================================

=== Example 1: Raw CSV Text ===
✓ Presentation created from raw text: example1_raw_csv.pptx
✓ Created: example1_raw_csv.pptx

=== Example 2: Tab-Separated Data ===
✓ Presentation created from raw text: example2_tsv.pptx
✓ Created: example2_tsv.pptx

... (more examples)

======================================================================
✓ ALL EXAMPLES COMPLETED SUCCESSFULLY!
======================================================================

Generated PPT Files:
  1. example1_raw_csv.pptx
  2. example2_tsv.pptx
  3. example3_pipe.pptx
  4. example4_json.pptx
  5. example5_dict.pptx
  6. example6_csv_file.pptx
  7. example8_auto_detect.pptx
  8. example9_space.pptx
  9. example10_dataframe.pptx
  10. example11_semicolon.pptx
======================================================================
```

---

## 🎯 Quick Validation Checklist

- [ ] Repository cloned successfully
- [ ] Dependencies installed without errors
- [ ] `python universal_examples.py` runs without errors
- [ ] 10 PPT files generated
- [ ] Each file is readable (can open in PowerPoint)
- [ ] Each file contains expected slides
- [ ] Charts display correctly
- [ ] Tables are formatted properly
- [ ] Statistics values are calculated
- [ ] Data preview shows correct data

---

## 📞 Support

If you encounter any issues:

1. Check the error message carefully
2. Review the Troubleshooting section
3. Check that all dependencies are installed
4. Ensure you have write permissions in the directory
5. Try running from a fresh Python environment

---

**Happy Testing! 🎉**
