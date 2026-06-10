"""Universal Data Canvas - Convert ANY raw data format to PowerPoint"""

from data_canvas import UniversalDataToPPT


def example_1_raw_csv_text():
    """Example 1: Convert raw CSV text directly"""
    print("\n=== Example 1: Raw CSV Text ===")
    
    raw_csv_text = """name,age,salary,department
Alice Johnson,28,75000,Sales
Bob Smith,35,85000,Engineering
Carol White,32,90000,Engineering
David Brown,41,95000,Management
Eve Davis,29,72000,Sales"""
    
    converter = UniversalDataToPPT(title="Raw CSV Data Report")
    success = converter.convert_text(raw_csv_text, "example1_raw_csv.pptx")
    
    if success:
        print("✓ Created: example1_raw_csv.pptx")


def example_2_raw_tab_separated():
    """Example 2: Convert raw Tab-Separated Values"""
    print("\n=== Example 2: Tab-Separated Data ===")
    
    raw_tsv_text = """Product\tQ1\tQ2\tQ3\tQ4
Widget A\t100\t120\t150\t180
Widget B\t200\t210\t205\t220
Widget C\t150\t160\t170\t185
Widget D\t80\t90\t95\t110"""
    
    converter = UniversalDataToPPT(title="Quarterly Sales Report")
    success = converter.convert_text(raw_tsv_text, "example2_tsv.pptx", delimiter='\t')
    
    if success:
        print("✓ Created: example2_tsv.pptx")


def example_3_raw_pipe_delimited():
    """Example 3: Convert raw pipe-delimited data"""
    print("\n=== Example 3: Pipe-Delimited Data ===")
    
    raw_pipe_text = """ID|Customer|Amount|Status
1001|Acme Corp|5000|Completed
1002|TechStart|7500|Pending
1003|DataFlow|3200|Completed
1004|CloudBase|8900|Processing
1005|NetSolutions|4100|Completed"""
    
    converter = UniversalDataToPPT(title="Sales Pipeline Report")
    success = converter.convert_text(raw_pipe_text, "example3_pipe.pptx", delimiter='|')
    
    if success:
        print("✓ Created: example3_pipe.pptx")


def example_4_raw_json_string():
    """Example 4: Convert raw JSON string"""
    print("\n=== Example 4: JSON String ===")
    
    json_string = """[
    {"name": "John", "age": 30, "score": 85, "city": "NYC"},
    {"name": "Jane", "age": 28, "score": 92, "city": "LA"},
    {"name": "Mike", "age": 35, "score": 78, "city": "Chicago"},
    {"name": "Sarah", "age": 26, "score": 95, "city": "Boston"}
]"""
    
    converter = UniversalDataToPPT(title="Student Performance Report")
    success = converter.convert_json_string(json_string, "example4_json.pptx")
    
    if success:
        print("✓ Created: example4_json.pptx")


def example_5_python_dict():
    """Example 5: Convert Python dictionary"""
    print("\n=== Example 5: Python Dictionary ===")
    
    data_dict = {
        'Month': ['January', 'February', 'March', 'April', 'May'],
        'Temperature': [32, 35, 45, 62, 75],
        'Humidity': [65, 60, 55, 50, 45],
        'Rainfall': [2.1, 1.8, 3.2, 2.5, 1.2]
    }
    
    converter = UniversalDataToPPT(title="Weather Analysis Report")
    success = converter.convert_dict(data_dict, "example5_dict.pptx")
    
    if success:
        print("✓ Created: example5_dict.pptx")


def example_6_csv_file():
    """Example 6: Convert CSV file"""
    print("\n=== Example 6: CSV File ===")
    
    converter = UniversalDataToPPT(title="Employee Data Analysis")
    success = converter.convert_file("examples/sample_data.csv", "example6_csv_file.pptx")
    
    if success:
        print("✓ Created: example6_csv_file.pptx")


def example_8_auto_delimiter_detection():
    """Example 8: Auto-detect delimiter"""
    print("\n=== Example 8: Auto Delimiter Detection ===")
    
    raw_mixed_text = """Product,Price,Stock,Category
Laptop,1200,45,Electronics
Mouse,25,150,Electronics
Keyboard,75,100,Electronics
Monitor,350,30,Electronics"""
    
    converter = UniversalDataToPPT(title="Inventory Report")
    success = converter.convert_text(raw_mixed_text, "example8_auto_detect.pptx")
    
    if success:
        print("✓ Created: example8_auto_detect.pptx")


def example_9_space_separated():
    """Example 9: Space-separated data"""
    print("\n=== Example 9: Space-Separated Data ===")
    
    raw_space_text = """Year Revenue Profit Growth
2020 1000000 200000 5
2021 1200000 300000 20
2022 1500000 400000 25
2023 1800000 500000 20
2024 2100000 600000 16.7"""
    
    converter = UniversalDataToPPT(title="Financial Report")
    success = converter.convert_text(raw_space_text, "example9_space.pptx")
    
    if success:
        print("✓ Created: example9_space.pptx")


def example_10_pandas_dataframe():
    """Example 10: Convert pandas DataFrame"""
    print("\n=== Example 10: Pandas DataFrame ===")
    
    import pandas as pd
    
    df = pd.DataFrame({
        'Product': ['Laptop', 'Desktop', 'Tablet', 'Phone'],
        'Q1_Sales': [1500, 800, 2000, 3500],
        'Q2_Sales': [1800, 950, 2200, 3800],
        'Q3_Sales': [2100, 1100, 2500, 4200],
        'Q4_Sales': [2800, 1400, 3000, 5000]
    })
    
    converter = UniversalDataToPPT(title="Quarterly Product Sales")
    success = converter.convert_dataframe(df, "example10_dataframe.pptx")
    
    if success:
        print("✓ Created: example10_dataframe.pptx")


def example_11_semicolon_separated():
    """Example 11: Semicolon-separated data"""
    print("\n=== Example 11: Semicolon-Separated Data ===")
    
    raw_semi_text = """Name;Department;Salary;Experience
John;Engineering;95000;8
Mary;Marketing;75000;5
Peter;Sales;70000;4
Lisa;Engineering;98000;10"""
    
    converter = UniversalDataToPPT(title="Employee Records")
    success = converter.convert_text(raw_semi_text, "example11_semicolon.pptx", delimiter=';')
    
    if success:
        print("✓ Created: example11_semicolon.pptx")


def main():
    """Run all examples"""
    print("=" * 70)
    print("     Data Canvas - Universal Data to PowerPoint Converter")
    print("=" * 70)
    print("\nThis tool converts ANY raw data format to professional presentations!")
    print("\nSupported Formats:")
    print("  ✓ Raw CSV, TSV, Pipe-delimited, Semicolon-separated, Space-separated")
    print("  ✓ JSON Strings & Files")
    print("  ✓ Python Dictionaries & DataFrames")
    print("  ✓ Excel Files (.xlsx, .xls)")
    print("  ✓ Text Files (.txt, .dat, .tsv)")
    print("=" * 70)
    
    try:
        example_1_raw_csv_text()
        example_2_raw_tab_separated()
        example_3_raw_pipe_delimited()
        example_4_raw_json_string()
        example_5_python_dict()
        example_6_csv_file()
        example_8_auto_delimiter_detection()
        example_9_space_separated()
        example_10_pandas_dataframe()
        example_11_semicolon_separated()
        
        print("\n" + "=" * 70)
        print("✓ ALL EXAMPLES COMPLETED SUCCESSFULLY!")
        print("=" * 70)
        print("\nGenerated PPT Files:")
        print("  1. example1_raw_csv.pptx")
        print("  2. example2_tsv.pptx")
        print("  3. example3_pipe.pptx")
        print("  4. example4_json.pptx")
        print("  5. example5_dict.pptx")
        print("  6. example6_csv_file.pptx")
        print("  7. example8_auto_detect.pptx")
        print("  8. example9_space.pptx")
        print("  9. example10_dataframe.pptx")
        print("  10. example11_semicolon.pptx")
        print("=" * 70)
        
    except Exception as e:
        print(f"\n✗ Error: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
