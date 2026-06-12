Data Canvas - AI-Powered Data to PowerPoint Agent

Convert raw data into professional PowerPoint presentations with AI-generated insights and recommendations.

🚀 Features

Supported Input Formats

- ✅ CSV Files
- ✅ Excel Files (.xls, .xlsx)
- ✅ JSON Strings
- ✅ Python Dictionaries
- ✅ Text Files (.txt, .tsv, .dat)
- ✅ Automatic Format Detection

AI Capabilities

- ✅ AI-Generated Insights
- ✅ Automated Executive Summary
- ✅ Recommendations Generation
- ✅ Multi-Slide AI Analysis
- ✅ Automated PowerPoint Creation
- ✅ End-to-End Data-to-Presentation Workflow

---

📋 Installation

git clone https://github.com/poojaofficial2203/data-canvas-pk.git

cd data-canvas-pk

pip install -r requirements.txt

---

🚀 Quick Demo

from agent import run_agent

run_agent("test.csv")

Output:

- Professional PowerPoint Presentation
- AI Insights Slides
- Recommendations Slide
- Automated Data Analysis

---

📊 Generated Presentation Includes

1. Title Slide
2. Data Overview
3. Data Preview
4. Column Summary
5. Statistical Analysis
6. AI Insights Slides
7. Recommendations Slide

---

🧠 Agent Workflow

Input Data (CSV / Excel / JSON)
                ↓
        Data Analysis
                ↓
    Statistical Processing
                ↓
     AI Insight Generation
                ↓
 Recommendation Generation
                ↓
     PowerPoint Creation
                ↓
 Professional Presentation

---

📂 Usage Examples

CSV File

from data_canvas import UniversalDataToPPT

converter = UniversalDataToPPT()
converter.convert_file("data.csv", "report.pptx")

Excel File

converter.convert_file("data.xlsx", "report.pptx")

JSON String

json_data = '[{"name":"John","age":30},{"name":"Jane","age":28}]'

converter.convert_json_string(
    json_data,
    "report.pptx"
)

Python Dictionary

data = {
    "Month": ["Jan", "Feb", "Mar"],
    "Sales": [1000, 1200, 1500]
}

converter.convert_dict(
    data,
    "report.pptx"
)

---

📈 Supported Data Formats

Format| Supported
CSV| ✅
Excel (.xls/.xlsx)| ✅
JSON| ✅
Text Files| ✅
Python Dictionaries| ✅

---

🎯 Use Cases

- Business Reporting
- Executive Presentations
- Automated Reporting
- Data Analysis
- Research Presentations
- Financial Reporting
- Data Storytelling

---

🏆 Hackathon Highlights

- AI-Powered Data Analysis
- AI Insight Generation
- Automated Recommendations
- PowerPoint Automation
- End-to-End Agent Workflow
- One-Command Presentation Creation

---

🔄 End-to-End Flow

Raw Data
   ↓
Data Processing
   ↓
AI Analysis
   ↓
AI Insights
   ↓
Recommendations
   ↓
PowerPoint Generation
   ↓
Final Presentation

---

🛠 Project Structure

data-canvas-pk/

├── data_canvas/
│   ├── ai_insights.py
│   ├── universal_converter.py
│   └── __init__.py
│
├── agent.py
├── README.md
├── requirements.txt
└── test.csv

---

🐛 Error Handling

success = converter.convert_file(
    "data.csv",
    "output.pptx"
)

if not success:
    print("Conversion failed")

---

📞 Support

For issues or questions:

1. Review the examples
2. Check the implementation guide
3. Open a GitHub issue

---

📄 License

Open Source – Free to use and modify.
