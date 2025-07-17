# 🧠 Name and Data Parser Web App

This project is a Streamlit-based web application designed to parse and clean unstructured name and data fields from Excel files—especially those containing complex entries like names, zip codes, SSNs, and plan numbers.

It was built to streamline manual data cleanup for both **mail** and **email** file formats, where entries are often inconsistently formatted and difficult to standardize by hand.

---

## 🚀 Features

- **Name Parsing**: Extracts first, middle, and last names, including handling of suffixes (e.g., Jr., III), particles (e.g., "van der"), and multiple middle names.
- **Field Cleanup**: Standardizes:
  - U.S. zip codes
  - Social Security Numbers (SSNs)
  - 401k plan ID numbers
- **Source Type Tagging**: Detects whether each row originates from a mail or email file.
- **Batch Processing**: Accepts entire Excel files and returns cleaned, structured data.
- **Browser-Based UI**: Built with [Streamlit](https://streamlit.io) for an easy-to-use, interactive interface.
- **Downloadable Results**: Cleaned files can be downloaded directly as CSVs.

---

## 🧰 Tech Stack

- **Python**
- **Pandas** – for data manipulation
- **Streamlit** – for the web UI
- **Regex & Heuristics** – for smart parsing
- **GitHub + Streamlit Cloud** – for version control and deployment

---

## 🔐 Data Privacy

This project was designed with **data privacy** in mind:
- Sensitive test data is excluded from version control via `.gitignore`
- Only public code and dummy sample files are stored in the repo
- No input data is ever stored on the server

---

## 🖥️ How to Use

1. Clone this repository or visit the [deployed Streamlit app](https://name-parser-app-2cxmbprxkjfkhr3vhbh2mj.streamlit.app/) 
2. Upload your Excel file
3. Select whether it's a **mail** or **email** file
4. Get your cleaned data and download it as a CSV
