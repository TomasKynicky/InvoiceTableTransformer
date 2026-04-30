# 🚀 InvoiceTableTransformer

<div align="center">

![Version](https://img.shields.io/badge/version-1.0.0-blue.svg)
![Python](https://img.shields.io/badge/python-3.14+-green.svg)
![Streamlit](https://img.shields.io/badge/streamlit-1.0+-red.svg)
![AI](https://img.shields.io/badge/powered%20by-Gemini%20AI-orange.svg)

**Transform your invoices and receipts into structured Excel data with the power of AI**

[Features](#-features) • [Quick Start](#-quick-start) • [Documentation](#-documentation) • [Contributing](#-contributing)

</div>

---

## 🎯 What is InvoiceTableTransformer?

InvoiceTableTransformer is a cutting-edge Python application that leverages **Google Gemini AI** to intelligently extract, parse, and transform unstructured invoice and receipt data into beautifully formatted Excel spreadsheets.

### ✨ Key Benefits

- **🤖 AI-Powered**: Uses state-of-the-art Gemini AI for accurate data extraction
- **⚡ Lightning Fast**: Parallel processing with multi-threading support
- **🎨 Smart Parsing**: Handles multiple number formats including Czech decimal notation
- **📊 Excel Ready**: Automatic column formatting and totals calculation
- **🌍 Multi-Format**: Supports PDF, JPG, PNG, BMP, TIFF, and GIF files

---

## 🚀 Features

<table>
<tr>
<td width="50%">

### Core Features
- ✅ **Multi-format Support**  
  Process PDFs, images (JPG, PNG, BMP, TIFF, GIF)

- ✅ **Parallel Processing**  
  ThreadPoolExecutor with 4 workers

- ✅ **Smart Amount Parsing**  
  Handles `2,975.00`, `2.975,00`, and implied decimals

- ✅ **Real-time Progress**  
  Live progress bar during batch processing

</td>
<td width="50%">

### Output Features
- ✅ **Excel Export**  
  Formatted spreadsheets with auto-sized columns

- ✅ **Data Preview**  
  Interactive table preview with Pandas

- ✅ **Line Item Extraction**  
  Quantities, unit prices, and totals

- ✅ **VAT Handling**  
  Separate columns for amounts with/without DPH

</td>
</tr>
</table>

---

## 🏗️ Architecture

```
┌─────────────────┐
│   Web Interface │
│   (Streamlit)   │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  File Upload    │
│  & Conversion   │◄── PDF → Images (poppler)
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│   Gemini AI     │
│  (Processing)   │◄── Structured JSON extraction
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  Data Normalization│
│  & Formatting   │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  Excel Export   │
│  (OpenPyXL)     │
└─────────────────┘
```

---

## ⚡ Quick Start

### Prerequisites

| Requirement | Description |
|------------|-------------|
| 🐍 Python | Version 3.14 or higher |
| 🔑 Gemini API Key | Get it from [Google AI Studio](https://aistudio.google.com/app/apikey) |
| 📦 Poppler | For PDF processing (system dependency) |

### Installation

<details>
<summary>📱 Click to expand installation steps</summary>

```bash
# 1. Clone the repository
cd /Users/tomaskynicky/Documents/Projects/Mamka

# 2. Create virtual environment
python -m venv venv

# 3. Activate virtual environment
# macOS/Linux:
source venv/bin/activate
# Windows:
# venv\Scripts\activate.bat

# 4. Install dependencies
pip install -r requirements.txt

# 5. Install poppler (for PDF support)
# macOS:
brew install poppler
# Ubuntu/Debian:
# sudo apt-get install poppler-utils
# Windows: Download poppler binaries
```

</details>

### Running the App

```bash
# Option 1: Using provided scripts (recommended)
./run.sh                    # macOS/Linux
# run.bat                   # Windows

# Option 2: Manual execution
streamlit run app.py
```

🌐 The app will automatically open at **http://localhost:8501**

### Environment Setup

```bash
export GEMINI_API_KEY="your-api-key-here"
```

---

## 📚 Documentation

### Extracted Data Fields

| Field | Description |
|-------|-------------|
| 📄 **Document Number** | Invoice or receipt number |
| 📅 **Date** | Document date (with timestamp for receipts) |
| 💰 **Total Amount** | Full amount including VAT |
| 💰 **Total without VAT** | Amount excluding DPH |
| 📋 **Line Items** | Individual items with quantity, price, and totals |

### Supported File Formats

```
📄 PDF  - Portable Document Format (converted page by page)
🖼️ JPG  - JPEG images
🖼️ PNG  - Portable Network Graphics
🖼️ BMP  - Bitmap images
🖼️ TIFF - Tagged Image File Format
🖼️ GIF  - Graphics Interchange Format
```

### Processing Pipeline

1. **Upload** → User uploads documents via Streamlit interface
2. **Convert** → PDFs are converted to images using pdf2image
3. **Analyze** → Each image sent to Gemini AI with optimized prompt
4. **Extract** → Gemini returns structured JSON with invoice data
5. **Normalize** → Data is cleaned and formatted (especially amounts)
6. **Export** → Results saved to Excel with proper formatting
7. **Download** → User previews and downloads the spreadsheet

---

## 🛠️ Tech Stack

<div align="center">

| Technology | Purpose | Version |
|-----------|---------|---------|
| ![Python](https://img.shields.io/badge/Python-3776A8?style=flat&logo=python&logoColor=white) | Core Language | 3.14+ |
| ![Streamlit](https://img.shields.io/badge/Streamlit-FF4B4B?style=flat&logo=streamlit&logoColor=white) | Web Framework | Latest |
| ![Gemini](https://img.shields.io/badge/Google%20Gemini-8E75B2?style=flat&logo=google&logoColor=white) | AI Processing | gemini-3-flash-preview |
| ![OpenPyXL](https://img.shields.io/badge/OpenPyXL-217346?style=flat&logo=microsoft-excel&logoColor=white) | Excel Generation | Latest |
| ![Pandas](https://img.shields.io/badge/Pandas-150458?style=flat&logo=pandas&logoColor=white) | Data Handling | Latest |

</div>

---

## 📁 Project Structure

```
InvoiceTableTransformer/
│
├── 🎨 app.py                 # Streamlit web interface & UI logic
├── ⚙️ invoice_processor.py   # Core AI processing engine (289 lines)
├── 📋 requirements.txt       # Python dependencies
├── 🚀 run.sh                 # macOS/Linux launcher
├── 🚀 run.bat                # Windows launcher
├── 📝 README.md              # This file
├── .gitignore                # Git ignore rules
└── 📦 venv/                  # Virtual environment
```

---

## ⚠️ Important Notes

<div align="center">

| 💡 Tip | ⚠️ Warning | 💰 Cost |
|--------|-----------|---------|
| Process files in batches for better performance | Requires internet for Gemini API | API usage may incur costs |
| Use high-quality scans for better accuracy | Poppler required for PDFs | Check [Google AI pricing](https://ai.google.dev/pricing) |
| Set API key as environment variable | No persistent storage | Monitor your usage quota |

</div>

---

## 🤝 Contributing

Contributions are welcome! Feel free to:
- 🐛 Report bugs
- 💡 Suggest features
- 🔧 Submit pull requests
- 📖 Improve documentation

---

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

---

<div align="center">

**Built with ❤️ and powered by Google Gemini AI**

[⬆ Back to top](#-invoicetabletransformer)

</div>
