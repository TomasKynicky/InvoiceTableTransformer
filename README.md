# InvoiceTableTransformer

AI-powered invoice and receipt processing tool that extracts structured data from documents and converts them to Excel spreadsheets.

## What it does

InvoiceTableTransformer accepts invoice and receipt files (PDF, JPG, PNG, BMP, TIFF, GIF) and uses Google Gemini AI to:
- Extract structured data from unstructured documents
- Parse invoice numbers, dates, amounts, and line items
- Convert extracted data into a formatted Excel spreadsheet
- Process multiple files in parallel

## Technologies

- **Python 3.14** with Streamlit web framework
- **Google Gemini AI** for document understanding
- **OpenPyXL** for Excel generation
- **Pillow** and **pdf2image** for document processing
- **Pandas** for data preview

## Features

- Multi-format support (PDF, JPG, PNG, BMP, TIFF, GIF)
- Parallel processing with ThreadPoolExecutor
- Smart amount parsing (handles various number formats including Czech decimal format)
- Real-time progress tracking
- Excel export with automatic formatting
- Data preview before download

## Installation

### Prerequisites
- Python 3.14+
- Google Gemini API key from https://aistudio.google.com/app/apikey
- Poppler (for PDF processing)

### Setup

```bash
# Create virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Install poppler (macOS)
brew install poppler
# On Ubuntu/Debian: sudo apt-get install poppler-utils
```

## Running the application

```bash
# Using provided script
./run.sh  # macOS/Linux
# or run.bat on Windows

# Manual execution
streamlit run app.py
```

The app opens at `http://localhost:8501`

## Environment setup

```bash
export GEMINI_API_KEY="your-api-key-here"
```

## Project structure

```
├── app.py                # Streamlit web interface
├── invoice_processor.py  # Core processing logic
├── requirements.txt      # Python dependencies
├── run.sh               # macOS/Linux run script
└── run.bat              # Windows run script
```

## How it works

1. Upload files via the web interface
2. Documents are converted to images (PDFs page by page)
3. Each image is processed by Gemini AI with a structured prompt
4. Extracted JSON data is normalized and formatted
5. Results are saved to Excel with proper column formatting
6. Preview and download the generated spreadsheet

## Note

This project uses Google Gemini API which may incur costs depending on your usage tier. Requires internet connection for API calls.
