# 📄 PDF to WELL Certification Excel Converter

A Flask web application that processes PDF files to extract the first 2 pages, converts them to Markdown using PyPDF2, and creates a comprehensive WELL certification Excel file with proper concept parsing, sub-points calculation, and percentage analysis.

## ✨ Features

- **Bulk PDF Upload**: Drag & drop multiple PDF files for processing
- **First 2 Pages Extraction**: Extracts only the first 2 pages from each PDF
- **PyPDF2 Conversion**: Reliable PDF to Markdown conversion using PyPDF2
- **WELL Certification Parsing**: Advanced parsing logic for WELL building certification data
- **Structured Excel Export**: Creates Excel files with:
  - Concept columns (Air, Water, Nourishment, Light, Movement, Thermal Comfort, Sound, Materials, Mind, Community, Innovation)
  - Sub-points calculation per concept
  - Percentage analysis
  - Individual part code tracking (A01.1, A01.2, etc.)
  - Project metadata (Project ID, Name, Date Certified, Total Points)

## 🏗️ Architecture

- **Backend**: Flask web framework with PyPDF2 for PDF processing
- **Frontend**: Modern HTML5/CSS3/JavaScript with drag & drop interface
- **Excel Generation**: OpenPyXL for WELL certification Excel creation
- **File Handling**: Secure file processing with automatic cleanup

## 🚀 Prerequisites

- Python 3.8 or higher
- pip (Python package installer)

## 📦 Installation

1. **Clone or download the project**
   ```bash
   git clone <repository-url>
   cd PDFconvert
   ```

2. **Create a virtual environment (recommended)**
   ```bash
   python -m venv venv
   
   # On Windows
   venv\Scripts\activate
   
   # On macOS/Linux
   source venv/bin/activate
   ```

3. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

## 🎯 Usage

### Starting the Application

**Option 1: Direct Python execution**
```bash
python app.py
```

**Option 2: Using the provided scripts**
```bash
# Windows
start_app.bat

# macOS/Linux
./start_app.sh
```

### Using the Web Interface

1. **Open your browser** and navigate to `http://localhost:5000`
2. **Upload PDFs**: Drag and drop PDF files or click to browse
3. **Review files**: Check the file list, remove unwanted files if needed
4. **Process files**: Click "Process Files" to start conversion
5. **Download results**: Download individual Excel files or export all at once

## 🔧 How It Works

### 1. PDF Processing
- **File Upload**: Multiple PDFs can be uploaded simultaneously
- **Page Extraction**: First 2 pages are extracted using PyPDF2
- **Text Extraction**: Text content is extracted from each page

### 2. Markdown Conversion
- **Content Processing**: Extracted text is formatted into Markdown structure
- **Page Organization**: Each page is clearly marked with headers
- **Text Cleaning**: Automatic formatting and structure preservation

### 3. Excel Export
- **Data Parsing**: Markdown content is parsed for structured data
- **Template Matching**: Uses intelligent parsing to identify project information
- **Excel Generation**: Creates organized Excel files with proper formatting

## 📁 Project Structure

```
PDFconvert/
├── app.py                          # Main Flask application
├── config.py                       # Configuration settings
├── requirements.txt                # Python dependencies
├── templates/
│   └── index.html                 # Web interface template
├── uploads/                        # Temporary upload storage
├── processed/                      # Generated files storage
├── start_app.bat                  # Windows startup script
├── start_app.sh                   # Unix startup script
├── .gitignore                     # Git ignore rules
└── README.md                      # This file
```

## 🌐 API Endpoints

- **GET /** - Main application interface
- **POST /upload** - Handle PDF file uploads and processing
- **GET /download-excel?file=<filename>** - Download generated Excel files using file parameter
- **POST /clear-session** - Clear session data

## ⚙️ Configuration

The application uses `config.py` for centralized configuration:

- **Server Settings**: Host, port, debug mode
- **File Paths**: Upload and processed directories
- **File Limits**: Maximum file size and allowed extensions
- **Security**: Secret key and session settings

## 🔍 Troubleshooting

### Common Issues

1. **Port already in use**
   - Change the port in `config.py`
   - Kill existing processes using the port

2. **File upload errors**
   - Check file size limits in `config.py`
   - Ensure PDF files are valid and not corrupted

3. **Excel download issues**
   - Check browser download settings
   - Verify file permissions in the `processed/` directory

### Performance Tips

- **Large files**: Process files in smaller batches
- **Memory usage**: Monitor system resources during bulk processing
- **Storage**: Ensure adequate disk space for temporary files

## 🛡️ Security Features

- **File validation**: Only PDF files are accepted
- **Secure filenames**: Automatic filename sanitization
- **Session management**: Secure session handling
- **File cleanup**: Automatic removal of temporary files

## 📄 License

This project is open source and available under the [MIT License](LICENSE).

## 🙏 Acknowledgments

- **PyPDF2**: PDF processing library
- **Flask**: Web framework
- **OpenPyXL**: Excel file generation
- **Pandas**: Data manipulation and analysis

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

### Development Setup

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## 📞 Support

If you encounter any issues or have questions:

1. Check the troubleshooting section above
2. Review the error logs in the terminal
3. Open an issue on the project repository
4. Contact the development team

---

**Note**: This application uses PyPDF2 for reliable PDF processing and text extraction. The conversion quality depends on the PDF structure and text encoding.
