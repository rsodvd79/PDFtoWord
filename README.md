# PDFtoWord

![Python](https://img.shields.io/badge/python-v3.8+-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)

A powerful Python script for extracting images from PDF files, performing OCR (Optical Character Recognition), translating text, and converting to DOCX format.

## Features

- 🔍 **PDF Image Extraction**: Automatically extracts all images from PDF files
- 🔄 **Smart Image Rotation**: Detects and corrects image orientation for better OCR accuracy
- 📝 **OCR Processing**: Converts images to text using Tesseract OCR
- 🌍 **Multi-language Support**: Supports Italian and English OCR with automatic translation
- 📄 **Multiple Output Formats**: Generates TXT, HTML, HOCR, and DOCX files
- 🔗 **Batch Processing**: Processes multiple PDF files in a single run
- 🎯 **Auto-detection**: Automatically finds Tesseract installation

## Prerequisites

### Required Software

1. **Python 3.8+**
2. **Tesseract OCR** - Download from [GitHub Releases](https://github.com/tesseract-ocr/tesseract/releases)
   - Windows: Install the `.exe` file
   - The script will automatically detect Tesseract installation

### Required Python Libraries

All dependencies are listed in `requirements.txt`:

```
PyMuPDF==1.23.0
Pillow==10.0.0
pytesseract==0.3.10
beautifulsoup4==4.12.2
python-docx==0.8.11
googletrans==4.0.0rc1
```

## Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/rsodvd79/PDFtoWord.git
   cd PDFtoWord
   ```

2. **Install Python dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Install Tesseract OCR:**
   - Download from [Tesseract releases](https://github.com/tesseract-ocr/tesseract/releases)
   - Follow the installation instructions for your operating system
   - The script will automatically detect the installation

## Usage

### Basic Usage

1. **Prepare your PDF files:**
   - Create a `PDF` folder in the project directory
   - Place your PDF files inside the `PDF` folder

2. **Run the script:**
   ```bash
   python PDFtoWord.py
   ```

3. **Select OCR language:**
   - Choose `1` for Italian (with English translation)
   - Choose `2` for English (with Italian translation)

### Output Structure

The script creates the following directory structure:

```
📁 IMG/
  └── 📁 [PDF_NAME]/
      ├── img_0001.png
      ├── img_0002.png
      └── ...

📁 TXT/
  └── 📁 [PDF_NAME]/
      ├── img_0001.png.txt
      ├── img_0002.png.txt
      └── ...

📁 HTML/
  └── 📁 [PDF_NAME]/
      ├── img_0001.png.html
      ├── img_0001.png.hocr
      ├── [PDF_NAME]_libro.txt.html
      └── ...

📁 risultato/
  ├── [PDF_NAME]_libro.txt
  ├── [PDF_NAME]_libro.docx
  ├── [PDF_NAME]_libro_translated_[LANG].txt
  ├── [PDF_NAME]_libro_translated_[LANG].docx
  ├── libro.txt (aggregated)
  ├── libro.docx (aggregated)
  ├── libro_translated_[LANG].txt (aggregated)
  └── libro_translated_[LANG].docx (aggregated)
```

## How It Works

1. **Image Extraction**: The script uses PyMuPDF to extract all images from PDF files
2. **Orientation Detection**: Uses Tesseract's OSD (Orientation and Script Detection) to detect and correct image rotation
3. **OCR Processing**: Converts images to text using Tesseract OCR
4. **Text Processing**: Merges hyphenated words that are split across lines
5. **Translation**: Uses Google Translate API to translate text to the target language
6. **Format Conversion**: Creates multiple output formats (TXT, HTML, HOCR, DOCX)

## Configuration

### Tesseract Path

The script automatically detects Tesseract installation in common locations:
- `C:\Program Files\Tesseract-OCR\tesseract.exe`
- `C:\Program Files (x86)\Tesseract-OCR\tesseract.exe`
- User AppData directories
- Current directory
- System PATH

If automatic detection fails, you can manually set the path in the script:
```python
pytesseract.pytesseract.tesseract_cmd = r'C:\Path\To\Your\tesseract.exe'
```

### Supported Languages

- **Italian** (`ita`) - with English translation
- **English** (`eng`) - with Italian translation

Additional languages can be added by modifying the script and ensuring the corresponding Tesseract language packs are installed.

## Error Handling

The script includes robust error handling for:
- Missing Tesseract installation
- Corrupted or unreadable images
- Translation service errors
- File I/O operations

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Author

**Davide Rosa** - [rsodvd79](https://github.com/rsodvd79)

## Acknowledgments

- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) for optical character recognition
- [PyMuPDF](https://github.com/pymupdf/PyMuPDF) for PDF processing
- [Google Translate](https://cloud.google.com/translate) for translation services

## Troubleshooting

### Common Issues

1. **Tesseract not found:**
   - Ensure Tesseract is properly installed
   - Check that the installation path is correct
   - Verify Tesseract is in your system PATH

2. **Poor OCR quality:**
   - Ensure images have good resolution
   - Check if images need manual rotation
   - Consider preprocessing images for better contrast

3. **Translation errors:**
   - Check internet connectivity
   - Verify Google Translate service availability
   - Consider rate limiting for large documents

4. **Memory issues with large PDFs:**
   - Process PDFs one at a time
   - Reduce image resolution if necessary
   - Consider splitting large PDFs into smaller files

## Support

If you encounter any issues or have questions, please [open an issue](https://github.com/rsodvd79/PDFtoWord/issues) on GitHub.
