# 🔍 Serial Number Validation System

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://python.org)
[![Flask](https://img.shields.io/badge/Flask-3.0+-green.svg)](https://flask.palletsprojects.com)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

A comprehensive web application for validating serial numbers against putaway records with advanced business rule validation, duplicate detection, and detailed reporting.

![Application Screenshot](docs/screenshot.png)

## ✨ Features

### 🎯 Core Functionality
- **Advanced Validation Logic**: Business rule validation with exception handling
- **Duplicate Detection**: Automatically identifies and reports duplicate serial numbers
- **Comprehensive Reporting**: Detailed validation results with statistics
- **Excel Integration**: Native Excel file processing with auto-fitted columns

### 🎨 User Experience
- **Drag & Drop Upload**: Simply drag Excel files into upload zones
- **Web-based Interface**: Modern, responsive design
- **Real-time Feedback**: Visual indicators for file selection and validation progress
- **Cross-platform**: Works on Windows, macOS, and Linux

### 📊 Reporting
- **Detailed Statistics**: Success rates, validation counts, error analysis
- **Duplicate Tracking**: Duplicates listed at the end of results
- **Excel Export**: Formatted reports with auto-fitted columns
- **Data Preview**: View results before downloading

## 🚀 Quick Start

### Prerequisites
- Python 3.8 or higher
- Web browser (Chrome, Firefox, Safari, Edge)

### Installation

1. **Clone the repository**:
   ```bash
   git clone https://github.com/joexgit2024/testingvalidation.git
   cd testingvalidation
   ```

2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the application**:
   ```bash
   python app.py
   ```
   
   Or on Windows, double-click `start_app.bat`

4. **Open in browser**:
   ```
   http://localhost:5000
   ```

## 📋 How It Works

### Validation Logic

The system applies sophisticated business rules to validate serial numbers:

| Test Result | Validation Rule | Expected Outcome |
|-------------|----------------|------------------|
| **PASS** | Part numbers match AND Rack Bin ends with 'P' | ✅ Valid |
| **FAIL/MAYBE** | Part numbers match AND Rack Bin does NOT end with 'P' | ✅ Valid |
| **Exception** | Rack Bin is 'Testwip' (case-insensitive) | ✅ Always Valid |
| **Missing** | Serial Number not found in putaway records | ❌ Invalid |

### File Requirements

Both Excel files must contain these columns:
- `Serial Number` - Unique identifier
- `Part Number` - Product part number
- `Test Result` - Test outcome (PASS/FAIL/MAYBE)
- `Rack Bin` - Storage location

## 🎯 Usage Guide

### Step 1: Upload Files
1. Open the web application
2. **Drag and drop** Excel files into upload zones, OR click to browse
3. Select CEVA putaway records file (left zone)
4. Select test records file (right zone)
5. Button turns green when both files are ready

### Step 2: View Results
- Review validation statistics
- Check for duplicate serial numbers
- Preview first 20 records
- Valid records highlighted in green, invalid in red

### Step 3: Download Results
- Click "Download Excel Report"
- Get comprehensive validation results
- Duplicates listed at the end
- Auto-fitted columns for easy reading

## 📁 Project Structure

```
testingvalidation/
├── app.py                    # Main Flask application
├── main.py                   # Legacy command-line version
├── requirements.txt          # Python dependencies
├── start_app.bat            # Windows startup script
├── README.md                # This file
├── .gitignore              # Git ignore rules
├── templates/              # HTML templates
│   ├── index.html          # Upload page
│   └── results.html        # Results page
├── static/                 # Static assets
│   └── css/               # Stylesheets
├── input/                 # Legacy input directory
├── output/                # Generated reports
└── uploads/               # Temporary uploads
```

## 🔧 API Reference

### Endpoints

- `GET /` - Main upload interface
- `POST /upload` - Process file uploads and validation
- `GET /download/<filename>` - Download validation results

### File Upload
- **Supported formats**: .xlsx, .xls
- **Maximum size**: 16MB per file
- **Validation**: Automatic file type checking

## 🛠️ Development

### Running in Development Mode
```bash
python app.py
```
- Debug mode enabled
- Auto-reload on file changes
- Detailed error messages

### Creating New Features
1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## 🚨 Troubleshooting

### Common Issues

**File Upload Problems**
- Ensure files are Excel format (.xlsx/.xls)
- Check file size (max 16MB)
- Verify required columns exist

**Validation Errors**
- Confirm exact column names
- Check for special characters
- Ensure data types are correct

**Browser Issues**
- Enable JavaScript
- Clear browser cache
- Try different browser

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request. For major changes, please open an issue first to discuss what you would like to change.

## 📞 Support

For questions or issues:
- Open an [issue](https://github.com/joexgit2024/testingvalidation/issues)
- Check existing documentation
- Review troubleshooting guide

## 🏆 Acknowledgments

- Built with [Flask](https://flask.palletsprojects.com/) web framework
- Excel processing powered by [pandas](https://pandas.pydata.org/) and [openpyxl](https://openpyxl.readthedocs.io/)
- UI styled with [Bootstrap](https://getbootstrap.com/)

---

**Made with ❤️ for efficient serial number validation**
