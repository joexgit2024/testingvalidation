# Installation Guide

## System Requirements

- **Operating System**: Windows 10/11, macOS 10.14+, or Linux
- **Python**: Version 3.8 or higher
- **Memory**: Minimum 512MB RAM
- **Storage**: 100MB free space
- **Browser**: Chrome, Firefox, Safari, or Edge

## Installation Methods

### Method 1: Quick Setup (Recommended)

1. **Download the project**:
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

### Method 2: Virtual Environment (Recommended for Development)

1. **Create virtual environment**:
   ```bash
   python -m venv venv
   ```

2. **Activate virtual environment**:
   - Windows: `venv\Scripts\activate`
   - macOS/Linux: `source venv/bin/activate`

3. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

4. **Run the application**:
   ```bash
   python app.py
   ```

### Method 3: Windows Batch File

1. Double-click `start_app.bat`
2. The application will start automatically

## Verification

After installation, verify the setup:

1. Open browser to `http://localhost:5000`
2. You should see the upload interface
3. Try uploading sample Excel files
4. Check that validation works correctly

## Troubleshooting

### Common Installation Issues

**Python not found**:
- Install Python from [python.org](https://python.org)
- Ensure Python is added to PATH

**Permission errors**:
- Run command prompt as administrator (Windows)
- Use `sudo` for installation commands (macOS/Linux)

**Port 5000 already in use**:
- Change port in `app.py`: `app.run(port=5001)`
- Or stop the service using port 5000

**Module import errors**:
- Verify virtual environment is activated
- Reinstall dependencies: `pip install -r requirements.txt --force-reinstall`

## Next Steps

After successful installation:
1. Read the [User Guide](user-guide.md)
2. Review [API Documentation](api.md)
3. Check [Troubleshooting Guide](troubleshooting.md)
