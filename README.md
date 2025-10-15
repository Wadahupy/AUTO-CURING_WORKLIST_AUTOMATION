# Auto Curing Worklist Helper

A Streamlit web application template for processing Excel files. This application provides a foundation for implementing custom Excel file processing, comparison, and analysis features.

## Features

- Excel file upload and preview
- Basic file statistics display
- Template for custom data processing
- Download processed results
- Extensible structure for adding new features

## Installation

1. Clone this repository
2. Create a virtual environment (recommended):
```powershell
python -m venv venv
.\venv\Scripts\activate
```

3. Install the required packages:
```powershell
pip install -r requirements.txt
```

## Running the Application

To run the Streamlit app:

```powershell
cd src
streamlit run app.py
```

The application will open in your default web browser at `http://localhost:8501`.

## Project Structure

```
├── requirements.txt    # Python dependencies
├── src/
│   ├── app.py         # Main Streamlit application
│   └── utils.py       # Utility functions for data processing
```

## Development

### Adding New Features

1. **Custom Processing Logic**: 
   - Modify the `process_excel_file()` function in `utils.py` to implement your specific processing requirements.

2. **File Comparison**:
   - Implement the `compare_excel_files()` function in `utils.py` to add file comparison capabilities.

3. **File Merging**:
   - Implement the `merge_excel_files()` function in `utils.py` to add file merging capabilities.

### Best Practices

- Add error handling for specific Excel file formats
- Implement input validation for uploaded files
- Add progress indicators for long-running operations
- Include data validation before processing
- Add logging for debugging purposes
