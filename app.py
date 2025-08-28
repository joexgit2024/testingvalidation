import os
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, flash, send_file, jsonify
from werkzeug.utils import secure_filename
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import tempfile
from datetime import datetime

app = Flask(__name__)
app.secret_key = 'your-secret-key-change-this'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def validate_row(row):
    """Validate a single row based on business logic"""
    if pd.isna(row['Part Number_putaway']):
        return 'NOT OK: Serial Number missing in putaway records'
    
    part_match = row['Part Number_test'] == row['Part Number_putaway']
    test_result = str(row['Test Result']).strip().upper()
    rack_bin = str(row['Rack Bin']).strip().upper()
    
    if test_result == 'PASS':
        rack_bin_ok = rack_bin.endswith('P')
        result = part_match and rack_bin_ok
        return 'OK' if result else 'NOT OK: Part mismatch or Rack Bin not ending with P'
    elif test_result in ['FAIL', 'MAYBE']:
        # Exception: if Rack Bin is 'Testwip' (case-insensitive), mark as OK
        if rack_bin == 'TESTWIP':
            return 'OK (Exception: Rack Bin is Testwip)'
        else:
            rack_bin_ok = not rack_bin.endswith('P')
            result = part_match and rack_bin_ok
            return 'OK' if result else 'NOT OK: Part mismatch or Rack Bin ends with P'
    else:
        return 'NOT OK: Unknown Test Result'

def check_duplicates(df_test, df_putaway):
    """Check for duplicate serial numbers in both dataframes"""
    dup_test = df_test['Serial Number'].duplicated()
    dup_putaway = df_putaway['Serial Number'].duplicated()
    
    test_duplicates = df_test[dup_test]['Serial Number'].tolist() if dup_test.any() else []
    putaway_duplicates = df_putaway[dup_putaway]['Serial Number'].tolist() if dup_putaway.any() else []
    
    return test_duplicates, putaway_duplicates

def format_dates(df_merged):
    """Format date columns"""
    for col in ['Testing Date_test', 'Testing Date_putaway']:
        if col in df_merged.columns:
            df_merged[col] = pd.to_datetime(df_merged[col], errors='coerce').dt.date
    return df_merged

def save_with_autofit(df_merged, output_path):
    """Save DataFrame to Excel with autofit columns"""
    df_merged.to_excel(output_path, index=False)
    wb = load_workbook(output_path)
    ws = wb.active
    
    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[col_letter].width = max_length + 2
    
    wb.save(output_path)

def process_validation(putaway_file, test_file):
    """Process the validation logic and return results"""
    try:
        # Load the Excel files
        df_putaway = pd.read_excel(putaway_file, engine='openpyxl')
        df_test = pd.read_excel(test_file, engine='openpyxl')
        
        # Check for duplicates
        test_duplicates, putaway_duplicates = check_duplicates(df_test, df_putaway)
        
        # Merge dataframes
        df_merged = pd.merge(df_test, df_putaway, on='Serial Number', how='left', suffixes=('_test', '_putaway'))
        
        # Apply validation logic
        df_merged['Validation Result'] = df_merged.apply(validate_row, axis=1)
        
        # Format dates
        df_merged = format_dates(df_merged)
        
        # Add duplicate information at the end
        if test_duplicates or putaway_duplicates:
            # Create summary rows for duplicates
            summary_rows = []
            
            if test_duplicates:
                summary = {col: '' for col in df_merged.columns}
                summary['Serial Number'] = 'DUPLICATES IN TEST RECORDS'
                summary['Validation Result'] = f'Duplicate Serial Numbers found: {", ".join(map(str, test_duplicates))}'
                summary_rows.append(summary)
            
            if putaway_duplicates:
                summary = {col: '' for col in df_merged.columns}
                summary['Serial Number'] = 'DUPLICATES IN PUTAWAY RECORDS'
                summary['Validation Result'] = f'Duplicate Serial Numbers found: {", ".join(map(str, putaway_duplicates))}'
                summary_rows.append(summary)
            
            # Add empty row for separation
            empty_row = {col: '' for col in df_merged.columns}
            df_merged = pd.concat([df_merged, pd.DataFrame([empty_row] + summary_rows)], ignore_index=True)
        
        # Generate statistics
        total_records = len(df_merged) - (2 if test_duplicates or putaway_duplicates else 0) - (1 if test_duplicates or putaway_duplicates else 0)  # Exclude summary rows
        ok_records = len(df_merged[df_merged['Validation Result'].str.startswith('OK', na=False)])
        not_ok_records = len(df_merged[df_merged['Validation Result'].str.startswith('NOT OK', na=False)])
        
        stats = {
            'total_records': total_records,
            'ok_records': ok_records,
            'not_ok_records': not_ok_records,
            'test_duplicates': test_duplicates,
            'putaway_duplicates': putaway_duplicates
        }
        
        return df_merged, stats, None
        
    except Exception as e:
        return None, None, str(e)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/sharepoint-info')
def sharepoint_info():
    """Provide information about SharePoint integration"""
    return jsonify({
        'status': 'info',
        'message': 'SharePoint direct access requires authentication setup',
        'instructions': [
            'Open the SharePoint link in a new tab',
            'Download the "Test Records.xlsx" file',
            'Return here and drag/drop the downloaded file'
        ],
        'alternative': 'Consider setting up SharePoint API with proper authentication for automation'
    })

@app.route('/sample-file')
def get_sample_file():
    """Provide a sample test file if available"""
    sample_path = os.path.join('input', 'Test Records.xlsx')
    if os.path.exists(sample_path):
        return send_file(sample_path, as_attachment=True, download_name='Sample_Test_Records.xlsx')
    else:
        return jsonify({
            'status': 'error',
            'message': 'Sample file not available. Please upload your own test records file.'
        }), 404

@app.route('/upload', methods=['POST'])
def upload_files():
    if 'putaway_file' not in request.files or 'test_file' not in request.files:
        flash('Please select both files')
        return redirect(request.url)
    
    putaway_file = request.files['putaway_file']
    test_file = request.files['test_file']
    
    if putaway_file.filename == '' or test_file.filename == '':
        flash('Please select both files')
        return redirect(url_for('index'))
    
    if not (allowed_file(putaway_file.filename) and allowed_file(test_file.filename)):
        flash('Only Excel files (.xlsx, .xls) are allowed')
        return redirect(url_for('index'))
    
    try:
        # Create temporary files for processing
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as putaway_temp:
            putaway_file.save(putaway_temp.name)
            putaway_temp_path = putaway_temp.name
        
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as test_temp:
            test_file.save(test_temp.name)
            test_temp_path = test_temp.name
        
        # Process validation
        df_merged, stats, error = process_validation(putaway_temp_path, test_temp_path)
        
        if error:
            flash(f'Error processing files: {error}')
            return redirect(url_for('index'))
        
        # Save results
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_filename = f'validation_results_{timestamp}.xlsx'
        output_path = os.path.join('output', output_filename)
        
        # Ensure output directory exists
        os.makedirs('output', exist_ok=True)
        
        save_with_autofit(df_merged, output_path)
        
        # Clean up temporary files
        os.unlink(putaway_temp_path)
        os.unlink(test_temp_path)
        
        return render_template('results.html', 
                             stats=stats, 
                             output_file=output_filename,
                             df_preview=df_merged.head(20).to_html(classes='table table-striped', table_id='preview-table'))
        
    except Exception as e:
        flash(f'Error processing files: {str(e)}')
        return redirect(url_for('index'))

@app.route('/download/<filename>')
def download_file(filename):
    try:
        return send_file(os.path.join('output', filename), as_attachment=True)
    except Exception as e:
        flash(f'Error downloading file: {str(e)}')
        return redirect(url_for('index'))

if __name__ == '__main__':
    # Ensure required directories exist
    os.makedirs('uploads', exist_ok=True)
    os.makedirs('output', exist_ok=True)
    
    app.run(debug=True, host='0.0.0.0', port=5000)
