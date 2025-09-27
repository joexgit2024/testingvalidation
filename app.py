import os
import subprocess
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, flash, send_file, jsonify
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment
import tempfile
from datetime import datetime

app = Flask(__name__)
app.secret_key = 'your-secret-key-change-this'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Email configuration
EMAIL_CONFIG = {
    'smtp_server': 'smtp.gmail.com',  # You may need to change this based on your email provider
    'smtp_port': 587,
    'sender_email': 'joe.xie@andrew.com',
    'sender_password': '',  # You'll need to set this up with app-specific password
    'to_email': 'Binoy.Thomas@Cevalogistics.com',
    'cc_emails': [
        'kiutau.taufa@cevalogistics.com',
        'jamaica.guevarra@cevalogistics.com', 
        'Martin.Caruana@andrew.com',
        'George.Yap@andrew.com'
    ]
}

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


def append_duplicate_summary_rows(df, test_duplicates, putaway_duplicates):
    """Append human-readable duplicate summaries to the provided DataFrame."""

    if not test_duplicates and not putaway_duplicates:
        return df

    summary_rows = []

    # Empty spacer row to separate the summaries from data
    empty_row = {column: '' for column in df.columns}
    summary_rows.append(empty_row)

    if test_duplicates:
        summary = {column: '' for column in df.columns}
        summary['Serial Number'] = 'DUPLICATES IN TEST RECORDS'
        summary['Validation Result'] = f'Duplicate Serial Numbers found: {", ".join(map(str, test_duplicates))}'
        summary_rows.append(summary)

    if putaway_duplicates:
        summary = {column: '' for column in df.columns}
        summary['Serial Number'] = 'DUPLICATES IN PUTAWAY RECORDS'
        summary['Validation Result'] = f'Duplicate Serial Numbers found: {", ".join(map(str, putaway_duplicates))}'
        summary_rows.append(summary)

    return pd.concat([df, pd.DataFrame(summary_rows)], ignore_index=True)

def format_dates(df_merged):
    """Format date columns"""
    for col in ['Testing Date_test', 'Testing Date_putaway']:
        if col in df_merged.columns:
            df_merged[col] = pd.to_datetime(df_merged[col], errors='coerce').dt.date
    return df_merged

def save_with_filtered_columns_and_summary(df_merged, output_path, stats):
    """Save DataFrame to Excel with filtered columns and summary sheet"""
    
    # Define the columns to keep in the specified order
    columns_to_keep = [
        'Testing Date_test',
        'Part Number_test', 
        'Serial Number',
        'Testing Date_putaway',
        'Part Number_putaway',
        'Rack Bin',
        'Test Result',
        'Validation Result'
    ]
    
    # Filter dataframe to only include specified columns
    # Only include columns that exist in the dataframe
    available_columns = [col for col in columns_to_keep if col in df_merged.columns]
    df_filtered = df_merged[available_columns].copy()
    
    # Create Excel writer object
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # Write main data sheet
        df_filtered.to_excel(writer, sheet_name='Validation Results', index=False)
        
        # Create summary sheet
        create_summary_sheet(writer, stats, df_filtered)
        
        # Get workbook and apply formatting
        workbook = writer.book
        
        # Format the main data sheet
        format_main_sheet(workbook['Validation Results'])
        
        # Format the summary sheet
        format_summary_sheet(workbook['Summary'])

def create_summary_sheet(writer, stats, df_filtered):
    """Create a summary sheet with statistics"""
    
    # Prepare summary data
    summary_data = []
    
    # Basic statistics
    summary_data.append(['Validation Summary', ''])
    summary_data.append(['', ''])
    summary_data.append(['Total Records Processed', stats['total_records']])
    summary_data.append(['Valid Records', stats['ok_records']])
    summary_data.append(['Invalid Records', stats['not_ok_records']])
    
    # Calculate success rate
    success_rate = (stats['ok_records'] / stats['total_records'] * 100) if stats['total_records'] > 0 else 0
    summary_data.append(['Success Rate', f'{success_rate:.1f}%'])
    summary_data.append(['', ''])
    
    # Duplicate information
    summary_data.append(['Duplicate Analysis', ''])
    summary_data.append(['', ''])
    if stats['test_duplicates']:
        summary_data.append(['Duplicates in Test Records', len(stats['test_duplicates'])])
    else:
        summary_data.append(['Duplicates in Test Records', 'None'])
        
    if stats['putaway_duplicates']:
        summary_data.append(['Duplicates in Putaway Records', len(stats['putaway_duplicates'])])
    else:
        summary_data.append(['Duplicates in Putaway Records', 'None'])
    
    summary_data.append(['', ''])
    
    # Validation breakdown
    summary_data.append(['Validation Breakdown', ''])
    summary_data.append(['', ''])
    
    # Count different validation results
    validation_counts = df_filtered['Validation Result'].value_counts()
    for result, count in validation_counts.items():
        if not result.startswith('---'):  # Skip separator rows
            summary_data.append([result, count])
    
    summary_data.append(['', ''])
    summary_data.append(['Report Generated', datetime.now().strftime('%Y-%m-%d %H:%M:%S')])
    summary_data.append(['Generated By', 'Serial Number Validation System'])
    
    # Create DataFrame and write to Excel
    summary_df = pd.DataFrame(summary_data, columns=['Metric', 'Value'])
    summary_df.to_excel(writer, sheet_name='Summary', index=False)

def format_main_sheet(worksheet):
    """Apply formatting to the main data sheet"""
    
    # Header formatting
    header_font = Font(bold=True, color='FFFFFF')
    header_fill = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
    header_alignment = Alignment(horizontal='center', vertical='center')
    
    # Apply header formatting
    for cell in worksheet[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
    
    # Auto-fit columns
    for column in worksheet.columns:
        max_length = 0
        column_letter = get_column_letter(column[0].column)
        for cell in column:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        worksheet.column_dimensions[column_letter].width = max_length + 2
    
    # Color code validation results
    for row in worksheet.iter_rows(min_row=2):
        validation_cell = row[-1]  # Last column should be Validation Result
        if validation_cell.value:
            if str(validation_cell.value).startswith('OK'):
                validation_cell.fill = PatternFill(start_color='D4EDDA', end_color='D4EDDA', fill_type='solid')
            elif str(validation_cell.value).startswith('NOT OK'):
                validation_cell.fill = PatternFill(start_color='F8D7DA', end_color='F8D7DA', fill_type='solid')

def format_summary_sheet(worksheet):
    """Apply formatting to the summary sheet"""
    
    # Title formatting
    title_font = Font(bold=True, size=14, color='FFFFFF')
    title_fill = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
    
    # Section header formatting  
    section_font = Font(bold=True, size=12, color='366092')
    
    # Apply formatting to specific cells
    for row_num, row in enumerate(worksheet.iter_rows(min_row=1), 1):
        cell_value = str(row[0].value) if row[0].value else ''
        
        if 'Summary' in cell_value or 'Analysis' in cell_value or 'Breakdown' in cell_value:
            row[0].font = section_font
        
        # Align all values to the left
        if len(row) > 1 and row[1].value is not None:
            row[1].alignment = Alignment(horizontal='left')
    
    # Auto-fit columns
    for column in worksheet.columns:
        max_length = 0
        column_letter = get_column_letter(column[0].column)
        for cell in column:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        worksheet.column_dimensions[column_letter].width = max_length + 2

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
        
        # Generate statistics before adding any summary rows
        total_records = len(df_merged)
        ok_records = df_merged['Validation Result'].str.startswith('OK', na=False).sum()
        not_ok_records = df_merged['Validation Result'].str.startswith('NOT OK', na=False).sum()

        # Append duplicate information for readability in the exported file
        df_merged = append_duplicate_summary_rows(df_merged, test_duplicates, putaway_duplicates)
        
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
        
        save_with_filtered_columns_and_summary(df_merged, output_path, stats)
        
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


def send_email_via_outlook(to_email, cc_emails, subject, body_text, attachment_path):
    """Send an email via Outlook using a temporary PowerShell script."""

    cc_string = ';'.join(cc_emails) if cc_emails else ''

    powershell_script = f'''
Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"
$outlook = New-Object -ComObject Outlook.Application
$mail = $outlook.CreateItem(0)
$mail.To = "{to_email}"
$mail.CC = "{cc_string}"
$mail.Subject = "{subject}"
$mail.Body = @"
{body_text}
"@
$mail.Attachments.Add("{attachment_path}")
$mail.Send()
'''

    script_path = None

    try:
        with tempfile.NamedTemporaryFile(mode='w', suffix='.ps1', delete=False) as script_file:
            script_file.write(powershell_script)
            script_path = script_file.name

        subprocess.run(
            ['powershell.exe', '-ExecutionPolicy', 'Bypass', '-File', script_path],
            check=True,
            capture_output=True
        )

    finally:
        if script_path and os.path.exists(script_path):
            os.unlink(script_path)

@app.route('/send-email', methods=['POST'])
def send_email():
    """Generate email with attachment using Windows shell integration"""
    try:
        data = request.get_json()
        output_file = data.get('output_file')
        stats = data.get('stats')
        
        if not output_file or not stats:
            return jsonify({'status': 'error', 'message': 'Missing required data'}), 400
        
        # Recipients
        to_email = EMAIL_CONFIG['to_email']
        cc_emails = EMAIL_CONFIG['cc_emails']
        
        # Email content
        subject = f"Antenna testing SN validation report - {datetime.now().strftime('%Y-%m-%d')}"
        
        # Create simple text body for mailto
        body_text = f"""Hello,

Please find attached the Serial Number Validation Report.

Summary:
- Total Records Processed: {stats.get('total_records', 'N/A')}
- Valid Records: {stats.get('ok_records', 'N/A')}
- Invalid Records: {stats.get('not_ok_records', 'N/A')}
- Report generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
- File: {output_file}

Please review the validation results and let me know if you have any questions.

Best regards,
Joe"""
        
        # Get full file path
        file_path = os.path.join('output', output_file)
        full_file_path = os.path.abspath(file_path)
        
        try:
            send_email_via_outlook(
                to_email=to_email,
                cc_emails=cc_emails,
                subject=subject,
                body_text=body_text,
                attachment_path=full_file_path
            )

            return jsonify({
                'status': 'success',
                'message': 'Email sent successfully with attachment!',
                'attachment_name': output_file,
                'sent_to': to_email,
                'cc_recipients': cc_emails
            })

        except Exception as ps_error:
            # Fallback to mailto link if PowerShell approach fails
            import urllib.parse
            cc_string = ';'.join(cc_emails) if cc_emails else ''
            mailto_url = f"mailto:{to_email}?cc={cc_string}&subject={urllib.parse.quote(subject)}&body={urllib.parse.quote(body_text)}"
            
            return jsonify({
                'status': 'partial_success',
                'message': 'Opening Outlook with email template. Please manually attach the file.',
                'mailto_url': mailto_url,
                'file_path': full_file_path,
                'attachment_name': output_file,
                'instructions': 'Please manually attach the downloaded Excel file to the email.',
                'error_detail': str(ps_error)
            })
        
    except Exception as e:
        return jsonify({'status': 'error', 'message': f'Error preparing email: {str(e)}'}), 500

if __name__ == '__main__':
    # Ensure required directories exist
    os.makedirs('uploads', exist_ok=True)
    os.makedirs('output', exist_ok=True)
    
    app.run(debug=True, host='0.0.0.0', port=5000)
