import os
import re
from datetime import datetime
from flask import Flask, request, render_template, session, send_file, jsonify
from werkzeug.utils import secure_filename
import PyPDF2
import pandas as pd
from openpyxl import load_workbook
import tempfile
import shutil
from config import *

app = Flask(__name__)
app.secret_key = SECRET_KEY
app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH

# Ensure directories exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def extract_first_two_pages(pdf_path):
    """Extract first 2 pages from PDF using PyPDF2"""
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            
            if len(pdf_reader.pages) == 0:
                return None, "PDF has no pages"
            
            # Create a new PDF writer
            pdf_writer = PyPDF2.PdfWriter()
            
            # Add first page
            pdf_writer.add_page(pdf_reader.pages[0])
            
            # Add second page if it exists
            if len(pdf_reader.pages) > 1:
                pdf_writer.add_page(pdf_reader.pages[1])
            
            # Save the extracted pages
            first_two_pages_path = os.path.join(PROCESSED_FOLDER, f"first_two_pages_{os.path.basename(pdf_path)}")
            with open(first_two_pages_path, 'wb') as output_file:
                pdf_writer.write(output_file)
            
            return first_two_pages_path, None
            
    except Exception as e:
        return None, f"Error extracting pages: {str(e)}"

def convert_to_markdown(pdf_path):
    """Convert PDF to markdown using PyPDF2 text extraction"""
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            
            markdown_content = []
            markdown_content.append(f"# PDF Document: {os.path.basename(pdf_path)}\n")
            
            for page_num, page in enumerate(pdf_reader.pages, 1):
                text = page.extract_text()
                if text.strip():
                    markdown_content.append(f"## Page {page_num}\n")
                    markdown_content.append(text)
                    markdown_content.append("\n")
            
            return '\n'.join(markdown_content), None
            
    except Exception as e:
        return None, f"Error converting to markdown: {str(e)}"

def create_combined_excel(results):
    """Create one Excel file with all PDF results as rows"""
    try:
        # Create DataFrame for Excel export
        excel_data = []
        
        for result in results:
            if result['status'] == 'success':
                # Extract basic info from filename
                filename = result['filename']
                file_size = os.path.getsize(os.path.join(UPLOAD_FOLDER, filename)) if os.path.exists(os.path.join(UPLOAD_FOLDER, filename)) else 0
                
                # Parse markdown content for structured data
                parsed_data = parse_markdown_content(result['markdown'])
                
                # Combine all data
                row_data = {
                    'Filename': filename,
                    'File Size (bytes)': file_size,
                    'Processing Date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'Project ID': parsed_data.get('project_id', 'N/A'),
                    'Project Name': parsed_data.get('project_name', 'N/A'),
                    'Date Certified': parsed_data.get('date_cert', 'N/A'),
                    'Total Points': parsed_data.get('total_points', 'N/A'),
                    'Markdown Content': result['markdown'][:1000] + '...' if len(result['markdown']) > 1000 else result['markdown'],
                    'Status': 'Success'
                }
                excel_data.append(row_data)
            else:
                # Handle failed conversions
                row_data = {
                    'Filename': result['filename'],
                    'File Size (bytes)': 0,
                    'Processing Date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'Project ID': 'N/A',
                    'Project Name': 'N/A',
                    'Date Certified': 'N/A',
                    'Total Points': 'N/A',
                    'Markdown Content': f"Error: {result['message']}",
                    'Status': 'Failed'
                }
                excel_data.append(row_data)
        
        # Create DataFrame
        df = pd.DataFrame(excel_data)
        
        # Generate Excel file
        excel_path = os.path.join(PROCESSED_FOLDER, f"combined_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        
        # Create Excel writer with formatting
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='PDF Results', index=False)
            
            # Get the workbook and worksheet
            workbook = writer.book
            worksheet = writer.sheets['PDF Results']
            
            # Auto-adjust column widths
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)  # Cap at 50 characters
                worksheet.column_dimensions[column_letter].width = adjusted_width
        
        return excel_path, None
        
    except Exception as e:
        return None, f"Error creating combined Excel: {str(e)}"

def parse_markdown_content(markdown_content):
    """Parse markdown content to extract structured data"""
    try:
        # Normalize text
        def normalize_text(s: str) -> str:
            s2 = s.replace("β", "")
            # Fix codes like "A01. 1" -> "A01.1"
            s2 = re.sub(r"(\b[AWNLVTSXMCI]\d{2})\.\s+(\d)\b", r"\1.\2", s2)
            # Fix spaced decimals like "0. 5", "1 . 5", "14. 5"
            s2 = re.sub(r"(\d)\s*\.\s*(\d)", r"\1.\2", s2)
            # Collapse whitespace
            s2 = re.sub(r"\s+", " ", s2)
            return s2
        
        text = normalize_text(markdown_content)
        
        # Extract header values
        project_id_match = re.search(r"(\d{10})\s*-\s*", text)
        project_id = project_id_match.group(1) if project_id_match else "Unknown"
        
        project_name_match = re.search(r"\d{10}\s*-\s*(.+?)\s*\(WELL", text)
        project_name = project_name_match.group(1).strip() if project_name_match else "Unknown Project"
        
        date_match = re.search(r"Date:\s*(\d{1,2}\s+[A-Za-z]{3},\s*\d{4})", text)
        date_cert = "Unknown"
        if date_match:
            try:
                date_cert = datetime.strptime(date_match.group(1).replace(",", ""), "%d %b %Y").strftime("%d/%m/%Y")
            except:
                date_cert = "Invalid Date"
        
        # Extract total points if available
        tot_m = re.search(r"Reviewer\s*-\s*Confirmed Total Points\s*([0-9]+(?:\.\d+)?)", text, re.IGNORECASE)
        total_points = tot_m.group(1) if tot_m else "N/A"
        
        return {
            'project_id': project_id,
            'project_name': project_name,
            'date_cert': date_cert,
            'total_points': total_points
        }
        
    except Exception as e:
        return {
            'project_id': 'Error',
            'project_name': 'Error',
            'date_cert': 'Error',
            'total_points': 'Error'
        }

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    if 'files' not in request.files:
        return jsonify({'error': 'No files provided'}), 400
    
    files = request.files.getlist('files')
    if not files or all(file.filename == '' for file in files):
        return jsonify({'error': 'No files selected'}), 400
    
    results = []
    
    for file in files:
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file_path = os.path.join(UPLOAD_FOLDER, filename)
            file.save(file_path)
            
            # Extract first 2 pages
            first_two_pages_path, extract_error = extract_first_two_pages(file_path)
            if extract_error:
                results.append({
                    'filename': filename,
                    'status': 'error',
                    'message': extract_error
                })
                continue
            
            # Convert to markdown
            markdown_content, convert_error = convert_to_markdown(first_two_pages_path)
            if convert_error:
                results.append({
                    'filename': filename,
                    'status': 'error',
                    'message': convert_error
                })
                continue
            
            results.append({
                'filename': filename,
                'status': 'success',
                'markdown': markdown_content,
                'message': 'Successfully processed'
            })
            
            # Clean up temporary files
            if os.path.exists(first_two_pages_path):
                os.remove(first_two_pages_path)
            if os.path.exists(file_path):
                os.remove(file_path)
    
    # Create combined Excel file
    excel_path, excel_error = create_combined_excel(results)
    if excel_error:
        return jsonify({'error': excel_error}), 500
    
    # Store results and Excel path in session
    session['results'] = results
    session['excel_path'] = excel_path
    
    return jsonify({
        'results': results,
        'excel_path': excel_path,
        'message': f'Successfully processed {len(results)} files. Combined Excel file created.'
    })

@app.route('/download-excel')
def download_excel():
    """Download the combined Excel file"""
    try:
        excel_path = session.get('excel_path')
        if not excel_path or not os.path.exists(excel_path):
            return jsonify({'error': 'Excel file not found. Please process files first.'}), 404
        
        # Get just the filename for download
        filename = os.path.basename(excel_path)
        
        return send_file(
            excel_path, 
            as_attachment=True, 
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        return jsonify({'error': f'Error downloading file: {str(e)}'}), 500

@app.route('/clear-session', methods=['POST'])
def clear_session():
    """Clear session and remove temporary files"""
    try:
        # Remove Excel file if it exists
        excel_path = session.get('excel_path')
        if excel_path and os.path.exists(excel_path):
            os.remove(excel_path)
        
        session.clear()
        return jsonify({'message': 'Session cleared and files removed'})
        
    except Exception as e:
        return jsonify({'error': f'Error clearing session: {str(e)}'}), 500

if __name__ == '__main__':
    app.run(debug=DEBUG, host=HOST, port=PORT)
