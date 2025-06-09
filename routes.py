import os
import uuid
import logging
from flask import render_template, request, redirect, url_for, session, flash, send_file, jsonify
from werkzeug.utils import secure_filename
from app import app
from excel_validator import KadamValidator
from validators.kadam_plus_validator import KadamPlusValidator

# Hardcoded credentials for MVP
VALID_CREDENTIALS = {
    'test@example.com': '123456',
    'admin@humana.org': 'password123',
    'user@humana.org': 'excel2024'
}

def allowed_file(filename):
    """Check if the uploaded file has an allowed extension"""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in {'xlsx', 'xls'}

@app.route('/')
def index():
    """Redirect to login page"""
    return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    """Handle user login"""
    if request.method == 'POST':
        email = request.form.get('email', '').strip()
        password = request.form.get('password', '')
        
        # Validate credentials
        if email in VALID_CREDENTIALS and VALID_CREDENTIALS[email] == password:
            session['logged_in'] = True
            session['user_email'] = email
            flash('Login successful!', 'success')
            return redirect(url_for('upload'))
        else:
            flash('Invalid email or password.', 'error')
    
    return render_template('login.html')

@app.route('/logout')
def logout():
    """Handle user logout"""
    session.clear()
    flash('You have been logged out.', 'info')
    return redirect(url_for('login'))

@app.route('/upload', methods=['GET', 'POST'])
def upload():
    """Handle file upload and processing"""
    # Check if user is logged in
    if not session.get('logged_in'):
        flash('Please log in to access this page.', 'error')
        return redirect(url_for('login'))
    
    if request.method == 'POST':
        # Check if file was uploaded
        if 'file' not in request.files:
            flash('No file selected.', 'error')
            return redirect(request.url)
        
        file = request.files['file']
        
        if file.filename == '':
            flash('No file selected.', 'error')
            return redirect(request.url)
        
        if file and allowed_file(file.filename):
            try:
                # Generate unique filename to avoid conflicts
                unique_id = str(uuid.uuid4())
                filename = secure_filename(file.filename)
                base_name = os.path.splitext(filename)[0]
                extension = os.path.splitext(filename)[1]
                
                # Save uploaded file
                upload_filename = f"{unique_id}_{filename}"
                upload_path = os.path.join(app.config['UPLOAD_FOLDER'], upload_filename)
                file.save(upload_path)
                
                # Get validation method from form
                validation_method = request.form.get('validation_method', 'kadam')
                
                # Process the file based on selected method
                if validation_method == 'kadam_plus':
                    validator = KadamPlusValidator()
                else:
                    validator = KadamValidator()
                
                output_files = validator.validate_excel(upload_path, unique_id, app.config['DOWNLOAD_FOLDER'])
                
                if output_files:
                    # Store file info in session for downloads
                    session['processed_files'] = {
                        'validated_output': output_files['validated_output'],
                        'validation_report': output_files['validation_report'],
                        'original_name': base_name
                    }
                    flash('File processed successfully! You can now download the results.', 'success')
                else:
                    flash('Error processing the file. Please check the file format and try again.', 'error')
                
                # Clean up uploaded file
                if os.path.exists(upload_path):
                    os.remove(upload_path)
                    
            except Exception as e:
                logging.error(f"Error processing file: {str(e)}")
                flash(f'Error processing file: {str(e)}', 'error')
        else:
            flash('Invalid file type. Please upload an Excel file (.xlsx or .xls).', 'error')
    
    return render_template('upload.html')

@app.route('/download/<file_type>')
def download_file(file_type):
    """Handle file downloads"""
    # Check if user is logged in
    if not session.get('logged_in'):
        flash('Please log in to access this page.', 'error')
        return redirect(url_for('login'))
    
    processed_files = session.get('processed_files')
    if not processed_files:
        flash('No processed files available for download.', 'error')
        return redirect(url_for('upload'))
    
    try:
        if file_type == 'output':
            file_path = processed_files['validated_output']
            download_name = f"{processed_files['original_name']}_Validated_Output.xlsx"
        elif file_type == 'report':
            file_path = processed_files['validation_report']
            download_name = f"{processed_files['original_name']}_Validation_Report.xlsx"
        else:
            flash('Invalid download type.', 'error')
            return redirect(url_for('upload'))
        
        if os.path.exists(file_path):
            return send_file(
                file_path,
                as_attachment=True,
                download_name=download_name,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        else:
            flash('File not found. Please process a file first.', 'error')
            return redirect(url_for('upload'))
            
    except Exception as e:
        logging.error(f"Error downloading file: {str(e)}")
        flash(f'Error downloading file: {str(e)}', 'error')
        return redirect(url_for('upload'))

@app.route('/clear_files')
def clear_files():
    """Clear processed files from session and disk"""
    if 'processed_files' in session:
        processed_files = session['processed_files']
        
        # Clean up files from disk
        for file_path in [processed_files.get('validated_output'), processed_files.get('validation_report')]:
            if file_path and os.path.exists(file_path):
                try:
                    os.remove(file_path)
                except Exception as e:
                    logging.error(f"Error removing file {file_path}: {str(e)}")
        
        # Clear from session
        del session['processed_files']
        flash('Files cleared successfully.', 'info')
    
    return redirect(url_for('upload'))

@app.errorhandler(413)
def too_large(e):
    """Handle file too large error"""
    flash('File is too large. Maximum size is 16MB.', 'error')
    return redirect(url_for('upload'))

@app.errorhandler(404)
def not_found(e):
    """Handle 404 errors"""
    return redirect(url_for('login'))
