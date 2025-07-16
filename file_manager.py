
import os
import glob
import logging
from pathlib import Path

class FileManager:
    def __init__(self, upload_folder='uploads', download_folder='downloads'):
        self.upload_folder = upload_folder
        self.download_folder = download_folder
        self.max_uploads = 16
        self.max_downloads_per_type = 16  # 16 validation outputs + 16 validation reports = 32 total
        
    def cleanup_old_files(self, folder_path, pattern, max_files):
        """Remove oldest files when exceeding max_files limit"""
        try:
            # Get all files matching pattern
            files = glob.glob(os.path.join(folder_path, pattern))
            
            if len(files) <= max_files:
                return
            
            # Sort by modification time (oldest first)
            files.sort(key=lambda x: os.path.getmtime(x))
            
            # Remove oldest files to stay within limit
            files_to_remove = files[:len(files) - max_files]
            
            for file_path in files_to_remove:
                try:
                    os.remove(file_path)
                    logging.info(f"Removed old file: {file_path}")
                except Exception as e:
                    logging.error(f"Error removing file {file_path}: {str(e)}")
                    
        except Exception as e:
            logging.error(f"Error in cleanup_old_files: {str(e)}")
    
    def manage_uploads(self):
        """Keep only the latest uploads within limit"""
        self.cleanup_old_files(
            self.upload_folder, 
            "*_*.xlsx",  # Pattern matches UUID_filename.xlsx
            self.max_uploads
        )
        
        # Also cleanup .xls files
        self.cleanup_old_files(
            self.upload_folder, 
            "*_*.xls", 
            self.max_uploads
        )
    
    def manage_downloads(self):
        """Keep only the latest download files within limit"""
        # Cleanup validation output files
        self.cleanup_old_files(
            self.download_folder,
            "*_Validated_Output.xlsx",
            self.max_downloads_per_type
        )
        
        # Cleanup validation report files
        self.cleanup_old_files(
            self.download_folder,
            "*_Validation_Report.xlsx", 
            self.max_downloads_per_type
        )
    
    def cleanup_all(self):
        """Run cleanup for both uploads and downloads"""
        self.manage_uploads()
        self.manage_downloads()
