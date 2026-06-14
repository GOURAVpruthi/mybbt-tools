"""File Manager tool — handles upload, list, zip, delete operations."""

import os
import zipfile
import shutil
from datetime import datetime


class FileManager:
    def __init__(self, upload_folder, output_folder):
        self.base_folder = os.path.join(upload_folder, 'file_manager')
        self.output_folder = output_folder
        os.makedirs(self.base_folder, exist_ok=True)

    def list_files(self):
        """Return list of all uploaded files with metadata."""
        files = []
        if not os.path.exists(self.base_folder):
            return files
        for fname in sorted(os.listdir(self.base_folder)):
            fpath = os.path.join(self.base_folder, fname)
            if os.path.isfile(fpath):
                stat = os.stat(fpath)
                ext = fname.rsplit('.', 1)[-1].lower() if '.' in fname else 'file'
                files.append({
                    'name': fname,
                    'size': stat.st_size,
                    'size_str': self._format_size(stat.st_size),
                    'modified': datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M'),
                    'ext': ext,
                    'type': self._get_file_type(ext)
                })
        return files

    def create_zip(self):
        """Zip all uploaded files and return path to zip."""
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        zip_name = f'collected_files_{ts}.zip'
        zip_path = os.path.join(self.output_folder, zip_name)
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for fname in os.listdir(self.base_folder):
                fpath = os.path.join(self.base_folder, fname)
                if os.path.isfile(fpath):
                    zf.write(fpath, fname)
        return zip_path

    def delete_file(self, filename):
        """Delete a file from the manager."""
        fpath = os.path.join(self.base_folder, filename)
        if os.path.exists(fpath):
            os.remove(fpath)
            return {'success': True, 'message': f'{filename} deleted successfully'}
        return {'success': False, 'error': 'File not found'}

    def get_file_path(self, filename):
        """Get full path for a file."""
        fpath = os.path.join(self.base_folder, filename)
        if os.path.exists(fpath):
            return fpath
        raise FileNotFoundError(f'{filename} not found')

    def _format_size(self, size_bytes):
        if size_bytes < 1024:
            return f'{size_bytes} B'
        elif size_bytes < 1024 * 1024:
            return f'{size_bytes / 1024:.1f} KB'
        elif size_bytes < 1024 * 1024 * 1024:
            return f'{size_bytes / (1024 * 1024):.1f} MB'
        return f'{size_bytes / (1024 * 1024 * 1024):.1f} GB'

    def _get_file_type(self, ext):
        types = {
            'pdf': 'PDF',
            'xlsx': 'Excel', 'xls': 'Excel', 'csv': 'CSV',
            'doc': 'Word', 'docx': 'Word',
            'jpg': 'Image', 'jpeg': 'Image', 'png': 'Image', 'gif': 'Image',
            'zip': 'Archive', 'rar': 'Archive',
            'txt': 'Text',
            'ppt': 'PowerPoint', 'pptx': 'PowerPoint',
        }
        return types.get(ext, 'File')
