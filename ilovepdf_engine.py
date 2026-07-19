import os
import requests
import json
import uuid
import time
from werkzeug.utils import secure_filename

ILOVEPDF_PUBLIC_KEY = os.environ.get('ILOVEPDF_PUBLIC_KEY', 'secret_key_f0e4d569bebb335f436dc622b763ca14_Qytuv0bc37628ae14d7f1f9a17dc209bb47af')

class ILovePDFError(Exception):
    pass

def get_auth_token():
    if not ILOVEPDF_PUBLIC_KEY:
        raise ILovePDFError("No API Key configured.")
    
    resp = requests.post('https://api.ilovepdf.com/v1/auth', data={'public_key': ILOVEPDF_PUBLIC_KEY})
    if resp.status_code == 200:
        return resp.json().get('token')
    else:
        raise ILovePDFError(f"Auth Failed: {resp.text}")

def start_task(tool, token):
    resp = requests.get(f'https://api.ilovepdf.com/v1/start/{tool}', headers={'Authorization': f'Bearer {token}'})
    if resp.status_code == 200:
        data = resp.json()
        return data['server'], data['task']
    raise ILovePDFError(f"Start Task Failed: {resp.text}")

def upload_file(server, task_id, token, file_path):
    with open(file_path, 'rb') as f:
        resp = requests.post(
            f'https://{server}/v1/upload',
            headers={'Authorization': f'Bearer {token}'},
            data={'task': task_id},
            files={'file': f}
        )
    if resp.status_code == 200:
        return resp.json().get('server_filename')
    raise ILovePDFError(f"Upload Failed: {resp.text}")

def process_task(server, task_id, token, tool, files_dict, **kwargs):
    payload = {
        'task': task_id,
        'tool': tool,
        'files': files_dict
    }
    payload.update(kwargs)
    
    resp = requests.post(
        f'https://{server}/v1/process',
        headers={'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'},
        data=json.dumps(payload)
    )
    if resp.status_code != 200:
        raise ILovePDFError(f"Process Failed: {resp.text}")
    return True

def download_file(server, task_id, token, output_dir):
    resp = requests.get(
        f'https://{server}/v1/download/{task_id}',
        headers={'Authorization': f'Bearer {token}'}
    )
    if resp.status_code == 200:
        # Check content disposition for filename or fallback
        cd = resp.headers.get('content-disposition', '')
        filename = f"result_{task_id}.pdf"
        if 'filename=' in cd:
            filename = cd.split('filename=')[1].strip('"')
            
        out_path = os.path.join(output_dir, secure_filename(filename))
        with open(out_path, 'wb') as f:
            f.write(resp.content)
        return out_path
    raise ILovePDFError(f"Download Failed: {resp.text}")


def process_hybrid(tool, file_paths, output_dir, **kwargs):
    """
    Attempts to process the files using iLovePDF API.
    If any error occurs (auth, limits, network), it raises ILovePDFError
    so the calling function can gracefully fallback to local Python tools.
    """
    token = get_auth_token()
    server, task_id = start_task(tool, token)
    
    files_dict = []
    for fp in file_paths:
        srv_fname = upload_file(server, task_id, token, fp)
        files_dict.append({
            'server_filename': srv_fname,
            'filename': os.path.basename(fp)
        })
        
    process_task(server, task_id, token, tool, files_dict, **kwargs)
    result_path = download_file(server, task_id, token, output_dir)
    return result_path
