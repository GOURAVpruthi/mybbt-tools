# Corporate Tools Suite

A complete professional productivity web application for corporate and GST workers.

## 🚀 Features

| Tool | Description |
|------|-------------|
| 📁 File Manager | Upload, organize, and batch download files as ZIP |
| 🗜️ PDF Compressor | Reduce PDF size (Low / Medium / High quality) |
| 🔗 PDF Merger | Combine multiple PDFs into one |
| 📊 Excel Consolidator | Merge multiple Excel/CSV files |
| 📋 GSTR-1 | Process outward supply data |
| 🔄 GSTR-2B Reco | Reconcile Purchase Register vs GSTR-2B |
| 📝 GSTR-3B | Monthly summary return processor |
| 📑 GSTR-9/9C | Annual return computation |
| 📊 PR vs 2B | Detailed 4-sheet reconciliation report |
| 🌐 GST Portal | Quick links to GST portal modules |

---

## 🛠️ Installation & Setup

### Step 1 — Install Python
Make sure Python 3.8+ is installed: https://python.org/downloads

### Step 2 — Install Dependencies
```bash
pip install -r requirements.txt
```

### Step 3 — Run Locally
```bash
python app.py
```
Open your browser at: **http://localhost:5000**

---

## 🌐 Deploy to Your Domain

### Option A — PythonAnywhere (Recommended for beginners)
1. Create free account at [pythonanywhere.com](https://www.pythonanywhere.com)
2. Upload the entire project folder
3. Create a new Web App → Flask → Python 3.10
4. Set source code path to your project folder
5. Set WSGI file to point to `app.py`
6. Add your custom domain in the "Web" tab

### Option B — VPS (Ubuntu/Debian)
```bash
# Install dependencies
sudo apt update && sudo apt install python3-pip nginx

# Upload your project
# Install requirements
pip3 install -r requirements.txt

# Run with Gunicorn
gunicorn --bind 0.0.0.0:5000 --workers 4 app:app

# Configure Nginx as reverse proxy (see nginx.conf below)
```

### Option C — Railway.app
1. Push code to GitHub
2. Create new project on [railway.app](https://railway.app)
3. Connect GitHub repo
4. Add environment variable: `PORT=5000`
5. Deploy and add custom domain

---

## 🔧 Integrating Your Python Code

Replace the placeholder implementations in the `tools/` folder with your own code:

| File | What to Replace |
|------|----------------|
| `tools/pdf_tools.py` | `compress()` and `merge()` methods |
| `tools/excel_tools.py` | `consolidate()` method |
| `tools/gst_tools.py` | All GST processing methods |
| `tools/file_manager.py` | Keep as-is or customize |

Each method should return a dict with `{'success': True/False, ...}` keys.

---

## 📁 Project Structure

```
corporate-tools-suite/
├── app.py                    # Flask backend
├── requirements.txt          # Dependencies
├── tools/
│   ├── file_manager.py       # File Manager logic
│   ├── pdf_tools.py          # PDF compress/merge
│   ├── excel_tools.py        # Excel consolidation
│   └── gst_tools.py          # All GST tools
├── templates/
│   ├── base.html             # Shared layout
│   ├── index.html            # Dashboard
│   ├── file_manager.html     # File Manager page
│   ├── pdf_tools.html        # PDF Tools page
│   ├── excel_tools.html      # Excel Tools page
│   └── gst_tools.html        # GST Tools page
├── static/
│   ├── css/style.css         # Design system
│   └── js/main.js            # Frontend logic
├── uploads/                  # Uploaded files (auto-created)
└── outputs/                  # Generated files (auto-created)
```

---

## 🔒 Security Notes

- Files are stored temporarily in `uploads/` and auto-cleaned after 24 hours
- No data is sent to external servers — all processing is local
- For production, set `debug=False` in `app.py`
- Use HTTPS for your domain (Let's Encrypt is free)

---

## 🧾 Nginx Config (for VPS deployment)

```nginx
server {
    listen 80;
    server_name yourdomain.com www.yourdomain.com;

    location / {
        proxy_pass http://127.0.0.1:5000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
    }

    client_max_body_size 100M;
}
```

---

## 📞 Support

Built with ❤️ for Indian corporate and GST professionals.
