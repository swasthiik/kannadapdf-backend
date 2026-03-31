# ಕನ್ನಡ PDF Backend — Setup Guide

## Step 1 — Install on your computer (Testing)

```bash
# Install LibreOffice (Ubuntu/Linux)
sudo apt-get install libreoffice

# Install Python packages
pip install -r requirements.txt

# Run the server
python app.py
```

Your API will run at: http://localhost:5000

---

## Step 2 — Deploy FREE on Render.com

1. Go to https://render.com and create free account
2. Click "New" → "Web Service"
3. Connect your GitHub account
4. Upload these backend files to a GitHub repo
5. Select the repo in Render
6. Render will auto-detect render.yaml and deploy
7. You get a FREE URL like: https://kannadapdf-backend.onrender.com

---

## Step 3 — Connect Frontend to Backend

In your index.html, find this section and replace YOUR_BACKEND_URL:

```javascript
const BACKEND_URL = 'https://kannadapdf-backend.onrender.com';

// Word to PDF
async function wordToPdf(file) {
  const formData = new FormData();
  formData.append('file', file);
  const res = await fetch(`${BACKEND_URL}/word-to-pdf`, {
    method: 'POST',
    body: formData
  });
  return await res.blob();
}

// PDF to Word
async function pdfToWord(file) {
  const formData = new FormData();
  formData.append('file', file);
  const res = await fetch(`${BACKEND_URL}/pdf-to-word`, {
    method: 'POST',
    body: formData
  });
  return await res.blob();
}

// PDF Split
async function pdfSplit(file, pages) {
  const formData = new FormData();
  formData.append('file', file);
  formData.append('pages', pages); // e.g. "1-3" or "1,3,5"
  const res = await fetch(`${BACKEND_URL}/pdf-split`, {
    method: 'POST',
    body: formData
  });
  return await res.blob();
}
```

---

## API Endpoints

| Method | URL | What it does |
|--------|-----|--------------|
| GET | / | Health check |
| POST | /word-to-pdf | Convert Word to PDF |
| POST | /pdf-to-word | Convert PDF to Word |
| POST | /pdf-split | Split PDF pages |

---

## Testing with curl

```bash
# Test Word to PDF
curl -X POST http://localhost:5000/word-to-pdf \
  -F "file=@test.docx" \
  --output converted.pdf

# Test PDF Split (pages 1 to 3)
curl -X POST http://localhost:5000/pdf-split \
  -F "file=@test.pdf" \
  -F "pages=1-3" \
  --output split.pdf
```

---

## Free Hosting Limits on Render.com

- 750 hours/month free (enough for 1 site)
- Sleeps after 15 min of inactivity (wakes up in ~30 sec)
- 512MB RAM
- Sufficient for your PDF site to start!
