# ResumeFit Pro

Professional ATS resume optimizer built with FastAPI. Users upload a DOCX resume, paste a job description, review ATS recommendations, preview the optimized resume, and download a formatted Word document.

The app includes a production-ready responsive interface, installable web app metadata, an app icon, service worker, health check endpoint, and Render deployment configuration.

## Local Run

```bat
pip install -r requirements.txt
copy .env.example .env
.\run_app.bat
```

Open:

```text
http://127.0.0.1:8006/
```

## Production Run

Set environment variables:

```text
GEMINI_API_KEY=your_key
GEMINI_MODEL=gemini-3-flash-preview
RAZORPAY_KEY_ID=your_razorpay_key_id
RAZORPAY_KEY_SECRET=your_razorpay_key_secret
DOWNLOAD_TOKEN_SECRET=use_a_long_random_secret
OWNER_EMAILS=sumitmishra7886@gmail.com
RESUME_PRICE_RUPEES=49
FREE_TRIAL_RESUMES=3
```

Downloads include 3 free resume unlocks per user email by default, then are payment-gated at Rs 49 per optimized resume. Emails listed in `OWNER_EMAILS` are exempt and receive a free download unlock after optimization.

Start command:

```bash
uvicorn app:app --host 0.0.0.0 --port $PORT
```

## Deploy On Render

1. Push this folder to a GitHub repository.
2. Open Render and choose **New Web Service**.
3. Connect the GitHub repository.
4. Render will detect `render.yaml`.
5. Add this environment variable when prompted:

```text
GEMINI_API_KEY=your_key
```

6. Deploy.

Health check:

```text
/health
```

## App Files

- `index.html` - user-facing web application
- `manifest.json` - installable app metadata
- `service-worker.js` - lightweight app shell caching
- `app-icon.svg` - app icon

## Docker

```bash
docker build -t resumefit-pro .
docker run -p 8000:8000 --env-file .env resumefit-pro
```

Then open:

```text
http://localhost:8000/
```
