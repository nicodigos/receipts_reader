# Receipts Database

Streamlit app for:

- browsing the receipts CSV stored in SharePoint,
- filtering and exporting the filtered table to styled Excel,
- loading and downloading merged PDFs for the currently filtered invoices,
- uploading new invoices and processing them with Google Cloud Vision plus OpenAI.

## Pages

- `Receipts_Database.py`: main page, focused on the cloud receipts database
- `pages/2_Upload_New_Invoices.py`: upload and process new invoices
- `invoice_app.py`: shared app logic

## Setup

1. Create and activate a virtual environment.
2. Install dependencies:

```powershell
pip install -r requirements.txt
```

3. Create the environment file:

```powershell
copy .env.example .env
```

4. Fill `.env` with valid values for:

- `OPENAI_API_KEY`
- `TENANT_ID`
- `CLIENT_ID`
- `CLIENT_SECRET`
- `REDIRECT_URI`
- `SP_HOSTNAME`
- `SP_SITE_PATH`
- `SP_DRIVE_NAME`
- `RECEIPTS_DATABASE_DIR`

5. Create `.streamlit/secrets.toml` from `.streamlit/secrets.toml.example` and add the Google service account there.

## Run

```powershell
streamlit run Receipts_Database.py
```

## Notes

- The database page fetches the CSV from SharePoint on each rerun.
- Uploaded invoices and the database CSV are stored in SharePoint, not in a local `data/` folder.
