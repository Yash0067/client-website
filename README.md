# Adarsh Boutique — automatic appointments with Excel download

This version does **not** use Google Sheets. The website sends each appointment request to the included Python server. The server saves it in a SQLite database and appends each new request directly to an Excel workbook. Set `APPOINTMENTS_XLSX` to use a specific workbook path.

## Run on your computer

1. Install Python 3.9 or newer. Unzip this folder and open a terminal inside it.
2. Install the required package:

   ```bash
   python -m pip install -r requirements.txt
   ```

3. Start the server. **Windows PowerShell:**

   ```powershell
   python server.py
   ```

   **macOS or Linux:**

   ```bash
   export ADMIN_PASSWORD='choose-a-long-unique-password'
   python3 server.py
   ```

4. Open **http://localhost:8000/**. Submit a test appointment, then open the configured workbook to confirm it was added.

When you stop the server, the website cannot accept requests until it is running again.

## Deploy online

- Host the whole folder on a Python-capable web service. Static hosting alone is not enough.
- Install `requirements.txt`; start with `python server.py`. The server listens on `0.0.0.0` and uses the `PORT` environment variable supplied by the host, or port 8000 if absent.
- Set `APPOINTMENTS_XLSX` if the workbook should be stored at a different path.
- Mount persistent storage and set `APP_DATA_DIR` to its path. The SQLite database is stored there. Without a persistent disk, redeploys/restarts on some hosts erase all bookings. Back it up securely and test restoration. Never expose or publish the database or admin credentials.
- Use HTTPS for a public deployment. A production deployment needs appropriate access controls, abuse/rate limiting, privacy and retention policies, monitoring, backups and a hosting configuration suitable for SQLite and the workbook.
- Open the public site and submit a test, then confirm the request was appended to the workbook. Delete test data only through a deliberate process; don't delete real bookings accidentally.
- Requests are preferences only. You must contact the customer to confirm. No availability checking or double-booking prevention is provided.
