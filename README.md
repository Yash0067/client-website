# Adarsh Boutique

The project is split into two independently deployable applications:

- `frontend/` is the responsive static website. Deploy it to Vercel.
- `backend/` is the Python appointment API. Deploy it to Render or Railway.

## Deploy the backend first

Deploy the `backend` folder with:

```text
Build command: pip install -r requirements.txt
Start command: python server.py
```

Configure persistent storage:

- `APP_DATA_DIR`: directory for SQLite data
- `APPOINTMENTS_XLSX`: full path to the Excel workbook

The backend exposes `POST /api/appointments` and `GET /healthz`.

## Deploy the frontend second

Edit `frontend/config.js` and set `appointmentEndpoint` to the backend URL ending in `/api/appointments`. Then deploy the `frontend` folder as a static site on Vercel.

The backend stores each booking in SQLite and appends it to the configured Excel workbook. Use persistent storage in production; temporary server files can be erased during redeploys.
