# Organization Chart Generator

Upload a Workday Excel file and generate an interactive org chart in your browser.

## Setup

```bash
cd /path/to/org_chart
python -m venv venv
source venv/bin/activate   # On Windows: venv\Scripts\activate
pip install -r requirements.txt
```

## Run locally (development)

```bash
python app.py
```

Or with the Flask CLI:

```bash
export FLASK_APP=app.py
flask run
```

Then open **http://127.0.0.1:5050** in your browser. (Port 5050 avoids conflict with macOS AirPlay. Set `PORT=5000` to use 5000.)

## Deploy so others on the internet can access (Render)

**[Render](https://render.com)** is a simple, professional host with a free tier. Your boss and others can open your app via a link like `https://org-chart-xxxx.onrender.com`.

### One-time setup

1. **Put the project on GitHub**  
   Create a repo, push this folder (e.g. `git init`, `git add .`, `git commit -m "Initial"`, `git remote add origin ...`, `git push -u origin main`).

2. **Sign up at [render.com](https://render.com)** (free account).

3. **Create a Web Service**  
   - Dashboard → **New +** → **Web Service**.  
   - Connect your GitHub account and select the `org_chart` repo.  
   - Use these settings:

   | Field | Value |
   |--------|--------|
   | **Name** | `org-chart` (or any name) |
   | **Region** | Choose nearest to your users |
   | **Branch** | `main` (or your default branch) |
   | **Root Directory** | *(leave blank)* |
   | **Runtime** | Python 3 |
   | **Build Command** | `pip install -r requirements.txt` |
   | **Start Command** | `gunicorn -w 4 -b 0.0.0.0:$PORT app:app` |

   - Click **Create Web Service**.

4. **Wait for the first deploy** (a few minutes). When it’s green, your app is live at `https://<your-service-name>.onrender.com`. Share that URL with your boss and colleagues.

### After the first deploy

- **Automatic deploys:** Every push to the connected branch triggers a new deploy.
- **Custom domain (optional):** In the service → **Settings** → **Custom Domain**, add your own domain and follow Render’s DNS instructions.
- **Free tier note:** The service may “spin down” after ~15 minutes of no traffic; the next visitor may wait 30–60 seconds while it starts. Paid plans keep it always on.

### Using the Blueprint (optional)

If your repo has `render.yaml`, you can use **New +** → **Blueprint** and point Render at the repo; it will create the web service from that file.

---

## Run production locally (Gunicorn)

To run with Gunicorn on your own machine:

```bash
gunicorn -w 4 -b 0.0.0.0:5050 app:app
```

- `-w 4`: 4 worker processes  
- `-b 0.0.0.0:5050`: bind to all interfaces on port 5050  
- `app:app`: module `app`, variable `app`

## Excel file

- Sheet name: **Org Chart**
- Required columns: **Unique Identifier**, **Name**
- Optional: **Reports To**, **Line Detail 1**, **Organization Name**
- Format: `.xlsx`
