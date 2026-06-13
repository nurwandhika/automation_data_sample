# Automate Data Sample

Simple tool to reconcile two CSVs with guided column matching.

## Deploying the backend

Quick options:

- Render: connect your GitHub repo and create a new Web Service. Render will use `requirements.txt` and `Procfile`.
- Any container host: build the included `Dockerfile` and push to your registry; the GH Action `publish-ghcr.yml` will publish to GitHub Container Registry on push to `main`.

After hosting the backend, set `window.APP_API_BASE` on the published frontend to point to the backend URL so the UI can call the API.
