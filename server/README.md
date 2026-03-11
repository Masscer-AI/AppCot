# Server setup

## Quick start (Git Bash on Windows)

```bash
cd server
bash setup.sh
python main.py
```

API will run at:

- `http://127.0.0.1:8009`

## Manual setup

```bash
cd server
python -m venv .venv
source .venv/Scripts/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python main.py
```

## Endpoint

- `POST /api/quotes/generate`
  - Body:
    ```json
    {
      "companyName": "Empresa S.A. de C.V.",
      "fullName": "Charly Chacon"
    }
    ```
  - Returns generated `.xlsx` file.
