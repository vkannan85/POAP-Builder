@echo off
cd /d %~dp0
if not exist .venv\Scripts\python.exe (
  echo First run detected. Installing local environment...
  py -3 -m venv .venv
  call .venv\Scripts\activate
  python -m pip install --upgrade pip
  pip install -r requirements.txt
) else (
  call .venv\Scripts\activate
)
start "" http://127.0.0.1:8000
python -m uvicorn app:app --host 127.0.0.1 --port 8000
pause