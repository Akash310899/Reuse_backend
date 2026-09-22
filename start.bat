@echo off
echo ===========================================
echo  Starting Reuse ^& Modify Tool Servers... 
echo ===========================================

set PYTHONPATH=%CD%

echo -> Starting FastAPI Backend on port 8000...
start /b .\.venv\Scripts\python.exe -m uvicorn backend.main:app --host 0.0.0.0 --port 8000 > backend.log 2>&1

echo -> Starting React Frontend on port 5173...
start /b npm.cmd run dev > frontend.log 2>&1

echo ===========================================
echo Both servers are starting in background!
echo UI is available at: http://localhost:5173
echo API is available at: http://localhost:8000
echo ===========================================
