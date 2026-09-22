#!/bin/bash

echo "==========================================="
echo " Starting Reuse & Modify Tool Servers... "
echo "==========================================="

# Setup absolute path for Python
export PYTHONPATH=$(pwd)

# Activate venv if available
if [ -d ".venv" ]; then
    source .venv/bin/activate
fi

# Start FastAPI Backend in the background
echo "-> Starting FastAPI Backend on port 8000..."
uvicorn backend.main:app --host 0.0.0.0 --port 8000 > backend.log 2>&1 &
BACKEND_PID=$!

# Wait a moment to ensure backend starts
sleep 2

# Start React Frontend in the background
echo "-> Starting React Frontend..."
npm run dev > frontend.log 2>&1 &
FRONTEND_PID=$!

echo "==========================================="
echo "✅ Both servers are running!"
echo "🌐 UI is available at: http://localhost:5173"
echo "⚙️  API is available at: http://localhost:8000"
echo "==========================================="
echo "Press CTRL+C to stop both servers."

# Trap CTRL+C to kill both background processes
trap "echo -e '\nStopping servers...'; kill $BACKEND_PID $FRONTEND_PID; exit" SIGINT SIGTERM

# Wait indefinitely until interrupted
wait

