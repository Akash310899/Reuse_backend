import sys
import os

# Add the root directory to the python path so Vercel can find the 'backend' folder
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from backend.main import app
