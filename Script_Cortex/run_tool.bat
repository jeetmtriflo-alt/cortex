@echo off
echo Installing required packages...
pip install -r requirements.txt
echo.
echo Starting Flask Server...
echo Open your browser and navigate to: http://localhost:5000
echo.
python app.py
pause

