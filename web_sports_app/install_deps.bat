@echo off
echo Installing dependencies in virtual environment...
pip install flask==2.3.3
pip install python-docx==1.2.0
pip install werkzeug==2.3.7
pip install boto3==1.40.40
echo.
echo Dependencies installed successfully!
echo You can now run: python app.py
pause