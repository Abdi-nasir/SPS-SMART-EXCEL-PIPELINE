@echo off
REM setup.bat - Auto-setup script for Windows

echo 🚀 Setting up SPS Excel Pipeline...

echo 📁 Creating required folders...
mkdir data\raw 2>nul
mkdir data\processed 2>nul
mkdir reports 2>nul

echo 📦 Installing dependencies...
py -m pip install -r requirements.txt

echo.
echo ✅ Setup complete!
echo.
echo To run the application:
echo   streamlit run src/app.py
echo.
echo Or with Docker:
echo   docker-compose up -d