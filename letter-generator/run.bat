@echo off
echo 📄 ISS Letter Generator
echo ========================

REM Load .env if exists
if exist .env (
    for /f "tokens=1,2 delims==" %%a in (.env) do (
        if not "%%a" == "" if not "%%a:~0,1%" == "#" set %%a=%%b
    )
    echo ✅ Loaded .env
)

REM Generate templates if not done
if not exist templates\Offer_Letter_Template.docx (
    echo ⚙️  Generating templates...
    python generate_templates.py
)

echo 🚀 Starting Streamlit app...
streamlit run app.py --server.maxUploadSize=200
pause
