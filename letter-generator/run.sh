#!/bin/bash
# run.sh — Launch the ISS Letter Generator
# Usage: ./run.sh

echo "📄 ISS Letter Generator"
echo "========================"

# Load .env if it exists
if [ -f .env ]; then
    export $(cat .env | grep -v '#' | xargs)
    echo "✅ Loaded .env"
fi

# Generate templates if not already done
if [ ! -f templates/Offer_Letter_Template.docx ]; then
    echo "⚙️  Generating templates..."
    python generate_templates.py
fi

echo "🚀 Starting Streamlit app..."
streamlit run app.py --server.maxUploadSize=200
