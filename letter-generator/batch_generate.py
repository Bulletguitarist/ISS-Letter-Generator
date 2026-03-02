"""
batch_generate.py — CLI alternative to the Streamlit UI

Usage:
  python batch_generate.py \
    --csv sample_data/sample_candidates.csv \
    --template templates/Offer_Letter_Template.docx \
    --type "Offer Letter" \
    --groq-key gsk_xxx \
    --output output/

Letter types: "Offer Letter" | "Internship Certificate" | "Letter of Recommendation (LOR)"
"""

import argparse
import os
import sys
import pandas as pd
from pathlib import Path
from datetime import datetime


def main():
    parser = argparse.ArgumentParser(description="ISS Bulk Letter Generator (CLI)")
    parser.add_argument("--csv",      required=True, help="Path to CSV file")
    parser.add_argument("--template", required=True, help="Path to DOCX template")
    parser.add_argument("--type",     required=True,
                        choices=["Offer Letter", "Internship Certificate", "Letter of Recommendation (LOR)"],
                        help="Letter type")
    parser.add_argument("--output",   default="output", help="Output directory")
    parser.add_argument("--groq-key", default=os.environ.get("GROQ_API_KEY", ""),
                        help="Groq API key (or set GROQ_API_KEY env var)")
    parser.add_argument("--model",    default="llama3-70b-8192", help="Groq model name (e.g. llama3-70b-8192, llama3-8b-8192, llama-3.1-8b-instant)")
    args = parser.parse_args()

    # ── Validate ──────────────────────────────────────────────────────────────
    if not Path(args.csv).exists():
        print(f"❌ CSV not found: {args.csv}")
        sys.exit(1)
    if not Path(args.template).exists():
        print(f"❌ Template not found: {args.template}")
        sys.exit(1)

    try:
        from docxtpl import DocxTemplate
    except ImportError:
        print("❌ docxtpl not installed. Run: pip install docxtpl")
        sys.exit(1)

    # ── Load data ─────────────────────────────────────────────────────────────
    df = pd.read_csv(args.csv)
    template_bytes = Path(args.template).read_bytes()
    out_dir = Path(args.output)
    out_dir.mkdir(parents=True, exist_ok=True)

    # ── Groq client ───────────────────────────────────────────────────────────
    groq_client = None
    if args.groq_key:
        try:
            from groq import Groq
            groq_client = Groq(api_key=args.groq_key)
            print("✅ Groq API connected.")
        except Exception as e:
            print(f"⚠️  Groq connection failed: {e}. Continuing without AI generation.")

    # ── Import app functions ──────────────────────────────────────────────────
    sys.path.insert(0, str(Path(__file__).parent))
    from app import build_context, replace_placeholders_in_docx, get_output_filename

    # ── Process ───────────────────────────────────────────────────────────────
    total    = len(df)
    success  = 0
    failures = []
    start    = datetime.now()

    print(f"\n📄 Generating {total} × '{args.type}' documents…\n")

    for i, (_, row) in enumerate(df.iterrows()):
        name = str(row.get("name", f"Candidate_{i+1}")).strip()
        try:
            ctx       = build_context(row.to_dict(), args.type, groq_client, args.model)
            doc_bytes = replace_placeholders_in_docx(template_bytes, ctx)
            filename  = get_output_filename(name, args.type)
            (out_dir / filename).write_bytes(doc_bytes)
            success += 1
            print(f"  [{i+1:4d}/{total}] ✅ {filename}")
        except Exception as e:
            failures.append((name, str(e)))
            print(f"  [{i+1:4d}/{total}] ❌ {name} — {e}")

    elapsed = (datetime.now() - start).total_seconds()
    print(f"\n{'─'*60}")
    print(f"✅ {success}/{total} generated in {elapsed:.1f}s → {out_dir.resolve()}")
    if failures:
        print(f"❌ {len(failures)} failures:")
        for name, err in failures:
            print(f"   • {name}: {err}")


if __name__ == "__main__":
    main()