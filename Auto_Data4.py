import os
import sys
import glob
import html
import argparse
from datetime import datetime
from dateutil import parser as dateparser
import pandas as pd
import validators

from dotenv import load_dotenv
load_dotenv()

# --- OpenAI ---
from openai import OpenAI
client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))
OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gpt-4o-mini")

# --- Colonnes attendues (adaptées à ton fichier) ---
REQUIRED_COLS = {
    "publication": ["media outlet", "publication", "media"],
    "published": ["published", "date", "publication_date"],
    "URL": ["url", "lien", "link"],
}
CONTENT_CANDIDATES = ["snippet", "content", "texte", "text", "body", "résumé", "summary"]
TITLE_CANDIDATES = ["article", "titre", "title"]

def find_col(df, candidates):
    cols_lower = {c.lower(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in cols_lower:
            return cols_lower[cand.lower()]
    return None

def coerce_date(x):
    if pd.isna(x) or str(x).strip() == "":
        return None
    try:
        return dateparser.parse(str(x), dayfirst=True, fuzzy=True)
    except Exception:
        return None

def validate_dataframe(df):
    issues = []
    col_map = {}
    for logical, names in REQUIRED_COLS.items():
        col = find_col(df, names)
        if not col:
            issues.append(f"Colonne requise manquante : {logical} (attendu parmi {names})")
        else:
            col_map[logical] = col

    content_col = find_col(df, CONTENT_CANDIDATES)
    title_col = find_col(df, TITLE_CANDIDATES)

    if not content_col:
        issues.append("⚠️ Pas de colonne de contenu trouvée (résumés plus pauvres).")

    if not title_col:
        issues.append("⚠️ Pas de colonne de titre trouvée.")

    if issues and not col_map:
        return issues, None, None, None

    # Vérifs de base
    pub_col = col_map["publication"]
    date_col = col_map["published"]
    url_col = col_map["URL"]

    parsed_dates = df[date_col].apply(coerce_date)
    df = df.copy()
    df["_parsed_date"] = parsed_dates

    empty_rows = df[df[pub_col].astype(str).str.strip().eq("")]
    if not empty_rows.empty:
        issues.append(f"{len(empty_rows)} ligne(s) avec '{pub_col}' vide.")

    invalid_urls = [i + 2 for i, val in df[url_col].items() if val and not validators.url(str(val))]
    if invalid_urls:
        issues.append(f"URLs invalides aux lignes {invalid_urls}")

    return issues, col_map, content_col, title_col

def smart_summarize(publication, date_str, title, content, url, max_words=60):
    base_context = content or title or publication or url or "Article de presse"
    prompt = (
        f"Résume cet article de presse en français, de manière claire et concise (max ~{max_words} mots).\n\n"
        f"CONTEXTE:\n{base_context}\n\n"
        "Attendu: un seul paragraphe, neutre et informatif."
    )
    resp = client.responses.create(model=OPENAI_MODEL, input=prompt)
    return resp.output_text.strip()

def html_escape(s):
    return html.escape("" if s is None else str(s))

def build_email_html(rows, title="Revue de presse"):
    parts = []
    parts.append('<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;line-height:1.45;">')
    parts.append(f'<h1 style="font-size:20px;margin-bottom:16px;">{html_escape(title)}</h1>')
    for r in rows:
        parts.append('<div style="margin-bottom:14px;padding:12px;border:1px solid #e6e6e6;border-radius:8px;">')
        parts.append(f'<div><b>Publication:</b> {html_escape(r["publication"])}</div>')
        parts.append(f'<div><b>Date:</b> {html_escape(r["date"])}</div>')
        parts.append(f'<div><b>Résumé:</b> {html_escape(r["summary"])}</div>')
        if r.get("url"):
            url_esc = html_escape(r["url"])
            parts.append(f'<div><b>Lien:</b> <a href="{url_esc}">{url_esc}</a></div>')
        parts.append('</div>')
    parts.append('</div>')
    return "\n".join(parts)

def find_excel_file(folder="data"):
    files = glob.glob(os.path.join(folder, "*.xlsx"))
    if not files:
        return None
    if len(files) == 1:
        return files[0]
    files = sorted(files, key=os.path.getmtime, reverse=True)
    print("📂 Plusieurs fichiers trouvés :")
    for i, f in enumerate(files, 1):
        print(f" {i}. {os.path.basename(f)} (modifié le {datetime.fromtimestamp(os.path.getmtime(f)).strftime('%d/%m/%Y %H:%M')})")
    while True:
        try:
            choice = int(input("👉 Entrez le numéro du fichier à utiliser : "))
            if 1 <= choice <= len(files):
                return files[choice - 1]
        except ValueError:
            pass
        print("⚠️ Choix invalide, réessayez.")

def send_via_outlook(subject, html_body, to=""):
    try:
        import win32com.client as win32
    except ImportError:
        print("❌ win32com.client non installé. Installez-le avec : pip install pywin32")
        return
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = to
    mail.Subject = subject
    mail.HTMLBody = html_body
    mail.Display()  # ouvre le mail en brouillon

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--excel", help="Chemin du fichier Excel")
    parser.add_argument("--sheet", help="Nom de la feuille", default="Articles")
    parser.add_argument("--title", help="Titre du mail", default="Revue de presse")
    parser.add_argument("--max-rows", type=int, default=None)
    parser.add_argument("--to", help="Destinataire (optionnel)", default="")
    args = parser.parse_args()

    if args.excel:
        excel_path = args.excel
    else:
        excel_path = find_excel_file("data")
        if not excel_path:
            print("❌ Aucun fichier Excel trouvé dans ./data")
            sys.exit(1)
    print(f"✅ Fichier sélectionné : {excel_path}")

    try:
        df = pd.read_excel(excel_path, sheet_name=args.sheet)
    except Exception as e:
        print(f"Erreur de lecture Excel: {e}")
        sys.exit(1)

    issues, col_map, content_col, title_col = validate_dataframe(df)
    if issues:
        print("=== Problèmes détectés ===")
        for it in issues:
            print(" -", it)

    if not col_map:
        print("❌ Colonnes requises introuvables.")
        sys.exit(1)

    pub_col = col_map["publication"]
    date_col = col_map["published"]
    url_col = col_map["URL"]

    if args.max_rows:
        df = df.head(args.max_rows)

    rows = []
    for _, row in df.iterrows():
        publication = str(row.get(pub_col, "")).strip()
        url = str(row.get(url_col, "")).strip() or None
        dt = row.get("_parsed_date") or coerce_date(row.get(date_col))
        date_out = dt.strftime("%d/%m/%Y") if isinstance(dt, datetime) else str(row.get(date_col, "")).strip()
        title_val = str(row.get(title_col, "")).strip() if title_col else publication
        content_val = str(row.get(content_col, "")).strip() if content_col else ""

        try:
            summary = smart_summarize(publication, date_out, title_val, content_val, url)
        except Exception as e:
            summary = title_val or "Résumé indisponible"
            print(f"[⚠] Résumé échoué : {e}")

        rows.append({
            "publication": publication,
            "date": date_out,
            "summary": summary,
            "url": url
        })

    html_out = build_email_html(rows, title=args.title)

    # Sauvegarde locale
    with open("rapport.html", "w", encoding="utf-8") as f:
        f.write(html_out)
    print("✅ Fichier HTML généré : rapport.html")

    # Ouvre dans Outlook
    send_via_outlook(args.title, html_out, args.to)

if __name__ == "__main__":
    main()
