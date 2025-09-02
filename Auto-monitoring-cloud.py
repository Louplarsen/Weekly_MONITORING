import os
import html
import unicodedata
from datetime import datetime
import pandas as pd
from dateutil import parser as dateparser
import validators
import streamlit as st
from dotenv import load_dotenv
from openai import OpenAI

# --- Chargement des variables d'environnement ---
load_dotenv()
client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))
OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gpt-4o-mini")

# --- Colonnes attendues ---
REQUIRED_COLS = {
    "publication": ["media outlet", "publication", "media", "journal"],
    "published": ["published", "date", "publication_date", "date de parution"],
    "URL": ["url", "lien", "link", "adresse web"],
}
CONTENT_CANDIDATES = ["snippet", "content", "texte", "text", "body", "résumé", "summary"]
TITLE_CANDIDATES = ["article", "titre", "title", "intitulé", "headline"]

# --- Normalisation des noms de colonnes ---
def normalize_colname(name: str) -> str:
    """Normalise un nom de colonne : minuscule, sans accents, sans espaces superflus."""
    name = str(name).strip().lower()
    name = ''.join(
        c for c in unicodedata.normalize("NFD", name)
        if unicodedata.category(c) != "Mn"
    )
    return name

def find_col(df, candidates):
    """Trouve la première colonne du dataframe qui correspond à une des candidates normalisées."""
    normalized_map = {normalize_colname(c): c for c in df.columns}
    for cand in candidates:
        norm_cand = normalize_colname(cand)
        if norm_cand in normalized_map:
            return normalized_map[norm_cand]
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

# --- Streamlit App ---
st.set_page_config(page_title="Revue de Presse", layout="wide")
st.title("📰 Générateur de Revue de Presse")

uploaded_file = st.file_uploader("📂 Uploader un fichier Excel", type=["xlsx"])
report_title = st.text_input("✏️ Titre du rapport", "Revue de presse")

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, sheet_name="Articles")
        issues, col_map, content_col, title_col = validate_dataframe(df)

        if issues:
            st.warning("⚠️ Problèmes détectés :")
            for i in issues:
                st.write("- " + i)

        if not col_map:
            st.error("❌ Colonnes requises introuvables")
        else:
            rows = []
            progress = st.progress(0)
            total = len(df)

            for idx, row in df.iterrows():
                publication = str(row.get(col_map["publication"], "")).strip()
                url = str(row.get(col_map["URL"], "")).strip() or None
                dt = row.get("_parsed_date") or coerce_date(row.get(col_map["published"]))
                date_out = dt.strftime("%d/%m/%Y") if isinstance(dt, datetime) else str(row.get(col_map["published"], "")).strip()
                title_val = str(row.get(title_col, "")).strip() if title_col else publication
                content_val = str(row.get(content_col, "")).strip() if content_col else ""

                try:
                    summary = smart_summarize(publication, date_out, title_val, content_val, url)
                except Exception as e:
                    summary = title_val or "Résumé indisponible"

                rows.append({
                    "publication": publication,
                    "date": date_out,
                    "summary": summary,
                    "url": url
                })

                progress.progress((idx + 1) / total)

            html_out = build_email_html(rows, title=report_title)

            st.subheader("📄 Rapport généré")
            st.markdown(html_out, unsafe_allow_html=True)

            st.download_button(
                "⬇️ Télécharger le rapport HTML",
                data=html_out,
                file_name="rapport.html",
                mime="text/html"
            )

    except Exception as e:
        st.error(f"Erreur lecture ou traitement Excel : {e}")
