import streamlit as st
import pandas as pd
import pdfplumber
import unidecode
from rapidfuzz import fuzz
import jinja2
import datetime
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill

def normalize(s):
    if pd.isna(s) or s is None:
        return ""
    return unidecode.unidecode(str(s)).replace(" ", "").lower()

def extract_pdf_lines(pdf_file):
    try:
        with pdfplumber.open(pdf_file) as pdf:
            lines = []
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    lines.extend(text.split("\n"))
            return lines
    except Exception:
        return []

def match_row_to_pdf(row, pdf_lines, fuzzy=True, fuzzy_threshold=85):
    results = []
    cmd = normalize(str(row.get("N° commande", "")))
    montant = normalize(str(row.get("Montant brut", "")))
    taux = normalize(str(row.get("Taux de facturation", "")))
    rubrique = normalize(str(row.get("Code rubrique", "")))
    unite = normalize(str(row.get("Unités", "")))
    semaine = normalize(str(row.get("Semaine finissant le", "")))

    for line in pdf_lines:
        norm_line = normalize(line)
        if cmd in norm_line:
            score = 0
            matched = []
            if montant and (montant in norm_line or (fuzzy and fuzz.partial_ratio(montant, norm_line) > fuzzy_threshold)):
                score += 1
                matched.append("montant")
            if taux and (taux in norm_line or (fuzzy and fuzz.partial_ratio(taux, norm_line) > fuzzy_threshold)):
                score += 1
                matched.append("taux")
            if rubrique and (rubrique in norm_line or (fuzzy and fuzz.partial_ratio(rubrique, norm_line) > fuzzy_threshold)):
                score += 1
                matched.append("rubrique")
            if unite and (unite in norm_line or (fuzzy and fuzz.partial_ratio(unite, norm_line) > fuzzy_threshold)):
                score += 1
                matched.append("unités")
            if semaine and (semaine in norm_line or (fuzzy and fuzz.partial_ratio(semaine, norm_line) > fuzzy_threshold)):
                score += 1
                matched.append("semaine")
            results.append((score, line, matched))
    if results:
        results = sorted(results, key=lambda x: -x[0])
        return results[0][1], results[0][0], results[0][2]
    return "", 0, []

def generate_html_report(results, columns, logo_url=None):
    template = jinja2.Template("""
    <!DOCTYPE html>
    <html lang="fr">
    <head>
        <meta charset="UTF-8">
        <title>Rapport de rapprochement Excel ↔ PDF</title>
        <style>
            body { font-family: Arial, sans-serif; margin: 2em;}
            h1 { color: #003366;}
            table { border-collapse: collapse; width: 100%; margin-top: 1em;}
            th, td { border: 1px solid #ddd; padding: 8px; }
            th { background-color: #f2f2f2; }
            tr:nth-child(even) { background-color: #f9f9f9;}
            .matched { background: #d4edda;}
            .unmatched { background: #f8d7da;}
            .score-high { color: #155724;}
            .score-low { color: #721c24;}
        </style>
    </head>
    <body>
        {% if logo_url %}
        <img src="{{ logo_url }}" alt="Logo" style="height:60px;"/>
        {% endif %}
        <h1>Rapport de rapprochement Excel ↔ PDF</h1>
        <p>Pour enregistrer ce rapport en PDF, ouvrez-le dans votre navigateur puis faites <b>Imprimer &gt; Enregistrer au format PDF</b>.</p>
        <table>
            <tr>
                {% for col in columns %}
                <th>{{ col }}</th>
                {% endfor %}
                <th>Fichier Excel</th>
                <th>PDF correspondant</th>
                <th>Ligne PDF</th>
                <th>Score</th>
                <th>Champs trouvés</th>
            </tr>
            {% for row in results %}
            <tr class="{{ 'matched' if row['Score correspondance'] > 0 else 'unmatched' }}">
                {% for col in columns %}
                <td>{{ row[col] }}</td>
                {% endfor %}
                <td>{{ row['Fichier Excel'] }}</td>
                <td>{{ row['PDF correspondant'] }}</td>
                <td>{{ row['Ligne PDF correspondante'] }}</td>
                <td class="{{ 'score-high' if row['Score correspondance'] > 2 else 'score-low' }}">{{ row['Score correspondance'] }}</td>
                <td>{{ row['Champs trouvés'] }}</td>
            </tr>
            {% endfor %}
        </table>
        <p style="margin-top:2em;font-size:small;color:#888;">Généré par BEELINE, le {{ now }}</p>
    </body>
    </html>
    """)
    return template.render(results=results, columns=columns, logo_url=logo_url, now=datetime.datetime.now().strftime("%d/%m/%Y %H:%M"))

def generate_excel_report(df, excel_columns):
    wb = Workbook()
    ws = wb.active
    ws.title = "Rapprochement"
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="003366")
    for col_num, col in enumerate(excel_columns, 1):
        cell = ws.cell(row=1, column=col_num, value=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for row_idx, row in enumerate(df.itertuples(index=False), 2):
        for col_num, value in enumerate(row, 1):
            cell = ws.cell(row=row_idx, column=col_num, value=value)
            cell.alignment = Alignment(horizontal="left", vertical="center")
    for col in ws.columns:
        max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max(12, min(40, max_length + 2))
    ws.auto_filter.ref = ws.dimensions
    file_stream = BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)
    return file_stream

st.title("Rapprochement multi-fichiers Excel ↔ PDF : Application robuste et rapport professionnel")

pdf_files = st.file_uploader("Importer un ou plusieurs PDF", type="pdf", accept_multiple_files=True)
excel_files = st.file_uploader("Importer un ou plusieurs fichiers Excel", type=["xlsx", "xls"], accept_multiple_files=True)
fuzzy = st.checkbox("Tolérance de rapprochement (fuzzy)", value=True)
fuzzy_threshold = st.slider("Seuil fuzzy (plus élevé = plus strict)", 70, 100, 85) if fuzzy else 100

if excel_files and pdf_files:
    st.info("Traitement en cours...")

    columns = ["Collaborateur", "N° commande", "Montant brut", "Montant net à payer au fournisseur", "Montant de la taxe",
               "Taux de facturation", "Code rubrique", "Unités", "Semaine finissant le"]
    excel_columns = columns + ["Fichier Excel", "PDF correspondant", "Ligne PDF correspondante", "Score correspondance", "Champs trouvés"]

    # Charger tous les PDF
    parsed_pdfs = []
    for pdf_file in pdf_files:
        lines = extract_pdf_lines(pdf_file)
        parsed_pdfs.append({'filename': pdf_file.name, 'lines': lines})

    results = []

    for excel_file in excel_files:
        df = pd.read_excel(excel_file)
        for idx, row in df.iterrows():
            best_pdf = ""
            best_line = ""
            best_score = -1
            best_matched = []
            for pdf in parsed_pdfs:
                matched_line, score, matched_fields = match_row_to_pdf(row, pdf['lines'], fuzzy, fuzzy_threshold)
                if score > best_score:
                    best_score = score
                    best_pdf = pdf['filename']
                    best_line = matched_line
                    best_matched = matched_fields
            results.append({
                **{col: row.get(col, "") for col in columns},
                "Fichier Excel": excel_file.name,
                "PDF correspondant": best_pdf if best_score > 0 else "",
                "Ligne PDF correspondante": best_line,
                "Score correspondance": best_score,
                "Champs trouvés": ', '.join(best_matched)
            })

    res_df = pd.DataFrame(results)
    # Forcer l'ordre des colonnes dans le DataFrame selon excel_columns
    res_df = res_df[excel_columns]

    st.subheader("Résultats du rapprochement")
    st.dataframe(res_df, use_container_width=True)

    with st.expander("🔍 Filtrer les résultats"):
        col1, col2 = st.columns(2)
        with col1:
            search_collab = st.text_input("Recherche par nom d'intérimaire")
            min_score = st.slider("Score minimum", 0, 5, 0)
        with col2:
            only_matched = st.checkbox("Afficher seulement les correspondances")
        filtered = res_df
        if search_collab:
            filtered = filtered[filtered["Collaborateur"].str.contains(search_collab, case=False, na=False)]
        filtered = filtered[filtered["Score correspondance"] >= min_score]
        if only_matched:
            filtered = filtered[filtered["Score correspondance"] > 0]
        st.dataframe(filtered, use_container_width=True)

    html_report = generate_html_report(filtered.to_dict(orient="records"), columns)
    st.subheader("📄 Générer un rapport HTML à imprimer en PDF")
    st.download_button(
        label="Télécharger le rapport HTML",
        data=html_report,
        file_name="rapport_rapprochement.html",
        mime="text/html"
    )
    st.markdown(
        """
        <div style="background:#eaf4fc;padding:12px;border-radius:6px;font-size:15px;">
        <b>Pour générer le PDF :</b><br>
        1. Téléchargez le rapport HTML ci-dessus.<br>
        2. Ouvrez-le dans votre navigateur.<br>
        3. Imprimez-le (Ctrl+P ou Cmd+P) et choisissez “Enregistrer au format PDF”.<br>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.subheader("📊 Générer un beau rapport Excel (formaté)")
    excel_file_bytes = generate_excel_report(filtered, excel_columns)
    st.download_button(
        label="Télécharger le rapport Excel",
        data=excel_file_bytes,
        file_name="rapport_rapprochement.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
