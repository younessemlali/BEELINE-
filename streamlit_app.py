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

def extract_invoice_info(pdf_lines):
    # Recherche du numéro de facture et du total net dans le PDF (synthèse)
    invoice_id = ""
    net_total = None
    for i, line in enumerate(pdf_lines):
        # Recherche du numéro de facture
        if "invoice id" in normalize(line) or "numéro" in normalize(line):
            # Exemple : "Invoice ID/Number / Numéro 4949S0001"
            parts = line.split()
            for part in parts:
                # Prend le premier qui ressemble à un numéro de facture (lettre+chiffre)
                if any(c.isdigit() for c in part) and any(c.isalpha() for c in part):
                    invoice_id = part.strip()
                    break
        # Recherche du total net
        if "invoice total" in normalize(line):
            # Exemple : "Invoice Total(EUR) 9.84 1.98 11.82"
            parts = line.replace(",", ".").split()
            # Cherche le premier nombre après le titre
            for idx, p in enumerate(parts):
                if "total" in normalize(p):
                    # Les 3 suivants sont net, tva, brut
                    try:
                        net_total = float(parts[idx + 1])
                    except:
                        pass
                    break
    return invoice_id, net_total

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

def generate_html_report(results, columns, synthese_factures, logo_url=None):
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
            .facture-ok { background: #d4edda; }
            .facture-ko { background: #f8d7da; }
        </style>
    </head>
    <body>
        {% if logo_url %}
        <img src="{{ logo_url }}" alt="Logo" style="height:60px;"/>
        {% endif %}
        <h1>Rapport de rapprochement Excel ↔ PDF</h1>
        <h2>Synthèse par facture</h2>
        <table>
            <tr>
                <th>Fichier Excel</th>
                <th>Fichier PDF</th>
                <th>Numéro de facture</th>
                <th>Total net PDF</th>
                <th>Total net Excel</th>
                <th>Statut</th>
            </tr>
            {% for syn in synthese_factures %}
            <tr class="{{ 'facture-ok' if syn['Statut'] == 'OK' else 'facture-ko' }}">
                <td>{{ syn['Fichier Excel'] }}</td>
                <td>{{ syn['Fichier PDF'] }}</td>
                <td>{{ syn['Numéro facture'] }}</td>
                <td>{{ "{:,.2f}".format(syn['Total net PDF']) if syn['Total net PDF'] is not none else '' }}</td>
                <td>{{ "{:,.2f}".format(syn['Total net Excel']) if syn['Total net Excel'] is not none else '' }}</td>
                <td>{{ syn['Statut'] }}</td>
            </tr>
            {% endfor %}
        </table>
        <h2>Rapprochement ligne à ligne</h2>
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
    return template.render(results=results, columns=columns, synthese_factures=synthese_factures, logo_url=logo_url, now=datetime.datetime.now().strftime("%d/%m/%Y %H:%M"))

def generate_excel_report(df, excel_columns, synthese_factures):
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

    # Ajout d'une feuille synthèse facture
    ws2 = wb.create_sheet("Synthèse Factures")
    syn_cols = ["Fichier Excel", "Fichier PDF", "Numéro facture", "Total net PDF", "Total net Excel", "Statut"]
    for col_num, col in enumerate(syn_cols, 1):
        cell = ws2.cell(row=1, column=col_num, value=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for row_idx, syn in enumerate(synthese_factures, 2):
        ws2.cell(row=row_idx, column=1, value=syn["Fichier Excel"])
        ws2.cell(row=row_idx, column=2, value=syn["Fichier PDF"])
        ws2.cell(row=row_idx, column=3, value=syn["Numéro facture"])
        ws2.cell(row=row_idx, column=4, value=syn["Total net PDF"])
        ws2.cell(row=row_idx, column=5, value=syn["Total net Excel"])
        ws2.cell(row=row_idx, column=6, value=syn["Statut"])
    for col in ws2.columns:
        max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        ws2.column_dimensions[col[0].column_letter].width = max(12, min(40, max_length + 2))
    ws2.auto_filter.ref = ws2.dimensions

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

    # Charger tous les PDF et extraire info facture
    parsed_pdfs = []
    pdf_infos = dict()
    for pdf_file in pdf_files:
        lines = extract_pdf_lines(pdf_file)
        invoice_id, net_total = extract_invoice_info(lines)
        parsed_pdfs.append({'filename': pdf_file.name, 'lines': lines})
        pdf_infos[pdf_file.name] = {
            "Numéro facture": invoice_id,
            "Total net PDF": net_total
        }

    # Calcul de la synthèse factures
    synthese_factures = []
    results = []

    # Correspondance par ordre d'upload : Excel 0 <-> PDF 0, etc.
    # Si nb fichiers différent, on fait au mieux.
    nb_match = min(len(excel_files), len(pdf_files))
    for i in range(nb_match):
        excel_file = excel_files[i]
        pdf_file = pdf_files[i]
        pdf_info = pdf_infos.get(pdf_file.name, {})
        df = pd.read_excel(excel_file)
        # Calcul du total net Excel
        try:
            total_net_excel = float(df["Montant net à payer au fournisseur"].fillna(0).sum())
        except Exception:
            total_net_excel = None
        total_net_pdf = pdf_info.get("Total net PDF")
        num_facture = pdf_info.get("Numéro facture", "")
        statut = "OK" if (total_net_pdf is not None and total_net_excel is not None and abs(total_net_pdf - total_net_excel) < 0.02) else "Erreur"
        synthese_factures.append({
            "Fichier Excel": excel_file.name,
            "Fichier PDF": pdf_file.name,
            "Numéro facture": num_facture,
            "Total net PDF": total_net_pdf,
            "Total net Excel": total_net_excel,
            "Statut": statut
        })

    # Ajout des lignes pour les fichiers restants (si déséquilibre nbre xlsx/pdf)
    if len(excel_files) > nb_match:
        for i in range(nb_match, len(excel_files)):
            excel_file = excel_files[i]
            df = pd.read_excel(excel_file)
            try:
                total_net_excel = float(df["Montant net à payer au fournisseur"].fillna(0).sum())
            except Exception:
                total_net_excel = None
            synthese_factures.append({
                "Fichier Excel": excel_file.name,
                "Fichier PDF": "",
                "Numéro facture": "",
                "Total net PDF": None,
                "Total net Excel": total_net_excel,
                "Statut": "Excel sans PDF"
            })
    if len(pdf_files) > nb_match:
        for i in range(nb_match, len(pdf_files)):
            pdf_file = pdf_files[i]
            pdf_info = pdf_infos.get(pdf_file.name, {})
            synthese_factures.append({
                "Fichier Excel": "",
                "Fichier PDF": pdf_file.name,
                "Numéro facture": pdf_info.get("Numéro facture", ""),
                "Total net PDF": pdf_info.get("Total net PDF"),
                "Total net Excel": None,
                "Statut": "PDF sans Excel"
            })

    # Rapprochement ligne à ligne (inchangé)
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

    html_report = generate_html_report(filtered.to_dict(orient="records"), columns, synthese_factures)
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

    st.subheader("📊 Générer un beau rapport Excel (2 feuilles : lignes et synthèse facture)")
    excel_file_bytes = generate_excel_report(filtered, excel_columns, synthese_factures)
    st.download_button(
        label="Télécharger le rapport Excel",
        data=excel_file_bytes,
        file_name="rapport_rapprochement.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
