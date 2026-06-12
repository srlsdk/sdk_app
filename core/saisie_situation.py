"""
core/saisie_situation.py
-------------------------
Extraction des factures de situation (clients) depuis un PDF mensuel
SADAKA, et mise à jour d'un fichier Excel "SUIVI DU CA" à partir d'un
modèle fourni.

Logique métier identique au script CLI d'origine (saisie_situation.py),
adaptée pour être appelée depuis Streamlit (entrées/sorties en mémoire).
"""

import io
import json
import base64
import re
import anthropic
from openpyxl import load_workbook

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURATION
# ─────────────────────────────────────────────────────────────────────────────

# Nom de la feuille à utiliser dans le fichier modèle. Si absent, la feuille
# active sera utilisée.
SHEET_NAME = "SUIVI DU CA"

# Première ligne où écrire les données (les lignes 1-3 sont les en-têtes)
FIRST_DATA_ROW = 4

# Mapping colonnes Excel -> clés JSON
COLUMN_MAPPING = {
    1: "client",
    2: "affaire",
    3: "date_facture",
    4: "num_situation",
    5: "num_facture",
    6: "montant_ht",
    7: "tva_montant",
    8: "montant_ttc",
    9: "rg_ht",
    11: "penalites",
    13: "prorata_ht",
}

SYSTEM_PROMPT = """
Tu es un expert en comptabilité BTP française. Tu analyses des factures de
situation de travaux (PDF) émises par l'entreprise SADAKA.

Pour CHAQUE facture présente dans le document, extrait les informations
suivantes et retourne UNIQUEMENT un tableau JSON valide (aucun texte avant
ou après) avec la structure suivante :

[
  {
    "client": "nom du destinataire (donneur d'ordre)",
    "affaire": "nom du chantier / affaire",
    "date_facture": "DD/MM/YYYY",
    "num_situation": "numéro de situation (chaîne, ex: '13', '2', 'N/A')",
    "num_facture": "numéro de facture (ex: '2026/34')",
    "montant_ht": 0.0,
    "tva_montant": 0.0,
    "montant_ttc": 0.0,
    "rg_ht": 0.0,
    "penalites": 0.0,
    "prorata_ht": 0.0
  }
]

Règles importantes :
- "montant_ht" et "tva_montant" correspondent à la situation HT du mois (colonne
  "Situation HT du mois" / dernière colonne du tableau), pas au total cumulé
  ni au montant du marché.
- "montant_ttc" correspond au net à payer / montant de la situation du mois
  (dernière colonne "MONTANT TTC NET A PAYER" ou équivalent), pas au cumulé.
- "rg_ht" = retenue de garantie du mois (situation du mois), 0 si absente.
- "prorata_ht" = part de prorata HT/TTC du mois, 0 si absente.
- "penalites" = pénalités du mois, 0 si absentes.
- Pour les factures sans tableau de situation (factures simples), utilise le
  montant HT, la TVA et le montant TTC de la facture ; rg_ht, penalites et
  prorata_ht = 0.
- Si une valeur n'est pas applicable, mets 0 pour les nombres et "N/A" pour
  num_situation si non précisé.
"""

USER_PROMPT = "Extrais toutes les factures présentes dans ce document en JSON."


def extraire_donnees_pdf(pdf_path: str, api_key: str) -> list:
    """Envoie le PDF à Claude et retourne la liste des factures extraites."""
    client = anthropic.Anthropic(api_key=api_key)

    with open(pdf_path, "rb") as f:
        pdf_b64 = base64.standard_b64encode(f.read()).decode("utf-8")

    message = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=4096,
        system=SYSTEM_PROMPT,
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "document",
                        "source": {
                            "type": "base64",
                            "media_type": "application/pdf",
                            "data": pdf_b64,
                        },
                    },
                    {"type": "text", "text": USER_PROMPT},
                ],
            }
        ],
    )

    contenu = message.content[0].text.strip()

    match = re.search(r"\[.*\]", contenu, re.DOTALL)
    json_clean = match.group(0) if match else contenu

    return json.loads(json_clean)


def generer_excel(template_path: str, factures: list) -> bytes:
    """
    Charge le fichier modèle, vide les anciennes lignes de données puis écrit
    une ligne par facture en conservant le style (bordures, formats) des
    lignes du modèle.

    Retourne le contenu xlsx en bytes.
    """
    wb = load_workbook(template_path)
    ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active

    last_existing_row = FIRST_DATA_ROW
    while ws.cell(row=last_existing_row, column=5).value:
        last_existing_row += 1
    last_existing_row -= 1
    if last_existing_row < FIRST_DATA_ROW:
        last_existing_row = FIRST_DATA_ROW

    for r in range(FIRST_DATA_ROW, last_existing_row + 1):
        for c in range(1, 15):
            ws.cell(row=r, column=c).value = None

    max_col = 14

    for i, facture in enumerate(factures):
        row_idx = FIRST_DATA_ROW + i

        if row_idx > last_existing_row:
            style_row = last_existing_row
            for c in range(1, max_col + 1):
                src = ws.cell(row=style_row, column=c)
                dst = ws.cell(row=row_idx, column=c)
                dst._style = src._style.copy() if hasattr(src._style, "copy") else src._style

        for col, key in COLUMN_MAPPING.items():
            cell = ws.cell(row=row_idx, column=col)
            value = facture.get(key)

            if col in (1, 2, 4, 5):
                cell.value = "" if value is None else str(value)
            elif col == 3:
                cell.value = value if value else None
            else:
                try:
                    cell.value = float(value) if value not in (None, "") else 0.0
                except (TypeError, ValueError):
                    cell.value = 0.0

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


def traiter(pdf_path: str, template_path: str, api_key: str, log=print) -> bytes:
    """
    Point d'entrée principal.
    pdf_path, template_path : chemins sur disque (temp dir).
    Retourne le contenu xlsx en bytes.
    """
    log("Lecture du PDF par Claude Haiku...")
    factures = extraire_donnees_pdf(pdf_path, api_key)
    log(f"{len(factures)} facture(s) extraite(s).")

    log("Génération du fichier Excel...")
    return generer_excel(template_path, factures)
