"""
Application Streamlit SDK — regroupe 3 outils :
  1. Saisie Facture   : extraction de factures fournisseurs (BL) depuis PDF/ZIP
  2. Saisie Situation : extraction des factures de situation clients (PDF)
                        (Génération automatique ou mise à jour gérée en arrière-plan)
  3. Suivi MO         : récap des jours travaillés par chantier (Excel "Suivi heures")

Clé API Anthropic : à renseigner dans .streamlit/secrets.toml
    ANTHROPIC_API_KEY = "sk-ant-..."
"""

import io
import zipfile
import tempfile
from pathlib import Path

import streamlit as st

from core import saisie_facture, saisie_situation, suivi_mo

st.set_page_config(page_title="Outils SDK", layout="wide")

st.title("Outils SDK")

tab1, tab2, tab3 = st.tabs(["Saisie Facture", "Saisie Situation", "Suivi MO"])


def get_api_key() -> str:
    key = st.secrets.get("ANTHROPIC_API_KEY", "")
    if not key:
        st.error("Clé API Anthropic manquante. Configurez ANTHROPIC_API_KEY dans les secrets.")
    return key


# ─────────────────────────────────────────────────────────────────────────────
# TAB 1 : SAISIE FACTURE
# ─────────────────────────────────────────────────────────────────────────────
with tab1:
    st.header("Saisie Facture (fournisseurs)")
    st.write("Envoyez un ou plusieurs PDF, ou un fichier ZIP contenant des PDF.")

    uploaded_files = st.file_uploader(
        "PDF ou ZIP",
        type=["pdf", "zip"],
        accept_multiple_files=True,
        key="facture_upload",
    )

    if st.button("Lancer l'extraction", key="facture_run"):
        api_key = get_api_key()

        if not uploaded_files:
            st.warning("Aucun fichier envoyé.")
        elif api_key:
            with tempfile.TemporaryDirectory() as tmpdir:
                tmp_path = Path(tmpdir)
                pdf_paths = []

                for f in uploaded_files:
                    if f.name.lower().endswith(".zip"):
                        with zipfile.ZipFile(io.BytesIO(f.read())) as zf:
                            for member in zf.namelist():
                                if member.endswith("/") or "__MACOSX" in member:
                                    continue
                                if not member.lower().endswith(".pdf"):
                                    continue
                                target = tmp_path / Path(member).name
                                with zf.open(member) as src, open(target, "wb") as dst:
                                    dst.write(src.read())
                                pdf_paths.append(target)
                    else:
                        target = tmp_path / f.name
                        with open(target, "wb") as dst:
                            dst.write(f.read())
                        pdf_paths.append(target)

                if not pdf_paths:
                    st.warning("Aucun PDF trouvé dans les fichiers envoyés.")
                else:
                    st.info(f"{len(pdf_paths)} PDF à traiter...")
                    log_area = st.empty()
                    logs = []

                    def log(msg):
                        logs.append(str(msg))
                        log_area.text("\n".join(logs[-30:]))

                    with st.spinner("Extraction en cours..."):
                        try:
                            result_bytes = saisie_facture.traiter(pdf_paths, api_key, log=log)
                            st.success("Extraction terminée.")
                            st.download_button(
                                "Télécharger le fichier Excel",
                                data=result_bytes,
                                file_name="saisie_facture.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            )
                        except Exception as e:
                            st.error(f"Erreur : {e}")


# ─────────────────────────────────────────────────────────────────────────────
# TAB 2 : SAISIE SITUATION
# ─────────────────────────────────────────────────────────────────────────────
with tab2:
    st.header("Saisie Situation (factures clients)")
    st.write("Envoyez le PDF de facturation mensuelle pour générer ou mettre à jour le suivi.")

    pdf_situation = st.file_uploader("PDF de facturation", type=["pdf"], key="situation_pdf")

    if st.button("Lancer l'extraction", key="situation_run"):
        api_key = get_api_key()

        if not pdf_situation:
            st.warning("Merci d'envoyer le PDF.")
        elif api_key:
            with tempfile.TemporaryDirectory() as tmpdir:
                tmp_path = Path(tmpdir)

                pdf_path = tmp_path / pdf_situation.name
                with open(pdf_path, "wb") as f:
                    f.write(pdf_situation.read())

                log_area = st.empty()
                logs = []

                def log(msg):
                    logs.append(str(msg))
                    log_area.text("\n".join(logs[-30:]))

                with st.spinner("Extraction en cours..."):
                    try:
                        result_bytes = saisie_situation.traiter(
                            str(pdf_path), api_key, log=log
                        )
                        st.success("Extraction terminée.")
                        st.download_button(
                            "Télécharger le fichier Excel mis à jour",
                            data=result_bytes,
                            file_name="suivi_facturation.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
                    except Exception as e:
                        st.error(f"Erreur : {e}")


# ─────────────────────────────────────────────────────────────────────────────
# TAB 3 : SUIVI MO
# ─────────────────────────────────────────────────────────────────────────────
with tab3:
    st.header("Suivi Main d'Œuvre")
    st.write("Envoyez le fichier Excel source (onglet 'Suivi heures').")

    mo_xlsx = st.file_uploader("Fichier Excel source", type=["xlsx"], key="mo_upload")
    mois_label = st.text_input("Libellé du mois (optionnel)", value="", key="mo_mois")

    if st.button("Générer le récap", key="mo_run"):
        if not mo_xlsx:
            st.warning("Merci d'envoyer un fichier Excel.")
        else:
            with tempfile.TemporaryDirectory() as tmpdir:
                tmp_path = Path(tmpdir) / mo_xlsx.name
                with open(tmp_path, "wb") as f:
                    f.write(mo_xlsx.read())

                with st.spinner("Génération en cours..."):
                    try:
                        result_bytes = suivi_mo.traiter(str(tmp_path), mois_label=mois_label, log=st.write)
                        st.success("Récap généré.")
                        st.download_button(
                            "Télécharger le récap",
                            data=result_bytes,
                            file_name="recap_chantiers.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
                    except Exception as e:
                        st.error(f"Erreur : {e}")
