# Outils SADAKA - Streamlit

Application Streamlit regroupant 3 outils internes :

1. **Saisie Facture** : extraction de factures fournisseurs (BL) depuis un ou
   plusieurs PDF, ou un fichier ZIP contenant des PDF. OCR de secours pour les
   PDF scannés.
2. **Saisie Situation** : extraction des factures de situation clients depuis
   le PDF de facturation mensuelle, et mise à jour d'un fichier Excel
   "SUIVI DU CA" à partir d'un modèle fourni.
3. **Suivi MO** : génère un récapitulatif des jours travaillés par chantier à
   partir d'un fichier Excel "Suivi heures".

## Structure du projet

```
sadaka_app/
├── streamlit_app.py       # page principale (3 onglets)
├── core/
│   ├── saisie_facture.py
│   ├── saisie_situation.py
│   └── suivi_mo.py
├── requirements.txt
├── packages.txt           # dépendances système (OCR)
├── .gitignore
└── .streamlit/
    └── secrets.toml.example
```

## Installation locale

```bash
python -m venv venv
source venv/bin/activate        # ou venv\Scripts\activate sur Windows
pip install -r requirements.txt
```

Pour l'OCR (PDF scannés), il faut aussi installer :
- **Tesseract OCR** (avec le pack langue français)
- **Poppler** (pour pdf2image)

et s'assurer qu'ils sont dans le `PATH` système.

## Configuration de la clé API

Copier `.streamlit/secrets.toml.example` en `.streamlit/secrets.toml` et
renseigner la clé Anthropic :

```toml
ANTHROPIC_API_KEY = "sk-ant-..."
```

Ce fichier **ne doit pas être commit** (il est dans `.gitignore`).

## Lancer l'application

```bash
streamlit run streamlit_app.py
```

## Déploiement (Streamlit Community Cloud)

1. Pousser ce dossier sur un repo GitHub (sans `.streamlit/secrets.toml`, ni
   `.cache/`, ni les fichiers `.xlsx` générés — voir `.gitignore`).
2. Sur [share.streamlit.io](https://share.streamlit.io), créer une nouvelle
   app en pointant sur ce repo, fichier principal `streamlit_app.py`.
3. Dans les **Settings > Secrets** de l'app, ajouter :
   ```toml
   ANTHROPIC_API_KEY = "sk-ant-..."
   ```
4. `requirements.txt` et `packages.txt` sont détectés automatiquement par
   Streamlit Cloud (le second installe poppler/tesseract via apt).

## Notes

- Le cache local (`.cache/`) du script CLI d'origine n'est pas repris dans
  la version Streamlit (environnement stateless).
- `OCR_WORKERS` est fixé à 4 dans `core/saisie_facture.py` ; ajuster selon
  les ressources disponibles sur l'environnement de déploiement.
