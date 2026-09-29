# Gestion Facture — Automatisation des attestations de cession (AT)

Application de bureau Windows qui automatise la préparation des **attestations de cession
sous admission temporaire (AT)** pour une entreprise de transformation de papier/carton.

À partir d'un dossier de PDF (fiches produit + facture de cession), l'application :

1. **extrait** les données des PDF (texte natif avec pdfminer, OCR Tesseract pour les scans) ;
2. **calcule** la composition papier et les poids brut / net / à déclarer de chaque produit ;
3. **impute** les quantités sur les déclarations douanières du **sommier** (base PostgreSQL),
   par ordre chronologique, en mettant à jour le reliquat de chaque déclaration ;
4. **génère** l'attestation et la fiche de calcul à partir de modèles Excel, exportées en **XLSX et PDF**.

## Fonctionnalités

- Interface moderne (CustomTkinter) : accueil, traitement, résultats, graphiques, paramètres
- Extraction automatique : client, code article, cannelure, composition, surface, poids, quantités
- Correction manuelle des types de papier avant imputation
- Imputation FIFO sur le sommier avec mise à jour des reliquats (`resteenkg`)
- Export Excel → PDF via Excel (xlwings), un dossier de résultats par jour
- Visualisation des données (barres, camemberts) et consultation de fichiers Excel
- Thème clair / sombre / système
- Toute la configuration machine (base de données, chemins, outils OCR) dans un fichier `.env`

## Prérequis

- **Windows** avec **Microsoft Excel** (requis par xlwings pour remplir les modèles et exporter en PDF)
- **Python 3.10+**
- **PostgreSQL**
- **Tesseract OCR** avec le pack de langue française — https://github.com/UB-Mannheim/tesseract/wiki
- **Poppler** pour Windows (conversion PDF → image) — https://github.com/oschwartz10612/poppler-windows/releases

## Installation

```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt

copy .env.example .env      # puis adapter les valeurs
```

Créer la base et la table :

```bash
createdb -U postgres sommier
psql -U postgres -d sommier -f schema.sql
```

Déposer les modèles Excel dans `templates/` (voir [templates/README.md](templates/README.md)).

## Configuration (`.env`)

| Variable | Description | Défaut |
|---|---|---|
| `DB_HOST` / `DB_PORT` | Serveur PostgreSQL | `localhost` / `5432` |
| `DB_NAME` | Base de données | `sommier` |
| `DB_USER` / `DB_PASSWORD` | Identifiants PostgreSQL | `postgres` / *(vide)* |
| `TESSERACT_CMD` | Chemin de `tesseract.exe` | `tesseract` du PATH, sinon `C:\Program Files\Tesseract-OCR\tesseract.exe` |
| `TESSERACT_LANG` | Langue OCR | `fra` |
| `POPPLER_PATH` | Dossier contenant `pdftoppm.exe` | PATH |
| `RESULT_DIR` | Dossier de sortie | `./resultat` |
| `ATTESTATION_TEMPLATE` | Modèle Excel de l'attestation | `./templates/attestation.xlsx` |
| `CALCUL_TEMPLATE` | Modèle Excel de la fiche de calcul | `./templates/calcul.xlsx` |
| `CALCUL_LOG_FILE` | Journal Excel des calculs | `./excelfiles/default.xlsx` |

Le dossier de résultats et le fichier journal peuvent aussi être modifiés depuis l'écran **Paramètres**.

## Utilisation

```bash
python app.py
```

1. Au démarrage, sélectionner le fichier Excel du **sommier** (colonnes : `bureau, nomenclature, regime,
   annee, declaration, date, frs, mottechnique, qualite, observation, qenkg, resteenkg, reste`) —
   il est chargé dans PostgreSQL.
2. Cliquer sur **Process** et choisir le dossier contenant les PDF d'une cession
   (fiches produit + facture de cession).
3. Vérifier / corriger les types de papier proposés.
4. Les attestations et fiches de calcul sont générées dans `RESULT_DIR`.

## Structure du projet

```
gestionfacture/
├── app.py                  # Interface graphique (point d'entrée)
├── main.py                 # Extraction PDF/OCR, calculs, imputation sur le sommier
├── export_attestation.py   # Remplissage du modèle d'attestation + export PDF
├── export_calcul.py        # Remplissage de la fiche de calcul + journal Excel
├── chart_viewer.py         # Visualiseur de graphiques autonome
├── config.py               # Configuration (lue depuis .env)
├── schema.sql              # Table PostgreSQL « sommier »
├── app.spec                # Build PyInstaller
├── requirements.txt
└── templates/              # Modèles Excel (non versionnés)
```

## Générer l'exécutable

```bash
pip install pyinstaller
pyinstaller app.spec        # → dist/GestionFacture.exe
```

Placer le fichier `.env` à côté de l'exécutable. Tesseract et Poppler doivent être installés sur
la machine cible (ou leurs chemins renseignés dans `.env`).

## Confidentialité

Les documents clients (factures, DUM, attestations, fiches produit, sommier) ne sont **jamais**
versionnés : le `.gitignore` fonctionne en liste blanche et n'autorise que le code source.
