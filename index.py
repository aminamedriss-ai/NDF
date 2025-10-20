import streamlit as st
import datetime
import base64
from googleapiclient.discovery import build
import os
import gspread
import gspread
from google.oauth2.service_account import Credentials
import unicodedata
import difflib
import re
import io
import openpyxl
from googleapiclient.http import MediaIoBaseDownload
import calendar
from datetime import datetime
from gspread_formatting import (
    set_data_validation_for_cell_range, 
    DataValidationRule, 
    BooleanCondition, 
    CellFormat, 
    Color, 
    TextFormat, 
    Borders, 
    Border, 
    format_cell_range
)
import time
import random
from gspread_formatting import ConditionalFormatRule, BooleanRule, CellFormat, Color
from google.oauth2 import service_account
from googleapiclient.discovery import build
st.set_page_config(
    page_title="Gestion des note de frais",
    page_icon="logo2.png",  # chemin local ou URL
    layout="wide"
)
st.markdown("""
<style>
/* 🖼 Logo centré */
.logo-container {
    display: flex;
    justify-content: center;
    align-items: center;
    margin-bottom: 20px;
}

.logo {
    width: 200px;
    height: auto;
}

/* 🌈 Arrière-plan personnalisé + forcer mode sombre */
html, body, .stApp {
    background: #1d2e4e !important;
    font-family: 'Segoe UI', sans-serif;
    color-scheme: dark !important; /* Empêche l'inversion automatique */
    color: white !important;
}

/* 🖍️ Titre centré et coloré */
.main > div > div > div > div > h1 {
    text-align: center;
    color: #00796B !important;
}

/* 🧼 Nettoyage des bordures Streamlit */
.css-18e3th9 {
    padding: 1rem 0.5rem;
}

/* 🎨 Sidebar */
section[data-testid="stSidebar"] {
    background-color: #1f3763 !important;
    color: white !important;
}

section[data-testid="stSidebar"] .css-1v3fvcr {
    color: white !important;
}

/* 🌈 Titres dans la sidebar */
section[data-testid="stSidebar"] h1, 
section[data-testid="stSidebar"] h2, 
section[data-testid="stSidebar"] h3 {
    color: #e01b36 !important;
}

/* 🎨 Barre supérieure */
header[data-testid="stHeader"] {
    background-color: #06dbae !important;
    color: white !important;
}

/* 🧪 Supprimer la transparence */
header[data-testid="stHeader"]::before {
    content: "";
    background: none !important;
}

/* 📱 Correction mobile : forcer couleurs partout */
h1, h2, h3, p, span, label {
    color: white !important;
}

/* 🔵 Boutons bleu foncé forcés */
.stButton button {
    background-color: #2b2c36 !important; /* Bleu foncé */
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    padding: 0.5rem 1rem !important;
    font-weight: bold !important;
    -webkit-appearance: none !important; /* Évite style par défaut mobile */
    appearance: none !important;
}

.stButton button:hover {
    background-color: #43444e !important; /* Bleu plus clair au survol */
    color: white !important;
}
</style>
""", unsafe_allow_html=True)


# 🖼️ Ajouter un logo (remplacer "logo.png" par ton fichier ou une URL)
with open("logo.png", "rb") as image_file:
    encoded = base64.b64encode(image_file.read()).decode()

st.markdown(
    f"""
    <div class="logo-container">
        <img class="logo" src="data:image/png;base64,{encoded}">
    </div>
    """,
    unsafe_allow_html=True
)
st.markdown(
    "<h1 style='text-align: center;'>Bienvenue sur l'application de gestion des notes de frais</h1>",
    unsafe_allow_html=True
)

SCOPE = ["https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive"]

# CREDENTIALS_FILE = "service_account.json"
DEST_SHEET_ID = "1jxjAstmnsWCuRaYwVIhW-Qh7pZvh-waw3BEQ2HDGvRM"
ROOT_FOLDER_ID = "12UkT_IjkNazYn9QCUOjgk_ZoUpkNtlMe"
VERIFIED_ROOT_ID = "1N96PnXaouIs1KqkaKHy_mOCP_gj7-sbP"  # 📂 racine des dossiers VERIFIED
ndf_root_id = "1KTRuCR59xLgKLCT1_AY3z-lgeh9JFmrb"
# === AUTH ===
creds = service_account.Credentials.from_service_account_info(
    st.secrets["gcp_service_account"],
    scopes=SCOPE
)
client = gspread.authorize(creds)
drive_service = build("drive", "v3", credentials=creds)
# drive_service = build("sheets", "v4", credentials=creds)
# client = drive_service.spreadsheets()
st.title("💰 Transfert Montant à rembourser")

# === Utils ===
def normalize(text):
    if not text:
        return ""
    return unicodedata.normalize("NFKD", str(text)).lower().strip()

def list_subfolders(folder_id):
    res = drive_service.files().list(
        q=f"'{folder_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false",
        fields="files(id, name)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True
    ).execute()
    return res.get("files", [])

def list_sheets_in_folder(folder_id):
    res = drive_service.files().list(
        q=f"'{folder_id}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false",
        fields="files(id, name)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True
    ).execute()
    return res.get("files", [])

def to_float(val):
    if val is None:
        return None
    # si c'est déjà numérique
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if not s:
        return None

    # DEBUG: garder la chaîne originale pour logs si besoin
    orig = s

    # nettoyer espaces (y compris NBSP)
    # nettoyer espaces (y compris NBSP et espace fine insécable U+202F)
    s = s.replace("\u00A0", "").replace("\u202F", "").replace(" ", "")

    # gère pourcentage (ex: "1,1%")
    if "%" in s:
        s2 = re.sub(r"[^\d,.\-]", "", s)
        s2 = s2.replace(",", ".")
        try:
            return float(s2) / 100.0
        except:
            return None

    # enlever lettres / currency
    s = re.sub(r"[^\d\-,.()]", "", s)

    # parenthèses -> négatif
    if "(" in s and ")" in s:
        s = s.replace("(", "-").replace(")", "")

    # garder seulement signes -, digits, ., ,
    s = re.sub(r"[^0-9\-,.]", "", s)

    # cas avec à la fois , et . -> heuristique
    if "," in s and "." in s:
        # si la première occurence est ',' avant '.', on suppose que ',' est séparateur de milliers
        if s.find(',') < s.find('.'):
            s = s.replace(',', '')
        else:
            # sinon on suppose que '.' est milliers et ',' est décimal
            s = s.replace('.', '').replace(',', '.')
    else:
        # si plusieurs ',' -> ce sont des milliers
        if s.count(',') > 1:
            s = s.replace(',', '')
        # si plusieurs '.' -> ce sont des milliers
        if s.count('.') > 1:
            s = s.replace('.', '')

        # si une seule ',' et aucune '.', on suppose virgule décimale européen
        if ',' in s and '.' not in s:
            s = s.replace(',', '.')

    # final parse
    try:
        return float(s)
    except:
        return None

def match_nom(nom_source, nom_dest):
    return bool(difflib.get_close_matches(nom_source, [nom_dest], n=1, cutoff=0.7))

def match_date(date_source, date_dest):
    return normalize(date_source) in normalize(date_dest) or normalize(date_dest) in normalize(date_source)
MOIS_MAP = {
    "janvier": "JANUARY",
    "février": "FEBRUARY",
    "mars": "MARCH",
    "avril": "APRIL",
    "mai": "MAY",
    "juin": "JUNE",
    "juillet": "JULY",
    "août": "AUGUST",
    "septembre": "SEPTEMBER",
    "octobre": "OCTOBER",
    "novembre": "NOVEMBER",
    "décembre": "DECEMBER"
}

def get_verified_id(verified_folders, mois_choisi):
    # Extraire le mois en français → convertir en anglais
    mois_fr = mois_choisi.split(".")[1].strip().lower()
    mois_en = MOIS_MAP.get(mois_fr, mois_fr).upper()

    # Construire le nom attendu
    target_name = f"VERIFIED TRAVEL EXPENSES {mois_en} 2025"

    for vf in verified_folders:
        if normalize(vf["name"]) == normalize(target_name):
            return vf["id"]
    return None


def find_verified_folder(mois_choisi):
    """Cherche récursivement le dossier VERIFIED TRAVEL EXPENSES {Mois}"""
    mois_norm = normalize(mois_choisi)
    mois_norm = re.sub(r"^\d+\.", "", mois_norm).strip()
    mois_en = MOIS_MAP.get(mois_norm, mois_norm)

    def recursive_search(folder_id, level=0):
        subfolders = list_subfolders(folder_id)
        st.write(f"📂 Vérification du dossier {folder_id} → {len(subfolders)} sous-dossiers trouvés")
        
        for f in subfolders:
            st.write("📂", "  " * level, f["name"])  # log
            name_norm = normalize(f["name"])

            if "verified" in name_norm and "travel" in name_norm and "expenses" in name_norm and mois_en in name_norm:
                return f["id"]

            result = recursive_search(f["id"], level+1)
            if result:
                return result
        return None


    return recursive_search(VERIFIED_ROOT_ID)


def list_employee_folders(verified_folder_id):
    """Retourne les sous-dossiers employés d’un dossier VERIFIED"""
    subfolders = list_subfolders(verified_folder_id)
    st.write(f"📂 {len(subfolders)} dossiers employés trouvés :", [sf["name"] for sf in subfolders])
    return subfolders

def find_verified_folder(mois_label):
    """
    Ex: '01. Janvier' -> doit retourner le dossier VERIFIED du mois précédent
    """
    # Extraire mois + année
    mois_num = int(mois_label.split(".")[0])
    year = datetime.now().year  # ou extraire autrement si tu stockes l'année
    month_name = calendar.month_name[mois_num].upper()

    # Calculer le mois précédent
    prev_month = mois_num - 1
    prev_year = year
    if prev_month == 0:
        prev_month = 12
        prev_year -= 1

    prev_month_name = calendar.month_name[prev_month].upper()

    # 🔎 Construire le libellé attendu
    target_name = f"VERIFIED TRAVEL EXPENSES {prev_month_name} {prev_year}"

    # Parcourir les sous-dossiers
    verified_folders = list_subfolders(VERIFIED_ROOT_ID)
    for vf in verified_folders:
        if normalize(target_name) in normalize(vf["name"]):
            return vf["id"]

    return None
def find_verified_for_month(mois_label, annee=2025):
    """
    Cherche dans TOUS les sous-dossiers de VERIFIED_ROOT_ID
    le dossier 'VERIFIED TRAVEL EXPENSES {MONTH} {YEAR}' correspondant.
    
    mois_label : ex. "01. Janvier"
    """
    # Extraire juste le mot "Janvier"
    mois_fr = mois_label.split(".")[1].strip().lower()
    mois_en = MOIS_MAP.get(mois_fr, mois_fr).upper()
    target_name = f"VERIFIED TRAVEL EXPENSES {mois_en} {annee}"
    # st.write("🔎 Recherche du dossier :", target_name)

    # Parcourir tous les sous-dossiers
    mois_folders = list_subfolders(VERIFIED_ROOT_ID)
    for mf in mois_folders:
        subfolders = list_subfolders(mf["id"])
        for sf in subfolders:
            if normalize(target_name) == normalize(sf["name"]):
                # st.success(f"✅ Trouvé dans {mf['name']} → {sf['name']}")
                return sf["id"]

    st.warning(f"⚠️ Aucun dossier nommé {target_name} trouvé")
    return None

def find_employee_folder(folders, employee_fullname):
    """Trouve le dossier employé avec la meilleure correspondance"""
    # Normaliser le nom complet de l'employé
    fullname_norm = normalize(employee_fullname)
    emp_parts = fullname_norm.split()
    best_match = None
    best_score = 0
    best_matching_parts = []

    for f in folders:
        folder_name_norm = normalize(f["name"])
        folder_parts = folder_name_norm.split()

        matching_parts = []
        score = 0
        
        # Compter les mots en commun (avec pondération)
        for emp_part in emp_parts:
            if any(emp_part in folder_part for folder_part in folder_parts):
                matching_parts.append(emp_part)
                # Pondérer: le premier mot (nom) compte double
                if emp_part == emp_parts[0]:
                    score += 2
                else:
                    score += 1
        
        # Bonus si l'ordre est respecté
        if len(matching_parts) >= 2:
            # Vérifier si l'ordre des mots correspondants est le même
            emp_matching_indices = [emp_parts.index(part) for part in matching_parts]
            if emp_matching_indices == sorted(emp_matching_indices):
                score += 1
        
        # st.write(f"   📁 '{f['name']}' → score: {score}, mots correspondants: {matching_parts}")

        # Mettre à jour la meilleure correspondance
        if score > best_score or (score == best_score and len(matching_parts) > len(best_matching_parts)):
            best_score = score
            best_match = f
            best_matching_parts = matching_parts

    # Seuil minimum pour accepter la correspondance
    if best_match and best_score >= 2:  # Au moins le nom + un autre mot ou le nom avec bonus
        # st.success(f"✅ Meilleure correspondance: '{best_match['name']}' (score: {best_score}, mots: {best_matching_parts})")
        return best_match
    elif best_match:
        # st.warning(f"⚠️ Correspondance faible: '{best_match['name']}' (score: {best_score}, mots: {best_matching_parts})")
        return None
    else:
        # st.warning(f"⚠️ Aucun dossier trouvé pour '{employee_fullname}'")
        return None


def debug_list_files_and_folders(folder_id):
    """Affiche tous les fichiers/sous-dossiers du dossier employé"""
    from googleapiclient.discovery import build
    drive = build("drive", "v3", credentials=creds)

    results = drive.files().list(
        q=f"'{folder_id}' in parents",
        fields="files(id, name, mimeType)"
    ).execute()
    items = results.get("files", [])

    if not items:
        st.warning("⚠️ Aucun fichier/sous-dossier trouvé dans ce dossier")
    

    return items
patterns = [
    r"travel\s*expense",     # Travel Expense ou Travel Expense:
    r"travel\s*expences",    # Travel expences
    r"expenses\s*in\s*dzd"   # Expenses in DZD 
    r"total\s*expense\s"     #Total expense
]

def matches_pattern(text):
    text = text.strip().lower()  # on nettoie les espaces et on met en minuscule
    for pat in patterns:
        if re.search(pat, text):
            return True
    return False

def download_xlsx(file_id):
    request = drive_service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    fh.seek(0)
    wb = openpyxl.load_workbook(fh, data_only=True)
    ws = wb.active
    rows = []
    for r in ws.iter_rows():
        row_cells = []
        for cell in r:
            row_cells.append({
                "value": cell.value,
                "number_format": getattr(cell, "number_format", None)
            })
        rows.append(row_cells)
    return rows
def find_and_sum_verified_amounts(emp_folder, employe):
    """Trouve tous les fichiers Travel Expense et somme leurs montants vérifiés"""
    items = debug_list_files_and_folders(emp_folder["id"])
    total_montant_verified = 0
    fichiers_trouves = []
    
    # Première passe : chercher seulement les fichiers Travel Expense
    for item in items:
        if item["mimeType"] not in ["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                                   "application/vnd.google-apps.spreadsheet"]:
            continue
            
        # Vérifier si c'est un fichier Travel Expense
        item_name_norm = normalize(item["name"])
        if not ("travel" in item_name_norm and "expense" in item_name_norm):
            continue
            
        st.info(f"📂 Analyse Travel Expense : {item['name']}")
        
        try:
            if item["mimeType"] == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
                vrows = download_xlsx(item["id"])
            else:
                gvals = client.open_by_key(item["id"]).sheet1.get_all_values()
                vrows = []
                for row in gvals:
                    vrows.append([{"value": cell, "number_format": None} for cell in row])
            
            montant_trouve = extract_montant_from_file(vrows, item["name"])
            
            if montant_trouve is not None:
                total_montant_verified += montant_trouve
                fichiers_trouves.append({
                    "nom": item["name"],
                    "montant": montant_trouve
                })
                st.success(f"   ✅ Montant trouvé : {montant_trouve}")
            else:
                st.warning(f"   ⚠️ Aucun montant trouvé dans ce fichier")
                
        except Exception as e:
            st.error(f"   ❌ Erreur lecture {item['name']} : {e}")
            continue
    
    # Si aucun Travel Expense trouvé OU aucun montant dans les Travel Expense
    # → Deuxième passe : vérifier TOUS les fichiers spreadsheet
    if not fichiers_trouves:
        # st.info("🔍 Aucun Travel Expense trouvé → vérification de tous les fichiers spreadsheet")
        
        for item in items:
            if item["mimeType"] not in ["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                                       "application/vnd.google-apps.spreadsheet"]:
                continue
                
            # Sauter les fichiers déjà vérifiés (Travel Expense)
            item_name_norm = normalize(item["name"])
            if "travel" in item_name_norm and "expense" in item_name_norm:
                continue
                
            st.info(f"📂 Analyse fichier : {item['name']}")
            
            try:
                if item["mimeType"] == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
                    vrows = download_xlsx(item["id"])
                else:
                    gvals = client.open_by_key(item["id"]).sheet1.get_all_values()
                    vrows = []
                    for row in gvals:
                        vrows.append([{"value": cell, "number_format": None} for cell in row])
                
                montant_trouve = extract_montant_from_file(vrows, item["name"])
                
                if montant_trouve is not None:
                    total_montant_verified += montant_trouve
                    fichiers_trouves.append({
                        "nom": item["name"],
                        "montant": montant_trouve
                    })
                    st.success(f"   ✅ Montant trouvé : {montant_trouve}")
                else:
                    st.warning(f"   ⚠️ Aucun montant trouvé dans ce fichier")
                    
            except Exception as e:
                st.error(f"   ❌ Erreur lecture {item['name']} : {e}")
                continue
    
    return total_montant_verified, fichiers_trouves

def extract_montant_from_file(rows, filename):
    """Extrait le montant d'un fichier avec la même logique qu'avant"""
    for r_index, row in enumerate(rows):
        for c_index, cell in enumerate(row):
            raw = cell["value"] if isinstance(cell, dict) else cell
            texte = normalize(str(raw)) if raw is not None else ""
            
            if matches_pattern(texte):
                found = False
                montant_trouve = None
                
                # 1) Chercher dans les colonnes suivantes
                for offset in range(1, 4):
                    if c_index + offset < len(row):
                        candidate = row[c_index + offset]["value"] if isinstance(row[c_index + offset], dict) else row[c_index + offset]
                        parsed = to_float(candidate)
                        if parsed is not None:
                            montant_trouve = parsed
                            found = True
                            break
                
                # 2) Si pas trouvé, chercher sur la ligne suivante
                if not found and r_index + 1 < len(rows):
                    next_row = rows[r_index + 1]
                    if c_index < len(next_row):
                        candidate = next_row[c_index]["value"]
                        parsed = to_float(candidate)
                        if parsed is not None:
                            montant_trouve = parsed
                            found = True
                    
                    if not found:
                        for cc, next_cell in enumerate(next_row):
                            candidate = next_cell["value"]
                            parsed = to_float(candidate)
                            if parsed is not None:
                                montant_trouve = parsed
                                found = True
                                break
                
                if found:
                    return montant_trouve
    return None

@st.cache_data
def charger_siemens(sheet_id: str, worksheet_name: str):
    sh = client.open_by_key(sheet_id).worksheet(worksheet_name)
    return sh.get_all_values()
def match_nom_employe(employe, texte_cellule):
    """
    Vérifie si le nom de l'employé est contenu dans la cellule,
    même si la cellule contient d'autres infos (ex: job title).
    """
    emp_norm = normalize(employe)
    cell_norm = normalize(texte_cellule)
    return emp_norm in cell_norm
def trouver_dossier_client(root_siemens_id, client_choisi):
    """Trouver le dossier spécifique du client dans le dossier racine avec correspondance exacte"""
    try:
        # Lister tous les dossiers dans le root
        dossiers_clients = list_subfolders(root_siemens_id)
        
        # st.write(f"🔍 Recherche du dossier pour '{client_choisi}' parmi {len(dossiers_clients)} dossiers...")
        
        # Afficher tous les dossiers disponibles pour debug
        # st.write(f"📂 Dossiers disponibles : {[d['name'] for d in dossiers_clients]}")
        
        # Chercher le dossier qui correspond exactement au client choisi
        for dossier in dossiers_clients:
            nom_dossier = dossier['name'].strip()
            nom_client = client_choisi.strip()
            
            # Vérifier si le nom du dossier correspond exactement au client choisi
            if nom_dossier.lower() == nom_client.lower():
                # st.success(f"✅ Dossier client trouvé : {dossier['name']} (correspondance exacte)")
                return dossier['id']
        
        # Si aucun dossier trouvé avec correspondance exacte, chercher une correspondance partielle
        st.warning(f"⚠️ Aucune correspondance exacte pour '{client_choisi}', recherche partielle...")
        for dossier in dossiers_clients:
            nom_dossier = dossier['name'].lower()
            nom_client = client_choisi.lower()
            
            # Vérifier si le nom du client est contenu dans le nom du dossier
            if nom_client in nom_dossier:
                st.warning(f"⚠️ Correspondance partielle trouvée : {dossier['name']}")
                return dossier['id']
        
        # Si aucun dossier trouvé
        st.error(f"❌ Aucun dossier trouvé pour le client '{client_choisi}'")
        st.info(f"📂 Dossiers disponibles : {[d['name'] for d in dossiers_clients]}")
        return None
        
    except Exception as e:
        st.error(f"❌ Erreur lors de la recherche du dossier client : {e}")
        return None
def traiter_ndf_siemens_optimise(root_siemens_id, client_choisi, sheet_siemens, dest_sheet):
    """
    Version avec debug complet et correction des types
    """
    st.info(f"🔍 Recherche du dossier client '{client_choisi}'...")
    dossier_client_id = trouver_dossier_client(root_siemens_id, client_choisi)
    if not dossier_client_id:
        st.error("❌ Impossible de trouver le dossier client. Arrêt du traitement.")
        return
    # 📂 Récupérer tous les fichiers
    st.info("📂 Recherche des fichiers NDF...")
    mois_folders = list_subfolders(dossier_client_id)
    ndf_files = []
    
    for mois in mois_folders:
        fichiers = list_sheets_in_folder(mois["id"])
        ndf_files.extend(fichiers)
        time.sleep(1)
    
    st.write(f"📂 {len(ndf_files)} fichiers trouvés dans {client_choisi}")


    # === LECTURE DES FICHIERS NDF ===
    lignes_ndf = []
    fichiers_lus = 0

    for i, file in enumerate(ndf_files):
        try:
            # ⏳ Gestion des quotas - Délais stratégiques
            if fichiers_lus > 0:
                if fichiers_lus % 3 == 0:
                    # Pause plus longue après 3 fichiers
                    wait_time = random.uniform(8, 12)  # 8-12 secondes
                    st.warning(f"⏳ Pause de {wait_time:.1f}s pour éviter les quotas API...")
                    time.sleep(wait_time)
                else:
                    # Pause courte entre les fichiers
                    time.sleep(random.uniform(3, 5))  # 3-5 secondes
            
            # === Lecture du fichier source ===
            source_sheet = client.open_by_key(file["id"]).sheet1
            values = source_sheet.get_all_values()

            prenom = values[9][1] if len(values) > 9 and len(values[9]) > 1 else ""
            nom = values[9][2] if len(values) > 9 and len(values[9]) > 2 else ""
            employe = f"{prenom} {nom}".strip()
            date = values[4][2] if len(values) > 4 and len(values[4]) > 2 else ""
            ref = values[5][2] if len(values) > 5 and len(values[5]) > 2 else ""

            # === 🔍 Chercher le montant brut dans la source ===
            target = "montant a rembourser"
            montant = None
            for r_index, row in enumerate(values):
                if len(row) > 4:
                    texte = normalize(row[4])
                    match = difflib.get_close_matches(texte, [target], n=1, cutoff=0.6)
                    if match:
                        montant_str = row[6] if len(row) > 6 else None
                        if montant_str:
                            # Convertir le montant en float
                            montant = to_float(montant_str)
                        st.write(f"Employé: {employe}")
                        st.success(f"✅ Montant trouvé dans la NDF :  {montant}")
                        break

            if not montant:
                st.error(f"❌ Impossible de trouver 'Montant à rembourser' dans {file['name']}")
                continue

            # Vérification que toutes les données sont présentes
            if employe and date and montant is not None and ref:
                lignes_ndf.append((employe, date, ref, montant, file['id']))
                st.success(f"✅ {employe} ({date}) [Ref: {ref}] → {montant} DZD")
            else:
                st.warning(f"⚠️ Données manquantes dans {file['name']} - Employé: '{employe}', Date: '{date}', Ref: '{ref}', Montant: '{montant}'")

        except Exception as e:
            if "429" in str(e) or "quota" in str(e).lower():
                st.error(f"🚨 QUOTA API ATTEINT pour {file['name']}")
                st.info("🔄 Attente de 60 secondes avant de continuer...")
                time.sleep(60)
                
                # Réessayer une fois après l'attente
                try:
                    source_sheet = client.open_by_key(file["id"]).sheet1
                    values = source_sheet.get_all_values()

                    prenom = values[9][1] if len(values) > 9 and len(values[9]) > 1 else ""
                    nom = values[9][2] if len(values) > 9 and len(values[9]) > 2 else ""
                    employe = f"{prenom} {nom}".strip()
                    date = values[4][2] if len(values) > 4 and len(values[4]) > 2 else ""
                    ref = values[5][2] if len(values) > 5 and len(values[5]) > 2 else ""

                    # === 🔍 Chercher le montant brut dans la source ===
                    target = "montant a rembourser"
                    montant = None
                    for r_index, row in enumerate(values):
                        if len(row) > 4:
                            texte = normalize(row[4])
                            match = difflib.get_close_matches(texte, [target], n=1, cutoff=0.6)
                            if match:
                                montant_str = row[6] if len(row) > 6 else None
                                if montant_str:
                                    montant = to_float(montant_str)
                                break

                    if not montant:
                        st.error(f"❌ Impossible de trouver 'Montant à rembourser' dans {file['name']}")
                        continue

                    if employe and date and montant is not None and ref:
                        lignes_ndf.append((employe, date, ref, montant, file['id']))
                        st.success(f"✅ {employe} ({date}) [Ref: {ref}] → {montant} DZD (après réessai)")
                    
                except Exception as e2:
                    st.error(f"❌ Échec définitif pour {file['name']} après réessai : {e2}")
            else:
                st.error(f"Erreur lecture NDF - {file['name']} : {e}")

    # Afficher le résumé
    st.write(f"📊 {len(lignes_ndf)} fichiers NDF traités avec succès sur {len(ndf_files)}")

    # === DÉDUPLICATION ===
    lignes_ndf_uniques = []
    seen = set()
    
    for ligne in lignes_ndf:
        employe, date, ref, montant, file_id = ligne
        cle_unique = (employe, date, ref, montant)
        
        if cle_unique not in seen:
            seen.add(cle_unique)
            lignes_ndf_uniques.append(ligne)
    # === LECTURE DES DONNÉES EXISTANTES ===
    st.info("📖 Lecture des données existantes...")
    
    try:
        donnees_siemens = sheet_siemens.get_all_values()
        donnees_dest = dest_sheet.get_all_values()
        
        st.success(f"✅ Données Siemens : {len(donnees_siemens)} lignes")
        st.success(f"✅ Données Global : {len(donnees_dest)} lignes")
        
    except Exception as e:
        st.error(f"❌ Erreur lecture données : {e}")
        return

    # === TRAITEMENT DES DONNÉES ===
    maj_siemens = []
    maj_global = []
    nouvelles_lignes = []
    lignes_global_traitees = set()

    # 1. Regroupement Siemens - CORRECTION DES TYPES
    regroupement_siemens = {}
    for employe, date, ref, montant, file_id in lignes_ndf_uniques:
        key = (employe, date)
        # S'assurer que montant est un nombre
        montant_float = float(montant) if not isinstance(montant, (int, float)) else montant
        regroupement_siemens[key] = regroupement_siemens.get(key, 0.0) + montant_float

    # 2. Mise à jour Siemens
    # st.info("🔍 Recherche des correspondances Siemens...")
    for (employe, date), montant_total in regroupement_siemens.items():
        for i, row in enumerate(donnees_siemens):
            if i >= len(donnees_siemens):
                break
                
            col_nom = row[5] if len(row) > 5 else ""
            col_date = row[12] if len(row) > 12 else ""

            if match_nom_employe(employe, col_nom) and match_date(normalize(date), normalize(col_date)):
                maj_siemens.append((i+1, float(montant_total)))  # Convertir en float
                # st.info(f"📌 Match Siemens : {employe} | {date} → {montant_total}")
                break

    for idx, (employe, date, ref, montant, file_id) in enumerate(lignes_ndf_uniques):
        cle_globale = (employe, date, ref)
   
        
        if cle_globale in lignes_global_traitees:
            st.warning("⚠️ Ligne déjà traitée - SKIP")
            continue
            
        found = False
        
        # Recherche dans Global avec debug détaillé
        for j, row in enumerate(donnees_dest):
            if j >= len(donnees_dest):
                break
                
            col_nom = row[5] if len(row) > 5 else ""
            col_date = row[2] if len(row) > 2 else ""

            # Debug du matching
            match_nom = match_nom_employe(employe, col_nom)
            match_dates = match_date(normalize(date), normalize(col_date))

            if match_nom and match_dates:
                # S'assurer que le montant est un nombre
                montant_float = float(montant) if not isinstance(montant, (int, float)) else montant
                maj_global.append((j+1, montant_float))
                lignes_global_traitees.add(cle_globale)
                found = True
                break

        # Nouvelle ligne si pas trouvé
        if not found:
            st.success("➕ **AUCUN MATCH TROUVÉ** → Création nouvelle ligne")
            
            # Récupérer période
            periode = ""
            if file_id:
                try:
                    source_sheet = client.open_by_key(file_id).sheet1
                    values = source_sheet.get_all_values()
                    if len(values) > 9 and len(values[9]) > 5:
                        periode = f"{values[9][4]} {values[9][5]}".strip()
                except:
                    pass
            
            # Créer la nouvelle ligne
            next_id = len(donnees_dest) + len(nouvelles_lignes) 
            ref_new = f"N°{next_id}/{client_choice}/{type_choice}/2025"

            # Convertir le montant en string pour l'insertion
            montant_str = str(montant) if not isinstance(montant, str) else montant

            nouvelle_ligne = [
                str(next_id),
                ref_new,
                date,
                type_choice,
                client_choice,
                employe,
                periode,
                montant_str,  # Utiliser la version string
                statut_choice,
                facturation_choice,
                "",
                commentaire
            ]
            
            nouvelles_lignes.append(nouvelle_ligne)
            lignes_global_traitees.add(cle_globale)

    
    if nouvelles_lignes:
        st.write("🎯 **NOUVELLES LIGNES PRÉPARÉES :**")
        for i, nl in enumerate(nouvelles_lignes):
            st.write(f"{i+1}. ID:{nl[0]} | Ref:{nl[1]} | Date:{nl[2]} | Employé:{nl[5]} | Montant:{nl[7]}")

    # === APPLICATIONS DES MISES À JOUR ===
    # Mise à jour Siemens
    if maj_siemens:
        try:
            appliquer_maj_siemens(sheet_siemens, maj_siemens, len(donnees_siemens))
            st.success(f"✅ {len(maj_siemens)} lignes Siemens mises à jour")
        except Exception as e:
            st.error(f"❌ Erreur mise à jour Siemens : {e}")

    # Mise à jour Global
    if maj_global:
        try:
            appliquer_maj_global(dest_sheet, maj_global, len(donnees_dest))
            st.success(f"✅ {len(maj_global)} lignes Global mises à jour")
        except Exception as e:
            st.error(f"❌ Erreur mise à jour Global : {e}")

    # NOUVELLES LIGNES - FORCER L'AJOUT MÊME EN CAS D'ERREUR
    if nouvelles_lignes:
        st.info("➕ AJOUT DES NOUVELLES LIGNES...")
        succes_ajout = 0
        
        for i, ligne in enumerate(nouvelles_lignes):
            try:
                st.write(f"🔄 Ajout ligne {i+1}/{len(nouvelles_lignes)}: {ligne[5]} | {ligne[2]} | {ligne[7]}")
                dest_sheet.append_row(ligne)
                succes_ajout += 1
                
                # ✅ APPLIQUER LE STYLE ET LES VALIDATIONS SUR LA NOUVELLE LIGNE
                derniere_ligne = len(dest_sheet.get_all_values())
                appliquer_style_ligne(dest_sheet, derniere_ligne)
                appliquer_validations_donnees(dest_sheet, derniere_ligne)
                
                # st.success(f"✅ Ligne {i+1} ajoutée avec succès")
                time.sleep(1)
                # ✏️ Écriture de la référence dans le fichier source (NDF)
                try:
                    source_sheet.update_cell(6, 3, ref)  # ligne 6, colonne C (3)
                    st.info(f"🔗 Référence '{ref}' écrite dans la NDF ({file['name']}) → C6")
                except Exception as e:
                    st.error(f"⚠️ Impossible d’écrire la référence dans {file['name']} : {e}")

            except Exception as e:
                st.error(f"❌ Erreur ajout ligne {i+1}: {e}")
        
        st.success(f"🎉 {succes_ajout}/{len(nouvelles_lignes)} nouvelles lignes ajoutées avec succès !")
    else:
        st.warning("⚠️ Aucune nouvelle ligne à ajouter")

# AJOUTER LES FONCTIONS DE STYLE ET VALIDATIONS SI ELLES N'EXISTENT PAS DÉJÀ
def appliquer_style_ligne(dest_sheet, ligne_num):
    """Applique le style sur une ligne spécifique"""
    try:
        fmt = CellFormat(
            textFormat=TextFormat(fontFamily="Baloo 2", bold=False),
            borders=Borders(
                top=Border(style="SOLID", color=Color(0, 0, 0)),
                bottom=Border(style="SOLID", color=Color(0, 0, 0)),
                left=Border(style="SOLID", color=Color(0, 0, 0)),
                right=Border(style="SOLID", color=Color(0, 0, 0)),
            )
        )
        format_cell_range(dest_sheet, f"A{ligne_num}:L{ligne_num}", fmt)
        st.success(f"🎨 Style appliqué sur la ligne {ligne_num}")
    except Exception as e:
        st.error(f"❌ Erreur application du style ligne {ligne_num}: {e}")

def appliquer_validations_donnees(dest_sheet, ligne_num):
    """Applique les validations de données sur une ligne spécifique"""
    try:
        # Validation client
        rule_client = DataValidationRule(
            BooleanCondition('ONE_OF_LIST', [
                'G+D', 'Epson', 'PMI', 'Siemens', 'Syngenta',
                'OS-Team', 'HH-Team', 'Cahros', 'Siemens Energy', 'Abott'
            ]),
            showCustomUi=True
        )
        set_data_validation_for_cell_range(dest_sheet, f"E{ligne_num}:E{ligne_num}", rule_client)

        # Validation statut
        rule_statut = DataValidationRule(
            BooleanCondition('ONE_OF_LIST', ['Payé', 'Non payé']),
            showCustomUi=True
        )
        set_data_validation_for_cell_range(dest_sheet, f"I{ligne_num}:I{ligne_num}", rule_statut)
        
        # Validation facturation
        rule_facturation = DataValidationRule(
            BooleanCondition('ONE_OF_LIST', [
                'Facturation Odoo', 'Facturation Note de débours',
                'Facturation Odoo + Note de débours', 'Sans facture'
            ]),
            showCustomUi=True
        )
        set_data_validation_for_cell_range(dest_sheet, f"J{ligne_num}:J{ligne_num}", rule_facturation)
        
        # Validation type
        rule_type = DataValidationRule(
            BooleanCondition('ONE_OF_LIST', ['NDF', 'FDM', 'FD']),
            showCustomUi=True
        )
        set_data_validation_for_cell_range(dest_sheet, f"D{ligne_num}:D{ligne_num}", rule_type)
        
        st.success(f"✅ Validations appliquées sur la ligne {ligne_num}")
    except Exception as e:
        st.error(f"❌ Erreur application des validations ligne {ligne_num}: {e}")
def get_verified_amount_from_sheet(sheet, employe, mois_choisi):
    """
    🔍 Recherche le montant correspondant à un employé et un mois dans la feuille 'Travel expenses'.
    """
    values = sheet.get_all_values()

    # --- Section Travel expenses ---
    start_idx = next(
        (i for i, row in enumerate(values) if any("travel expenses" in c.lower() for c in row if c)), None
    )
    if start_idx is None:
        raise ValueError("Section 'Travel expenses' introuvable dans la feuille")

    end_idx = next(
        (i for i, row in enumerate(values[start_idx + 1:], start=start_idx + 1)
         if any("allowance" in c.lower() for c in row if c)),
        len(values)
    )

    table = values[start_idx:end_idx]
    if not table or len(table) < 2:
        raise ValueError("Section 'Travel expenses' vide ou mal délimitée")

    # --- En-tête ---
    header_row_idx = next(
        (i for i, row in enumerate(table) if any("name" in c.strip().lower() for c in row if c)), None
    )
    if header_row_idx is None:
        raise ValueError("Colonne 'Name' introuvable dans la section Travel expenses")

    header = table[header_row_idx]
    data_rows = table[header_row_idx + 1:]
    name_col = next((i for i, c in enumerate(header) if "name" in c.strip().lower()), None)

    # --- Normalisation du mois ---
    mois_map = {
    "décembre": (4, 5),
    "decembre": (4, 5),
    "janvier": (6, 7),
    "février": (8, 9),
    "fevrier": (8, 9),
    "mars": (10, 11),
    "avril": (12, 13),
    "mai": (14, 15),
    "juin": (16, 17),
    "juillet": (18, 19),
    "août": (20, 21),
    "aout": (20, 21),
    "septembre": (22, 23),
    "octobre": (24, 25),
    "novembre": (26, 27),
}


    # 🧠 Nouvelle logique pour extraire le mot du mois proprement
    mois_key = normalize(mois_choisi).lower()
    import re
    mois_key = re.sub(r"[^a-zàâçéèêëîïôûùüÿñæœ\s-]", "", mois_key)  # enlève chiffres et ponctuation
    mots = mois_key.split()
    mois_key = next((m for m in mots if m in mois_map.keys()), None)

    if not mois_key:
        raise ValueError(f"Mois '{mois_choisi}' non reconnu pour la vérification (valeur nettoyée: '{mois_key}')")

    col1, col2 = mois_map[mois_key]

    # --- Recherche de l'employé ---
    for row in data_rows:
        if len(row) > name_col and match_nom(normalize(employe), normalize(row[name_col])):
            
            raw1 = row[col1] if len(row) > col1 else ""
            raw2 = row[col2] if len(row) > col2 else ""
            montant1 = to_float(raw1)
            montant2 = to_float(raw2)
            st.write(f"[DEBUG] {employe} : col1={col1}, val1='{raw1}' -> {montant1} | col2={col2}, val2='{raw2}' -> {montant2}")
            total = (montant1 or 0) + (montant2 or 0)
            st.write(f"[DEBUG] Mois choisi : {mois_key} → colonnes {col1} (M) et {col2} (C)")

            return total

    return None

def  traiter_fichiers_ndf_G_D(mois_id, mois_choisi, client_choice, type_choice, statut_choice, facturation_choice, commentaire, dest_sheet, verified_sheet, annee=2025):
    """
    Traite tous les fichiers NDF d'un mois donné et effectue les vérifications
    """
    fichiers = list_sheets_in_folder(mois_id)
    st.write(f"📂 {len(fichiers)} fichiers trouvés dans {mois_choisi}")
    
    for file in fichiers:
        try:
            # === Lecture du fichier source ===
            source_sheet = client.open_by_key(file["id"]).sheet1
            values = source_sheet.get_all_values()

            prenom = values[9][1] if len(values) > 9 and len(values[9]) > 1 else ""
            nom = values[9][2] if len(values) > 9 and len(values[9]) > 2 else ""
            employe = f"{prenom} {nom}".strip()
            date = values[4][2] if len(values) > 4 and len(values[4]) > 2 else ""

            # === 🔍 Chercher le montant brut dans la source ===
            target = "montant a rembourser"
            montant = None
            for r_index, row in enumerate(values):
                if len(row) > 4:
                    texte = normalize(row[4])
                    match = difflib.get_close_matches(texte, [target], n=1, cutoff=0.6)
                    if match:
                        montant = row[6] if len(row) > 6 else None
                        st.write(f"Employé: {employe}")
                        st.success(f"✅ Montant trouvé dans la NDF : '{montant}'")
                        break

            if not montant:
                st.error(f"❌ Impossible de trouver 'Montant à rembourser' dans {file['name']}")
                continue

            montant_brut = to_float(montant)

            # === Vérification dans VERIFIED ===
            total_montant_verified = get_verified_amount_from_sheet(verified_sheet, employe, mois_choisi)

            if total_montant_verified is None:
                st.warning(f"⚠️ Aucun montant VERIFIED trouvé pour {employe}")
                continue

            # === Comparaison ===
            st.write("---")
            st.subheader(f"📊 {employe}")
            st.write(f"Montant NDF : {montant_brut} | Montant VERIFIED : {total_montant_verified}")

            if abs(montant_brut - total_montant_verified) < 0.01:
                dest_values = dest_sheet.get_all_values()
                updated = False

                for i, row in enumerate(dest_values):
                    if len(row) >= 6:
                        nom_dest = normalize(row[5])
                        date_dest = normalize(row[2])

                        if match_nom(normalize(employe), nom_dest) and match_date(normalize(date), date_dest):
                            # ✅ Mise à jour du montant VERIFIED (colonne H = index 7)
                            dest_sheet.update_cell(i + 1, 8, str(total_montant_verified))
                            updated = True
                            break

                if updated:
                    st.success(f"✅ {file['name']} → {employe} mis à jour dans la feuille DEST avec montant VERIFIED = {total_montant_verified}")
                else:
                    st.warning(f"⚠️ Aucun matching trouvé dans DEST → ajout d'une nouvelle ligne pour {employe}")

                    # 🕐 Essayer de récupérer la période
                    periode = ""
                    if len(values) > 9 and len(values[9]) > 5:
                        periode = f"{values[9][4]} {values[9][5]}".strip()
                    if not periode:
                        periode = mois_choisi

                    # 🆕 Préparer la nouvelle ligne
                    next_id = len(dest_values)
                    ref = f"N°{next_id}/{client_choice}/{type_choice}/{annee}"

                    nouvelle_ligne = [
                        str(next_id),          # A: ID
                        ref,                   # B: Référence
                        date,                  # C: Date
                        type_choice,           # D
                        client_choice,         # E
                        employe,               # F
                        periode,               # G
                        str(total_montant_verified),  # H: Montant
                        statut_choice,         # I
                        facturation_choice,    # J
                        "",                    # K (vide par défaut)
                        commentaire            # L
                    ]

                    # ➕ Ajout dans la feuille
                    dest_sheet.append_row(nouvelle_ligne)
                    st.success(f"➕ Nouvelle ligne ajoutée pour {employe} ({periode}) avec montant VERIFIED = {total_montant_verified}")

                    # 🔢 Dernière ligne ajoutée
                    last_row = len(dest_sheet.get_all_values())

                    # 🎨 Appliquer style et validation
                    appliquer_validations_donnees(dest_sheet, last_row)
                    appliquer_style_ligne(dest_sheet, last_row)
                    # ✏️ Écriture de la référence dans le fichier source (NDF)
                    try:
                        source_sheet.update_cell(6, 3, ref)  # ligne 6, colonne C (3)
                        st.info(f"🔗 Référence '{ref}' écrite dans la NDF ({file['name']}) → C6")
                    except Exception as e:
                        st.error(f"⚠️ Impossible d’écrire la référence dans {file['name']} : {e}")


            else:
                # 🔴 Cas NON CONCORDANT
                delta = round(abs(montant_brut - total_montant_verified), 2)
                commentaire_non_concordant = f"❌ NON CONCORDANT : NDF={montant_brut} / VERIFIED={total_montant_verified} / Δ={delta}"
                st.error(commentaire_non_concordant)

                # On ajoute quand même la ligne dans la feuille DEST
                dest_values = dest_sheet.get_all_values()
                next_id = len(dest_values)
                ref = f"N°{next_id}/{client_choice}/{type_choice}/{annee}"

                # 🕐 Essayer de récupérer la période
                periode = ""
                if len(values) > 9 and len(values[9]) > 5:
                    periode = f"{values[9][4]} {values[9][5]}".strip()
                if not periode:
                    periode = mois_choisi

                # 🆕 Nouvelle ligne avec le commentaire d'erreur
                nouvelle_ligne = [
                    str(next_id),                # A: ID
                    ref,                         # B: Référence
                    date,                        # C: Date
                    type_choice,                 # D
                    client_choice,               # E
                    employe,                     # F
                    periode,                     # G
                    str(total_montant_verified), # H: Montant VERIFIED
                    statut_choice,               # I
                    facturation_choice,          # J
                    "",                          # K (vide)
                    commentaire_non_concordant   # L: commentaire
                ]

                # ➕ Ajout de la ligne
                dest_sheet.append_row(nouvelle_ligne)
                st.warning(f"⚠️ Ligne ajoutée pour {employe} (non concordant)")

                # 🎨 Mise en forme jaune
                last_row = len(dest_sheet.get_all_values())
                appliquer_style_ligne(dest_sheet, last_row, couleur="JAUNE")

                # Validation et autres formats
                appliquer_validations_donnees(dest_sheet, last_row)
                # ✏️ Écriture de la référence dans le fichier source (NDF)
                try:
                    source_sheet.update_cell(6, 3, ref)  # ligne 6, colonne C (3)
                    st.info(f"🔗 Référence '{ref}' écrite dans la NDF ({file['name']}) → C6")
                except Exception as e:
                    st.error(f"⚠️ Impossible d’écrire la référence dans {file['name']} : {e}")


        except Exception as e:
            st.error(f"Erreur sur {file['name']} : {e}")

def appliquer_validations_donnees(dest_sheet, ligne_num):
    """Applique les validations de données sur une ligne spécifique"""
    # Validation client
    rule_client = DataValidationRule(
        BooleanCondition('ONE_OF_LIST', [
            'G+D', 'Epson', 'PMI', 'Siemens', 'Syngenta',
            'OS-Team', 'HH-Team', 'Cahros', 'Siemens Energy', 'Abott'
        ]),
        showCustomUi=True
    )
    set_data_validation_for_cell_range(dest_sheet, f"E{ligne_num}:E{ligne_num}", rule_client)

    # Validation statut
    rule_statut = DataValidationRule(
        BooleanCondition('ONE_OF_LIST', ['Payé', 'Non payé']),
        showCustomUi=True
    )
    set_data_validation_for_cell_range(dest_sheet, f"I{ligne_num}:I{ligne_num}", rule_statut)
    
    # Validation facturation
    rule_facturation = DataValidationRule(
        BooleanCondition('ONE_OF_LIST', [
            'Facturation Odoo', 'Facturation Note de débours',
            'Facturation Odoo + Note de débours', 'Sans facture'
        ]),
        showCustomUi=True
    )
    set_data_validation_for_cell_range(dest_sheet, f"J{ligne_num}:J{ligne_num}", rule_facturation)
    
    # Validation type
    rule_type = DataValidationRule(
        BooleanCondition('ONE_OF_LIST', ['NDF', 'FDM', 'FD']),
        showCustomUi=True
    )
    set_data_validation_for_cell_range(dest_sheet, f"D{ligne_num}:D{ligne_num}", rule_type)

def appliquer_style_ligne(dest_sheet, ligne_num, couleur="BLANC"):
    """
    Applique un style sur une ligne spécifique.
    couleur : "VERT" (par défaut) ou "JAUNE" pour les cas non concordants.
    """
    # 🎨 Choix de la couleur de fond
    if couleur.upper() == "JAUNE":
        bg_color = Color(1, 1, 0.6)   # Jaune clair
    else:
        bg_color = Color(1, 1, 1)  # Vert clair (par défaut)

    # ✏️ Définition du format
    fmt = CellFormat(
        backgroundColor=bg_color,
        textFormat=TextFormat(fontFamily="Baloo 2", bold=False),
        borders=Borders(
            top=Border(style="SOLID", color=Color(0, 0, 0)),
            bottom=Border(style="SOLID", color=Color(0, 0, 0)),
            left=Border(style="SOLID", color=Color(0, 0, 0)),
            right=Border(style="SOLID", color=Color(0, 0, 0)),
        )
    )

    # 🧾 Application du style sur la ligne complète (A à L)
    format_cell_range(dest_sheet, f"A{ligne_num}:L{ligne_num}", fmt)
# Fonctions auxiliaires
def appliquer_maj_siemens(worksheet, mises_a_jour, nb_lignes):
    if not mises_a_jour:
        return
    
    updates = {}
    for ligne, valeur in mises_a_jour:
        updates[ligne] = valeur
    
    try:
        cell_list = worksheet.range(f"AY1:AY{nb_lignes}")
        for i, cell in enumerate(cell_list):
            ligne_num = i + 1
            if ligne_num in updates:
                cell.value = updates[ligne_num]
        
        worksheet.update_cells(cell_list, value_input_option="USER_ENTERED")
    except Exception as e:
        st.error(f"❌ Erreur mise à jour Siemens : {e}")

def appliquer_maj_global(worksheet, mises_a_jour, nb_lignes):
    if not mises_a_jour:
        return
    
    updates = {}
    for ligne, valeur in mises_a_jour:
        updates[ligne] = valeur
    
    try:
        cell_list = worksheet.range(f"H1:H{nb_lignes}")
        for i, cell in enumerate(cell_list):
            ligne_num = i + 1
            if ligne_num in updates:
                cell.value = updates[ligne_num]
        
        worksheet.update_cells(cell_list, value_input_option="USER_ENTERED")
    except Exception as e:
        st.error(f"❌ Erreur mise à jour Global : {e}")
# === Destination sheet ===
dest_sheet = client.open_by_key(DEST_SHEET_ID).sheet1
# 2️⃣ Lister les clients (sous-dossiers de NDF)
clients = list_subfolders(ndf_root_id)
client_names = [c["name"] for c in clients]
# client_choisi = st.selectbox("👔 Choisir un client", client_names)
client_choice = st.sidebar.selectbox("🏢 Choisir le client :", 
                                                ["G+D", "Epson", "PMI", "Siemens", "Syngenta", "OS-Team", "HH-Team", "Cahros", "Siemens Energy", "Abott"])

# 3️⃣ Trouver l’ID du dossier du client choisi
client_id = next((c["id"] for c in clients if c["name"] == client_choice), None)

# 4️⃣ Lister les mois dans ce client
mois_folders = list_subfolders(client_id)
mois_names = [m["name"] for m in mois_folders]
mois_choisi = st.selectbox("📅 Choisir le mois", mois_names)

# 5️⃣ Trouver l’ID du mois choisi
mois_id = next((m["id"] for m in mois_folders if m["name"] == mois_choisi), None)

# 6️⃣ Charger les fichiers du dossier mois
fichiers = list_sheets_in_folder(mois_id)
st.write(f"📂 {len(fichiers)} fichiers trouvés dans {client_choice} / {mois_choisi}")

# === Inputs utilisateur ===

facturation_choice = st.sidebar.selectbox("🧾 Choisir le type de facturation :", 
                                                    ["Facturation Odoo", "Facturation Note de débours", 
                                                    "Facturation Odoo + Note de débours", "Sans facture"])

statut_choice = st.sidebar.selectbox("💳 Statut de paiement :", ["Non payé", "Payé"])
type_choice = st.sidebar.selectbox("Type :", ["NDF", "FD", "FDM"])
commentaire = st.sidebar.text_input("📝 Commentaire :", "")

sheet_siemens = client.open_by_key("1ZI726DLcpqsho3ZVx-ofx825DcE1vSqaCn2FlT-cFcI").worksheet("Feuille 3")
sheet_global = client.open_by_key("1jxjAstmnsWCuRaYwVIhW-Qh7pZvh-waw3BEQ2HDGvRM").worksheet("2025")
root_id = "1KTRuCR59xLgKLCT1_AY3z-lgeh9JFmrb"

# === Étape 2 : transfert avec vérification ===
if st.button("🔄 Récupérer et transférer"):
    if client_choice == "Siemens" or client_choice == "Siemens Energy":
        try:
            traiter_ndf_siemens_optimise(
                root_siemens_id=root_id,
                client_choisi=client_choice,  # ⬅️ AJOUTER LE CLIENT CHOISI
                sheet_siemens=sheet_siemens,
                dest_sheet=sheet_global
            )
        except Exception as e:
            st.error(f"❌ Erreur globale du traitement : {e}")
            st.info("🔄 Le traitement a été interrompu. Vous pouvez réessayer dans quelques minutes.")

        st.success("🎉 Traitement terminé !")
    elif client_choice == "G+D":
        VERIFIED_SHEET_ID = "1Rv4zNx7Q9OxBxTnFGP1oRW47fZyfP7Oxdn25w0UM9EU"
        verified_sheet = client.open_by_key(VERIFIED_SHEET_ID).sheet1

        # Dans votre code principal, remplacez le bloc par :
        traiter_fichiers_ndf_G_D(
    mois_id=mois_id,
    mois_choisi=mois_choisi,
    client_choice=client_choice,
    type_choice=type_choice,
    statut_choice=statut_choice,
    facturation_choice=facturation_choice,
    commentaire=commentaire,
    dest_sheet=dest_sheet,
    verified_sheet=verified_sheet,  # 👈 ici !
    annee=2025
)

    else:
        fichiers = list_sheets_in_folder(mois_id)
        st.write(f"📂 {len(fichiers)} fichiers trouvés dans {mois_choisi}")
    
        for file in fichiers:
            try:
                # === Lecture du fichier source ===
                source_sheet = client.open_by_key(file["id"]).sheet1
                values = source_sheet.get_all_values()

                prenom = values[9][1] if len(values) > 9 and len(values[9]) > 1 else ""
                nom = values[9][2] if len(values) > 9 and len(values[9]) > 2 else ""
                employe = f"{prenom} {nom}".strip()
                date = values[4][2] if len(values) > 4 and len(values[4]) > 2 else ""

                # === 🔍 Chercher le montant brut dans la source ===
                target = "montant a rembourser"
                montant = None
                for r_index, row in enumerate(values):
                    if len(row) > 4:
                        texte = normalize(row[4])
                        match = difflib.get_close_matches(texte, [target], n=1, cutoff=0.6)
                        if match:
                            montant = row[6] if len(row) > 6 else None
                            st.write(f"Employé: {employe}")
                            st.success(f"✅ Montant trouvé dans la NDF : '{montant}'")
                            break

                if not montant:
                    st.error(f"❌ Impossible de trouver 'Montant à rembourser' dans {file['name']}")
                    continue

                montant_brut = to_float(montant)

                # === Vérification dans VERIFIED ===
                 # Tolérance pour les floats
                dest_values = dest_sheet.get_all_values()
                updated = False
                for i, row in enumerate(dest_values):
                        if len(row) >= 6:
                            nom_dest = normalize(row[5])
                            date_dest = normalize(row[2])

                            if match_nom(normalize(employe), nom_dest) and match_date(normalize(date), date_dest):
                                dest_sheet.update_cell(i+1, 8, montant)
                                updated = True
                                break

                if updated:
                       
                            st.success(f"✅ {file['name']} → {employe} : {montant_brut} (VERIFIED ok)")
                else:
                        # === Aucun matching trouvé → création d'une nouvelle ligne ===
                        st.warning(f"⚠️ Pas de correspondance dans destination pour {employe} ({file['name']}), création d'une nouvelle ligne")

                    

                        # === Récupération de la période depuis le fichier source (E10 + F10 fusionnées) ===
                        periode = ""
                        if len(values) > 9 and len(values[9]) > 5:
                            periode = f"{values[9][4]} {values[9][5]}".strip()  # colonne E et F ligne 10

                        # === Construction de la nouvelle ligne ===
                        next_id = len(dest_values)  # ID auto
                        ref = f"N°{next_id}/{client_choice}/{type_choice}/2025"  # Référence générée

                        nouvelle_ligne = [
                            str(next_id),                        
                            ref,                                 # Référence
                            date if date else "",                # Date source
                            type_choice,                             
                            client_choice,                       # Client choisi
                            employe,                             # Employé
                            periode if periode else mois_choisi, # Période (soit fichier source, soit mois choisi)
                            str(montant_brut),         # Montant TTC (vérifié)
                            statut_choice,                       # Statut payé / non payé
                            facturation_choice,                  # Facturation
                            "",                                  # N° débours (vide)
                            commentaire                          # Commentaire saisi
                        ]

                        # ➕ Ajout dans la feuille destination
                        dest_sheet.append_row(nouvelle_ligne)
                        st.success(f"➕ Nouvelle ligne ajoutée pour {employe} ({periode if periode else mois_choisi}) avec montant {montant_brut}")

                        # ✅ Ajouter la ligne
                    

                        # Récupérer l’index de la dernière ligne insérée
                        last_row = len(dest_sheet.get_all_values())

                        # ✅ Menus déroulants (uniquement sur la nouvelle ligne)
                        rule_client = DataValidationRule(
                            BooleanCondition('ONE_OF_LIST', [
                                'G+D', 'Epson', 'PMI', 'Siemens', 'Syngenta',
                                'OS-Team', 'HH-Team', 'Cahros', 'Siemens Energy', 'Abott'
                            ]),
                            showCustomUi=True
                        )
                        set_data_validation_for_cell_range(dest_sheet, f"E{last_row}:E{last_row}", rule_client)

                        rule_statut = DataValidationRule(
                            BooleanCondition('ONE_OF_LIST', ['Payé', 'Non payé']),
                            showCustomUi=True
                        )
                        set_data_validation_for_cell_range(dest_sheet, f"I{last_row}:I{last_row}", rule_statut)
                        
                        rule_facturation = DataValidationRule(
                            BooleanCondition('ONE_OF_LIST', [
                                'Facturation Odoo', 'Facturation Note de débours',
                                'Facturation Odoo + Note de débours', 'Sans facture'
                            ]),
                            showCustomUi=True
                        )
                        set_data_validation_for_cell_range(dest_sheet, f"J{last_row}:J{last_row}", rule_facturation)
                        rule_Type = DataValidationRule(
                            BooleanCondition('ONE_OF_LIST', ['NDF', 'FDM', 'FD']),
                            showCustomUi=True
                        )
                        set_data_validation_for_cell_range(dest_sheet, f"D{last_row}:D{last_row}", rule_Type)
                        # ✅ Style uniquement sur la ligne insérée
                        fmt = CellFormat(
                            textFormat=TextFormat(fontFamily="Baloo 2", bold=False),
                            borders=Borders(
                                top=Border(style="SOLID", color=Color(0, 0, 0)),
                                bottom=Border(style="SOLID", color=Color(0, 0, 0)),
                                left=Border(style="SOLID", color=Color(0, 0, 0)),
                                right=Border(style="SOLID", color=Color(0, 0, 0)),
                            )
                        )

                        format_cell_range(dest_sheet, f"A{last_row}:L{last_row}", fmt)  

            except Exception as e:
                st.error(f"Erreur sur {file['name']} : {e}")

