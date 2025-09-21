import os
import re
import sys
import threading
import unicodedata
import requests
import json
import base64
from datetime import datetime, timedelta
from difflib import SequenceMatcher
from ics import Calendar, Event
from icalendar import Calendar as ICal
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from dotenv import load_dotenv
from bs4 import BeautifulSoup
import pytz

# --- Timeout global (5 minutes) ---
def kill_script():
    print("⏱️ Timeout global dépassé, arrêt du script.")
    sys.exit(1)

watchdog = threading.Timer(300, kill_script)
watchdog.start()

# --- Chargement des variables d'environnement ---
load_dotenv()

# MyKomunoté
myk_username = os.getenv("MYK_USERNAME")
myk_password = os.getenv("MYK_PASSWORD")
myk_base_url = os.getenv("MYK_BASE_URL")
myk_api_endpoint = os.getenv("MYK_API_ENDPOINT")
myk_module_agenda = os.getenv("MYK_MODULE_AGENDA")
myk_action_agenda = os.getenv("MYK_ACTION_AGENDA")
myk_login_selector = os.getenv("MYK_LOGIN_SELECTOR")
myk_menu_selector = os.getenv("MYK_MENU_SELECTOR")
myk_calendar_selector = os.getenv("MYK_CALENDAR_SELECTOR")
myk_class_schedule_selector = os.getenv("MYK_CLASS_SCHEDULE_SELECTOR")
myk_obligatory_class_selector = os.getenv("MYK_OBLIGATORY_CLASS_SELECTOR")

# ADE
ade_base_url = os.getenv("ADE_BASE_URL")
ade_resources = os.getenv("ADE_RESOURCES")
ade_project_id = os.getenv("ADE_PROJECT_ID")

# UNESS
uness_base_url = os.getenv("UNESS_BASE_URL")
uness_id_ue_code_encoded = os.getenv("UNESS_ID_UE_CODE")  
uness_id_ue_code_str = base64.b64decode(uness_id_ue_code_encoded).decode("utf-8")
ue_to_uness = json.loads(uness_id_ue_code_str)

# --- Vérification des variables d'env ---
required_vars = [
    myk_username, myk_password, myk_base_url, myk_api_endpoint, myk_module_agenda,
    myk_action_agenda, myk_login_selector, myk_menu_selector, myk_calendar_selector,
    myk_class_schedule_selector, myk_obligatory_class_selector,
    ade_base_url, ade_resources, ade_project_id, uness_base_url, uness_id_ue_code_str
]

if not all(required_vars):
    raise ValueError("⚠️ Certaines variables d'environnement sont manquantes !")

# --- Définition des dates dynamiques ---
tz_paris = pytz.timezone("Europe/Paris")
start_dt = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0, tzinfo=tz_paris)
end_dt = start_dt + timedelta(days=14)

# Dates pour MyKomunoté (ISO avec fuseau)
mk_start = start_dt.isoformat()
mk_end = end_dt.isoformat()

# Dates pour ADE (yyyy-mm-dd)
ade_start = start_dt.strftime("%Y-%m-%d")
ade_end = end_dt.strftime("%Y-%m-%d")
ade_url = f"{ade_base_url}?resources={ade_resources}&projectId={ade_project_id}&calType=ical&firstDate={ade_start}&lastDate={ade_end}"

# --- Récupération de l'ICS ADE ---
ade_resp = requests.get(ade_url, timeout=30)
if ade_resp.status_code != 200:
    raise RuntimeError(f"❌ Erreur lors de la récupération ADE ICS : {ade_resp.status_code}")

ade_cal = ICal.from_ical(ade_resp.text)
ade_events = [
    {
        "summary": str(component.get("SUMMARY")),
        "location": str(component.get("LOCATION", "")),
        "start": component.get("DTSTART").dt,
        "end": component.get("DTEND").dt
    }
    for component in ade_cal.walk() if component.name == "VEVENT"
]

# --- Connexion à MyKomunoté et récupération JSON ---
content = []
try:
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()

        page.goto(myk_base_url, timeout=20000)

        # Connexion
        page.fill('input[name="CODE"]', myk_username)
        page.fill('input[name="MOT_DE_PASSE"]', myk_password)
        page.click(myk_login_selector)

        # Navigation vers l’agenda
        page.wait_for_selector(myk_menu_selector, timeout=10000)
        page.click(myk_menu_selector)
        page.wait_for_selector(myk_calendar_selector, timeout=10000)
        page.click(myk_calendar_selector)
        page.wait_for_selector(f'text="{myk_class_schedule_selector}"', timeout=10000)
        page.click(f'text="{myk_class_schedule_selector}"')
        page.wait_for_load_state("networkidle")

        # Récupération JSON
        response = page.request.post(
            f"{myk_base_url}/{myk_api_endpoint}",
            data={
                "sCheminVersFichier": myk_module_agenda,
                "bEstUneClasse": "true",
                "sAction": myk_action_agenda,
                "sA2NINF": "",
                "startDate": mk_start,
                "endDate": mk_end
            }
        )
        content = response.json()
        import pprint
        pprint.pprint(content[0])
        sys.exit()


except PlaywrightTimeoutError as e:
    raise RuntimeError(f"⏳ Timeout Playwright : {e}")
except Exception as e:
    raise RuntimeError(f"❌ Erreur Playwright : {e}")

# --- Fonctions utilitaires ---
def normalize(s: str) -> str:
    return unicodedata.normalize("NFD", s).encode("ascii", "ignore").decode("utf-8").lower() if s else ""

def strip_html(text: str) -> str:
    return re.sub(r"<.*?>", "", text).strip() if text else ""

def clean_text(text: str) -> str:
    if not text:
        return ""
    text = strip_html(text)
    text = text.replace("\r", "").replace("\n", " - ")
    return text.strip()

# --- Nouvelle logique de matching ADE ---
STOPWORDS_FR = {
    "de", "du", "des", "le", "la", "les", "un", "une", "en", "au", "aux", 
    "pour", "par", "dans", "sur", "avec", "et", "ou", "à", "d'", "l'", "chez"
}

def nettoyer_intitule(titre: str) -> str:
    if not titre:
        return ""
    titre_clean = re.sub(r"l\d\sifmem.*?ue\s*\d+(\.\d+)?", "", titre, flags=re.IGNORECASE)
    titre_clean = normalize(titre_clean)
    mots = [m for m in re.findall(r"\w+", titre_clean) if m not in STOPWORDS_FR]
    return " ".join(mots)

def extraire_type_cours(titre: str) -> str:
    titre_norm = titre.upper()
    types_possibles = ["CM", "TD", "CM", "TPG", "TP"]
    for t in types_possibles:
        if re.search(r'[\s\-]*' + re.escape(t) + r'[\s\-]*', titre_norm):
            return t
    return ""

def trouver_salle_ade(cours_mk: dict, ade_events: list, seuil_similarite: float = 0.3) -> str:
    """
    Retourne la salle ADE correspondant à un cours MyKom.
    """
    # --- Date du cours MyKom ---
    dt_mk_start = datetime.fromisoformat(cours_mk["start"])
    
    # --- UE_CODE nettoyé ---
    ue_code = re.sub(r'^ue\s*:\s*', '', strip_html(cours_mk.get("UE_CODE", "")), flags=re.IGNORECASE).lower()
    
    # --- Intitulé et type du cours MyKom ---
    intitule_mk = strip_html(cours_mk.get("INTITULE", ""))
    type_mk = extraire_type_cours(strip_html(cours_mk.get("TYPE_COURS", "")))
    intitule_mk_clean = nettoyer_intitule(intitule_mk)

    if not ue_code:
        return "Non précisée"

    # --- Candidats ADE pour la même date ---
    candidats = [ade for ade in ade_events if ade["start"].date() == dt_mk_start.date()]
    if not candidats:
        return "Non précisée"

    # --- Filtrage par UE_CODE si possible ---
    candidats_ue = [ade for ade in candidats if ue_code in ade["summary"].lower()]
    if candidats_ue:
        candidats = candidats_ue

    meilleur_score = -1
    meilleure_salle = "Non précisée"

    for ade in candidats:
        # Intitulé et type du cours ADE
        ade_title = ade["summary"]
        ade_type = extraire_type_cours(ade_title)
        ade_title_clean = nettoyer_intitule(ade_title)

        # --- Calcul du score ---
        sim_intitule = SequenceMatcher(None, intitule_mk_clean, ade_title_clean).ratio()
        type_bonus = 0.3 if type_mk and ade_type and type_mk == ade_type else 0
        diff_minutes = abs((ade["start"] - dt_mk_start).total_seconds()) / 60
        horaire_bonus = 0.1 if diff_minutes < 30 else 0

        score = sim_intitule * 0.6 + type_bonus + horaire_bonus

        # --- Mise à jour du meilleur score ---
        if score > meilleur_score:
            meilleur_score = score
            meilleure_salle = ade.get("location", "Non précisée")

    return meilleure_salle if meilleur_score >= seuil_similarite else "Non précisée"


# --- Création du calendrier fusionné ---
final_cal = Calendar()

for cours in content:
    if not isinstance(cours, dict):
        continue

    e = Event()
    type_cours = clean_text(cours.get("TYPE_COURS", "")).strip("- ")
    intitule = clean_text(cours.get("INTITULE", ""))

    title_html = cours.get("title", "")
    cours["obligatoire"] = False
    if title_html:
        soup = BeautifulSoup(title_html, "html.parser")
        if soup.find("i", class_=myk_obligatory_class_selector):
            cours["obligatoire"] = True

    prefix = "★ " if cours.get("obligatoire") else ""
    e.name = f"{prefix}{type_cours} - {intitule}"

    start_dt = datetime.fromisoformat(cours["start"])
    end_dt = datetime.fromisoformat(cours["end"])
    e.begin = tz_paris.localize(start_dt)
    e.end = tz_paris.localize(end_dt)

    description_parts = [
        f"Le {start_dt.strftime('%d/%m/%Y')} de {start_dt.strftime('%H:%M')} à {end_dt.strftime('%H:%M')}",
        f"Cours obligatoire : <b>{'Oui' if cours.get('obligatoire') else 'Non'}</b>"
    ]
    if type_cours:
        description_parts.append(f"Type cours : <b>{type_cours}</b>")
    if intitule:
        description_parts.append(f"Intitulé : <b>{intitule}</b>")
    if cours.get("GROUPE"):
        raw_groupe = strip_html(str(cours["GROUPE"])).strip()
        match = re.match(r'^Gpe\s*:\s*(?:Gr\s*)?(.+)$', raw_groupe, flags=re.IGNORECASE)
        clean_groupe = match.group(1).strip() if match else raw_groupe
        description_parts.append(f"Groupe : <b>{clean_groupe}</b>")
    if cours.get("MEMBRE_PERSO"):
        clean_formateur = re.sub(r"^Formateur\(s\)\s*:\s*", "", strip_html(str(cours["MEMBRE_PERSO"])))
        description_parts.append(f"Formateur(s) : <b>{clean_formateur}</b>")
    if cours.get("INTERVENANT"):
        clean_intervenant = re.sub(r"^Intervenant\(s\)\s*:\s*", "", strip_html(str(cours["INTERVENANT"])))
        description_parts.append(f"Intervenant(s) : <b>{clean_intervenant}</b>")

    description_parts.append("")

    ue_code_clean = None
    if cours.get("UE_CODE"):
        ue_code_clean = re.sub(r'^(UE\s*:)\s*', '', strip_html(str(cours['UE_CODE'])))
        description_parts.append(f"UE code : {ue_code_clean}")
    if cours.get("UE_LIBE"):
        description_parts.append(f"UE libellé : {strip_html(cours['UE_LIBE'])}")

    if ue_code_clean in ue_to_uness:
        uness_id = ue_to_uness[ue_code_clean]
        description_parts.append(f"Lien UNESS : <b>{uness_base_url}?id={uness_id}</b>")

    description_clean = [
        re.sub(r"<br\s*/?>", " - ", part, flags=re.IGNORECASE).replace("\n", " - ").replace("\r", "")
        for part in description_parts
    ]
    e.description = "\n".join(description_clean)
    e.location = trouver_salle_ade(cours, ade_events)
    final_cal.events.add(e)

# --- Sauvegarde ICS ---
with open("edt_global.ics", "w", encoding="utf-8") as f:
    f.writelines(final_cal)

print("✅ Fichier edt_global.ics généré avec succès !")
watchdog.cancel()


