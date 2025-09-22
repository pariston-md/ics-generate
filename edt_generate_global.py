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

# --- Vérification des variables d'env ---
required_vars = [
    myk_username, myk_password, myk_base_url, myk_api_endpoint, myk_module_agenda,
    myk_action_agenda, myk_login_selector, myk_menu_selector, myk_calendar_selector,
    myk_class_schedule_selector, myk_obligatory_class_selector,
    ade_base_url, ade_resources, ade_project_id
]

if not all(required_vars):
    raise ValueError("⚠️ Certaines variables d'environnement sont manquantes !")

# --- Définition des dates dynamiques ---
tz_paris = pytz.timezone("Europe/Paris")
start_dt = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0, tzinfo=tz_paris)
end_dt = start_dt + timedelta(days=5)

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
with open("ade.json", "w", encoding="utf-8") as f:
    json.dump(ade_events, f, ensure_ascii=False, indent=2, default=str)

print("✅ Fichier ade.json généré avec succès !")

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
    with open("mykomu.json", "w", encoding="utf-8") as f:
        json.dump(content, f, ensure_ascii=False, indent=2, default=str)

    print("✅ Fichier mykomu.json généré avec succès !")


except PlaywrightTimeoutError as e:
    raise RuntimeError(f"⏳ Timeout Playwright : {e}")
except Exception as e:
    raise RuntimeError(f"❌ Erreur Playwright : {e}")

# --- Fonctions utilitaires ---

from collections import Counter
from difflib import SequenceMatcher
from datetime import datetime
import re
import unicodedata
from scipy.optimize import linear_sum_assignment
import numpy as np

# --- Fonctions utilitaires (inchangées sauf ajustement petits bugs) ---
def clean_text(text: str) -> str: 
    if not text: 
        return "" 
    text = strip_html(text) 
    text = text.replace("\r", "").replace("\n", " - ") 
    return text.strip()

import re
import unicodedata
from difflib import SequenceMatcher
from collections import Counter
from datetime import datetime
import numpy as np
import itertools

# --- utilitaires déjà présents ---
def strip_html(text: str) -> str:
    return re.sub(r"<.*?>", "", text).strip() if text else ""

def enlever_accents(txt: str) -> str:
    return ''.join(
        c for c in unicodedata.normalize('NFD', txt)
        if unicodedata.category(c) != 'Mn'
    )

def normaliser_texte(txt: str) -> str:
    txt = re.sub(r"<.*?>", " ", txt)          
    txt = enlever_accents(txt)                
    txt = txt.lower()
    txt = re.sub(r"[^\w\s]", " ", txt)        
    txt = re.sub(r"\b(td|cm|ue|obligatoire|gr\d+|l\d+|ifmem)\b", " ", txt)
    txt = re.sub(r"\s+", " ", txt).strip()
    return txt

def normaliser_ue_code(code: str) -> str:
    """Ex: 3.02 -> 3.2, 04.01 -> 4.1"""
    if not code:
        return ""
    parts = code.split(".")
    parts = [str(int(p)) for p in parts if p.isdigit()]
    return ".".join(parts)

def extraire_ue_code_mk(cours_mk: dict) -> str:
    raw = cours_mk.get("UE_CODE", "")
    m = re.search(r"UE\s*:\s*([\d.]+)", raw)
    return normaliser_ue_code(m.group(1)) if m else ""

def extraire_ue_code_ade(summary: str) -> str:
    m = re.search(r"UE\s*([\d.]+)", summary)
    return normaliser_ue_code(m.group(1)) if m else ""

def similarite_tokens(mk_tokens, ade_tokens, poids_mots=None) -> float:
    matches, poids_total = 0, 0
    for mk in mk_tokens:
        poids = poids_mots.get(mk, 1) if poids_mots else 1
        poids_total += poids
        for ade in ade_tokens:
            ratio = SequenceMatcher(None, mk, ade).ratio()
            if ratio >= 0.7:
                matches += poids
                break
    return matches / poids_total if poids_total > 0 else 0

def calculer_poids_mots(cours_mk_du_jour, cours_ade_du_jour):
    all_tokens = []
    for c in cours_mk_du_jour:
        all_tokens += normaliser_texte(strip_html(c.get("INTITULE",""))).split()
    for c in cours_ade_du_jour:
        all_tokens += normaliser_texte(c["summary"]).split()

    freq = Counter(all_tokens)
    poids = {tok: 1/freq[tok] for tok in freq}
    return poids

# --- matching global sans scipy (bruteforce car peu de cours par jour) ---
def matcher_cours_journee(cours_mk_du_jour, cours_ade_du_jour, seuil=0.2):
    if not cours_mk_du_jour or not cours_ade_du_jour:
        return {}

    poids_mots = calculer_poids_mots(cours_mk_du_jour, cours_ade_du_jour)

    n_mk = len(cours_mk_du_jour)
    n_ade = len(cours_ade_du_jour)
    score_matrix = np.zeros((n_mk, n_ade))

    for i, mk in enumerate(cours_mk_du_jour):
        mk_tokens = normaliser_texte(strip_html(mk.get("INTITULE",""))).split()
        ue_mk = extraire_ue_code_mk(mk)
        dt_mk_start = datetime.fromisoformat(mk["start"])
        dt_mk_end = datetime.fromisoformat(mk["end"])

        for j, ade in enumerate(cours_ade_du_jour):
            ade_tokens = normaliser_texte(ade["summary"]).split()
            score = similarite_tokens(mk_tokens, ade_tokens, poids_mots)

            # Bonus UE identique
            ue_ade = extraire_ue_code_ade(ade["summary"])
            if ue_mk and ue_mk == ue_ade:
                score += 0.5  # bonus renforcé

            # Bonus si même jour (même sans la même heure)
            if ade["start"].date() == dt_mk_start.date():
                score += 0.1

            # Bonus horaires exacts
            if (ade["start"].time() == dt_mk_start.time() and
                ade["end"].time() == dt_mk_end.time()):
                score += 0.2

            score_matrix[i, j] = score

    # --- DEBUG : afficher les scores ---
    date_jour = datetime.fromisoformat(cours_mk_du_jour[0]["start"]).date()
    print(f"\n=== Matching du {date_jour} ===")
    for i, mk in enumerate(cours_mk_du_jour):
        mk_title = strip_html(mk.get("INTITULE",""))
        print(f"\nCours MyKomu {i}: {mk_title}")
        for j, ade in enumerate(cours_ade_du_jour):
            ade_title = ade["summary"]
            print(f"   -> ADE {j}: {ade_title} | score={score_matrix[i,j]:.2f}")

    # --- Appariement bruteforce (permutations) ---
    best_perm, best_score = None, -1
    for perm in itertools.permutations(range(n_ade), min(n_mk, n_ade)):
        total = sum(score_matrix[i, j] for i, j in enumerate(perm))
        if total > best_score:
            best_score = total
            best_perm = perm

    mapping = {}
    for i, j in enumerate(best_perm):
        if score_matrix[i, j] >= seuil:
            mapping[i] = cours_ade_du_jour[j].get("location", "Non précisée")
        else:
            mapping[i] = "Non précisée"

    return mapping





# --- Création du calendrier fusionné ---
final_cal = Calendar()

# Récupération de toutes les dates distinctes dans MyKomu
dates_jours = sorted(set(datetime.fromisoformat(c["start"]).date() for c in content if isinstance(c, dict)))

for date_jour in dates_jours:
    # Sélection cours du jour côté MyKomu et ADE
    cours_mk_du_jour = [c for c in content if isinstance(c, dict) and datetime.fromisoformat(c["start"]).date() == date_jour]
    cours_ade_du_jour = [a for a in ade_events if a["start"].date() == date_jour]

    # Matching global
    mapping = matcher_cours_journee(cours_mk_du_jour, cours_ade_du_jour)

    # Génération des events ICS
    for idx, cours in enumerate(cours_mk_du_jour):
        e = Event()
        type_cours = clean_text(cours.get("TYPE_COURS", "")).strip("- ")
        intitule = clean_text(cours.get("INTITULE", ""))

        # Détection cours obligatoire
        title_html = cours.get("title", "")
        cours["obligatoire"] = False
        if title_html:
            soup = BeautifulSoup(title_html, "html.parser")
            if soup.find("i", class_=myk_obligatory_class_selector):
                cours["obligatoire"] = True

        prefix = "★ " if cours.get("obligatoire") else ""
        e.name = f"{prefix}{type_cours} - {intitule}"

        # Dates
        start_dt = datetime.fromisoformat(cours["start"])
        end_dt = datetime.fromisoformat(cours["end"])
        e.begin = tz_paris.localize(start_dt)
        e.end = tz_paris.localize(end_dt)

        # Description
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

        if cours.get("UE_CODE"):
            ue_code_clean = re.sub(r'^(UE\s*:)\s*', '', strip_html(str(cours['UE_CODE'])))
            description_parts.append(f"UE code : {ue_code_clean}")
        if cours.get("UE_LIBE"):
            description_parts.append(f"UE libellé : {strip_html(cours['UE_LIBE'])}")

        description_clean = [
            re.sub(r"<br\s*/?>", " - ", part, flags=re.IGNORECASE).replace("\n", " - ").replace("\r", "")
            for part in description_parts
        ]
        e.description = "\n".join(description_clean)

        # Salle issue du matching global
        e.location = mapping.get(idx, "Non précisée")

        final_cal.events.add(e)

# --- Sauvegarde ICS ---
with open("edt_global.ics", "w", encoding="utf-8") as f:
    f.writelines(final_cal)

print("✅ Fichier edt_global.ics généré avec succès !")
watchdog.cancel()
