import re
from ics import Calendar
from bs4 import BeautifulSoup

# --- Charger le calendrier global ---
with open("edt_global.ics", "r", encoding="utf-8") as f:
    global_cal = Calendar(f.read())

# --- Définition des groupes ---
letters = ["A", "B", "C", "D"]        # Quart promo
subgroups = ["alpha", "beta", "gamma"] # Tiers promo
demi_map = {"A": "1/2", "B": "1/2", "C": "2/2", "D": "2/2"}  # Demi-promo selon lettre

# --- Préparer les 12 calendriers ---
ics_grouped = {}
for letter in letters:
    for sub in subgroups:
        combo_name = f"{letter}_{demi_map[letter]}_{sub}"
        ics_grouped[combo_name] = Calendar()

# --- Ajouter les événements appropriés ---
for event in global_cal.events:
    desc = event.description or ""
    
    # Extraire la valeur du champ "Groupe"
    match = re.search(r"Groupe\s*:\s*(.*)", desc, re.IGNORECASE)
    event_group = match.group(1).strip() if match else "Promotion entière"
    
    # Nettoyer HTML et espaces
    event_group = BeautifulSoup(event_group, "html.parser").get_text().strip()
    event_group_lower = event_group.lower()

    # Ajouter l'événement aux ICS correspondants
    for combo_name, cal in ics_grouped.items():
        letter, demi, sub = combo_name.split("_")
        if event_group_lower in [letter.lower(), demi.lower(), sub.lower(), "promotion entière"]:
            cal.events.add(event)

# --- Sauvegarde des ICS ---
for combo_name, cal in ics_grouped.items():
    letter, _, sub = combo_name.split("_")  # ignore le demi-groupe
    filename = f"edt_{letter}_{sub}.ics"
    with open(filename, "w", encoding="utf-8") as f:
        f.writelines(cal)

print("✅ 12 ICS créés à partir du global avec les groupes corrects !")