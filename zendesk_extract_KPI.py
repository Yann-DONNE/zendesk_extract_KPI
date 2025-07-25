import requests
from requests.auth import HTTPBasicAuth
from collections import defaultdict
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Alignment
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from tqdm import tqdm
import time

print("Script d'extraction KPI Zendesk")
print("Auteur : Yann Donne")
print("Date : 2025")
print("Version : 1.2 (Plage de dates + Stats mensuelles)")

# Chrono début
start_time = time.time()

# Paramètres Zendesk
SUBDOMAIN = 'ton domaine'
EMAIL = 'TonEmail@adresse.com'
API_TOKEN = 'Ton API Token'

auth = HTTPBasicAuth(f'{EMAIL}/token', API_TOKEN)

# Demande date de début
date_input = input("Entrez la date de **début** d'extraction (JJ/MM/YYYY) ou appuyez sur Entrée pour utiliser le 01/01 de l'année en cours : ")
if date_input.strip() == "":
    start_date = datetime(datetime.now().year, 1, 1).strftime('%Y-%m-%d')
else:
    try:
        date_obj = datetime.strptime(date_input, '%d/%m/%Y')
        start_date = date_obj.strftime('%Y-%m-%d')
    except ValueError:
        print("❌ Format de date invalide. Utilisation du 01/01 de l'année en cours.")
        start_date = datetime(datetime.now().year, 1, 1).strftime('%Y-%m-%d')

# Demande date de fin
end_input = input("Entrez la date de **fin** d'extraction (JJ/MM/YYYY) ou appuyez sur Entrée pour utiliser aujourd'hui : ")
if end_input.strip() == "":
    end_date = datetime.now().strftime('%Y-%m-%d')
else:
    try:
        end_obj = datetime.strptime(end_input, '%d/%m/%Y')
        end_date = end_obj.strftime('%Y-%m-%d')
    except ValueError:
        print("❌ Format de date invalide. Utilisation de la date du jour.")
        end_date = datetime.now().strftime('%Y-%m-%d')

print(f"📅 Extraction des tickets créés entre le : {start_date} et {end_date}")

types_cibles = {'incident', 'question', 'problem', 'task'}

# Ajout plage de dates dans la requête
url_tickets = f"https://{SUBDOMAIN}.zendesk.com/api/v2/search.json?query=type:ticket created>{start_date} created<{end_date}"

def get_all_tickets_filtered(url, types_cibles):
    all_tickets = []
    while url:
        print("Récupération : ", url)
        response = requests.get(url, auth=auth)
        if response.status_code != 200:
            raise Exception(f"Erreur API: {response.status_code} {response.text}")
        data = response.json()
        filtered = [t for t in data.get('results', []) if (t.get('type') or '').lower() in types_cibles]
        all_tickets.extend(filtered)
        url = data.get('next_page')
    return all_tickets

def get_ticket_metrics(ticket_id):
    url = f"https://{SUBDOMAIN}.zendesk.com/api/v2/tickets/{ticket_id}/metrics.json"
    try:
        response = requests.get(url, auth=auth, timeout=10)
        if response.status_code == 200:
            data = response.json().get('ticket_metric', {})
            first_reply_time = data.get('reply_time_in_minutes', {}).get('calendar')
            resolution_time = data.get('full_resolution_time_in_minutes', {}).get('calendar')
            return ticket_id, first_reply_time, resolution_time
        else:
            print(f"Erreur récupération métriques ticket {ticket_id}: {response.status_code}")
            return ticket_id, None, None
    except Exception as e:
        print(f"Exception sur ticket {ticket_id}: {e}")
        return ticket_id, None, None

print("🔍 Chargement des tickets filtrés...")
tickets = get_all_tickets_filtered(url_tickets, types_cibles)
print(f"✔ {len(tickets)} tickets récupérés après filtrage.")

# Fusionner 'problem' dans 'incident'
for ticket in tickets:
    if ticket.get('type') == 'problem':
        ticket['type'] = 'incident'

all_types = ['incident', 'question', 'task']

tag_data = defaultdict(lambda: {
    'types': defaultdict(int),
    'total': 0
})

for ticket in tickets:
    ttype = ticket.get('type') or 'inconnu'
    tags = ticket.get('tags', [])
    for tag in tags:
        if tag.startswith('com'):
            tag_data[tag]['types'][ttype] += 1
            tag_data[tag]['total'] += 1

def sort_com_tags(tags):
    def extract_number(tag):
        match = re.search(r'com(\d+)', tag)
        return int(match.group(1)) if match else float('inf')
    return sorted(tags, key=extract_number)

com_tags_sorted = sort_com_tags([tag for tag in tag_data if tag.startswith('com')])

# Analyse délais
delai_first_reply = {'0-1h': 0, '1-8h': 0, '8-24h': 0, '>24h': 0}
delai_resolution = {'0-5h': 0, '5-24h': 0, '1-7j': 0, '7-30j': 0, '>30j': 0}

max_workers = 10

print("⏳ Récupération des métriques tickets...")
with ThreadPoolExecutor(max_workers=max_workers) as executor:
    futures = {executor.submit(get_ticket_metrics, ticket['id']): ticket for ticket in tickets}
    for future in tqdm(as_completed(futures), total=len(futures), desc="Traitement des tickets"):
        ticket_id, first_reply_time, resolution_time = future.result()
        if first_reply_time is not None:
            if first_reply_time <= 60:
                delai_first_reply['0-1h'] += 1
            elif first_reply_time <= 480:
                delai_first_reply['1-8h'] += 1
            elif first_reply_time <= 1440:
                delai_first_reply['8-24h'] += 1
            else:
                delai_first_reply['>24h'] += 1
        if resolution_time is not None:
            if resolution_time <= 300:
                delai_resolution['0-5h'] += 1
            elif resolution_time <= 1440:
                delai_resolution['5-24h'] += 1
            elif resolution_time <= 10080:
                delai_resolution['1-7j'] += 1
            elif resolution_time <= 43200:
                delai_resolution['7-30j'] += 1
            else:
                delai_resolution['>30j'] += 1

total_first_reply = sum(delai_first_reply.values())
total_resolution = sum(delai_resolution.values())

# Analyse Satisfaction
satisfaction_counts = {'good': 0, 'bad': 0}
total_notes = 0
for ticket in tickets:
    satisfaction = ticket.get('satisfaction_rating')
    if satisfaction and isinstance(satisfaction, dict):
        score = satisfaction.get('score')
        if score in ('good', 'bad'):
            satisfaction_counts[score] += 1
            total_notes += 1
pct_satisfaction = round((satisfaction_counts['good'] / total_notes) * 100, 1) if total_notes > 0 else 0.0

# Analyse nombre de tickets par type
tickets_par_type = defaultdict(int)
for ticket in tickets:
    ttype = ticket.get('type') or 'inconnu'
    tickets_par_type[ttype] += 1

# --- Partie nouvelle : Extraction par mois et type + satisfaction par mois ---

# Mois en français (index 1 = janvier)
mois_fr = [
    "Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
    "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"
]

# Structure de stockage par mois : { 'YYYY-MM' : { 'incident': int, 'question': int, 'task': int, 'total': int, 'good': int, 'bad': int } }
stats_mensuelles = defaultdict(lambda: {'incident': 0, 'question': 0, 'task': 0, 'total': 0, 'good': 0, 'bad': 0})

for ticket in tickets:
    created_at = ticket.get('created_at')
    if created_at:
        # Format '2025-07-25T13:25:31Z' -> extraire YYYY-MM
        mois_cle = created_at[:7]
        ttype = ticket.get('type')
        if ttype == 'problem':
            ttype = 'incident'
        if ttype in all_types:
            stats_mensuelles[mois_cle][ttype] += 1
            stats_mensuelles[mois_cle]['total'] += 1
            satisfaction = ticket.get('satisfaction_rating')
            if satisfaction and isinstance(satisfaction, dict):
                score = satisfaction.get('score')
                if score in ('good', 'bad'):
                    stats_mensuelles[mois_cle][score] += 1

# Création Excel
wb = Workbook()

# Onglet "Tickets par Tags"
ws_tags = wb.active
ws_tags.title = "Tickets par Tags"
header = ["Tag"] + all_types + ["Total tickets"]
ws_tags.append(header)
for tag in com_tags_sorted:
    data = tag_data[tag]
    row = [tag]
    for ttype in all_types:
        row.append(data['types'].get(ttype, 0))
    row.append(data['total'])
    ws_tags.append(row)

# Onglet "Délai 1ère Prise"
ws_delai = wb.create_sheet(title="Délai 1ère Prise")
ws_delai.append(["Délai", "% Tickets", "Nombre estimé"])
for categorie, count in delai_first_reply.items():
    pct = round((count / total_first_reply) * 100) if total_first_reply > 0 else 0
    ws_delai.append([categorie, f"{pct}%", count])

# Onglet "Délai Résolution Complète"
ws_resol = wb.create_sheet(title="Délai Résolution Complète")
ws_resol.append(["Délai", "% Tickets", "Nombre estimé"])
for categorie, count in delai_resolution.items():
    pct = round((count / total_resolution) * 100) if total_resolution > 0 else 0
    ws_resol.append([categorie, f"{pct}%", count])

# Onglet "Satisfaction"
ws_satisfaction = wb.create_sheet(title="Satisfaction")
ws_satisfaction.append(["Indicateur", "Valeur"])
ws_satisfaction.append(["% Satisfaction Globale", f"{pct_satisfaction}%"])
ws_satisfaction.append(["Nombre 'Good'", satisfaction_counts['good']])
ws_satisfaction.append(["Nombre 'Bad'", satisfaction_counts['bad']])

# Onglet "Tickets par Type"
ws_type = wb.create_sheet(title="Tickets par Type")
ws_type.append(["Type de ticket", "Nombre"])
total_tickets = sum([tickets_par_type[t] for t in all_types])
for ttype in all_types:
    count = tickets_par_type.get(ttype, 0)
    pct = round((count / total_tickets) * 100) if total_tickets > 0 else 0
    ws_type.append([ttype, f"{count} ({pct}%)"])
ws_type.append(["Total", total_tickets])

# --- Nouvel onglet "Tickets par Mois" avec taux satisfaction et nombre avis ---

ws_mois = wb.create_sheet(title="Tickets par Mois")

# Entête
entete = ["Mois", "Incidents", "Questions", "Tasks", "Total tickets", "% Satisfaction", "Nb Avis"]
ws_mois.append(entete)

# Centrage et style
align_center = Alignment(horizontal='center', vertical='center')

# Tri des mois par ordre chronologique
mois_tries = sorted(stats_mensuelles.keys())

for mois_cle in mois_tries:
    data = stats_mensuelles[mois_cle]
    # Mois et année pour afficher en français (ex: '2025-07' -> 'Juillet 2025')
    annee, mois_num = mois_cle.split('-')
    mois_fr_str = mois_fr[int(mois_num)-1] + " " + annee

    # Calcul % satisfaction good
    total_avis = data['good'] + data['bad']
    pct_sat = (data['good'] / total_avis * 100) if total_avis > 0 else 0.0

    # Formater nombre + %
    def format_nb_pct(nb, total):
        pct = (nb / total * 100) if total > 0 else 0
        return f"{nb} ({int(pct)}%)"

    ligne = [
        mois_fr_str,
        format_nb_pct(data['incident'], data['total']),
        format_nb_pct(data['question'], data['total']),
        format_nb_pct(data['task'], data['total']),
        data['total'],
        f"{round(pct_sat, 1)}%",
        total_avis
    ]
    ws_mois.append(ligne)

# Centrer les données
for row in ws_mois.iter_rows(min_row=2, max_row=ws_mois.max_row, min_col=1, max_col=len(entete)):
    for cell in row:
        cell.alignment = align_center

# Sauvegarde Excel
wb.save("tickets_par_tags_et_delais.xlsx")

end_time = time.time()
elapsed_time = round(end_time - start_time, 2)
print(f"\n✔ Fichier Excel généré : tickets_par_tags_et_delais.xlsx en {elapsed_time} secondes.")
input("Appuyez sur Entrée pour fermer...")

