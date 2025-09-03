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

print("Script d'extraction KPI Zendesk (version robuste)")
print("Auteur : Yann Donne")
print("Date : 2025")
print("Version : 2.0 (Incremental API + Gestion erreurs)")

# Chrono début
start_time = time.time()

# Paramètres Zendesk
SUBDOMAIN = 'ton domaine'
EMAIL = 'ton email'
API_TOKEN = 'ta clé API'

auth = HTTPBasicAuth(f'{EMAIL}/token', API_TOKEN)

# --- Dates d'extraction ---
date_input = input("Entrez la date de **début** (JJ/MM/YYYY) ou Entrée pour 01/01 : ")
if date_input.strip() == "":
    start_date = datetime(datetime.now().year, 1, 1)
else:
    try:
        start_date = datetime.strptime(date_input, '%d/%m/%Y')
    except ValueError:
        print("❌ Format invalide → utilisation du 01/01 de l'année en cours.")
        start_date = datetime(datetime.now().year, 1, 1)

end_input = input("Entrez la date de **fin** (JJ/MM/YYYY) ou Entrée pour aujourd'hui : ")
if end_input.strip() == "":
    end_date = datetime.now()
else:
    try:
        end_date = datetime.strptime(end_input, '%d/%m/%Y')
    except ValueError:
        print("❌ Format invalide → utilisation d'aujourd'hui.")
        end_date = datetime.now()

print(f"📅 Extraction des tickets du {start_date.date()} au {end_date.date()}")

# --- Fonction récupération tickets via Incremental API ---
def get_tickets_incremental(start_date, end_date):
    url = f"https://{SUBDOMAIN}.zendesk.com/api/v2/incremental/tickets.json?start_time={int(start_date.timestamp())}"
    all_tickets = []
    while url:
        try:
            resp = requests.get(url, auth=auth, timeout=30)
            if resp.status_code == 429:
                retry_after = int(resp.headers.get("Retry-After", 10))
                print(f"⏳ Rate limit atteint → pause {retry_after}s...")
                time.sleep(retry_after)
                continue
            if resp.status_code != 200:
                print(f"❌ Erreur API {resp.status_code}: {resp.text}")
                break

            data = resp.json()
            for t in data.get("tickets", []):
                created = datetime.strptime(t["created_at"], "%Y-%m-%dT%H:%M:%SZ")
                if start_date <= created <= end_date:
                    all_tickets.append(t)

            if data.get("end_of_stream"):
                break
            url = data.get("next_page")
        except Exception as e:
            print("❌ Exception pendant récupération:", e)
            break

    return all_tickets

# --- Chargement tickets ---
print("🔍 Chargement tickets...")
tickets = get_tickets_incremental(start_date, end_date)
print(f"✔ {len(tickets)} tickets récupérés")

# --- Nettoyage et regroupements ---
types_cibles = {"incident", "question", "problem", "task"}
for t in tickets:
    if t.get("type") == "problem":
        t["type"] = "incident"

all_types = ["incident", "question", "task"]

tag_data = defaultdict(lambda: {
    "types": defaultdict(int),
    "total": 0
})

# --- NOUVEAU: sets pour uniques par type (tickets ayant au moins un tag com)
unique_tagged_by_type = {t: set() for t in all_types}
unique_tagged_total = set()

# ✅ Correction appliquée : on ne compte que les types valides pour les totaux par tag
for ticket in tickets:
    ttype = ticket.get("type")
    tid = ticket.get("id")
    tags = ticket.get("tags", [])
    has_com = False
    for tag in tags:
        if tag.startswith("com") and ttype in all_types:
            tag_data[tag]["types"][ttype] += 1
            tag_data[tag]["total"] += 1
            has_com = True
    if has_com and ttype in all_types and tid is not None:
        unique_tagged_by_type[ttype].add(tid)
        unique_tagged_total.add(tid)

def sort_com_tags(tags):
    def extract_number(tag):
        match = re.search(r'com(\d+)', tag)
        return int(match.group(1)) if match else float('inf')
    return sorted(tags, key=extract_number)

com_tags_sorted = sort_com_tags([tag for tag in tag_data if tag.startswith("com")])

# --- Récupération des métriques ---
def get_ticket_metrics(ticket_id, retries=3):
    url = f"https://{SUBDOMAIN}.zendesk.com/api/v2/tickets/{ticket_id}/metrics.json"
    try:
        response = requests.get(url, auth=auth, timeout=10)
        if response.status_code == 200:
            data = response.json().get("ticket_metric", {})
            first_reply_time = data.get("reply_time_in_minutes", {}).get("calendar")
            resolution_time = data.get("full_resolution_time_in_minutes", {}).get("calendar")
            return ticket_id, first_reply_time, resolution_time
        elif response.status_code == 429:
            if retries <= 0:
                return ticket_id, None, None
            retry_after = int(response.headers.get("Retry-After", 5))
            time.sleep(retry_after)
            return get_ticket_metrics(ticket_id, retries - 1)
        else:
            return ticket_id, None, None
    except Exception:
        if retries <= 0:
            return ticket_id, None, None
        time.sleep(2)
        return get_ticket_metrics(ticket_id, retries - 1)

delai_first_reply = {'0-1h': 0, '1-8h': 0, '8-24h': 0, '>24h': 0}
delai_resolution = {'0-5h': 0, '5-24h': 0, '1-7j': 0, '7-30j': 0, '>30j': 0}

# Pour cohérence : on rattache les métriques au type du ticket
ticket_lookup = {ticket['id']: ticket for ticket in tickets if 'id' in ticket}

# Counters par type pour métriques (permet de calculer "sans métrique" par type)
metrics_first_by_type = defaultdict(int)
metrics_resolution_by_type = defaultdict(int)
tickets_with_first_reply = set()
tickets_with_resolution = set()

max_workers = 5  # réduit pour éviter surcharge API

print("⏳ Récupération des métriques tickets...")
with ThreadPoolExecutor(max_workers=max_workers) as executor:
    futures = {executor.submit(get_ticket_metrics, tid): tid for tid in ticket_lookup.keys()}
    for future in tqdm(as_completed(futures), total=len(futures), desc="Traitement des tickets"):
        try:
            ticket_id, first_reply_time, resolution_time = future.result()
        except Exception:
            # Sécurité : si une future lève, on ignore ce ticket (compte comme sans métrique)
            continue

        ticket_obj = ticket_lookup.get(ticket_id)
        if not ticket_obj:
            continue
        ttype = ticket_obj.get('type')

        # On ne prend en compte que les types valides
        if ttype in all_types:
            # First reply
            if first_reply_time is not None:
                tickets_with_first_reply.add(ticket_id)
                metrics_first_by_type[ttype] += 1
                if first_reply_time <= 60:
                    delai_first_reply['0-1h'] += 1
                elif first_reply_time <= 480:
                    delai_first_reply['1-8h'] += 1
                elif first_reply_time <= 1440:
                    delai_first_reply['8-24h'] += 1
                else:
                    delai_first_reply['>24h'] += 1
            # Resolution
            if resolution_time is not None:
                tickets_with_resolution.add(ticket_id)
                metrics_resolution_by_type[ttype] += 1
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

total_first_reply = sum(delai_first_reply.values())  # nombre tickets avec first_reply catégorisé
total_resolution = sum(delai_resolution.values())  # nombre tickets avec resolution catégorisé

# --- Satisfaction globale ---
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

# --- Tickets par type ---
tickets_par_type = defaultdict(int)
for ticket in tickets:
    ttype = ticket.get('type') or 'inconnu'
    tickets_par_type[ttype] += 1

# total des tickets considérés dans les stats (seulement les 3 types)
total_tickets = sum([tickets_par_type[t] for t in all_types])

# --- Stats mensuelles ---
mois_fr = [
    "Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
    "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"
]

stats_mensuelles = defaultdict(lambda: {'incident': 0, 'question': 0, 'task': 0, 'total': 0, 'good': 0, 'bad': 0})

for ticket in tickets:
    created_at = ticket.get('created_at')
    if created_at:
        mois_cle = created_at[:7]
        ttype = ticket.get('type')
        if ttype in all_types:
            stats_mensuelles[mois_cle][ttype] += 1
            stats_mensuelles[mois_cle]['total'] += 1
            satisfaction = ticket.get('satisfaction_rating')
            if satisfaction and isinstance(satisfaction, dict):
                score = satisfaction.get('score')
                if score in ('good', 'bad'):
                    stats_mensuelles[mois_cle][score] += 1

# --- Création Excel ---
wb = Workbook()
try:
    # Onglet Tags
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

    # --- Résumé de cohérence : tickets uniques avec >=1 com tag + sans tag
    ws_tags.append([])  # ligne vide
    ws_tags.append(["--- Résumé cohérence ---"])
    summary_header = ["Indicateur"] + all_types + ["Total"]
    ws_tags.append(summary_header)

    # Unique tickets ayant au moins 1 com tag (par type)
    unique_row = ["Tickets avec >=1 tag 'com'"]
    for t in all_types:
        unique_row.append(len(unique_tagged_by_type.get(t, set())))
    unique_row.append(len(unique_tagged_total))
    ws_tags.append(unique_row)

    # Tickets sans tag 'com' (par type)
    sans_tag_row = ["Tickets SANS tag 'com'"]
    sum_sans_tag = 0
    for t in all_types:
        sans = tickets_par_type.get(t, 0) - len(unique_tagged_by_type.get(t, set()))
        if sans < 0:
            sans = 0
        sans_tag_row.append(sans)
        sum_sans_tag += sans
    sans_tag_row.append(sum_sans_tag)
    ws_tags.append(sans_tag_row)

    # Totaux pour vérifier cohérence
    total_row = ["Total tickets (par type)"]
    sum_tot = 0
    for t in all_types:
        val = tickets_par_type.get(t, 0)
        total_row.append(val)
        sum_tot += val
    total_row.append(sum_tot)
    ws_tags.append(total_row)

    # Onglet délai 1ère prise
    ws_delai = wb.create_sheet(title="Délai 1ère Prise")
    ws_delai.append(["Délai", "% Tickets (sur total)", "Nombre estimé"])
    # On affiche les catégories existantes puis on ajoute "Sans métrique" pour cohérence
    for categorie, count in delai_first_reply.items():
        pct = round((count / total_tickets) * 100, 1) if total_tickets > 0 else 0
        ws_delai.append([categorie, f"{pct}%", count])

    # Ligne "Sans métrique" (tickets de types valides sans first_reply)
    sum_cats_first = sum(delai_first_reply.values())
    sans_metric_first = total_tickets - sum_cats_first
    if sans_metric_first < 0:
        sans_metric_first = 0
    pct_sans_first = round((sans_metric_first / total_tickets) * 100, 1) if total_tickets > 0 else 0
    ws_delai.append(["Sans métrique (aucun first_reply)", f"{pct_sans_first}%", sans_metric_first])

    # Onglet délai résolution
    ws_resol = wb.create_sheet(title="Délai Résolution Complète")
    ws_resol.append(["Délai", "% Tickets (sur total)", "Nombre estimé"])
    for categorie, count in delai_resolution.items():
        pct = round((count / total_tickets) * 100, 1) if total_tickets > 0 else 0
        ws_resol.append([categorie, f"{pct}%", count])

    sum_cats_res = sum(delai_resolution.values())
    sans_metric_res = total_tickets - sum_cats_res
    if sans_metric_res < 0:
        sans_metric_res = 0
    pct_sans_res = round((sans_metric_res / total_tickets) * 100, 1) if total_tickets > 0 else 0
    ws_resol.append(["Sans métrique (aucune resolution)", f"{pct_sans_res}%", sans_metric_res])

    # Onglet Satisfaction
    ws_satisfaction = wb.create_sheet(title="Satisfaction")
    ws_satisfaction.append(["Indicateur", "Valeur"])
    ws_satisfaction.append(["% Satisfaction Globale", f"{pct_satisfaction}%"])
    ws_satisfaction.append(["Nombre 'Good'", satisfaction_counts['good']])
    ws_satisfaction.append(["Nombre 'Bad'", satisfaction_counts['bad']])

    # Onglet tickets par type
    ws_type = wb.create_sheet(title="Tickets par Type")
    ws_type.append(["Type de ticket", "Nombre"])
    total_tickets = sum([tickets_par_type[t] for t in all_types])
    for ttype in all_types:
        count = tickets_par_type.get(ttype, 0)
        pct = round((count / total_tickets) * 100, 1) if total_tickets > 0 else 0
        ws_type.append([ttype, f"{count} ({pct}%)"])
    ws_type.append(["Total", total_tickets])

    # Onglet mensuel
    ws_mois = wb.create_sheet(title="Tickets par Mois")
    entete = ["Mois", "Incidents", "Questions", "Tasks", "Total tickets", "% Satisfaction", "Nb Avis"]
    ws_mois.append(entete)
    align_center = Alignment(horizontal='center', vertical='center')

    mois_tries = sorted(stats_mensuelles.keys())
    for mois_cle in mois_tries:
        data = stats_mensuelles[mois_cle]
        annee, mois_num = mois_cle.split('-')
        mois_fr_str = mois_fr[int(mois_num)-1] + " " + annee
        total_avis = data['good'] + data['bad']
        pct_sat = (data['good'] / total_avis * 100) if total_avis > 0 else 0.0
        ligne = [
            mois_fr_str,
            data['incident'],
            data['question'],
            data['task'],
            data['total'],
            f"{round(pct_sat, 1)}%",
            total_avis
        ]
        ws_mois.append(ligne)

    for row in ws_mois.iter_rows(min_row=2, max_row=ws_mois.max_row, min_col=1, max_col=len(entete)):
        for cell in row:
            cell.alignment = align_center

finally:
    wb.save("tickets_par_tags_et_delais.xlsx")
    print("✔ Fichier Excel généré : tickets_par_tags_et_delais.xlsx")

elapsed = round(time.time() - start_time, 2)
print(f"⏱ Terminé en {elapsed} secondes")
input("Appuyez sur Entrée pour fermer...")




