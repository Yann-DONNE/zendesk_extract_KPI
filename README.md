# zendesk_extract_KPI
🎯 Objectif du script :
Ce script permet d’extraire automatiquement depuis Zendesk les indicateurs suivants sur la période de l'année en cours ou sur une période définie de date à date :
📊 Répartition des tickets par Tag commercial (tous ce qui commence par "COM")
⏱️ Délai de première prise en charge des tickets
⏳ Délai de résolution complète
✅ Taux de service (satisfaction)
🗂️ Répartition des tickets par type (incident, tâche, question)

⚙️ Configuration requise pour une utilisation personnelle :
Si tu souhaites adapter ce script à ton propre environnement Zendesk, tu devras modifier les paramètres suivants dans le code Python (à l'aide d'un éditeur tel que Visual Studio Code) :
	https://code.visualstudio.com/docs/?dv=win64user

# Paramètres Zendesk à personnaliser dans le script : 
SUBDOMAIN = 'TonDomaineZendesk'         # Modifier si le sous-domaine a changé
EMAIL = 'ton_email@exemple.com' # Remplacer par l’adresse e-mail liée à ton compte Zendesk
API_TOKEN = 'ton_token_API'     # Remplacer par la clé API générée dans Zendesk
🖥️ Étapes d'installation :
1️⃣ Installer Python :
Télécharge et installe Python depuis :
	 https://www.python.org/downloads/windows/
⚠️ Pendant l'installation, coche l'option "Add Python to PATH".

2️⃣ Installer les bibliothèques nécessaires :
Ouvre l’invite de commande (cmd) et tape :
pip install requests openpyxl tqdm python-dotenv
3️⃣ Lancer le script :
1.	Soit en double-cliquant sur le fichier .py
2.	Soit en ligne de commande :
python zendesk_extract_KPI.py
3.	Soit en faisant clic droit et « ouvrir avec Python »
4️⃣ Laissez-vous guider pas à pas par l’application
À la fin de l’exécution, un fichier Excel (.xlsx) sera automatiquement créé dans le même dossier que zendesk_extract_KPI.
L’application vous demandera d’appuyer sur Entrée pour se fermer.
Si elle se ferme toute seule, cela signifie qu’une erreur est survenue pendant le traitement. Dans ce cas, il suffit de relancer l’application tout simplement.

"Ce script est conçu pour extraire uniquement les tickets dont le type est : problem, task, question ou incident. Pensez à l'adapter si vous utilisez d'autres types ou une classification différente dans Zendesk.
