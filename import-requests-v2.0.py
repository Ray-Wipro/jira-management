import requests
from datetime import datetime

# === CONFIGURAZIONE ===
JIRA_URL = "https://ejlog.atlassian.net"        # <-- Modifica con il tuo dominio
USERNAME = "rraimondi@ferrettogroup.com"        # <-- Tuo utente Atlassian
API_TOKEN = "ATATT3xFfGF0rFdHpQs25sHiT1VUn4-jmZl3nsdGWBB_hHAUMa1rePZsj758F_s7d0lWi7C0tCD4a-e8EOTQTvFZO_jtl3WYU0gG5tNv8bg2RkUDV24oYBapEB7-MQF7opXFmFREB1UCMDqRN_-ZZAkQL8xFceDjOowuNTx53CSyB1-HO1WJhUc=9EAB0E42"                  # <-- Generato da https://id.atlassian.com/manage/api-tokens

# === JQL ===
JQL = 'status in ("Da Gestire", "In corso", "Stand by Cliente", "Stand by Interno") ORDER BY priority DESC, duedate ASC, created ASC'

# === PARAMETRI RICHIESTA ===
url = f"{JIRA_URL}/rest/api/3/search"
headers = {"Accept": "application/json"}
auth = (USERNAME, API_TOKEN)

params = {
    "jql": JQL,
    "fields": "summary,status,priority,created,duedate,project",
    "maxResults": 1000
}

# === ESECUZIONE ===
response = requests.get(url, headers=headers, params=params, auth=auth)

if response.status_code != 200:
    print("Errore:", response.status_code)
    print(response.text)
    exit()

issues = response.json().get("issues", [])
print(f"Trovate {len(issues)} issue")

# === CATEGORIZZAZIONE PER PRIORITÀ ===
priorities = {
    "High": [],
    "Medium": [],
    "Low": [],
    "Nessuna": []
}

def parse_date(date_str):
    return datetime.strptime(date_str, "%Y-%m-%d") if date_str else None

for issue in issues:
    fields = issue["fields"]
    key = issue["key"]
    project = fields["project"]["name"]
    title = fields["summary"]
    status = fields["status"]["name"]
    priority = fields.get("priority", {}).get("name", "Nessuna")
    duedate = fields.get("duedate")  # può essere None
    created = fields["created"][:10]

    priorities.setdefault(priority, []).append({
        "cliente": project,
        "titolo": title,
        "stato": status,
        "scadenza": duedate,
        "creazione": created
    })

# === ORDINAMENTO E STAMPA ===
output_lines = []

for prio_label in ["High", "Medium", "Low", "Nessuna"]:
    blocco = priorities.get(prio_label, [])
    if not blocco:
        continue

    output_lines.append("##############################")
    output_lines.append(f"# {prio_label.upper()} PRIORITY")
    output_lines.append("##############################\n")

    # ordina per scadenza (se c'è) poi per creazione
    blocco.sort(key=lambda x: (
        parse_date(x["scadenza"]) or datetime.max,
        parse_date(x["creazione"])
    ))

    for item in blocco:
        scad = f", scad. {item['scadenza']}" if item["scadenza"] else ""
        line = f"{item['cliente']} - {item['titolo']} ({item['stato']}{scad})"
        output_lines.append(line)

    output_lines.append("\n")

# === SCRITTURA SU FILE ===
with open("elenco_attivita.txt", "w", encoding="utf-8") as f:
    f.write("\n".join(output_lines))

print("✅ File 'elenco_attivita.txt' generato correttamente.")
