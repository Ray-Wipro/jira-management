import requests

# === CONFIGURAZIONE ===
JIRA_URL = "https://ejlog.atlassian.net"        # <-- Modifica con il tuo dominio
USERNAME = "rraimondi@ferrettogroup.com"        # <-- Tuo utente Atlassian
API_TOKEN = "ATATT3xFfGF0rFdHpQs25sHiT1VUn4-jmZl3nsdGWBB_hHAUMa1rePZsj758F_s7d0lWi7C0tCD4a-e8EOTQTvFZO_jtl3WYU0gG5tNv8bg2RkUDV24oYBapEB7-MQF7opXFmFREB1UCMDqRN_-ZZAkQL8xFceDjOowuNTx53CSyB1-HO1WJhUc=9EAB0E42"                  # <-- Generato da https://id.atlassian.com/manage/api-tokens

# === JQL per tutti i progetti ===
JQL = 'status in ("Da Gestire", "In corso", "Stand by Cliente", "Stand by Interno") ORDER BY project, priority DESC, created ASC'

# === Parametri API ===
url = f"{JIRA_URL}/rest/api/3/search"
headers = {
    "Accept": "application/json"
}
auth = (USERNAME, API_TOKEN)

params = {
    "jql": JQL,
    "fields": "summary,status,priority,created",
    "maxResults": 1000
}

# === ESECUZIONE RICHIESTA ===
response = requests.get(url, headers=headers, params=params, auth=auth)

if response.status_code == 200:
    issues = response.json().get("issues", [])
    print(f"Trovate {len(issues)} issue")

    with open("elenco_attivita.txt", "w", encoding="utf-8") as f:
        f.write("Progetto | Key | Titolo | Stato | Priorità | Data Creazione\n")
        f.write("-" * 90 + "\n")

        for issue in issues:
            key = issue["key"]
            project = key.split("-")[0]
            summary = issue["fields"]["summary"]
            status = issue["fields"]["status"]["name"]
            priority = issue["fields"].get("priority", {}).get("name", "Nessuna")
            created = issue["fields"]["created"][:10]
            f.write(f"{project} | {key} | {summary} | {status} | {priority} | {created}\n")

    print("✅ File 'elenco_attivita.txt' generato con successo.")
else:
    print(f"❌ Errore nella richiesta: {response.status_code}")
    print(response.text)