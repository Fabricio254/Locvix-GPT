import os
import requests
import json

FULLTRACK_API_KEY = os.environ.get("FULLTRACK_API_KEY", "530fdb8a61907a2f9904477a335f1a8eee0ea5d9")
FULLTRACK_SECRET_KEY = os.environ.get("FULLTRACK_SECRET_KEY", "8c002f8a04533e2f2ed428820f89dca5b3e9996a")
FULLTRACK_BASE = "https://ws.fulltrack2.com"

NOME_ALVO = "I/MO ZOOMLION QY 60V"

# Testa todos os endpoints conhecidos
endpoints = [
    "/events/all",
    "/vehicles/all",
    "/events/interval",  # requer params
    "/events/last",
    "/events/last/1276697",  # id do veículo
]

for ep in endpoints:
    url = FULLTRACK_BASE + ep
    print(f"\n--- {url} ---")
    try:
        if ep == "/events/interval":
            # Exemplo: últimos 7 dias
            from datetime import datetime, timedelta
            dt_ini = (datetime.now() - timedelta(days=7)).strftime("%d/%m/%Y")
            dt_fim = datetime.now().strftime("%d/%m/%Y")
            params = {"id": "1276697", "data_ini": dt_ini, "data_fim": dt_fim}
            r = requests.get(url, headers={"apikey": FULLTRACK_API_KEY, "secretkey": FULLTRACK_SECRET_KEY}, params=params, timeout=15)
        else:
            r = requests.get(url, headers={"apikey": FULLTRACK_API_KEY, "secretkey": FULLTRACK_SECRET_KEY}, timeout=15)
        r.raise_for_status()
        data = r.json()
        # Procura pelo veículo alvo
        found = False
        if isinstance(data, dict):
            for k, v in data.items():
                if isinstance(v, list):
                    for item in v:
                        nome = (item.get("ras_vei_veiculo") or item.get("nome") or "").strip()
                        if nome.upper() == NOME_ALVO.upper():
                            print(json.dumps(item, indent=2, ensure_ascii=False))
                            found = True
        if not found:
            print(f"Equipamento '{NOME_ALVO}' não encontrado neste endpoint.")
    except Exception as e:
        print(f"Erro ao consultar {ep}: {e}")
