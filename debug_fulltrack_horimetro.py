import os
import requests

FULLTRACK_API_KEY = os.environ.get("FULLTRACK_API_KEY", "530fdb8a61907a2f9904477a335f1a8eee0ea5d9")
FULLTRACK_SECRET_KEY = os.environ.get("FULLTRACK_SECRET_KEY", "8c002f8a04533e2f2ed428820f89dca5b3e9996a")
FULLTRACK_BASE = "https://ws.fulltrack2.com"

# Nome do equipamento a debugar
NOME_ALVO = "I/MO ZOOMLION QY 60V"

response = requests.get(
    f"{FULLTRACK_BASE}/events/all",
    headers={"apikey": FULLTRACK_API_KEY, "secretkey": FULLTRACK_SECRET_KEY},
    timeout=15,
)
response.raise_for_status()
dados = response.json().get("data") or []

for evento in dados:
    nome = (evento.get("ras_vei_veiculo") or "").strip()
    if nome.upper() == NOME_ALVO.upper():
        print(f"Equipamento: {nome}")
        print(f"ras_eve_horimetro: {evento.get('ras_eve_horimetro')}")
        print(f"Evento completo: {evento}")
        break
else:
    print(f"Equipamento '{NOME_ALVO}' não encontrado na resposta da API.")
