import pandas as pd
import requests
import json
import time
from datetime import datetime, timedelta

# =============================
# CONFIGURAÇÕES
# =============================
ARQUIVO_EXCEL = 'embarques.xlsx'

# =============================
# 1) LOGIN PARA OBTER TOKEN
# =============================
login_url = "https://ws-tms.lincros.com/api/auth/login"
login_payload = {
    "login": "thiagonunes910@hotmail.com",
    "senha": "Lincros@25"
}
login_headers = {
    "accept": "text/plain",
    "Content-Type": "application/json"
}

print("🔐 Realizando login...")
resp = requests.post(login_url, json=login_payload, headers=login_headers)

if resp.status_code != 200:
    print("❌ Falha no login:", resp.status_code, resp.text)
    exit()

token = "B43Xe6ZwE6vxjOG2LPOIJZ0Z0ktMg1xEQ8ZYnoU3I8" # token novo fixo
print("✅ Login recebido com sucesso!")


data_hoje = datetime.now()
data_vencimento = (data_hoje + timedelta(days=30)).strftime("%Y-%m-%d")

# =============================
# 2) CRIAR EMBARQUES
# =============================
url_criar_embarque = "https://ws-tms.lincros.com/api/embarque/criarAsync"

# Ler planilha
df = pd.read_excel(ARQUIVO_EXCEL)

# Garantir colunas de controle
for col in ["Protocolo", "Embarque", "FreteSPOT"]:
    if col not in df.columns:
        df[col] = None

embarques_json = []
linhas_processadas = []  # Armazena os índices que foram processados nesta execução

print("\n🔍 Analisando linhas para criar embarque...\n")

for idx, row in df.iterrows():
    protocolo = row["Protocolo"]
    embarque = row["Embarque"]

    # Pular se já tiver protocolo ou número de embarque
    if pd.notna(protocolo) or pd.notna(embarque):
        print(f"🚫 Linha {idx + 2}: Já processada — pulando.")
        continue

    # Marcar para processar
    linhas_processadas.append(idx)
    print(f"✅ Linha {idx + 2}: Criando embarque...")

    try:
        chave_acesso = str(row["Documento Chave Acesso"]).strip()
        if chave_acesso.lower() in ["nan", ""]:
            chave_acesso = None
    except:
        chave_acesso = None

    # Montar payload do embarque
    embarque_data = {
        "cnpjUnidade": str(row["CNPJ Unidade"]).strip(),
        "calcularcarga": False,
        "agruparConhecimentos": True,
        "remetente": {
            "cnpj": str(row["Remetente CNPJ"]).strip(),
            "marcadores": ["DEVOLUCAO"]
        },
        "destinatario": {
            "cnpj": str(row["Destinatário CNPJ"]).strip(),
            "marcadores": ["DEVOLUCAO"]
        },
        "transportadora": {
            "cnpj": str(row["Transportadora CNPJ"]).strip(),
            "marcadores": ["DEVOLUCAO"]
        },
        "cepOrigem": int(row["CEP Origem"]),
        "cepDestino": int(row["CEP Destino"]),
        "documentos": [
            {
                "tipoDocumento": 0,
                "cnpjEmissor": int(row["CNPJ Emissor"]),
                "numeroDocumento": int(row["Nota Fiscal"]),
                "serie": int(row["Série NF"]),
                "chaveAcesso": chave_acesso,
                "tipoConhecimento": 3,
                "grupoVeiculo": "3517"
            }
        ],
        
        "marcadores": ["DEVOLUCAO"],
        "grupoVeiculo": "3517",
        "observacao": str(row.get("Observação", "")).strip(),
        "identificador": str(row.get("Identificador", "")).strip(),
        "motoristas": [
            {
                "documento": str(row.get("Motorista Documento", "")).strip(),
                "nome": str(row.get("Motorista Nome", "")).strip(),
                "tipoDocumento": int(row.get("Motorista Tipo Documento", 1))
            }
        ]
    }
    embarques_json.append(embarque_data)

# Enviar embarques
if embarques_json:
    payload_envio = {"embarques": embarques_json}
    headers_envio = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {token}"
    }

    print(f"\n📤 Enviando {len(embarques_json)} embarque(s)...")
    response = requests.post(url_criar_embarque, json=payload_envio, headers=headers_envio)
    print("Status code:", response.status_code)

    if response.status_code == 200:
        protocolos = response.json().get("protocolo", [])
        if protocolos and len(protocolos) == len(embarques_json):
            # Atualiza apenas as linhas processadas
            for i, idx in enumerate(linhas_processadas):
                df.at[idx, "Protocolo"] = protocolos[i]
            print(f"✅ {len(protocolos)} protocolo(s) vinculados.")
        else:
            print("⚠️ Número de protocolos não corresponde ao esperado.")
    else:
        print("❌ Falha ao criar embarques:", response.text)
        exit()
else:
    print("✅ Nenhum embarque novo para criar.")

# Salvar protocolos
df.to_excel(ARQUIVO_EXCEL, index=False)
print(f"📄 Planilha salva com protocolos atualizados.")


# =============================
# 3) BUSCAR OID EMBARQUE PELO PROTOCOLO
# =============================
print("\n⏳ Aguardando 15 segundos antes de buscar os OIDs...")
time.sleep(15)

def obter_token_busca():
    url = "https://ws-tms.lincros.com/api/auth/login"
    headers = {
        "accept": "text/plain",
        "Content-Type": "application/json"
    }
    payload = {
        "login": "thiagonunes910@hotmail.com",
        "senha": "Lincros@25"
    }
    try:
        resp = requests.post(url, json=payload, headers=headers)
        if resp.status_code == 200:
            return resp.text.strip()
        else:
            print("❌ Falha ao obter token de busca:", resp.text)
            return None
    except Exception as e:
        print("❌ Erro de conexão:", e)
        return None

def buscar_oid(token, protocolo):
    url = "https://ws-tms.lincros.com/api/embarque/recuperarDados"
    headers = {
        "accept": "application/json",
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    try:
        resp = requests.post(url, json={"protocolo": protocolo}, headers=headers)
        if resp.status_code == 200:
            dados = resp.json()
            return dados.get("embarque", {}).get("oidEmbarque")
        else:
            print(f"❌ Erro na API ({protocolo}): {resp.status_code} - {resp.text[:200]}")
            return None
    except Exception as e:
        print(f"❌ Exceção ao buscar OID ({protocolo}): {e}")
        return None

# Obter token
token_busca = obter_token_busca()
if not token_busca:
    print("❌ Não foi possível continuar a busca do OID.")
    exit()

# Buscar OID para cada linha com protocolo novo
print("\n🔄 Buscando OID dos embarques...\n")
for idx, row in df.iterrows():
    if pd.isna(row["Protocolo"]) or not pd.isna(row["Embarque"]):
        continue  # já tem ou não tem protocolo

    try:
        protocolo = int(row["Protocolo"])
        oid = buscar_oid(token_busca, protocolo)
        if oid:
            df.at[idx, "Embarque"] = int(oid)
            print(f"✔️ Linha {idx + 2}: Embarque {oid} vinculado.")
        else:
            df.at[idx, "Embarque"] = None
    except Exception as e:
        print(f"❌ Erro ao processar linha {idx + 2}: {e}")
        df.at[idx, "Embarque"] = None

# Salvar
df.to_excel(ARQUIVO_EXCEL, index=False)
print(f"📄 Planilha atualizada com números de embarque.")


# =============================
# 5) SALVAR FINAL
# =============================
df.to_excel(ARQUIVO_EXCEL, index=False)
print(f"\n✅ Processo completo! Planilha final salva em: {ARQUIVO_EXCEL}")
print("🎉 Embarque → Protocolo → Embarque ID  → Concluído!")