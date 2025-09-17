import pandas as pd
import streamlit as st
import requests
import json
import time
from datetime import datetime, timedelta
import io


# =======================
# CONFIGURAÇÕES DA PÁGINA INICIAL
# =======================

st.set_page_config(
    page_title="Criação de Embarques de Devolução no TMS - 17.Setembro.2025",
    layout="wide"
)

st.title("🚚 Criador de Embarques em Massa - TMS - v1")
st.write(
    """
    **Desenvolvedor:** Thiago Nunes e Rafael Góis
    \\
    **Descrição:** Com esta aplicação você será capaz de criar embarques de *devolução* em massa no TMS Lincros, a partir de dados em uma planilha Excel.
    """
)

ARQUIVO_EXCEL = st.file_uploader(
    "📂 Faça upload do arquivo Excel com as informações do embarque",
    type=["xlsx"]
)

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file, engine="openpyxl")
    df.columns = df.columns.str.lower().str.strip()

    required_columns = {"Protocolo", "CNPJ Unidade", "Calcular Carga", "Agrupar Conhecimentos", "CEP Origem", "CEP Destino", "Data Embarque", "Remetente CNPJ", "Remetente Nome", "Destinatário CNPJ", "Destinatário Nome", "Transportadora CNPJ", "Transportadora Nome", "CNPJ Emissor", "Nota Fiscal", "Série NF", "Documento Chave Acesso", "Pedido Série", "Pedido Número", "Motorista Documento", "Motorista Nome", "Motorista Tipo Documento", "Observação", "Identificador", "Embarque"}
    if not required_columns.issubset(df.columns):
        st.error(f"Arquivo faltando colunas obrigatórias: {required_columns}")
        st.stop()

    st.success("✅ Arquivo carregado com sucesso!")

# =======================
    # Botão para gerar os embarques
# =======================
    if st.button("🚀 Gerar Cargas"):
        with st.spinner("⏳ Executando otimização... isso pode levar alguns minutos...")

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
            
            st.write("🔐 Realizando login...")
            resp = requests.post(login_url, json=login_payload, headers=login_headers)
            
            if resp.status_code != 200:
                st.write("❌ Falha no login:", resp.status_code, resp.text)
                exit()
            
            token = "B43Xe6ZwE6vxjOG2LPOIJZ0Z0ktMg1xEQ8ZYnoU3I8" # token novo fixo
            st.write("✅ Login recebido com sucesso!")
            
            
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
            
            st.write("\n🔍 Analisando linhas para criar embarque...\n")
            
            for idx, row in df.iterrows():
                protocolo = row["Protocolo"]
                embarque = row["Embarque"]
            
                # Pular se já tiver protocolo ou número de embarque
                if pd.notna(protocolo) or pd.notna(embarque):
                    st.write(f"🚫 Linha {idx + 2}: Já processada — pulando.")
                    continue
            
                # Marcar para processar
                linhas_processadas.append(idx)
                st.write(f"✅ Linha {idx + 2}: Criando embarque...")
            
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
            
                st.write(f"\n📤 Enviando {len(embarques_json)} embarque(s)...")
                response = requests.post(url_criar_embarque, json=payload_envio, headers=headers_envio)
                st.write("Status code:", response.status_code)
            
                if response.status_code == 200:
                    protocolos = response.json().get("protocolo", [])
                    if protocolos and len(protocolos) == len(embarques_json):
                        # Atualiza apenas as linhas processadas
                        for i, idx in enumerate(linhas_processadas):
                            df.at[idx, "Protocolo"] = protocolos[i]
                        st.write(f"✅ {len(protocolos)} protocolo(s) vinculados.")
                    else:
                        st.write("⚠️ Número de protocolos não corresponde ao esperado.")
                else:
                    st.write("❌ Falha ao criar embarques:", response.text)
                    exit()
            else:
                st.write("✅ Nenhum embarque novo para criar.")
            
            # Salvar protocolos
            df.to_excel(ARQUIVO_EXCEL, index=False)
            st.write(f"📄 Planilha salva com protocolos atualizados.")
            
            
            # =============================
            # 3) BUSCAR OID EMBARQUE PELO PROTOCOLO
            # =============================
            st.write("\n⏳ Aguardando 15 segundos antes de buscar os OIDs...")
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
                        st.write("❌ Falha ao obter token de busca:", resp.text)
                        return None
                except Exception as e:
                    st.write("❌ Erro de conexão:", e)
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
                        st.write(f"❌ Erro na API ({protocolo}): {resp.status_code} - {resp.text[:200]}")
                        return None
                except Exception as e:
                    st.write(f"❌ Exceção ao buscar OID ({protocolo}): {e}")
                    return None
            
            # Obter token
            token_busca = obter_token_busca()
            if not token_busca:
                st.write("❌ Não foi possível continuar a busca do OID.")
                exit()
            
            # Buscar OID para cada linha com protocolo novo
            st.write("\n🔄 Buscando OID dos embarques...\n")
            for idx, row in df.iterrows():
                if pd.isna(row["Protocolo"]) or not pd.isna(row["Embarque"]):
                    continue  # já tem ou não tem protocolo
            
                try:
                    protocolo = int(row["Protocolo"])
                    oid = buscar_oid(token_busca, protocolo)
                    if oid:
                        df.at[idx, "Embarque"] = int(oid)
                        st.write(f"✔️ Linha {idx + 2}: Embarque {oid} vinculado.")
                    else:
                        df.at[idx, "Embarque"] = None
                except Exception as e:
                    st.write(f"❌ Erro ao processar linha {idx + 2}: {e}")
                    df.at[idx, "Embarque"] = None
            
            # Salvar
            df.to_excel(ARQUIVO_EXCEL, index=False)
            st.write(f"📄 Planilha atualizada com números de embarque.")


# =============================
# 5) SALVAR FINAL
# =============================
            df.to_excel(ARQUIVO_EXCEL, index=False)
            st.write(f"\n✅ Processo completo! Planilha final salva em: {ARQUIVO_EXCEL}")
            
            st.write("🎉 Embarque → Protocolo → Embarque ID  → Concluído!")



###############################
# 6) DOWNLOAD ARQUIVO FINAL
###############################



            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                ARQUIVO_EXCEL.to_excel(writer, sheet_name='BASE DE EMBARQUES', index=False)
    
            st.download_button(
                "📥 Baixar Excel",
                data=output.getvalue(),
                file_name="EMBARQUES_GERADOS TMS.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
