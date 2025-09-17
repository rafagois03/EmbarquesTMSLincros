import pandas as pd
import streamlit as st
import requests
import json
import time
from datetime import datetime, timedelta
import io

# =======================
# CONFIGURAÇÕES DA PÁGINA
# =======================

st.set_page_config(
    page_title="Criação de Embarques de Devolução no TMS - 17.Setembro.2025",
    layout="wide"
)

st.title("🚚 Criador de Embarques em Massa - TMS - v1")
st.write(
    """
    **Desenvolvedor:** Thiago Nunes e Rafael Góis  
    **Descrição:** Com esta aplicação você será capaz de criar embarques de *devolução* em massa no TMS Lincros, a partir de dados em uma planilha Excel.
    """
)
# =======================
# DOWNLOAD TEMPLATE
# =======================

needed_columns = {
    "Protocolo", "CNPJ Unidade", "Calcular Carga", "Agrupar Conhecimentos", "CEP Origem",
    "CEP Destino", "Data Embarque", "Remetente CNPJ", "Remetente Nome", "Destinatário CNPJ",
    "Destinatário Nome", "Transportadora CNPJ", "Transportadora Nome", "CNPJ Emissor", "Nota Fiscal",
    "Série NF", "Documento Chave Acesso", "Pedido Série", "Pedido Número", "Motorista Documento",
    "Motorista Nome", "Motorista Tipo Documento", "Observação", "Identificador", "Embarque", "Link TMS", "FreteSPOT"}


st.write("📥 **Não sabe como montar o arquivo?** Baixe o modelo abaixo e preencha:")

# Cria um modelo vazio com as colunas obrigatórias
modelo_df = pd.DataFrame(columns=list(needed_columns))

# Salva em memória
output = io.BytesIO()
with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
    modelo_df.to_excel(writer, sheet_name='Modelo', index=False)

st.download_button(
    label="⬇️ Baixar Modelo de Planilha",
    data=output.getvalue(),
    file_name="MODELO_EMBARQUES_TMS.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)


# =======================
# UPLOAD DA PLANILHA
# =======================

required_columns = {
    "protocolo", "cnpj unidade", "calcular carga", "agrupar conhecimentos",
    "cep origem", "cep destino", "data embarque", "remetente cnpj", "remetente nome",
    "destinatário cnpj", "destinatário nome", "transportadora cnpj", "transportadora nome",
    "cnpj emissor", "nota fiscal", "série nf", "documento chave acesso",
    "motorista documento", "motorista nome", "motorista tipo documento", "observação",
    "identificador", "embarque"
    }

ARQUIVO_EXCEL = st.file_uploader(
    "📂 Faça upload do arquivo Excel com as informações do embarque",
    type=["xlsx"]
)

if ARQUIVO_EXCEL is not None:
    df = pd.read_excel(ARQUIVO_EXCEL, engine="openpyxl")
# novo trecho imprimir colunas
    st.write("🔍 **Colunas detectadas pelo pandas (ANTES de atribuir nomes):**")
    st.code(list(df.columns))
    st.write("Valores únicos da primeira linha (linha 0 do DataFrame):")
    st.code(df.iloc[0].tolist() if len(df) > 0 else "Arquivo vazio")
    
    df.columns = df.columns.str.strip().str.lower()  #deixando cabeçalho encontrado com letra minúscula
    

    # Verifica quais colunas estão faltando
    missing_columns = required_columns - set(df.columns)
    
    if missing_columns:
        st.error(f"❌ Arquivo Excel está faltando as seguintes colunas obrigatórias:\n\n{', '.join(sorted(missing_columns))}")
        st.write("💡 Dica: As colunas devem estar na ordem correta. Baixe o modelo para garantir compatibilidade.")
        st.stop()

    st.success("✅ Arquivo carregado com sucesso!")

    # =======================
    # BOTÃO PARA GERAR EMBARQUES
    # =======================
    if st.button("🚀 Gerar Cargas"):
        with st.spinner("⏳ Executando... isso pode levar alguns minutos..."):

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
                st.error(f"❌ Falha no login: {resp.status_code} - {resp.text}")
                st.stop()

            token = "B43Xe6ZwE6vxjOG2LPOIJZ0Z0ktMg1xEQ8ZYnoU3I8"  # token fixo
            st.write("✅ Login realizado com sucesso!")

            data_hoje = datetime.now()
            data_vencimento = (data_hoje + timedelta(days=30)).strftime("%Y-%m-%d")

            # =============================
            # 2) CRIAR EMBARQUES
            # =============================
            url_criar_embarque = "https://ws-tms.lincros.com/api/embarque/criarAsync"

            # Garantir colunas de controle
            for col in ["Protocolo", "Embarque", "FreteSPOT"]:
                if col not in df.columns:
                    df[col] = None

            embarques_json = []
            linhas_processadas = []

            st.write("\n🔍 Analisando linhas para criar embarque...\n")

            for idx, row in df.iterrows():
                protocolo = row["Protocolo"]
                embarque = row["Embarque"]

                if pd.notna(protocolo) or pd.notna(embarque):
                    st.write(f"🚫 Linha {idx + 2}: Já processada — pulando.")
                    continue

                linhas_processadas.append(idx)
                st.write(f"✅ Linha {idx + 2}: Criando embarque...")

                try:
                    chave_acesso = str(row["Documento Chave Acesso"]).strip()
                    if chave_acesso.lower() in ["nan", ""]:
                        chave_acesso = None
                except:
                    chave_acesso = None

                # Montar payload
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
                st.write(f"Status code: {response.status_code}")

                if response.status_code == 200:
                    protocolos = response.json().get("protocolo", [])
                    if protocolos and len(protocolos) == len(embarques_json):
                        for i, idx in enumerate(linhas_processadas):
                            df.at[idx, "Protocolo"] = protocolos[i]
                        st.write(f"✅ {len(protocolos)} protocolo(s) vinculados.")
                    else:
                        st.warning("⚠️ Número de protocolos não corresponde ao esperado.")
                else:
                    st.error(f"❌ Falha ao criar embarques: {response.text}")
                    st.stop()
            else:
                st.write("✅ Nenhum embarque novo para criar.")

            # =============================
            # 3) BUSCAR OID EMBARQUE PELO PROTOCOLO
            # =============================
            # Barra de progresso de 15 segundos
            st.write("⏳ Aguardando processamento do servidor (15 segundos)...")
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            for i in range(15):
                time.sleep(1)
                progress = (i + 1) / 15
                progress_bar.progress(int(progress * 100))
                status_text.text(f"Aguardando... {i + 1}/15 segundos")
            
            status_text.text("✅ Tempo de espera concluído!")
            time.sleep(0.5)  # pequena pausa visual
            progress_bar.empty()  # opcional: remove a barra depois
            status_text.empty()   # opcional: remove o texto depois

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
                        st.error(f"❌ Falha ao obter token de busca: {resp.text}")
                        return None
                except Exception as e:
                    st.error(f"❌ Erro de conexão: {e}")
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
                        st.error(f"❌ Erro na API ({protocolo}): {resp.status_code} - {resp.text[:200]}")
                        return None
                except Exception as e:
                    st.error(f"❌ Exceção ao buscar OID ({protocolo}): {e}")
                    return None

            token_busca = obter_token_busca()
            if not token_busca:
                st.error("❌ Não foi possível continuar a busca do OID.")
                st.stop()

            st.write("\n🔄 Buscando OID dos embarques...\n")
            for idx, row in df.iterrows():
                if pd.isna(row["Protocolo"]) or not pd.isna(row["Embarque"]):
                    continue

                try:
                    protocolo = int(row["Protocolo"])
                    oid = buscar_oid(token_busca, protocolo)
                    if oid:
                        df.at[idx, "Embarque"] = int(oid)
                        st.write(f"✔️ Linha {idx + 2}: Embarque {oid} vinculado.")
                    else:
                        df.at[idx, "Embarque"] = None
                except Exception as e:
                    st.error(f"❌ Erro ao processar linha {idx + 2}: {e}")
                    df.at[idx, "Embarque"] = None

            # =============================
            # 4) DOWNLOAD DO ARQUIVO ATUALIZADO
            # =============================
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name='BASE DE EMBARQUES', index=False)

            st.success("🎉 Embarque → Protocolo → Embarque ID → Concluído!")
            st.download_button(
                label="📥 Baixar Excel Atualizado",
                data=output.getvalue(),
                file_name="EMBARQUES_GERADOS_TMS.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )










