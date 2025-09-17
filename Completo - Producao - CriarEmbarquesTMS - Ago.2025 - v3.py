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

st.title("🚚 Criador de Embarques em Massa - TMS Lincros")
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
    if st.button("🚀 Gerar Embarques de Devolução!"):
        with st.spinner("⏳ Executando... isso pode levar alguns minutos..."):

            # =============================
            # 1) LOGIN PARA OBTER TOKEN
            # =============================
            login_url = "https://ws-tms.lincros.com/api/auth/login"

            if "lincros" not in st.secrets:
                st.error("❌ Secrets não configurados. Configure em ⚙️ Settings > Secrets no Streamlit Cloud.")
                st.stop()

            login_payload = {
                "login": st.secrets["lincros"]["login"],
                "senha": st.secrets["lincros"]["senha"]
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

            token = st.secrets["lincros"]["token"]
            st.write("✅ Login realizado com sucesso!")

            data_hoje = datetime.now()
            data_vencimento = (data_hoje + timedelta(days=30)).strftime("%Y-%m-%d")

            # =============================
            # 2) CRIAR EMBARQUES
            # =============================
            url_criar_embarque = "https://ws-tms.lincros.com/api/embarque/criarAsync"

            # Garantir colunas de controle
            for col in ["protocolo", "embarque", "fretespot"]:
                if col not in df.columns:
                    df[col] = None

            embarques_json = []
            linhas_processadas = []

            st.write("\n🔍 Analisando linhas para criar embarque...\n")

            for idx, row in df.iterrows():
                protocolo = row["protocolo"]
                embarque = row["embarque"]

                if pd.notna(protocolo) or pd.notna(embarque):
                    st.write(f"🚫 Linha {idx + 2}: Já processada — pulando.")
                    continue

                linhas_processadas.append(idx)
                st.write(f"✅ Linha {idx + 2}: Criando embarque...")

                try:
                    chave_acesso = str(row["documento chave acesso"]).strip()
                    if chave_acesso.lower() in ["nan", ""]:
                        chave_acesso = None
                except:
                    chave_acesso = None

                # Montar payload
                embarque_data = {
                    "cnpjUnidade": str(row["cnpj unidade"]).strip(),
                    "calcularcarga": False,
                    "agruparConhecimentos": True,
                    "remetente": {
                        "cnpj": str(row["remetente cnpj"]).strip(),
                        "marcadores": ["DEVOLUCAO"]
                    },
                    "destinatario": {
                        "cnpj": str(row["destinatário cnpj"]).strip(),
                        "marcadores": ["DEVOLUCAO"]
                    },
                    "transportadora": {
                        "cnpj": str(row["transportadora cnpj"]).strip(),
                        "marcadores": ["DEVOLUCAO"]
                    },
                    "cepOrigem": int(row["cep origem"]),
                    "cepDestino": int(row["cep destino"]),
                    "documentos": [
                        {
                            "tipoDocumento": 0,
                            "cnpjEmissor": int(row["cnpj emissor"]),
                            "numeroDocumento": int(row["nota fiscal"]),
                            "serie": int(row["série nf"]),
                            "chaveAcesso": chave_acesso,
                            "tipoConhecimento": 3,
                            "grupoVeiculo": "3517"
                        }
                    ],
                    "marcadores": ["DEVOLUCAO"],
                    "grupoVeiculo": "3517",
                    "observacao": str(row.get("observação", "")).strip(),
                    "identificador": str(row.get("identificador", "")).strip(),
                    "motoristas": [
                        {
                            "documento": str(row.get("motorista documento", "")).strip(),
                            "nome": str(row.get("motorista nome", "")).strip(),
                            "tipoDocumento": int(row.get("motorista tipo documento", 1))
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
                            df.at[idx, "protocolo"] = protocolos[i]
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
           # Calcula o tempo total de espera: 7 segundos por embarque processado
            total_embarques = len(linhas_processadas)  # <-- assumindo que você tem essa lista
            tempo_total = max(5, total_embarques * 7)  # mínimo de 5s para evitar espera muito curta
            
            if total_embarques > 0:
                st.write(f"⏳ Aguardando processamento do servidor ({tempo_total} segundos)...")
                progress_bar = st.progress(0)
                status_text = st.empty()
            
                for i in range(tempo_total):
                    time.sleep(1)
                    progress = (i + 1) / tempo_total
                    progress_bar.progress(int(progress * 100))
                    status_text.text(f"Aguardando... {i + 1}/{tempo_total} segundos")
            
                status_text.text("✅ Tempo de espera concluído!")
                time.sleep(0.5)
                progress_bar.empty()
                status_text.empty()
            else:
                st.info("ℹ️ Nenhum embarque criado — pulando espera.")

            def obter_token_busca():
                url = "https://ws-tms.lincros.com/api/auth/login"
                headers = {
                    "accept": "text/plain",
                    "Content-Type": "application/json"
                }
                payload = {
                    "login": st.secrets["lincros"]["login"],
                    "senha": st.secrets["lincros"]["senha"]
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
                if pd.isna(row["protocolo"]) or not pd.isna(row["embarque"]):
                    continue

                try:
                    protocolo = int(row["protocolo"])
                    oid = buscar_oid(token_busca, protocolo)
                    if oid:
                        df.at[idx, "embarque"] = int(oid)
                        st.write(f"✔️ Linha {idx + 2}: Embarque {oid} vinculado.")
                    else:
                        df.at[idx, "embarque"] = None
                except Exception as e:
                    st.error(f"❌ Erro ao processar linha {idx + 2}: {e}")
                    df.at[idx, "embarque"] = None

            # =============================
            # 4) DOWNLOAD DO ARQUIVO ATUALIZADO
            # =============================
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name='BASE DE EMBARQUES', index=False)

            st.success("🎉 Embarque → Protocolo → Embarque ID → Concluído!")
            st.balloons()  # 🎈 Aqui! Comemoramos o sucesso do código... deu trabalho demais bixo kkk!    
            st.markdown(
                """
                📌 **Acompanhe o processamento ou visualize as cargas no TMS:**
                [👉 Acessar TMS - Lista de Importação de Embarques](https://generalmills-tms.lincros.com/default/cadastro/importacaoArquivo/listarImportacaoEmbarque.xhtml?s=1)
                """,
                unsafe_allow_html=True
            )
            
            st.download_button(
                label="📥 Baixar Excel Atualizado",
                data=output.getvalue(),
                file_name="EMBARQUES_GERADOS_TMS.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )















