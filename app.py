import streamlit as st
import re
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from io import BytesIO
import time

# Conecta com Google Sheets via credenciais
def conectar_planilha_google(credenciais_json, sheet_url):
    scopes = ['https://www.googleapis.com/auth/spreadsheets']
    creds = Credentials.from_service_account_file(credenciais_json, scopes=scopes)
    cliente = gspread.authorize(creds)
    planilha = cliente.open_by_url(sheet_url)
    return planilha

# Lê planilha pública via CSV exportado
def ler_planilha_publica(sheet_url):
    id_match = re.search(r"/d/([a-zA-Z0-9-_]+)", sheet_url)
    gid_match = re.search(r"gid=([0-9]+)", sheet_url)
    if not id_match:
        raise ValueError("URL da planilha inválida.")
    sheet_id = id_match.group(1)
    gid = gid_match.group(1) if gid_match else "0"
    url_csv = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
    df = pd.read_csv(url_csv)
    return df

# Gera planilha Excel com o resultado
def gerar_excel_para_download(df_resultado):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_resultado.to_excel(writer, index=False, sheet_name='DIVERGENCIAS')
    output.seek(0)
    return output

# Interface Streamlit
def main():
    st.title("🔍 Conciliação Automática de Boletos - ISS")

    st.markdown("Preencha os dados do boleto manualmente e informe o link da planilha **Base**.")

    valor_total_guia = st.text_input("💰 Total Guia", help="Valor total do boleto (ex: 148,51)")
    vencimento_guia = st.text_input("📅 Vencimento (dd/mm/aaaa)", help="Data de vencimento da guia")
    juros_multa_guia = st.text_input("⚠️ Juros/Multa", help="Informe juros/multa se houver, senão 0")

    url_planilha = st.text_input("🔗 URL da planilha Google Sheets")
    arquivo_cred = st.file_uploader("🔐 credentials.json (opcional, para acesso via API)", type=["json"])
    filial = st.number_input("🏢 Número da Filial", min_value=1, step=1)

    # Estado da execução
    if 'executado' not in st.session_state:
        st.session_state.executado = False

    botao_label = "✅ Verificar Conciliação" if not st.session_state.executado else "🔁 Verificar Novamente"
    botao_cor = "primary" if not st.session_state.executado else "secondary"

    if st.button(botao_label, type="primary"):
        if not url_planilha or not filial:
            st.error("A URL da planilha e o número da filial são obrigatórios.")
            return

        try:
            valor_total = float(valor_total_guia.replace('.', '').replace(',', '.')) if valor_total_guia else None
            juros_multa = float(juros_multa_guia.replace('.', '').replace(',', '.')) if juros_multa_guia else 0.0
        except Exception:
            st.error("Erro ao interpretar os valores numéricos do boleto. Use formato 1234,56")
            return

        vencimento = vencimento_guia.strip() if vencimento_guia else None

        with st.spinner("🔄 Processando conciliação..."):
            time.sleep(1.2)  # Simula carregamento realista

            try:
                if arquivo_cred:
                    planilha = conectar_planilha_google(arquivo_cred, url_planilha)
                    worksheet = planilha.worksheet("Base")
                    dados = worksheet.get_all_records()
                    df_base = pd.DataFrame(dados)
                else:
                    df_base = ler_planilha_publica(url_planilha)

                col_essenciais = [
                    'FILIAL',
                    'NUM_TITULO_ISS',
                    'RAZAO_SOCIAL_PREFEITURA',
                    'VLR_SERVICO',
                    'VLR_ISS',
                    'CONVÊNIO',
                    'RESPONSÁVEL',
                    'DATA_EMISSAO',
                    'MES_ANO'
                ]

                col_faltantes = [col for col in col_essenciais if col not in df_base.columns]
                if col_faltantes:
                    st.error(f"Faltam colunas na planilha: {', '.join(col_faltantes)}")
                    return

                df_filial = df_base[df_base['FILIAL'] == filial]
                if df_filial.empty:
                    st.warning("Nenhum registro encontrado para essa filial.")
                    return

                for col in ['VLR_SERVICO', 'VLR_ISS']:
                    df_filial[col] = df_filial[col].astype(str).str.replace('.', '').str.replace(',', '.').astype(float)

                soma_iss = df_filial['VLR_ISS'].sum()
                valor_comparar = valor_total if valor_total is not None else soma_iss
                diferenca = round(valor_comparar - soma_iss, 2)

                df_filial['VALOR GUIA (INFORMADO)'] = valor_total if valor_total is not None else 'Não informado'
                df_filial['JUROS/MULTA (INFORMADO)'] = juros_multa
                df_filial['VENCIMENTO (INFORMADO)'] = vencimento if vencimento else 'Não informado'
                df_filial['SOMA VLR_ISS (PLANILHA)'] = soma_iss
                df_filial['DIFERENÇA FINAL'] = diferenca

                st.success("✅ Conciliação concluída com sucesso!")

                st.subheader("📄 Resultado da Conciliação")
                st.dataframe(df_filial)

                excel = gerar_excel_para_download(df_filial)
                st.download_button(
                    label="⬇️ Baixar Excel da Conciliação",
                    data=excel,
                    file_name=f"conciliacao_filial_{filial}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                st.session_state.executado = True

            except Exception as e:
                st.error(f"Erro ao processar: {e}")

if __name__ == "__main__":
    main()
