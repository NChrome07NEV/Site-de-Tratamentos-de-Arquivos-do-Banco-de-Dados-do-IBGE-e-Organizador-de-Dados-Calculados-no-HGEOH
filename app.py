import io
import re
import pandas as pd
import streamlit as st

# Configuração da página do aplicativo
st.set_page_config(
    page_title="Organizador e Formator de Arquivos Brutos das Estações Gravimétricas Gerados no Site do Banco de Dados do IBGE ",
    layout="wide",
)

# ==========================================
# CONTROLE DE ACESSO / SEGURANÇA
# ==========================================
st.title("Acesso Restrito - Organizador e Formatador de Arquivos Brutos das Estações Gravimétricas Gerados no Site do Banco de Dados do IBGE (RJ)")

SENHA_DO_PROJETO = "SIRAB_HGEO2026"

senha_digitada = st.text_input(
    "Digite a senha de acesso para liberar a ferramenta:", type="password"
)

if senha_digitada != SENHA_DO_PROJETO:
    if senha_digitada != "":
        st.error("Senha incorreta, tente novamente.")
    st.warning("Por favor, insira a senha fornecida pelo responsável pelo projeto para acessar o sistema.")
    st.stop()

st.success("Acesso autorizado com sucesso! Seja bem-vindo ao sistema.")

# ==========================================
# ESTRUTURA DO APLICATIVO EM 2 ETAPAS (ABAS)
# ==========================================
aba1, aba2 = st.tabs([
    "Etapa 1: Formatação de Coordenadas & Gerar TXT para o hgeoHNOR2020",
    "Etapa 2: Cruzamento pós-hgeoHNOR2020 & Relatórios Regionais (9 Arquivos)"
])

# Função auxiliar para formatar coordenadas
def formatar_coordenada(valor):
    s = str(valor).strip().replace(",", ".")
    if "." in s:
        partes = s.split(".")
        if len(partes[0].replace("-", "")) == 2:
            return s if s.startswith("-") else f"-{s}"
    s_limpo = s.replace("-", "").replace(".", "")
    if len(s_limpo) >= 2:
        graus = s_limpo[:2]
        decimais = s_limpo[2:]
        if not decimais:
            decimais = "00"
        return f"-{graus}.{decimais}"
    else:
        return s

# Função auxiliar para gerar Excel em memória
def converter_para_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Dados")
    return output.getvalue()

# Dicionário de Mapeamento Oficial dos 89 Municípios do RJ
mapa_regioes = {
    # Região Metropolitana
    'Rio de Janeiro': 'Região Metropolitana', 'Niterói': 'Região Metropolitana',
    'São Gonçalo': 'Região Metropolitana', 'Duque de Caxias': 'Região Metropolitana',
    'Nova Iguaçu': 'Região Metropolitana', 'Belford Roxo': 'Região Metropolitana',
    'São João de Meriti': 'Região Metropolitana', 'Nilópolis': 'Região Metropolitana',
    'Itaboraí': 'Região Metropolitana', 'Maricá': 'Região Metropolitana',
    'Magé': 'Região Metropolitana', 'Guapimirim': 'Região Metropolitana',
    'Itaguaí': 'Região Metropolitana', 'Seropédica': 'Região Metropolitana',
    'Paracambi': 'Região Metropolitana', 'Japeri': 'Região Metropolitana',
    'Queimados': 'Região Metropolitana', 'Tanguá': 'Região Metropolitana',
    'Rio Bonito': 'Região Metropolitana',

    # Região Médio Paraíba
    'Volta Redonda': 'Região Médio Paraíba', 'Barra Mansa': 'Região Médio Paraíba',
    'Resende': 'Região Médio Paraíba', 'Itatiaia': 'Região Médio Paraíba',
    'Piraí': 'Região Médio Paraíba', 'Pinheiral': 'Região Médio Paraíba',
    'Porto Real': 'Região Médio Paraíba', 'Quatis': 'Região Médio Paraíba',
    'Rio Claro': 'Região Médio Paraíba',

    # Região Costa Verde
    'Angra dos Reis': 'Região Costa Verde', 'Paraty': 'Região Costa Verde',
    'Mangaratiba': 'Região Costa Verde',

    # Região Serrana
    'Petrópolis': 'Região Serrana', 'Teresópolis': 'Região Serrana',
    'Nova Friburgo': 'Região Serrana', 'São José do Vale do Rio Preto': 'Região Serrana',
    'Sumidouro': 'Região Serrana', 'Bom Jardim': 'Região Serrana',
    'Duas Barras': 'Região Serrana', 'Carmo': 'Região Serrana',
    'Cachoeiras de Macacu': 'Região Serrana', 'Cantagalo': 'Região Serrana',

    # Região Norte Fluminense
    'Campos dos Goytacazes': 'Região Norte Fluminense', 'Macaé': 'Região Norte Fluminense',
    'Carapebus': 'Região Norte Fluminense', 'Quissamã': 'Região Norte Fluminense',
    'Conceição de Macabu': 'Região Norte Fluminense', 'São João da Barra': 'Região Norte Fluminense',
    'São Francisco de Itabapoana': 'Região Norte Fluminense',

    # Região Noroeste Fluminense
    'Itaperuna': 'Região Noroeste Fluminense', 'Bom Jesus do Itabapoana': 'Região Noroeste Fluminense',
    'Natividade': 'Região Noroeste Fluminense', 'Varre-Sai': 'Região Noroeste Fluminense',
    'Porciúncula': 'Região Noroeste Fluminense', 'Laje do Muriaé': 'Região Noroeste Fluminense',
    'Miracema': 'Região Noroeste Fluminense', 'Santo Antônio de Pádua': 'Região Noroeste Fluminense',
    'Aperibé': 'Região Noroeste Fluminense', 'Cambuci': 'Região Noroeste Fluminense',
    'Itaocara': 'Região Noroeste Fluminense', 'Italva': 'Região Noroeste Fluminense',
    'São José de Ubá': 'Região Noroeste Fluminense', 'Cardoso Moreira': 'Região Noroeste Fluminense',

    # Região Baixadas Litorâneas
    'Cabo Frio': 'Região Baixadas Litorâneas', 'Araruama': 'Região Baixadas Litorâneas',
    'Saquarema': 'Região Baixadas Litorâneas', 'Arraial do Cabo': 'Região Baixadas Litorâneas',
    'Armação dos Búzios': 'Região Baixadas Litorâneas', 'São Pedro da Aldeia': 'Região Baixadas Litorâneas',
    'Iguaba Grande': 'Região Baixadas Litorâneas', 'Casimiro de Abreu': 'Região Baixadas Litorâneas',
    'Silva Jardim': 'Região Baixadas Litorâneas', 'Rio das Ostras': 'Região Baixadas Litorâneas',

    # Região Centro-Sul Fluminense
    'Valença': 'Região Centro-Sul Fluminense', 'Vassouras': 'Região Centro-Sul Fluminense',
    'Barra do Piraí': 'Região Centro-Sul Fluminense', 'Três Rios': 'Região Centro-Sul Fluminense',
    'Paraíba do Sul': 'Região Centro-Sul Fluminense', 'Sapucaia': 'Região Centro-Sul Fluminense',
    'Areal': 'Região Centro-Sul Fluminense', 'Comendador Levy Gasparian': 'Região Centro-Sul Fluminense',
    'Mendes': 'Região Centro-Sul Fluminense', 'Miguel Pereira': 'Região Centro-Sul Fluminense',
    'Paty do Alferes': 'Região Centro-Sul Fluminense', 'Engenheiro Paulo de Frontin': 'Região Centro-Sul Fluminense',
    'Rio das Flores': 'Região Centro-Sul Fluminense', 'Santa Maria Madalena': 'Região Centro-Sul Fluminense',
    'Trajano de Moraes': 'Região Centro-Sul Fluminense', 'Macuco': 'Região Centro-Sul Fluminense',
    'Cordeiro': 'Região Centro-Sul Fluminense', 'São Sebastião do Alto': 'Região Centro-Sul Fluminense',
    'São Fidélis': 'Região Centro-Sul Fluminense'
}

# ==========================================
# ABA 1: TRATAMENTO INICIAL (GERAR TXT E EXCEL ORGANIZADO)
# ==========================================
with aba1:
    st.subheader("1️° - Upload do arquivo CSV bruto baixado do Banco de Dados do IBGE")
    uploaded_bruto = st.file_uploader(
        "Selecione o arquivo CSV bruto do IBGE:", type=["csv"], key="bruto"
    )

    if uploaded_bruto is not None:
        try:
            try:
                df_bruto = pd.read_csv(uploaded_bruto, sep=None, engine="python", encoding="utf-8")
            except Exception:
                uploaded_bruto.seek(0)
                df_bruto = pd.read_csv(uploaded_bruto, sep=None, engine="python", encoding="latin1")

            st.info(f"Arquivo bruto lido com sucesso; Total de estações: **{len(df_bruto)}**")

            # Identificação de colunas
            col_id = [c for c in df_bruto.columns if "id" in c.lower() or "ponto" in c.lower() or "estacao" in c.lower()][0]
            col_lat = [c for c in df_bruto.columns if "lat" in c.lower()][0]
            col_lon = [c for c in df_bruto.columns if "lon" in c.lower()][0]
            col_alt = [c for c in df_bruto.columns if "alt" in c.lower() or "h" in c.lower()][0]

            # Formatação de coordenadas
            df_bruto[col_lat] = df_bruto[col_lat].apply(formatar_coordenada)
            df_bruto[col_lon] = df_bruto[col_lon].apply(formatar_coordenada)

            # 1. Gerar TXT formatado para o hgeoHNOR2020
            txt_buffer = io.StringIO()
            for _, row in df_bruto.iterrows():
                ponto = str(row[col_id]).strip()
                lat = str(row[col_lat]).strip()
                lon = str(row[col_lon]).strip()
                alt = str(row[col_alt]).strip().replace(",", ".")
                txt_buffer.write(f"{ponto} {lat} {lon} {alt}\n")

            st.markdown("### Downloads da Etapa 1:")
            
            # Botão TXT
            st.download_button(
                label="Baixar Arquivo .TXT formatado (Entrada para o hgeoHNOR2020)",
                data=txt_buffer.getvalue(),
                file_name="Estacoes_Entrada_HGEOHNOR2020.txt",
                mime="text/plain",
            )

            # Botão Excel Organizado
            excel_bruto_bytes = converter_para_excel_bytes(df_bruto)
            st.download_button(
                label="Baixar Planilha Excel (.XLSX) Bruta Formatada e Organizada",
                data=excel_bruto_bytes,
                file_name="00_IBGE_Estacoes_Tratadas_Organizadas.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as e:
            st.error(f"Erro ao processar o arquivo bruto: {e}")

# ==========================================
# ABA 2: CRUZAMENTO PÓS-CALCULO & RELATÓRIOS REGIONAIS
# ==========================================
with aba2:
    st.subheader("2️° - Cruzamento de Dados e Geração dos 9 Relatórios Regionais")
    
    col1, col2 = st.columns(2)
    with col1:
        uploaded_ibge_org = st.file_uploader(
            "1. Selecione a planilha organizada (.xlsx) gerada na Etapa 1:", type=["xlsx"], key="ibge_org"
        )
    with col2:
        uploaded_hgeo_out = st.file_uploader(
            "2. Selecione o arquivo de saída calculado pelo hgeoHNOR2020 (.csv ou .txt):", type=["csv", "txt"], key="hgeo_out"
        )

    if uploaded_ibge_org is not None and uploaded_hgeo_out is not None:
        try:
            df_excel_org = pd.read_excel(uploaded_ibge_org)
            
            try:
                df_hgeo = pd.read_csv(uploaded_hgeo_out, sep=None, engine="python", encoding="utf-8")
            except Exception:
                uploaded_hgeo_out.seek(0)
                df_hgeo = pd.read_csv(uploaded_hgeo_out, sep=None, engine="python", encoding="latin1")

            st.info("Arquivos carregados com sucesso: Iniciando cruzamento e regionalização...")

            col_id_excel = [c for c in df_excel_org.columns if "estacao" in c.lower() or "id" in c.lower() or "ponto" in c.lower()][0]
            col_id_hgeo = [c for c in df_hgeo.columns if "id" in c.lower() or "ponto" in c.lower() or "estacao" in c.lower()][0]

            df_excel_org[col_id_excel] = df_excel_org[col_id_excel].astype(str).str.strip()
            df_hgeo[col_id_hgeo] = df_hgeo[col_id_hgeo].astype(str).str.strip()

            # Merge
            df_completo = pd.merge(df_hgeo, df_excel_org, left_on=col_id_hgeo, right_on=col_id_excel, how="left")

            # Organização de colunas
            colunas_para_frente = []
            for col_alvo in ["MUNICIPIO", "MEDICAOPLAN", "UF", "TEMA"]:
                if col_alvo in df_completo.columns:
                    colunas_para_frente.append(col_alvo)

            outras_colunas = [c for c in df_completo.columns if c != col_id_hgeo and c not in colunas_para_frente]
            ordem_final = [col_id_hgeo] + colunas_para_frente + outras_colunas
            df_completo = df_completo[ordem_final]

            # Ordenação cronológica decrescente
            if "MEDICAOPLAN" in df_completo.columns:
                df_completo["data_temp"] = pd.to_datetime(df_completo["MEDICAOPLAN"], errors="coerce")
                df_completo = df_completo.sort_values(by="data_temp", ascending=False).drop(columns=["data_temp"])

            # Mapeamento regional
            if "MUNICIPIO" in df_completo.columns:
                df_completo["REGIAO_OFICIAL"] = (
                    df_completo["MUNICIPIO"]
                    .astype(str)
                    .str.strip()
                    .map(mapa_regioes)
                    .fillna("Outras Regiões")
                )
            else:
                df_completo["REGIAO_OFICIAL"] = "Não Identificada"

            st.subheader("Lobby de Downloads dos Relatórios Finais")

            # 1. Download Relatório Geral
            geral_bytes = converter_para_excel_bytes(df_completo)
            st.download_button(
                label="Baixar Relatório Geral do Estado do Rio de Janeiro (Ordenado)",
                data=geral_bytes,
                file_name="01_Relatorio_Geral_Estado_Rio_Ordenado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            # 2. Downloads das 8 Regiões
            st.markdown("### Relatórios Regionais das 8 Regiões de Governo do RJ:")
            regioes_nomes = {
                "Região Metropolitana": "02_Regiao_Metropolitana.xlsx",
                "Região Médio Paraíba": "03_Regiao_Medio_Paraiba.xlsx",
                "Região Costa Verde": "04_Regiao_Costa_Verde.xlsx",
                "Região Serrana": "05_Regiao_Serrana.xlsx",
                "Região Norte Fluminense": "06_Regiao_Norte_Fluminense.xlsx",
                "Região Noroeste Fluminense": "07_Regiao_Noroeste_Fluminense.xlsx",
                "Região Baixadas Litorâneas": "08_Regiao_Baixadas_Litoraneas.xlsx",
                "Região Centro-Sul Fluminense": "09_Regiao_Centro_Sul_Fluminense.xlsx",
            }

            for regiao_nome, nome_arquivo in regioes_nomes.items():
                df_regiao = df_completo[df_completo["REGIAO_OFICIAL"] == regiao_nome]
                regiao_bytes = converter_para_excel_bytes(df_regiao)
                st.download_button(
                    label="Baixar {} ({} estações)".format(regiao_nome, len(df_regiao)),
                    data=regiao_bytes,
                    file_name=nome_arquivo,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=nome_arquivo,
                )

        except Exception as e:
            st.error(f"Erro ao processar o cruzamento pós-hgeoHNOR2020: {e}")
