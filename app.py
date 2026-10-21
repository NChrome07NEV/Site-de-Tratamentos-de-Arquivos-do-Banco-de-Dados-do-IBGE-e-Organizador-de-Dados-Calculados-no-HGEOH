import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.styles import Font, Alignment
#Sexto Teste: após algumas edições, adições e correções do programa, a ideia permanece a mesma de gerar os arquivos gerados para uma melhor organização dos dados
#gerados pelo banco de dados do IBGE e pelo arquivo gerado na conversão de altitudes hgeohnor2020. Dessa vez foi necessário fazer o import de mais uma biblioteca
#openpyxl para a organização e formatação dos arquivos gerados nas duas etapas do programa.

# Configuração da página Streamlit
st.set_page_config(page_title="Organizador e Formator de Arquivos Brutos das Estações Gravimétricas Gerados no Site do Banco de Dados do IBGE ", layout="wide")

# ==========================================
# CONTROLE DE ACESSO POR SENHA
# ==========================================
st.title("Acesso Restrito - Organizador e Formatador de Arquivos Brutos das Estações Gravimétricas Gerados no Site do Banco de Dados do IBGE (RJ)")

SENHA_DO_PROJETO = "314159SIRAB_HGEO2026981"

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
# FUNÇÕES AUXILIARES
# ==========================================
def aplicar_formatacao_excel(buffer_excel):
    """
    Aplica formatação Times New Roman, tamanho 12, centralização de texto
    e autofit de largura de colunas em uma planilha Excel contida em buffer.
    """
    buffer_excel.seek(0)
    wb = openpyxl.load_workbook(buffer_excel)
    
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                cell.font = Font(name='Times New Roman', size=12)
                cell.alignment = Alignment(horizontal='center', vertical='center')
        
        # Ajuste automático de largura de colunas
        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col:
                val_str = str(cell.value or '')
                if len(val_str) > max_len:
                    max_len = len(val_str)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)
            
    out_buffer = io.BytesIO()
    wb.save(out_buffer)
    out_buffer.seek(0)
    return out_buffer

def formatar_coordenada(valor):
    s = str(valor).strip().replace(',', '.')
    if '.' in s:
        partes = s.split('.')
        if len(partes[0].replace('-', '')) == 2:
            return s
    s_limpo = s.replace('-', '').replace('.', '')
    if len(s_limpo) >= 2:
        graus = s_limpo[:2]
        decimais = s_limpo[2:]
        if not decimais:
            decimais = "00"
        return f"-{graus}.{decimais}"
    else:
        return s

# ==========================================
# INTERFACE PRINCIPAL
# ==========================================
st.title("Processador e Organizador de Dados do IBGE")

aba1, aba2 = st.tabs(["Etapa 1: Formatação do CSV IBGE", "Etapa 2: Cruzamento e Agrupamento Regional"])

# ==========================================
# ABA 1: ETAPA 1
# ==========================================
with aba1:
    st.header("Etapa 1: Leitura do CSV e Geração dos Arquivos Tratados")
    
    csv_file = st.file_uploader("Selecione o arquivo BRUTO (.csv) baixado do IBGE", type=["csv"], key="csv_stage1")
    
    if csv_file is not None:
        try:
            try:
                df = pd.read_csv(csv_file, sep=None, engine='python', encoding='utf-8')
            except Exception:
                csv_file.seek(0)
                df = pd.read_csv(csv_file, sep=None, engine='python', encoding='latin1')
                
            st.success(f"Arquivo '{csv_file.name}' carregado com sucesso! Total de linhas: {len(df)}")
            
            col_id = 'ESTACAO'
            col_lat = 'LATITUDE'
            col_lon = 'LONGITUDE'
            
            if col_lat in df.columns and col_lon in df.columns:
                df[col_lat] = df[col_lat].apply(formatar_coordenada)
                df[col_lon] = df[col_lon].apply(formatar_coordenada)
            
            # Ordenação por data se a coluna MEDICAOPLAN existir
            col_data = 'MEDICAOPLAN' if 'MEDICAOPLAN' in df.columns else None
            if col_data:
                df['data_temp'] = pd.to_datetime(df[col_data], dayfirst=True, errors='coerce')
                df = df.sort_values(by='data_temp', ascending=False).drop(columns=['data_temp'])
            
            # Geração do TXT contendo unicamente ESTACAO, LATITUDE e LONGITUDE
            if col_id in df.columns and col_lat in df.columns and col_lon in df.columns:
                df_txt = df[[col_id, col_lat, col_lon]]
            else:
                df_txt = df.iloc[:, :3]
                
            txt_buffer = io.StringIO()
            df_txt.to_csv(txt_buffer, sep=' ', index=False, header=False)
            
            # Geração da planilha Excel formatada
            excel_buffer_raw = io.BytesIO()
            with pd.ExcelWriter(excel_buffer_raw, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            
            excel_buffer_formatted = aplicar_formatacao_excel(excel_buffer_raw)
            
            st.subheader("Downloads dos Arquivos Gerados (Etapa 1)")
            col_d1, col_d2 = st.columns(2)
            
            with col_d1:
                st.download_button(
                    label="Baixar Arquivo Entradas HGEOHNOR2020 (.txt)",
                    data=txt_buffer.getvalue(),
                    file_name="Estacoes_Entrada_HGEOHNOR2020.txt",
                    mime="text/plain"
                )
            with col_d2:
                st.download_button(
                    label="Baixar Planilha de Dados Organizados (.xlsx)",
                    data=excel_buffer_formatted,
                    file_name="00_IBGE_Estacoes_Tratadas_Organizadas.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Erro ao processar arquivo: {str(e)}")

# ==========================================
# ABA 2: ETAPA 2
# ==========================================
with aba2:
    st.header("Etapa 2: Cruzamento de Dados e Separação por Regiões do RJ")
    
    file_ibge_csv = st.file_uploader("1. Selecione o arquivo de saída gerado pelo site do HGEOHNOR2020 (.csv)", type=["csv"], key="csv_stage2")
    file_excel_ref = st.file_uploader("2. Selecione a planilha de dados organizados gerada na Etapa 1 (.xlsx)", type=["xlsx"], key="xlsx_stage2")
    
    if file_ibge_csv is not None and file_excel_ref is not None:
        try:
            try:
                df_ibge = pd.read_csv(file_ibge_csv, sep=None, engine='python', encoding='utf-8')
            except Exception:
                file_ibge_csv.seek(0)
                df_ibge = pd.read_csv(file_ibge_csv, sep=None, engine='python', encoding='latin1')
                
            df_excel = pd.read_excel(file_excel_ref)
            
            col_id_ibge_candidates = [c for c in df_ibge.columns if 'id' in c.lower() or 'ponto' in c.lower() or 'estacao' in c.lower()]
            col_id_ibge = col_id_ibge_candidates[0] if col_id_ibge_candidates else df_ibge.columns[0]
            col_id_excel = 'ESTACAO' if 'ESTACAO' in df_excel.columns else df_excel.columns[0]
            
            df_ibge[col_id_ibge] = df_ibge[col_id_ibge].astype(str).str.strip()
            df_excel[col_id_excel] = df_excel[col_id_excel].astype(str).str.strip()
            
            df_completo = pd.merge(df_ibge, df_excel, left_on=col_id_ibge, right_on=col_id_excel, how='left')
            
            colunas_para_frente = [c for c in ['MUNICIPIO', 'MEDICAOPLAN', 'UF', 'TEMA'] if c in df_completo.columns]
            outras_colunas = [c for c in df_completo.columns if c != col_id_ibge and c not in colunas_para_frente]
            ordem_final = [col_id_ibge] + colunas_para_frente + outras_colunas
            df_completo = df_completo[ordem_final]
            
            col_data = 'MEDICAOPLAN' if 'MEDICAOPLAN' in df_completo.columns else None
            if col_data:
                df_completo['data_temp'] = pd.to_datetime(df_completo[col_data], dayfirst=True, errors='coerce')
                df_completo = df_completo.sort_values(by='data_temp', ascending=False).drop(columns=['data_temp'])
                
            mapa_regioes = {
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
                'Volta Redonda': 'Região Médio Paraíba', 'Barra Mansa': 'Região Médio Paraíba',
                'Resende': 'Região Médio Paraíba', 'Itatiaia': 'Região Médio Paraíba',
                'Piraí': 'Região Médio Paraíba', 'Pinheiral': 'Região Médio Paraíba',
                'Porto Real': 'Região Médio Paraíba', 'Quatis': 'Região Médio Paraíba',
                'Rio Claro': 'Região Médio Paraíba',
                'Angra dos Reis': 'Região Costa Verde', 'Paraty': 'Região Costa Verde',
                'Mangaratiba': 'Região Costa Verde',
                'Petrópolis': 'Região Serrana', 'Teresópolis': 'Região Serrana',
                'Nova Friburgo': 'Região Serrana', 'São José do Vale do Rio Preto': 'Região Serrana',
                'Sumidouro': 'Região Serrana', 'Bom Jardim': 'Região Serrana',
                'Duas Barras': 'Região Serrana', 'Carmo': 'Região Serrana',
                'Cachoeiras de Macacu': 'Região Serrana', 'Cantagalo': 'Região Serrana',
                'Campos dos Goytacazes': 'Região Norte Fluminense', 'Macaé': 'Região Norte Fluminense',
                'Carapebus': 'Região Norte Fluminense', 'Quissamã': 'Região Norte Fluminense',
                'Conceição de Macabu': 'Região Norte Fluminense', 'São João da Barra': 'Região Norte Fluminense',
                'São Francisco de Itabapoana': 'Região Norte Fluminense',
                'Itaperuna': 'Região Noroeste Fluminense', 'Bom Jesus do Itabapoana': 'Região Noroeste Fluminense',
                'Natividade': 'Região Noroeste Fluminense', 'Varre-Sai': 'Região Noroeste Fluminense',
                'Porciúncula': 'Região Noroeste Fluminense', 'Laje do Muriaé': 'Região Noroeste Fluminense',
                'Miracema': 'Região Noroeste Fluminense', 'Santo Antônio de Pádua': 'Região Noroeste Fluminense',
                'Aperibé': 'Região Noroeste Fluminense', 'Cambuci': 'Região Noroeste Fluminense',
                'Itaocara': 'Região Noroeste Fluminense', 'Italva': 'Região Noroeste Fluminense',
                'São José de Ubá': 'Região Noroeste Fluminense', 'Cardoso Moreira': 'Região Noroeste Fluminense',
                'Cabo Frio': 'Região Baixadas Litorâneas', 'Araruama': 'Região Baixadas Litorâneas',
                'Saquarema': 'Região Baixadas Litorâneas', 'Arraial do Cabo': 'Região Baixadas Litorâneas',
                'Armação dos Búzios': 'Região Baixadas Litorâneas', 'São Pedro da Aldeia': 'Região Baixadas Litorâneas',
                'Iguaba Grande': 'Região Baixadas Litorâneas', 'Casimiro de Abreu': 'Região Baixadas Litorâneas',
                'Silva Jardim': 'Região Baixadas Litorâneas', 'Rio das Ostras': 'Região Baixadas Litorâneas',
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
            
            if 'MUNICIPIO' in df_completo.columns:
                df_completo['REGIAO_OFICIAL'] = df_completo['MUNICIPIO'].astype(str).str.strip().map(mapa_regioes).fillna('Outras Regiões')
            else:
                df_completo['REGIAO_OFICIAL'] = 'Outras Regiões'
                
            st.subheader("Relatórios Regionais Gerados")
            
            # Relatório Geral
            b_geral = io.BytesIO()
            with pd.ExcelWriter(b_geral, engine='openpyxl') as writer:
                df_completo.to_excel(writer, index=False)
            b_geral_fmt = aplicar_formatacao_excel(b_geral)
            
            st.download_button(
                label=f"Baixar 01_Relatorio_Geral_Estado_Rio_Ordenado.xlsx ({len(df_completo)} estações)",
                data=b_geral_fmt,
                file_name="01_Relatorio_Geral_Estado_Rio_Ordenado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            regioes_nomes = {
                'Região Metropolitana': '02_Regiao_Metropolitana.xlsx',
                'Região Médio Paraíba': '03_Regiao_Medio_Paraiba.xlsx',
                'Região Costa Verde': '04_Regiao_Costa_Verde.xlsx',
                'Região Serrana': '05_Regiao_Serrana.xlsx',
                'Região Norte Fluminense': '06_Regiao_Norte_Fluminense.xlsx',
                'Região Noroeste Fluminense': '07_Regiao_Noroeste_Fluminense.xlsx',
                'Região Baixadas Litorâneas': '08_Regiao_Baixadas_Litoraneas.xlsx',
                'Região Centro-Sul Fluminense': '09_Regiao_Centro_Sul_Fluminense.xlsx'
            }
            
            for regiao_nome, nome_arquivo in regioes_nomes.items():
                df_regiao = df_completo[df_completo['REGIAO_OFICIAL'] == regiao_nome]
                b_reg = io.BytesIO()
                with pd.ExcelWriter(b_reg, engine='openpyxl') as writer:
                    df_regiao.to_excel(writer, index=False)
                b_reg_fmt = aplicar_formatacao_excel(b_reg)
                
                st.download_button(
                    label=f"Baixar {nome_arquivo} ({len(df_regiao)} estações)",
                    data=b_reg_fmt,
                    file_name=nome_arquivo,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Erro no processamento da Etapa 2: {str(e)}")
