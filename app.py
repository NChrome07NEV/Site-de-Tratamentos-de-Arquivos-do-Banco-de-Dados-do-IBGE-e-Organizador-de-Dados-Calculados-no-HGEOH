import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.styles import Font, Alignment

# Configuração da página sem emojis
st.set_page_config(
    page_title="Organizador e Formator de Arquivos Brutos das Estações Gravimétricas Gerados no Site do Banco de Dados do IBGE ",
    layout="wide"
)

st.title("Acesso Restrito - Organizador e Formatador de Arquivos Brutos das Estações Gravimétricas Gerados do Banco de Dados do IBGE (RJ)")

# Função genérica para estilizar qualquer arquivo Excel com OpenPyXL
def aplicar_estilo_excel(df, nome_aba='Sheet1'):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name=nome_aba, index=False)
        workbook = writer.book
        worksheet = writer.sheets[nome_aba]
        
        # Configuração de estilo: Times New Roman, Tamanho 12, Centralizado
        fonte_customizada = Font(name='Times New Roman', size=12)
        alinhamento_central = Alignment(horizontal='center', vertical='center')
        
        for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
            for cell in row:
                cell.font = fonte_customizada
                cell.alignment = alinhamento_central
                
        # Auto-ajuste de largura das colunas (equivalente ao duplo clique entre colunas)
        for col in worksheet.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)
            
    return output.getvalue()

# Função para ordenar dataframe por data
def ordenar_por_data(df):
    # Procura colunas de data disponíveis (ex: VISITA, MEDICAOPLAN, etc.)
    colunas_data = [col for col in ['VISITA', 'MEDICAOPLAN', 'CALCULOGRA', 'MEDICAOALT'] if col in df.columns]
    
    if colunas_data:
        col_ref = colunas_data[0]
        # Converte temporariamente para datetime para ordenar corretamente
        df_temp = df.copy()
        df_temp['_data_sort'] = pd.to_datetime(df_temp[col_ref], format='%d/%m/%Y', errors='coerce')
        df_sorted = df_temp.sort_values(by='_data_sort').drop(columns=['_data_sort'])
        return df_sorted
    return df

# Função para gerar o arquivo TXT estrito com 3 colunas (ESTACAO, LATITUDE, LONGITUDE)
def gerar_txt_estacoes(df):
    df_txt = df[['ESTACAO', 'LATITUDE', 'LONGITUDE']].copy()
    df_txt = df_txt.dropna(subset=['ESTACAO', 'LATITUDE', 'LONGITUDE'])
    return df_txt.to_csv(index=False, header=False, sep=' ')

# --- ETAPA 1: Processamento Inicial ---
st.header("Etapa 1: Organização dos Dados do IBGE")
# Exemplo de fluxo para a Etapa 1:
# df_etapa1 = processar_dados_ibge(...)
# df_etapa1_ordenado = ordenar_por_data(df_etapa1)
# excel_etapa1 = aplicar_estilo_excel(df_etapa1_ordenado, nome_aba='Sheet1')
# txt_etapa1 = gerar_txt_estacoes(df_etapa1_ordenado)

# --- ETAPA 2: Processamento dos Dados do HGEOH ---
st.header("Etapa 2: Organização dos Dados do HGEOH")
# Exemplo de fluxo para a Etapa 2:
# df_etapa2 = processar_dados_hgeoh(...)
# df_etapa2_ordenado = ordenar_por_data(df_etapa2)
# excel_etapa2 = aplicar_estilo_excel(df_etapa2_ordenado, nome_aba='Sheet1')
