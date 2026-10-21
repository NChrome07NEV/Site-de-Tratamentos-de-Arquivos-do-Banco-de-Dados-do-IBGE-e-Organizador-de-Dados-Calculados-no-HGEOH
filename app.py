import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.styles import Font, Alignment
#Oitavo teste:
##Realizando o experimento total, agora aqui está a tentativa de juntar os dois fluxos de trabalhos em um único programa com o intuito de gerar um site e facilitar
#o trabalho do projeto e talvez outros possíveis usos desse programa.
#Dessa vez, foi implementado uma nova biblioteca 'openpyxl' para formatar o texto e estrutura das células no arquivo excel da primeira etapa.

# Configuração da página sem emojis
st.set_page_config(
    page_title="Processador de Estações Geodésicas - IBGE",
    layout="wide"
)

st.title("Processador de Estações Geodésicas - IBGE")

# ... (Seu fluxo de upload/busca dos dados mantido) ...

def formatar_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Nome correto da aba: Sheet1
        df.to_excel(writer, sheet_name='Sheet1', index=False)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        
        # Estilo solicitado: Times New Roman, Tamanho 12, Centralizado
        fonte_customizada = Font(name='Times New Roman', size=12)
        alinhamento_central = Alignment(horizontal='center', vertical='center')
        
        for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
            for cell in row:
                cell.font = fonte_customizada
                cell.alignment = alinhamento_central
                
        # Auto-ajuste de largura das colunas
        for col in worksheet.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)
            
    return output.getvalue()

def gerar_txt_estacoes(df):
    # Garante a seleção estrita das 3 colunas: ESTACAO, LATITUDE, LONGITUDE
    df_txt = df[['ESTACAO', 'LATITUDE', 'LONGITUDE']].copy()
    
    # Remove linhas onde alguma das coordenadas ou estação seja nula
    df_txt = df_txt.dropna(subset=['ESTACAO', 'LATITUDE', 'LONGITUDE'])
    
    # Formatação limpa separada por espaço
    txt_content = df_txt.to_csv(index=False, header=False, sep=' ')
    return txt_content

# Exemplo de geração dos botões de download:
# 1. Download do Excel
# excel_data = formatar_excel(df_processado)
# st.download_button(
#     label="Baixar Dados Organizados (.xlsx)",
#     data=excel_data,
#     file_name="Dados Organizados - EG.xlsx",
#     mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
# )

# 2. Download do TXT Entrada HGEOHNOR2020
# txt_data = gerar_txt_estacoes(df_processado)
# st.download_button(
#     label="Baixar Estações - LAT - LON (.txt)",
#     data=txt_data,
#     file_name="Estacoes_Entrada_HGEOHNOR2020.txt",
#     mime="text/plain"
# )
