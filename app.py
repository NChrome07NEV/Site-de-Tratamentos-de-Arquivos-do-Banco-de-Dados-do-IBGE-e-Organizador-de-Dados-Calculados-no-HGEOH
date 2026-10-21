import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.styles import Font, Alignment
#Nono Teste: após algumas edições, adições e correções do programa, a ideia permanece a mesma de gerar os arquivos gerados para uma melhor organização dos dados
#gerados pelo banco de dados do IBGE e pelo arquivo gerado na conversão de altitudes hgeohnor2020. Dessa vez foi necessário fazer o import de mais uma biblioteca
#openpyxl para a organização e formatação dos arquivos gerados nas duas etapas do programa e alinhamento / centralização dos aspectos visuais do site. Além disso, 
#agora o programa vai reorganizar novamente os arquivos por data, porém para cada tipo de medicao presente dos arquivos. Antes era gerado apenas o tipo de 
#medicaoplan para cada arquivo, mas talvez seja interessante olhar para cada dado de forma individual para outros processos futuros para os outros cálculos por PVC
#ou definição do numero geopotencial.
#Problema gerado anteriormente: alto processamento do CPU/RAM e fazendo com que o site crashe. Para resolver, foi adicionado uma aba que o usuário possa selecionar
#o que quiser gerar e reorganizado algumas linhas de código para que o programe baixe um por vez ao invés de gerar todos de uma vez só. Além disso, também foi corrigido
#o problema de duplicidade nas colunas da segunda etapa. Além disso, também foi feito mais uma otimização através do st.spinner para que seja acompanhado o progresso de
#cruzamento de dados.

# Configuração da página Streamlit
st.set_page_config(page_title="Organizador e Formatador de Arquivos Brutos das Estações Gravimétricas do Rio de Janeiro - BD - IBGE", layout="wide")

# ==========================================
# ESTILO CSS PARA CENTRALIZAR TÍTULOS E TEXTOS
# ==========================================
st.markdown("""
    <style>
    h1, h2, h3, p, .stMarkdown {
        text-align: center;
    }
    .stTextInput, .stFileUploader, .stTabs, .stButton, .stSelectbox, .stRadio {
        text-align: left;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# CONTROLE DE ACESSO POR SENHA
# ==========================================
st.title("Acesso Restrito - Organizador e Formatador de Arquivos Brutos das Estações Gravimétricas Geradas no Site do Banco de Dados do IBGE (RJ)")

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
def criar_excel_multidatas(df_base):
    """
    Cria um buffer Excel contendo 8 abas com ordenações específicas e formatação visual.
    """
    buffer_out = io.BytesIO()
    with pd.ExcelWriter(buffer_out, engine='openpyxl') as writer:
        df_base.to_excel(writer, sheet_name='Geral', index=False)
        
        topicos = {
            'Medicaoplan': ['MEDICAOPLAN', 'MEDICAO_PLAN', 'DATA_MEDICAO_PLAN'],
            'Datcalcplan': ['DATCALCPLAN', 'DAT_CALC_PLAN', 'DATA_CALC_PLAN'],
            'Medicaoalt': ['MEDICAOALT', 'MEDICAO_ALT', 'DATA_MEDICAO_ALT'],
            'Calculoalt': ['CALCULOALT', 'CALCULO_ALT', 'DATA_CALCULO_ALT'],
            'Medicaogra': ['MEDICAOGRA', 'MEDICAO_GRA', 'DATA_MEDICAO_GRA'],
            'Calculogra': ['CALCULOGRA', 'CALCULO_GRA', 'DATA_CALCULO_GRA'],
            'Visita': ['VISITA', 'DATA_VISITA', 'ULTIMA_VISITA']
        }
        
        for nome_aba, col_candidates in topicos.items():
            col_encontrada = next((c for c in col_candidates if c in df_base.columns), None)
            df_copia = df_base.copy()
            if col_encontrada:
                df_copia['temp_dt'] = pd.to_datetime(df_copia[col_encontrada], dayfirst=True, errors='coerce')
                df_copia = df_copia.sort_values(by='temp_dt', ascending=False).drop(columns=['temp_dt'])
            
            df_copia.to_excel(writer, sheet_name=nome_aba, index=False)
            
    buffer_out.seek(0)
    wb = openpyxl.load_workbook(buffer_out)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                cell.font = Font(name='Times New Roman', size=12)
                cell.alignment = Alignment(horizontal='center', vertical='center')
        
        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col:
                val_str = str(cell.value or '')
                if len(val_str) > max_len:
                    max_len = len(val_str)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)
            
    final_buffer = io.BytesIO()
    wb.save(final_buffer)
    final_buffer.seek(0)
    return final_buffer

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
st.title("Processador e Organizador de Dados dos Arquivos Gerados pelo Siteistungs IBGE - Banco de Dados")

aba1, aba2 = st.tabs(["Etapa 1: Formatação do CSV IBGE", "Etapa 2: Cruzamento e Agrupamento Regional"])

# ==========================================
# ABA 1: ETAPA 1
# ==========================================
with aba1:
    st.header("Etapa 1: Leitura do CSV e Geração dos Arquivos Tratados")
    
    csv_file = st.file_uploader("Selecione o arquivo BRUTO baixado do IBGE", key="csv_stage1")
    
    if csv_file is not None:
        try:
            with st.spinner("Processando arquivo da Etapa 1..."):
                try:
                    df = pd.read_csv(csv_file, sep=None, engine='python', encoding='utf-8')
                except Exception:
                    csv_file.seek(0)
                    df = pd.read_csv(csv_file, sep=None, engine='python', encoding='latin1')
                    
                col_id = 'ESTACAO'
                col_lat = 'LATITUDE'
                col_lon = 'LONGITUDE'
                
                if col_lat in df.columns and col_lon in df.columns:
                    df[col_lat] = df[col_lat].apply(formatar_coordenada)
                    df[col_lon] = df[col_lon].apply(formatar_coordenada)
                
                if col_id in df.columns and col_lat in df.columns and col_lon in df.columns:
                    df_txt = df[[col_id, col_lat, col_lon]]
                else:
                    df_txt = df.iloc[:, :3]
                    
                txt_buffer = io.StringIO()
                df_txt.to_csv(txt_buffer, sep=' ', index=False, header=False)
                
                excel_buffer_multisheet = criar_excel_multidatas(df)
                
            st.success(f"Arquivo '{csv_file.name}' carregado e formatado com sucesso! Total de linhas: {len(df)}")
            
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
                    label="Baixar Planilha com 8 Abas de Dados Organizados (.xlsx)",
                    data=excel_buffer_multisheet,
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
    
    file_ibge_csv = st.file_uploader("1. Selecione o arquivo de saída gerado pelo site do HGEOHNOR2020", key="csv_stage2")
    file_excel_ref = st.file_uploader("2. Selecione a planilha de dados organizados gerada na Etapa 1 (.xlsx)", key="xlsx_stage2")
    
    if file_ibge_csv is not None and file_excel_ref is not None:
        try:
            with st.spinner("Realizando o cruzamento inteligente e mapeamento regional..."):
                try:
                    df_ibge = pd.read_csv(file_ibge_csv, sep=None, engine='python', encoding='utf-8')
                except Exception:
                    file_ibge_csv.seek(0)
                    df_ibge = pd.read_csv(file_ibge_csv, sep=None, engine='python', encoding='latin1')
                    
                df_excel = pd.read_excel(file_excel_ref, sheet_name=0)
                
                col_id_ibge_candidates = [c for c in df_ibge.columns if 'id' in c.lower() or 'ponto' in c.lower() or 'estacao' in c.lower()]
                col_id_ibge = col_id_ibge_candidates[0] if col_id_ibge_candidates else df_ibge.columns[0]
                col_id_excel = 'ESTACAO' if 'ESTACAO' in df_excel.columns else df_excel.columns[0]
                
                df_ibge[col_id_ibge] = df_ibge[col_id_ibge].astype(str).str.strip()
                df_excel[col_id_excel] = df_excel[col_id_excel].astype(str).str.strip()
                
                if col_id_excel in df_ibge.columns and col_id_excel != col_id_ibge:
                    df_ibge = df_ibge.drop(columns=[col_id_excel])
                    
                df_completo = pd.merge(df_ibge, df_excel, left_on=col_id_ibge, right_on=col_id_excel, how='left')
                    
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
                
            st.success("Arquivos cruzados e prontos para exportação!")
            st.markdown("---")
            st.subheader("Opções de Geração de Relatórios da Etapa 2")
            
            tipo_relatorio = st.radio(
                "Selecione o tipo de relatório que deseja gerar e baixar:",
                [
                    "Gerar relatório geral",
                    "Gerar relatório de todas as regiões do Rio de Janeiro",
                    "Gerar relatório de uma região específica"
                ]
            )
            
            if tipo_relatorio == "Gerar relatório geral":
                with st.spinner("Gerando arquivo Excel Geral (8 abas)..."):
                    b_geral_multisheet = criar_excel_multidatas(df_completo)
                st.download_button(
                    label=f"Baixar 01_Relatorio_Geral_Estado_Rio_Multidatas.xlsx ({len(df_completo)} estações)",
                    data=b_geral_multisheet,
                    file_name="01_Relatorio_Geral_Estado_Rio_Multidatas.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            elif tipo_relatorio == "Gerar relatório de todas as regiões do Rio de Janeiro":
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
                    if len(df_regiao) > 0:
                        b_reg_multisheet = criar_excel_multidatas(df_regiao)
                        st.download_button(
                            label=f"Baixar {nome_arquivo} ({regiao_nome} - {len(df_regiao)} estações)",
                            data=b_reg_multisheet,
                            file_name=nome_arquivo,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        st.info(f"Nenhuma estação encontrada para a {regiao_nome}.")
                        
            elif tipo_relatorio == "Gerar relatório de uma região específica":
                lista_regioes_unicas = sorted(list(mapa_regioes.values())) + ['Outras Regiões']
                regiao_escolhida = st.selectbox("Escolha a região específica:", lista_regioes_unicas)
                
                df_regiao_esp = df_completo[df_completo['REGIAO_OFICIAL'] == regiao_escolhida]
                
                if len(df_regiao_esp) > 0:
                    with st.spinner(f"Preparando relatório para a {regiao_escolhida}..."):
                        b_esp_multisheet = criar_excel_multidatas(df_regiao_esp)
                        nome_arq_esp = f"Relatorio_{regiao_escolhida.replace(' ', '_')}.xlsx"
                    st.download_button(
                        label=f"Baixar {nome_arq_esp} ({len(df_regiao_esp)} estações)",
                        data=b_esp_multisheet,
                        file_name=nome_arq_esp,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning(f"Não há estações cadastradas para a região selecionada ({regiao_escolhida}).")
                    
        except Exception as e:
            st.error(f"Erro no processamento da Etapa 2: {str(e)}")
