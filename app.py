import streamlit as st
import pandas as pd
import os
import re
from datetime import datetime
from io import BytesIO
import base64
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.chart import BarChart, Reference, Series
from openpyxl.chart.label import DataLabelList

DIAS_SEMANA = {
    0: 'Segunda-feira',
    1: 'Terça-feira',
    2: 'Quarta-feira',
    3: 'Quinta-feira',
    4: 'Sexta-feira',
    5: 'Sábado',
    6: 'Domingo'
}

MESES_ABREV = {
    1: 'jan', 2: 'fev', 3: 'mar', 4: 'abr', 5: 'mai', 6: 'jun',
    7: 'jul', 8: 'ago', 9: 'set', 10: 'out', 11: 'nov', 12: 'dez'
}


# Configurar página do Streamlit
st.set_page_config(page_title="WhatsApp Bot Analytics", layout="wide")

# CONFIGURAÇÕES GLOBAIS
COR_PRINCIPAL = "#00CEF7"
COR_CABECALHO = "#FFC000"
BORDA = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)
PALETA_CORES = ["003F5C", "2F4B7C", "665191", "A05195", "D45087", "F95D6A", "FF7C43", "FFA600"]

# Título e descrição
st.title("📊 WhatsApp Bot Analytics")
st.markdown("Ferramenta para análise de dados de atendimento do bot WhatsApp")

# Função para processar os arquivos CSV
def processar_arquivos_csv(arquivos_uploaded, nome_base_saida):
    dfs = []
    
    for arquivo_uploaded in arquivos_uploaded:
        # Salvar o arquivo temporário
        bytes_data = arquivo_uploaded.read()
        with open(arquivo_uploaded.name, "wb") as f:
            f.write(bytes_data)
        
        # Ler o CSV
        df_temp = pd.read_csv(arquivo_uploaded.name, encoding="latin1", sep=";")
        dfs.append(df_temp)
        
        # Remover o arquivo temporário
        os.remove(arquivo_uploaded.name)
    
    # Concatenar DataFrames
    df = pd.concat(dfs, ignore_index=True) if len(dfs) > 1 else dfs[0]
    df = df.drop_duplicates(subset="Id. Atendimento", keep="first")
    
    # Tratamentos iniciais
    df["Data"] = pd.to_datetime(df["Data"], errors="coerce", dayfirst=True)
    df["Hora"] = df["Data"].dt.strftime("%H")
    df["Dia da Semana"] = df["Data"].dt.dayofweek.map(DIAS_SEMANA)
    df["Período"] = df["Data"].dt.month.map(MESES_ABREV) + "/" + df["Data"].dt.strftime("%y")
    
    def obter_trimestre(data):
        if pd.isna(data):
            return None
        mes = data.month
        ano = data.year
        if mes <= 3:
            return f"1° Trimestre de {ano}"
        elif mes <= 6:
            return f"2° Trimestre de {ano}"
        elif mes <= 9:
            return f"3° Trimestre de {ano}"
        else:
            return f"4° Trimestre de {ano}"
    
    df["Trimestre"] = df["Data"].apply(obter_trimestre)
    
    # Extrair tags
    def extrair_tag(valor, tag):
        pattern = rf"(?<=\b{tag}:)(.*?)(?=\s*\|)"
        resultado = re.search(pattern, valor)
        return resultado.group(1).strip() if resultado else None
    
    # Coletar todas as tags únicas
    tags = set()
    for propriedade in df['Propriedades']:
        tags.update(re.findall(r'tag_[\w_]+', str(propriedade)))
    
    # Criar novas colunas para cada tag
    for tag in tags:
        df[tag] = df['Propriedades'].apply(lambda x: extrair_tag(str(x), tag))
    
    # Criar aba "U.U e Rec."
    resultados = []
    meses = df['Período'].dropna().unique()
    
    for mes in meses:
        df_mes = df[df['Período'] == mes]
        total_atendimentos = df_mes['Id. Atendimento'].nunique()
        df_unicos = df_mes.drop_duplicates(subset=['Identificação'])
        total_unicos = df_unicos['Id. Atendimento'].nunique()
        taxa_recontato = (total_atendimentos - total_unicos) / total_unicos * 100 if total_unicos else 0
        media_atendimentos_por_usuario = total_atendimentos / total_unicos if total_unicos else 0
        resultados.append({
            'Período': mes,
            'Total de atendimentos': total_atendimentos,
            'Total únicos por período': total_unicos,
            'Taxa de recontato (%)': taxa_recontato,
            'Média de atendimentos por usuário único': media_atendimentos_por_usuario
        })
    
    df_resultados = pd.DataFrame(resultados)
    df_resultados['Taxa de recontato (%)'] = df_resultados['Taxa de recontato (%)'].apply(lambda x: f"{x:.2f}%")
    df_resultados['Média de atendimentos por usuário único'] = df_resultados['Média de atendimentos por usuário único'].apply(lambda x: f"{x:.2f}")
    
    return df, df_resultados, tags


# FUNÇÃO PARA CRIAR GRÁFICOS NATIVOS DO EXCEL
def criar_grafico_excel(ws, tag_base, colunas_valores, inicio_linha, inicio_coluna, tabela_dados):
    muitas_categorias = len(tabela_dados) > 6
    chart = BarChart()
    chart.type = "bar" if muitas_categorias else "col"
    chart.style = 10
    
    # Remover títulos dos eixos
    chart.x_axis.title = None
    chart.y_axis.title = None
    chart.x_axis.tickLblPos = 'low'
    
    # Forçar a exibição dos eixos, mas sem linhas de grade
    from openpyxl.chart.axis import ChartLines
    chart.x_axis.majorGridlines = None
    chart.y_axis.majorGridlines = None
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    
    chart.title = None
    
    # Aumentar largura para evitar problemas de legenda
    chart.width = 25
    chart.height = 15
    
    # Mover legenda para fora da área do gráfico
    chart.legend.position = 'r'
    chart.legend.overlay = False
    
    num_categorias = len(tabela_dados)
    num_colunas = len(colunas_valores)
    
    # Dados para o gráfico - mantendo a lógica original
    data = Reference(ws, min_col=inicio_coluna+1, max_col=inicio_coluna+num_colunas, 
                     min_row=inicio_linha, max_row=inicio_linha+num_categorias)
    
    # Categorias (eixo x)
    categorias = Reference(ws, min_col=inicio_coluna, max_col=inicio_coluna,
                          min_row=inicio_linha+1, max_row=inicio_linha+num_categorias)
    
    # Adiciona dados ao gráfico
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categorias)
    
    # Configura rótulos de dados (mostrar somente valores)
    for serie in chart.series:
        dLbls = DataLabelList()
        dLbls.showVal = True
        dLbls.showCatName = False
        dLbls.showSerName = False
        dLbls.showLegendKey = False
        serie.dLbls = dLbls
    
    # Cores das barras
    for i, serie in enumerate(chart.series):
        serie.graphicalProperties.solidFill = PALETA_CORES[i % len(PALETA_CORES)]
    
    # Posiciona o gráfico ao lado da tabela
    celula_grafico = get_column_letter(inicio_coluna + num_colunas + 2) + str(inicio_linha)
    ws.add_chart(chart, celula_grafico)
    
    return 20  # Valor fixo para espaçamento mais consistente

# FUNÇÃO PARA SALVAR NO EXCEL E CRIAR GRÁFICOS
def salvar_no_excel(df, df_resultados, tabelas, tags_base):
    # Criar arquivo Excel em memória
    output = BytesIO()
    
    # Salvar primeira versão do Excel com os dados brutos
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="atendimentos", index=False)
        df_resultados.to_excel(writer, sheet_name="U.U e Rec", index=False)
    
    # Carregar o workbook para adicionar formatação e gráficos
    output.seek(0)
    wb = load_workbook(output)
    
    # Formata abas principais
    for aba in ["atendimentos", "U.U e Rec"]:
        if aba in wb.sheetnames:
            ws = wb[aba]
            ws.auto_filter.ref = ws.dimensions
            ws.freeze_panes = "A2"
            for cell in ws[1]:
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color=COR_CABECALHO[1:], fill_type="solid")
    
    # Configura aba resumo
    if "resumo" in wb.sheetnames:
        ws_resumo = wb["resumo"]
        ws_resumo.delete_rows(1, ws_resumo.max_row)
    else:
        ws_resumo = wb.create_sheet("resumo")
    
    linha_atual = 1
    
    for tabela, tag_base in zip(tabelas, tags_base):
        # Título da seção
        ws_resumo.cell(linha_atual, 1).value = f"Análise de {tag_base}"
        ws_resumo.cell(linha_atual, 1).font = Font(bold=True, size=14)
        ws_resumo.cell(linha_atual, 1).fill = PatternFill(start_color="DDDDDD", fill_type="solid")
        colunas_totais = len(tabela.columns)
        for col in range(2, colunas_totais + 1):
            ws_resumo.cell(linha_atual, col).fill = PatternFill(start_color="DDDDDD", fill_type="solid")
        
        linha_atual += 1
        
        # Verifica se a coluna tag_base existe no DataFrame
        if tag_base in tabela.columns:
            df_sem_total = tabela.copy()
            df_sem_total = df_sem_total[~df_sem_total[tag_base].astype(str).str.contains("TOTAL", case=False, na=False)]

            colunas_a_manter = [tag_base]
            for coluna in df_sem_total.columns:
                if coluna != tag_base and coluna != "TOTAL":
                    colunas_a_manter.append(coluna)
            
            # Filtra o DataFrame para manter apenas as colunas necessárias
            df_sem_total = df_sem_total[colunas_a_manter]
            
            # Memoriza início da tabela para referência do gráfico
            inicio_linha_tabela = linha_atual
            
            # Escreve tabela completa (com TOTAL) na planilha
            for r_idx, row in enumerate(dataframe_to_rows(tabela, index=False, header=True)):
                for c_idx, value in enumerate(row):
                    ws_resumo.cell(row=linha_atual, column=c_idx+1, value=value)
                    cell = ws_resumo.cell(row=linha_atual, column=c_idx+1)
                    cell.border = BORDA
                    
                    # Destaca cabeçalho
                    if r_idx == 0:
                        cell.font = Font(bold=True)
                        cell.fill = PatternFill(start_color=COR_CABECALHO[1:], fill_type="solid")
                        cell.alignment = Alignment(horizontal='center')
                    
                    # Destaca linha de total
                    if r_idx > 0 and "TOTAL" in str(row[0]):
                        cell.font = Font(bold=True)
                        cell.fill = PatternFill(start_color="EEEEEE", fill_type="solid")
                    
                    # Centraliza valores numéricos
                    if c_idx > 0 and r_idx > 0:
                        cell.alignment = Alignment(horizontal='center')
                
                linha_atual += 1
            
            # Adiciona gráfico nativo do Excel
            altura_grafico = criar_grafico_excel(
                ws_resumo, 
                tag_base, 
                df_sem_total.columns[1:].tolist(),
                inicio_linha_tabela, 
                1, 
                df_sem_total
            )
            
            linha_atual += altura_grafico
        else:
            ws_resumo.cell(linha_atual, 1).value = "Erro: Coluna não encontrada"
            linha_atual += 1
        
        linha_atual += 2

    for column in ws_resumo.columns:
        max_length = 0
        column_letter = get_column_letter(column[0].column)
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2)
        ws_resumo.column_dimensions[column_letter].width = min(adjusted_width, 30)
    
    # Salvar o arquivo Excel final
    output_final = BytesIO()
    wb.save(output_final)
    output_final.seek(0)
    
    return output_final

# Função para gerar link de download
def get_binary_file_downloader_html(bin_file, file_label='Arquivo'):
    data = bin_file.read()
    b64 = base64.b64encode(data).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{file_label}">📥 Baixar {file_label}</a>'
    return href

# Interface principal
tabs = st.tabs(["Upload de Arquivos", "Análise de Dados", "Visualização de Resultados"])

# Aba 1: Upload de Arquivos
with tabs[0]:
    st.header("1️⃣ Upload de Arquivos CSV")
    
    st.info("📁 Carregue os arquivos CSV do bot de WhatsApp.")
    
    uploaded_files = st.file_uploader("Selecione os arquivos CSV", 
                                     type=["csv"], 
                                     accept_multiple_files=True)
    
    nome_base_saida = st.text_input("Nome do arquivo de saída (sem extensão)", 
                                   value="bot_wpp_tratado")
    
    arquivo_saida = f"{nome_base_saida}.xlsx"
    
    if uploaded_files:
        st.success(f"✅ {len(uploaded_files)} arquivo(s) carregado(s).")
        
        if st.button("Processar Arquivos", key="processar"):
            with st.spinner("Processando arquivos..."):
                # Armazenar os DataFrames na sessão
                df, df_resultados, tags = processar_arquivos_csv(uploaded_files, nome_base_saida)
                
                # Armazenar na sessão
                st.session_state['df'] = df
                st.session_state['df_resultados'] = df_resultados
                st.session_state['tags'] = list(tags)
                st.session_state['nome_arquivo'] = arquivo_saida
                
                # Mostrar preview
                st.subheader("Preview dos dados processados")
                st.dataframe(df.head())
                
                # Mostrar tags detectadas
                st.subheader("Tags Detectadas")
                st.write(", ".join(tags))
                
                st.success("✅ Processamento concluído! Prossiga para a aba 'Análise de Dados'.")

# Aba 2: Análise de Dados
with tabs[1]:
    st.header("2️⃣ Análise de Dados")
    
    if 'df' not in st.session_state:
        st.warning("⚠️ Primeiro faça o upload e processamento dos arquivos na aba anterior.")
    else:
        st.info("📊 Configure as análises que deseja realizar.")
        
        # Inicializar lista para armazenar as análises
        if 'analises' not in st.session_state:
            st.session_state['analises'] = []
        
        # Container para adicionar nova análise
        with st.container():
            st.subheader("Nova Análise")
            
            col1, col2 = st.columns(2)
            
            with col1:
                tag_base = st.selectbox("Tag para análise", 
                                       options=st.session_state['tags'],
                                       key="nova_tag")
                
                col_valores = st.selectbox("Coluna para contar", 
                                          options=st.session_state['df'].columns,
                                          index=list(st.session_state['df'].columns).index("Id. Atendimento") if "Id. Atendimento" in st.session_state['df'].columns else 0,
                                          key="nova_col_valores")
            
            with col2:
                segmentar = st.checkbox("Segmentar os dados", key="nova_segmentar")
                
                coluna_segmento = None
                if segmentar:
                    coluna_segmento = st.selectbox("Coluna para segmentar", 
                                                 options=st.session_state['df'].columns,
                                                 key="nova_col_segmento")
            
            if st.button("Adicionar Análise", key="add_analise"):
                df_filtrado = st.session_state['df'][
                    st.session_state['df'][tag_base].notna() & 
                    (st.session_state['df'][tag_base] != "")
                ]
                
                # Cria tabela dinâmica
                if segmentar and coluna_segmento:
                    tabela_total = pd.pivot_table(
                        df_filtrado,
                        index=tag_base,
                        columns=coluna_segmento,
                        values=col_valores,
                        aggfunc="nunique",
                        margins=True,
                        margins_name="TOTAL",
                        fill_value=0
                    ).reset_index()
                else:
                    tabela_total = pd.pivot_table(
                        df_filtrado,
                        index=tag_base,
                        values=col_valores,
                        aggfunc="nunique",
                        margins=True,
                        margins_name="TOTAL"
                    ).reset_index()
                
                # Adicionar à lista de análises
                analise = {
                    'tag_base': tag_base,
                    'tabela': tabela_total,
                    'col_valores': col_valores,
                    'segmentado': segmentar,
                    'col_segmento': coluna_segmento if segmentar else None
                }
                
                st.session_state['analises'].append(analise)
                st.success(f"✅ Análise para {tag_base} adicionada!")
        
        # Listagem das análises já configuradas
        if st.session_state['analises']:
            st.subheader("Análises Configuradas")
            
            for i, analise in enumerate(st.session_state['analises']):
                expander = st.expander(f"Análise {i+1}: {analise['tag_base']}")
                with expander:
                    st.write(f"**Tag base:** {analise['tag_base']}")
                    st.write(f"**Coluna de valores:** {analise['col_valores']}")
                    if analise['segmentado']:
                        st.write(f"**Segmentado por:** {analise['col_segmento']}")
                    
                    st.write("**Preview da tabela:**")
                    st.dataframe(analise['tabela'].head())
                    
                    if st.button("Remover", key=f"remover_{i}"):
                        st.session_state['analises'].pop(i)
                        st.rerun()
            
            if st.button("Gerar Excel com Análises", key="gerar_excel"):
                if st.session_state['analises']:
                    with st.spinner("Gerando arquivo Excel..."):
                        tabelas = [analise['tabela'] for analise in st.session_state['analises']]
                        tags_base = [analise['tag_base'] for analise in st.session_state['analises']]
                        
                        excel_bytes = salvar_no_excel(
                            st.session_state['df'],
                            st.session_state['df_resultados'],
                            tabelas,
                            tags_base
                        )
                        
                        st.session_state['excel_bytes'] = excel_bytes
                        st.session_state['download_ready'] = True
                        
                        st.success("✅ Excel gerado com sucesso! Vá para a aba 'Visualização de Resultados' para baixar.")
                else:
                    st.warning("⚠️ Adicione pelo menos uma análise antes de gerar o Excel.")

# Aba 3: Visualização de Resultados
with tabs[2]:
    st.header("3️⃣ Visualização e Download")
    
    if 'download_ready' not in st.session_state or not st.session_state['download_ready']:
        st.warning("⚠️ Configure e gere suas análises nas abas anteriores primeiro.")
    else:
        st.success("✅ Seu arquivo Excel está pronto para download!")
        
        st.markdown(
            get_binary_file_downloader_html(
                st.session_state['excel_bytes'], 
                st.session_state['nome_arquivo']
            ),
            unsafe_allow_html=True
        )
        
        st.info("📌 O arquivo Excel contém:")
        st.markdown("""
        - Aba **atendimentos**: Todos os dados processados, incluindo as tags extraídas
        - Aba **U.U e Rec**: Análise de usuários únicos e taxa de recontato
        - Aba **resumo**: Tabelas e gráficos das análises configuradas
        """)
        
        # Visualizar análises
        st.subheader("Resumo das Análises")
        
        for i, analise in enumerate(st.session_state['analises']):
            expander = st.expander(f"Análise {i+1}: {analise['tag_base']}")
            with expander:
                st.write(f"**Tag base:** {analise['tag_base']}")
                st.write(f"**Coluna de valores:** {analise['col_valores']}")
                if analise['segmentado']:
                    st.write(f"**Segmentado por:** {analise['col_segmento']}")
                
                st.write("**Tabela completa:**")
                st.dataframe(analise['tabela'])
                
                # Gerar visualização simplificada
                df_chart = analise['tabela'].copy()
                if "TOTAL" in df_chart[analise['tag_base']].values:
                    df_chart = df_chart[df_chart[analise['tag_base']] != "TOTAL"]
                
                # Simplificar para visualização
                if len(df_chart) > 1:
                    st.write("**Visualização:**")
                    
                    if analise['segmentado']:
                        st.bar_chart(df_chart.set_index(analise['tag_base']))
                    else:
                        # Remover a coluna TOTAL se existir
                        cols_to_plot = [col for col in df_chart.columns if col != analise['tag_base']]
                        st.bar_chart(df_chart.set_index(analise['tag_base'])[cols_to_plot])

# Rodapé
st.markdown("---")

st.markdown("📊 **WhatsApp Bot Analytics** | R.R")
