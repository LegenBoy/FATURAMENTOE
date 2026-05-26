import streamlit as st
import pandas as pd
import os
import re

# ==========================================
# CONFIGURAÇÃO DA PÁGINA
# ==========================================
st.set_page_config(page_title="Portal de Faturamento Ecommerce", layout="wide")
st.title("📦 Portal Ecommerce - Faturamento Automático")

# ==========================================
# CONFIGURAÇÃO DA BASE DE DADOS LOCAL
# ==========================================
PASTA_BD = "dados_sistema" # Mudamos o nome para evitar qualquer conflito
ARQUIVO_LOTES = os.path.join(PASTA_BD, "lotes_pendentes.xlsx")
ARQUIVO_FINALIZADOS = os.path.join(PASTA_BD, "finalizados.xlsx")

# Cria a pasta automaticamente de forma segura
try:
    if not os.path.exists(PASTA_BD):
        os.makedirs(PASTA_BD)
except FileExistsError:
    pass # Se o ambiente na nuvem acusar que já existe, ele ignora e segue funcionando 

def carregar_bd(caminho):
    """Carrega o ficheiro Excel se existir, senão retorna um DataFrame vazio."""
    if os.path.exists(caminho):
        df = pd.read_excel(caminho)
        df.columns = df.columns.astype(str).str.strip().str.upper()
        return df
    return pd.DataFrame()

def salvar_bd(df, caminho):
    """Guarda o DataFrame num ficheiro Excel."""
    df.to_excel(caminho, index=False)

# Inicializa a Base de Dados carregando os ficheiros reais para a sessão
if 'bd_lotes' not in st.session_state:
    st.session_state['bd_lotes'] = carregar_bd(ARQUIVO_LOTES)

if 'bd_finalizados' not in st.session_state:
    st.session_state['bd_finalizados'] = carregar_bd(ARQUIVO_FINALIZADOS)

# ==========================================
# FUNÇÕES DE IDENTIFICAÇÃO AUTOMÁTICA
# ==========================================
def identificar_tipo_arquivo(df):
    colunas = df.columns.tolist() # Já normalizadas para UPPER e STRIP no loop de upload
    if 'DOCAS' in colunas and 'BATIDA' in colunas:
        return 'cubagem'
    elif 'LOTE' in colunas and 'PEDIDO_ECOMMERCE' in colunas:
        return 'lotes_geral'
    elif 'FILIAL' in colunas and 'N.F. DE SAIDA' in colunas and 'TIPO' in colunas:
        # Verifica a primeira linha para descobrir se é a nota 555 ou 551
        if not df.empty:
            tipo_nota = str(df['TIPO'].iloc[0]).strip()
            if tipo_nota == '555': return 'faturamento_555'
            elif tipo_nota == '551': return 'faturamento_551'
    return 'desconhecido'

# ==========================================
# INTERFACE DE UPLOAD NO MENU LATERAL
# ==========================================
st.sidebar.header("📤 Upload de Relatórios")
arquivos_upados = st.sidebar.file_uploader("Arraste os ficheiros do dia (CSV ou Excel)", accept_multiple_files=True)

dados = {
    'cubagem': pd.DataFrame(),
    'lotes_geral': pd.DataFrame(),
    'faturamento_555': pd.DataFrame(),
    'faturamento_551': pd.DataFrame()
}

# Lendo e categorizando os ficheiros carregados
if arquivos_upados:
    for arquivo in arquivos_upados:
        try:
            if arquivo.name.endswith('.csv'):
                df = pd.read_csv(arquivo, encoding='utf-8') 
            else:
                df = pd.read_excel(arquivo)
            
            # Limpeza preventiva de colunas (Remove espaços e padroniza para maiúsculo)
            df.columns = df.columns.astype(str).str.strip().str.upper()
            tipo = identificar_tipo_arquivo(df)
            if tipo != 'desconhecido':
                dados[tipo] = df
                st.sidebar.success(f"✅ {tipo.replace('_', ' ').upper()} carregado!")
            else:
                st.sidebar.warning(f"⚠️ Não consegui identificar: {arquivo.name}")
        except Exception as e:
            st.sidebar.error(f"Erro ao ler {arquivo.name}: {e}")

# ==========================================
# LÓGICA PRINCIPAL (O CÉREBRO)
# ==========================================
if not dados['cubagem'].empty and not dados['lotes_geral'].empty:
    st.divider()
    
    # 1. Unir os Lotes carregados hoje com os que estavam pendentes na Base de Dados
    df_lotes_hoje = dados['lotes_geral']
    df_lotes_historico = st.session_state['bd_lotes']
    
    df_lotes_combinado = pd.concat([df_lotes_historico, df_lotes_hoje]).drop_duplicates(subset=['LOTE', 'PEDIDO_ECOMMERCE'])
    st.session_state['bd_lotes'] = df_lotes_combinado # Atualiza a memória temporal
    
    # 2. Filtrar Lotes que já foram finalizados anteriormente para não duplicar
    if not st.session_state['bd_finalizados'].empty:
        pedidos_finalizados = set(st.session_state['bd_finalizados']['PEDIDO'].astype(str).tolist())
        df_lotes_combinado = df_lotes_combinado[~df_lotes_combinado['PEDIDO_ECOMMERCE'].astype(str).isin(pedidos_finalizados)]

    # 3. Extrair Filiais e Metadados da Cubagem (Data, Rota, Ordem)
    def extrair_ax_cidade(texto):
        try:
            if '-' in texto:
                ax = texto.split('-')[0].strip().lstrip('0')
                cidade = texto.split('-')[1].split('/')[0].strip()
                return f"{ax} - {cidade}"
            return texto
        except: return texto

    filiais_info = {}
    df_cubagem = dados['cubagem']
    data_cubagem = str(df_cubagem['DATA'].iloc[0]) if 'DATA' in df_cubagem.columns else "N/D"
    
    # Identificar coluna de rota
    col_rota = 'ROTAS' if 'ROTAS' in df_cubagem.columns else ('ROTA' if 'ROTA' in df_cubagem.columns else None)
    
    for idx, row in df_cubagem.iterrows():
        rota_nome = str(row.get(col_rota, 'N/D'))
        for col in df_cubagem.columns:
            if 'filial' in col.lower() and 'cubagem' in col.lower():
                celula = str(row[col])
                if '-' in celula:
                    # Mantemos o código limpo para o cruzamento, mas guardamos o display formatado
                    cod_cruzamento = celula.split('-')[0].strip().lstrip('0')
                    display_filial = extrair_ax_cidade(celula)
                    ordem_filial = col.split('/')[0].strip().replace('filial', 'Filial ')
                    filiais_info[cod_cruzamento] = {
                        "Codigo": cod_cruzamento,
                        "Display": display_filial,
                        "Data": data_cubagem,
                        "Rota/Ordem": f"{rota_nome} ({ordem_filial})"
                    }
    
    # 4. Cruzamento de Dados (Lotes vs Cubagem vs Notas Fiscais)
    faturamento_view = []
    lotes_sobrando_amarelo = []
    filiais_com_pedido = set()
    
    for idx, row in df_lotes_combinado.iterrows():
        filial_lote = str(row.get('FILIAL', '')).strip().lstrip('0')
        if filial_lote in filiais_info:
            filiais_com_pedido.add(filial_lote)
            pedido = str(row.get('PEDIDO_ECOMMERCE', '')).strip()
            lote_num = str(row.get('LOTE', '')).strip()
            cod_produto = str(row.get('PRODUTO', 'N/D')).strip()
            
            status_555 = "NÃO FATURADO"
            status_551 = "BLOQUEADO"
            st_nf_555 = None
            st_nf_551 = None
            
            # Procura a Nota 555 (pelo Lote)
            if not dados['faturamento_555'].empty:
                match_555 = dados['faturamento_555'][dados['faturamento_555']['LOTE'].astype(str).str.strip() == lote_num]
                if not match_555.empty:
                    nf_555 = str(match_555['N.F. DE SAIDA'].iloc[0])
                    status_555 = nf_555
                    status_551 = "PRONTO P/ FATURAR"
                    if 'STATUS' in match_555.columns:
                        st_nf_555 = match_555['STATUS'].iloc[0]
            
            # Procura a Nota 551 (pelo Pedido Ecommerce)
            if not dados['faturamento_551'].empty and status_555 != "NÃO FATURADO":
                # Busca a linha onde o pedido está contido em qualquer coluna de texto
                mask = dados['faturamento_551'].astype(str).apply(lambda col: col.str.contains(pedido, na=False, flags=re.IGNORECASE)).any(axis=1)
                match_551 = dados['faturamento_551'][mask]
                if not match_551.empty:
                    nf_551 = str(match_551['N.F. DE SAIDA'].iloc[0])
                    status_551 = nf_551
                    if 'STATUS' in match_551.columns:
                        st_nf_551 = match_551['STATUS'].iloc[0]

            faturamento_view.append({
                "Data": filiais_info[filial_lote]["Data"],
                "Rota/Ordem": filiais_info[filial_lote]["Rota/Ordem"],
                "Filial (AX - Cidade)": filiais_info[filial_lote]["Display"],
                "Produto": cod_produto,
                "Lote": lote_num,
                "Pedido": pedido,
                "NF 555": status_555,
                "ST 555": st_nf_555,
                "Entrada": False,
                "NF 551": status_551,
                "ST 551": st_nf_551,
                "Impresso": False,
                "Ticket": False,
                "Cliente": row.get('CLIENTE', '')
            })
        else:
            lotes_sobrando_amarelo.append(row)

    df_fat_final = pd.DataFrame(faturamento_view)
    
    # 5. Identificar filiais na cubagem que não têm pedido
    # Esta lógica agora será usada para colorir a aba de Cubagem
    codigos_pendentes = set(df_lotes_combinado['FILIAL'].astype(str).str.strip().str.lstrip('0').unique())
    codigos_finalizados = set()
    if not st.session_state['bd_finalizados'].empty:
        codigos_finalizados = set(st.session_state['bd_finalizados']['FILIAL'].astype(str).str.strip().str.lstrip('0').unique())

    def colorir_cubagem(df_c):
        st_df = pd.DataFrame('', index=df_c.index, columns=df_c.columns)
        for col in df_c.columns:
            if 'FILIAL' in col and 'CUBAGEM' in col:
                for row_idx in df_c.index:
                    val = str(df_c.at[row_idx, col])
                    if '-' in val:
                        code = val.split('-')[0].strip().lstrip('0')
                        if code in codigos_pendentes:
                            continue # Branco (pendente)
                        elif code in codigos_finalizados:
                            st_df.at[row_idx, col] = 'background-color: #c6efce; color: #006100;' # Verde
                        else:
                            st_df.at[row_idx, col] = 'background-color: #ffeb9c; color: #9c5700;' # Amarelo
        return st_df

    # ==========================================
    # ABAS DE VISUALIZAÇÃO
    # ==========================================
    tab_pendentes, tab_cubagem, tab_lotes_geral, tab_finalizados = st.tabs([
        "📋 Faturamentos Pendentes", 
        "🚛 Planilha de Cubagem", 
        "📦 Lotes Geral (Estoque)", 
        "✅ Histórico de Finalizados"
    ])

    # --- TAB PENDENTES ---
    def colorir_texto_status(row):
        styles = [''] * len(row)
        # Cores para Status 555/551
        for col in ['NF 555', 'NF 551']:
            val = row[col]
            idx = row.index.get_loc(col)
            if str(val).strip().isdigit():
                styles[idx] = 'color: #006100; font-weight: bold;' # Verde
                # Lógica do Status 6 (Fica vermelho se for 6)
                st_col = 'ST 555' if col == 'NF 555' else 'ST 551'
                if str(row[st_col]) in ['6', '6.0']:
                    styles[idx] = 'color: #9c0006; font-weight: bold;'
            elif val == "PRONTO P/ FATURAR":
                styles[idx] = 'color: #9c5700; font-weight: bold;' # Amarelo
            elif val in ["NÃO FATURADO", "BLOQUEADO"]:
                styles[idx] = 'color: #9c0006; font-weight: bold;' # Vermelho
        return styles

    with tab_pendentes:
        st.subheader("Painel de Faturamento do Dia")
        if not df_fat_final.empty:
            # Funcionalidade de Cópia em Massa das NFs 551
            nfs_para_imprimir = df_fat_final[df_fat_final['NF 551'].str.isdigit()]['NF 551'].tolist()
            if nfs_para_imprimir:
                with st.expander("🖨️ Ações de Impressão em Massa"):
                    lista_nfs_str = ", ".join(nfs_para_imprimir)
                    st.text_area("Notas 551 prontas para copiar:", value=lista_nfs_str, height=70)
                    if st.button("Marcar Todas como Impressas"):
                        st.info("Para marcar como impresso, utilize a caixa de seleção na coluna 'Impresso' abaixo.")

            # Configuração de colunas para centralizar
            config_col = {
                "Data": st.column_config.Column(width="small"),
                "NF 555": st.column_config.Column("NF 555", help="Número da Nota Fiscal 555"),
                "NF 551": st.column_config.Column("NF 551", help="Número da Nota Fiscal 551"),
                "ST 555": None, "ST 551": None, # Oculta colunas de status técnico
                "Entrada": st.column_config.CheckboxColumn("Entrada", help="Entrada realizada no sistema?"),
                "Impresso": st.column_config.CheckboxColumn("Impresso", help="Página impressa?"),
                "Ticket": st.column_config.CheckboxColumn("Ticket", help="Abrir ticket para este cliente"),
            }
            
            # Usamos data_editor para permitir os checkboxes
            df_editavel = st.data_editor(
                df_fat_final.style.apply(colorir_texto_status, axis=1), 
                use_container_width=True,
                column_config=config_col,
                hide_index=True,
                key="editor_faturamento"
            )
            df_fat_final = df_editavel # Atualiza com os valores marcados pelo usuário
        else:
            st.info("Nenhum pedido da tabela de Lotes corresponde à Cubagem de hoje.")

    # --- TAB CUBAGEM ---
    with tab_cubagem:
        st.subheader("Status dos Pedidos na Cubagem")
        if not dados['cubagem'].empty:
            st.dataframe(
                dados['cubagem'].style.apply(colorir_cubagem, axis=None),
                use_container_width=True,
                hide_index=True
            )
        else:
            st.info("Carregue a planilha de cubagem para visualizar.")

    # --- TAB LOTES GERAL ---
    with tab_lotes_geral:
        st.subheader("Lotes em Estoque (Fora do Carregamento de Hoje)")
        if lotes_sobrando_amarelo:
            df_sobras = pd.DataFrame(lotes_sobrando_amarelo)
            st.dataframe(df_sobras.style.set_properties(**{'background-color': '#ffeb9c', 'color': 'black'}), use_container_width=True, hide_index=True)
        else:
            st.info("Não há lotes pendentes fora da cubagem atual.")

    # --- TAB FINALIZADOS ---
    with tab_finalizados:
        st.subheader("Histórico de Pedidos Concluídos")
        if not st.session_state['bd_finalizados'].empty:
            st.dataframe(st.session_state['bd_finalizados'], use_container_width=True, hide_index=True)
        else:
            st.write("Nenhum pedido foi finalizado ainda.")

    # ==========================================
    # FUNÇÃO DE FINALIZAR E GUARDAR NO HISTÓRICO
    # ==========================================
    st.divider()
    col1, col2 = st.columns([1, 3])
    with col1:
        if st.button("🚀 Finalizar Faturamentos e Limpar", type="primary", use_container_width=True):
            if not df_fat_final.empty:
                # 1. Isolar quem está 100% faturado (NF 555 e 551 existem e estão impressas/com entrada)
                finalizados = df_fat_final[
                    (df_fat_final['NF 555'].str.isdigit()) & 
                    (df_fat_final['NF 551'].str.isdigit()) &
                    (df_fat_final['Impresso'] == True)
                ]

                if not finalizados.empty:
                    # 2. Adicionar ao Histórico de Finalizados e guardar no Excel
                    df_historico = pd.concat([st.session_state['bd_finalizados'], finalizados])
                    st.session_state['bd_finalizados'] = df_historico
                    salvar_bd(df_historico, ARQUIVO_FINALIZADOS)
                    
                    # 3. Remover estes finalizados da tabela de Lotes Pendentes e atualizar o Excel
                    lotes_para_remover = finalizados['Lote'].astype(str).tolist()
                    df_lotes_atualizado = st.session_state['bd_lotes'][~st.session_state['bd_lotes']['LOTE'].astype(str).isin(lotes_para_remover)]
                    
                    st.session_state['bd_lotes'] = df_lotes_atualizado
                    salvar_bd(df_lotes_atualizado, ARQUIVO_LOTES)
                    
                    st.success(f"🎉 {len(finalizados)} pedidos finalizados com sucesso! Movidos para o histórico e limpos da vista diária.")
                    st.balloons()
                else:
                    st.warning("⚠️ Nenhum pedido tem os dois faturamentos (555 e 551) concluídos. Nada a finalizar neste momento.")
            else:
                st.warning("A tabela está vazia.")

else:
    st.info("👈 Por favor, faça o upload dos relatórios de Lotes Geral e Cubagem no menu lateral esquerdo para começar.")
