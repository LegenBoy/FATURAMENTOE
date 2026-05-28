import streamlit as st
import pandas as pd
import os
import re
import gspread # Importar a biblioteca gspread
import json

# ==========================================
# CONFIGURAÇÃO DA PÁGINA
# ==========================================
st.set_page_config(page_title="Portal de Faturamento Ecommerce", layout="wide")
st.title("📦 Portal Ecommerce - Faturamento Automático")

# ==========================================
# CONFIGURAÇÃO DA BASE DE DADOS LOCAL
# ==========================================
# PASTA_BD = "dados_sistema" # Não será mais usada para armazenamento persistente

# Nomes das planilhas no Google Sheets
PLANILHA_LOTES = "lotes_pendentes_ecommerce" # Nome da sua planilha de lotes
PLANILHA_FINALIZADOS = "finalizados_ecommerce" # Nome da sua planilha de finalizados
PLANILHA_CUBAGEM_ATIVA = "faturamento_cubagem_ativa"
PLANILHA_555_ATIVA = "faturamento_555_ativa"
PLANILHA_551_ATIVA = "faturamento_551_ativa"

# Definir os cabeçalhos padrão para cada planilha
DEFAULT_HEADERS_LOTES = [
    "ROTA", "FILIAL", "CIDADE", "LOTE", "PEDIDO_ECOMMERCE",
    "PEDIDO_SITE", "PRODUTO", "DESCRICAO", "QUANTIDADE",
    "CUBTOTAL_PRODUTO", "CLIENTE", "DATA_PAGAMENTO",
    "ENTRADA", "IMPRESSO", "TICKET"
]

DEFAULT_HEADERS_FINALIZADOS = [
    "N° LOTE", "ROTA", "AX - CIDADE", "PEDIDO CLIENTE ECOMMERCE",
    "CLIENTE", "NÚMERO NF 555", "NÚMERO NF 551", "CÓD PRODUTO",
    "DATA PLANILHA DE CUBAGEM", "TICKET"
]

# Garante a criação da pasta de dados de forma robusta
# Removemos a criação de diretórios locais para dados persistentes

# Garante a criação da pasta de dados de forma robusta
# Removemos a criação de diretórios locais para dados persistentes

# Configuração do Google Sheets
@st.cache_resource(ttl=3600) # Cache a conexão para evitar reconexões frequentes
def get_gsheet_client():
    """Conecta ao Google Sheets usando Streamlit Secrets."""
    try:
        # Carrega as credenciais do Streamlit Secrets
        # O conteúdo do seu arquivo JSON da conta de serviço deve estar em st.secrets["gcp_service_account"]
        # Certifique-se de que o JSON está formatado corretamente nos secrets
        creds_data = st.secrets["gcp_service_account"]
        
        # Se creds_data for uma string (JSON puro), converte para dicionário
        if isinstance(creds_data, str):
            creds = json.loads(creds_data)
        else:
            # Caso já seja um objeto dict-like do Streamlit
            creds = dict(creds_data)
            
        gc = gspread.service_account_from_dict(creds)
        return gc
    except Exception as e:
        st.error(f"Erro ao conectar ao Google Sheets. Verifique suas credenciais em `st.secrets['gcp_service_account']`: {e}")
        st.stop() # Para a execução do app se não conseguir conectar

gc = get_gsheet_client()

def carregar_bd(caminho):
    """Carrega dados de uma planilha do Google Sheets.
    Se a planilha não existir ou estiver vazia, cria com cabeçalhos padrão."""
    
    default_headers = []
    if caminho == PLANILHA_LOTES:
        default_headers = DEFAULT_HEADERS_LOTES
    elif caminho == PLANILHA_FINALIZADOS:
        default_headers = DEFAULT_HEADERS_FINALIZADOS

    try:
        sh = gc.open(caminho)
        worksheet = sh.get_worksheet(0)
        records = worksheet.get_all_records()
        
        if not records: # Planilha existe mas está vazia (sem cabeçalhos ou dados)
            if default_headers:
                worksheet.update([default_headers])
                st.info(f"Planilha '{caminho}' estava vazia. Cabeçalhos padrão criados.")
                return pd.DataFrame(columns=[h.upper() for h in default_headers])
            else:
                return pd.DataFrame() # Fallback if no default headers defined
        
        df = pd.DataFrame(records)
        if not df.empty:
            df.columns = df.columns.astype(str).str.strip().str.upper()
            # Garante que todas as colunas esperadas existem (retrocompatibilidade)
            for header in default_headers:
                if header.upper() not in df.columns:
                    df[header.upper()] = ""
        return df
    except gspread.exceptions.SpreadsheetNotFound:
        st.warning(f"Planilha '{caminho}' não encontrada. Criando uma nova...")
        try:
            sh = gc.create(caminho)
            # Compartilha a planilha com a conta de serviço
            sh.share(st.secrets["gcp_service_account"]["client_email"], perm_type='user', role='writer')
            worksheet = sh.get_worksheet(0)
            
            if default_headers:
                worksheet.update([default_headers])
                st.success(f"Planilha '{caminho}' criada com cabeçalhos padrão.")
                return pd.DataFrame(columns=[h.upper() for h in default_headers])
            else:
                return pd.DataFrame()
        except Exception as create_e:
            st.error(f"Erro ao criar e configurar a planilha '{caminho}': {create_e}")
            return pd.DataFrame()
    except Exception as e:
        st.error(f"Erro ao carregar dados da planilha '{caminho}': {e}")
        return pd.DataFrame()

def salvar_bd(df, caminho):
    """Salva o DataFrame em uma planilha do Google Sheets."""
    try:
        sh = gc.open(caminho)
        worksheet = sh.get_worksheet(0) # Pega a primeira aba disponível (independente do nome)
        
        # Converte o DataFrame para string para evitar erros de serialização JSON (como Timestamps)
        # Substituímos 'nan' por vazio para a planilha ficar limpa
        df_salvar = df.astype(str).replace(['nan', 'None', '<NA>'], '')

        worksheet.clear()
        worksheet.update([df_salvar.columns.values.tolist()] + df_salvar.values.tolist())
        st.success(f"Dados salvos na planilha '{caminho}' com sucesso!")
    except gspread.exceptions.SpreadsheetNotFound:
        st.error(f"Erro: Planilha '{caminho}' não encontrada para salvar. Verifique o nome ou crie-a manualmente.")
    except Exception as e:
        st.error(f"Erro ao salvar dados na planilha '{caminho}': {e}")

# Inicializa a Base de Dados carregando os ficheiros reais para a sessão
if 'bd_lotes' not in st.session_state:
    st.session_state['bd_lotes'] = carregar_bd(PLANILHA_LOTES)

if 'bd_finalizados' not in st.session_state:
    st.session_state['bd_finalizados'] = carregar_bd(PLANILHA_FINALIZADOS)

# Inicializa ou recupera os dados dos arquivos anexados para sobreviver ao F5
if 'dados_ativos' not in st.session_state:
    st.session_state['dados_ativos'] = {
        'cubagem': carregar_bd(PLANILHA_CUBAGEM_ATIVA),
        'faturamento_555': carregar_bd(PLANILHA_555_ATIVA),
        'faturamento_551': carregar_bd(PLANILHA_551_ATIVA)
    }

# ==========================================
# FUNÇÕES DE IDENTIFICAÇÃO AUTOMÁTICA
# ==========================================
def identificar_tipo_arquivo(df):
    colunas = df.columns.tolist() # Já normalizadas para UPPER e STRIP no loop de upload
    colunas = [str(c).upper() for c in colunas]
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

dados = st.session_state['dados_ativos']
# lotes_geral é processado e mesclado, não precisa de persistência separada
lotes_geral_temp = pd.DataFrame()

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
                if tipo == 'lotes_geral':
                    lotes_geral_temp = df
                else:
                    dados[tipo] = df
                    salvar_bd(df, f"faturamento_{tipo}_ativa")
                st.sidebar.success(f"✅ {tipo.replace('_', ' ').upper()} carregado!")
            else:
                st.sidebar.warning(f"⚠️ Não consegui identificar: {arquivo.name}")
        except Exception as e:
            st.sidebar.error(f"Erro ao ler {arquivo.name}: {e}")

if st.sidebar.button("🗑️ Limpar Arquivos do Dia"):
    for t in ['cubagem', 'faturamento_555', 'faturamento_551']:
        dados[t] = pd.DataFrame()
        salvar_bd(pd.DataFrame(), f"faturamento_{t}_ativa")
    st.rerun()

# ==========================================
# LÓGICA PRINCIPAL (O CÉREBRO)
# ==========================================
if not dados['cubagem'].empty:
    st.divider()
    
    # 1. Unir os Lotes carregados hoje com os que estavam pendentes na Base de Dados
    df_lotes_hoje = lotes_geral_temp
    df_lotes_historico = st.session_state['bd_lotes']
    
    if not df_lotes_hoje.empty:
        df_lotes_combinado = pd.concat([df_lotes_historico, df_lotes_hoje]).drop_duplicates(subset=['LOTE', 'PEDIDO_ECOMMERCE'], keep='last')
        st.session_state['bd_lotes'] = df_lotes_combinado
        salvar_bd(df_lotes_combinado, PLANILHA_LOTES)
    else:
        df_lotes_combinado = df_lotes_historico
    
    # 2. Filtrar Lotes que já foram finalizados anteriormente para não duplicar
    if not st.session_state['bd_finalizados'].empty:
        pedidos_finalizados = set(st.session_state['bd_finalizados']['PEDIDO CLIENTE ECOMMERCE'].astype(str).tolist())
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

            # Lógica de automação: se a 551 está faturada, a entrada da 555 já ocorreu
            is_551_faturado = False
            v551 = str(status_551).strip().upper()
            if v551 not in ["NÃO FATURADO", "BLOQUEADO", "PRONTO P/ FATURAR", "NAN", "NONE", "N/D", ""]:
                try:
                    is_551_faturado = float(v551) > 0
                except: pass

            # Recupera estados salvos diretamente do DataFrame (que veio do Google Sheets)
            val_entrada = row.get('ENTRADA')
            if val_entrada == "" or pd.isna(val_entrada):
                val_entrada = is_551_faturado
            
            val_impresso = row.get('IMPRESSO', False)
            val_ticket = row.get('TICKET', "")
            
            # Extrair Rota, Cidade e AX para as colunas finais
            rota_ordem_full = filiais_info[filial_lote]["Rota/Ordem"] # Ex: "AZ 01 (Filial 1)"
            rota_val = rota_ordem_full # Mantém a ordem da loja (Filial 1, 2, etc)
            
            display_filial_full = filiais_info[filial_lote]["Display"] # Ex: "064 - COLINA"
            ax_val = display_filial_full.split('-')[0].strip() if '-' in display_filial_full else ""
            cidade_val = display_filial_full.split('-')[1].strip() if '-' in display_filial_full else display_filial_full


            faturamento_view.append({
                "Data Planilha de Cubagem": filiais_info[filial_lote]["Data"],
                "Rota": rota_val,
                "AX - Cidade": f"{ax_val} - {cidade_val}",
                "N° Lote": lote_num,
                "Pedido Cliente Ecommerce": pedido,
                "Cód Produto": cod_produto,
                "Cliente": row.get('CLIENTE', ''),
                "Número NF 555": status_555,
                "ST 555": st_nf_555,
                "Entrada": bool(val_entrada) if val_entrada != "" else is_551_faturado,
                "Número NF 551": status_551,
                "ST 551": st_nf_551,
                "Impresso": bool(val_impresso),
                "Ticket": str(val_ticket) if pd.notna(val_ticket) else "",
            }) 
        else:
            lotes_sobrando_amarelo.append(row)

    df_fat_final = pd.DataFrame(faturamento_view)
    
    # 5. Identificar filiais na cubagem que não têm pedido
    # Esta lógica agora será usada para colorir a aba de Cubagem
    codigos_pendentes = set(df_lotes_combinado['FILIAL'].astype(str).str.strip().str.lstrip('0').unique())
    codigos_finalizados = set()
    if not st.session_state['bd_finalizados'].empty:
        # Extrai o código AX da coluna 'AX - CIDADE' (ex: "064 - COLINA" -> "64") para bater com a cubagem
        codigos_finalizados = set(st.session_state['bd_finalizados']['AX - CIDADE'].astype(str).str.split('-').str[0].str.strip().str.lstrip('0').unique())

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
        def is_nf_val(val):
            v = str(val).strip().upper()
            if v in ["NÃO FATURADO", "BLOQUEADO", "PRONTO P/ FATURAR", "NAN", "NONE", "N/D", "", "NULL"]:
                return False
            try:
                # Se conseguir converter para número e não for zero, é uma NF
                return float(v) > 0
            except:
                return False

        styles = [''] * len(row)
        for col in ['Número NF 555', 'Número NF 551']:
            val = str(row[col]).strip().upper()
            idx = row.index.get_loc(col)

            if is_nf_val(val):
                styles[idx] = 'color: #00A300; font-weight: bold;' # Verde forte para NF faturada
                # Lógica do Status 6 (Fica vermelho se for 6)
                st_col = 'ST 555' if col == 'Número NF 555' else 'ST 551'
                if str(row.get(st_col, '')).strip() in ['6', '6.0']:
                    styles[idx] = 'color: #E60000; font-weight: bold;'
            elif val == "PRONTO P/ FATURAR":
                styles[idx] = 'color: #B8860B; font-weight: bold;' # Amarelo Dourado
            elif val in ["NÃO FATURADO", "BLOQUEADO"]:
                styles[idx] = 'color: #E60000; font-weight: bold;' # Vermelho
        return styles

    with tab_pendentes:
        st.subheader("Painel de Faturamento do Dia")
        if not df_fat_final.empty:
            # Lógica de Filtro e Cópia
            def is_nf_val_static(val):
                v = str(val).strip().upper()
                if v in ["NÃO FATURADO", "BLOQUEADO", "PRONTO P/ FATURAR", "NAN", "NONE", "N/D", ""]: return False
                try: return float(v) > 0
                except: return False

            col_btn1, col_btn2 = st.columns([1, 1])
            with col_btn1:
                if st.button("🔍 Filtrar Notas 551 p/ Imprimir", use_container_width=True):
                    st.session_state['filtro_print'] = True
            with col_btn2:
                if st.button("📋 Mostrar Todos os Pedidos", use_container_width=True):
                    st.session_state['filtro_print'] = False

            if st.session_state.get('filtro_print', False):
                df_fat_final = df_fat_final[df_fat_final.apply(lambda r: not r['Impresso'] and is_nf_val_static(r['Número NF 551']), axis=1)]

            # --- Lógica para o botão de Copiar e Marcar Impressas ---
            # Identifica NFs 551 faturadas, NÃO impressas e sem Status 6
            nfs_para_copiar_e_marcar = []
            nfs_bloqueadas_status6 = []
            
            # Itera sobre o DataFrame completo (faturamento_view) para encontrar todas as notas elegíveis
            # Isso garante que o botão apareça mesmo se a tabela principal estiver filtrada
            for _, r in pd.DataFrame(faturamento_view).iterrows():
                st_551 = str(r.get('ST 551', '')).strip()
                is_err_6 = st_551 in ['6', '6.0']
                
                if is_nf_val_static(r['Número NF 551']):
                    nf_limpa = str(r['Número NF 551']).split('.')[0].strip()
                    if is_err_6:
                        nfs_bloqueadas_status6.append(nf_limpa)
                    elif not r['Impresso']:
                        if nf_limpa.isdigit():
                            nfs_para_copiar_e_marcar.append(nf_limpa)

            # Garante que a lista de NFs seja única e ordenada para ser o mais enxuta possível
            nfs_para_copiar_e_marcar = sorted(list(set(nfs_para_copiar_e_marcar)))

            with st.expander("🖨️ Copiar Notas para Impressão", expanded=True):
                if nfs_para_copiar_e_marcar:
                    st.info(f"Foram encontradas {len(nfs_para_copiar_e_marcar)} notas prontas para impressão.")
                    
                    # Botão para copiar e marcar, habilitado se houver notas para copiar
                    if st.button("📋 Copiar e Marcar como Impressas", key="copy_and_mark_btn"):
                        # Marcar como impressas
                        for nf in nfs_para_copiar_e_marcar:
                            # Localiza o registro original na memória para marcar
                            for _, r_orig in pd.DataFrame(faturamento_view).iterrows(): # Itera sobre o DF base
                                if str(r_orig['Número NF 551']).split('.')[0].strip() == nf:
                                    # Atualiza no BD de Lotes Principal
                                    lote_id = str(r_orig['N° Lote'])
                                    ped_id = str(r_orig['Pedido Cliente Ecommerce'])
                                    idx_bd = st.session_state['bd_lotes'][(st.session_state['bd_lotes']['LOTE'].astype(str) == lote_id) & (st.session_state['bd_lotes']['PEDIDO_ECOMMERCE'].astype(str) == ped_id)].index
                                    if not idx_bd.empty:
                                        st.session_state['bd_lotes'].at[idx_bd[0], 'IMPRESSO'] = True
                        
                        salvar_bd(st.session_state['bd_lotes'], PLANILHA_LOTES)
                        st.success("Notas marcadas como impressas! Clique em Sincronizar para atualizar a visualização.")

                    st.write("Copie as NFs 551 abaixo (uma por linha para o TOTVS):")
                    st.code("\n".join(nfs_para_copiar_e_marcar), language="text")
                else:
                    st.info("Nenhuma NF 551 faturada e não impressa encontrada para copiar.")

            if nfs_bloqueadas_status6:
                st.error(f"🚨 **ALERTA DE STATUS 6:** As notas {', '.join(nfs_bloqueadas_status6)} estão com erro/canceladas e foram bloqueadas para impressão.")

            # Configuração de colunas para centralizar
            config_col = {
                "Data Planilha de Cubagem": st.column_config.Column(width="small"),
                "Rota": st.column_config.Column(width="small"),
                "AX - Cidade": st.column_config.Column(width="medium"),
                "N° Lote": st.column_config.Column("N° Lote"),
                "Pedido Cliente Ecommerce": st.column_config.Column("Pedido Cliente Ecommerce"),
                "Cód Produto": st.column_config.Column("Cód Produto"),
                "Cliente": st.column_config.Column("Cliente"),
                "Número NF 555": st.column_config.Column("Número NF 555", help="Número da Nota Fiscal 555"),
                "Número NF 551": st.column_config.Column("Número NF 551", help="Número da Nota Fiscal 551"),
                "ST 555": None, "ST 551": None, # Oculta colunas de status técnico
                "Entrada": st.column_config.CheckboxColumn("Entrada", help="Entrada realizada no sistema?"),
                "Impresso": st.column_config.CheckboxColumn("Impresso", help="Página impressa?"),
                "Ticket": st.column_config.TextColumn("N° Ticket", help="Digite o número do ticket para finalizar com erro/TI"),
            }
            
            with st.form("form_faturamento"):
                # Usamos data_editor para permitir os checkboxes
                df_editavel = st.data_editor(
                    df_fat_final.style.apply(colorir_texto_status, axis=1), 
                    use_container_width=True,
                    column_config=config_col,
                    hide_index=True,
                    key="editor_faturamento"
                )
                
                col_sync, col_finalize = st.columns([1, 1])
                with col_sync:
                    sync_btn = st.form_submit_button("🔄 Sincronizar e Salvar Marcações", use_container_width=True)
                with col_finalize:
                    finalize_btn = st.form_submit_button("🚀 Finalizar Faturamentos e Limpar", type="primary", use_container_width=True)

                if sync_btn:
                    # --- LÓGICA DE PERSISTÊNCIA E SINCRONIZAÇÃO DE DUPLICADOS ---
                    edits = st.session_state.get("editor_faturamento", {}).get("edited_rows", {})
                    if edits:
                        for row_idx_str, row_changes in edits.items():
                            row_idx = int(row_idx_str)
                            lote_ref = df_editavel.iloc[row_idx]['N° Lote']
                            ped_ref = df_editavel.iloc[row_idx]['Pedido Cliente Ecommerce']
                            
                            # Localiza o índice real no banco de dados principal
                            idx_match = st.session_state['bd_lotes'][
                                (st.session_state['bd_lotes']['LOTE'].astype(str) == str(lote_ref)) & 
                                (st.session_state['bd_lotes']['PEDIDO_ECOMMERCE'].astype(str) == str(ped_ref))
                            ].index

                            for col_name, novo_valor in row_changes.items():
                                col_maiuscula = col_name.upper()
                                if col_maiuscula in st.session_state['bd_lotes'].columns:
                                    if not idx_match.empty:
                                        st.session_state['bd_lotes'].at[idx_match[0], col_maiuscula] = novo_valor
                                
                                # Propagação para duplicados
                                if col_name in ['Entrada', 'Impresso']:
                                    nf_col = 'Número NF 555' if col_name == 'Entrada' else 'Número NF 551'
                                    nf_valor = df_editavel.iloc[row_idx][nf_col]
                                    if str(nf_valor).split('.')[0].isdigit():
                                        # Encontra todos com a mesma NF no banco principal e atualiza
                                        target_nf_col = 'LOTE' if col_name == 'Entrada' else 'PEDIDO_ECOMMERCE' # Ajuste conforme lógica de cruzamento
                                        # Simplificado: busca no banco de lotes onde a NF bateria
                                        st.session_state['bd_lotes'].loc[st.session_state['bd_lotes']['LOTE'].astype(str) == str(nf_valor), col_maiuscula] = novo_valor
                        
                        salvar_bd(st.session_state['bd_lotes'], PLANILHA_LOTES)
                        st.rerun() # Atualiza a tela após processar tudo

                if finalize_btn:
                    if not df_editavel.empty:
                        # Função auxiliar para checar se o valor é uma NF válida
                        def is_numeric_nf(val):
                            v = str(val).split('.')[0].strip()
                            return v.isdigit() and int(v) > 0

                        # 1. Isolar quem está 100% faturado OU possui Ticket aberto para TI
                        mask_finalizar = (
                            (df_editavel['Número NF 555'].apply(is_numeric_nf) & 
                             (df_editavel['Número NF 551'].apply(is_numeric_nf)) & 
                             (df_editavel['Impresso'] == True)) |
                            (df_editavel['Ticket'].astype(str).str.strip() != "")
                        )
                        finalizados_raw = df_editavel[mask_finalizar].copy()

                        if not finalizados_raw.empty:
                            # Define as colunas conforme o padrão da planilha de histórico
                            final_columns = [
                                "N° Lote", "Rota", "AX - Cidade", "Pedido Cliente Ecommerce", "Cliente",
                                "Número NF 555", "Número NF 551", "Cód Produto", "Data Planilha de Cubagem", "Ticket"
                            ]
                            
                            # Seleciona as colunas e normaliza para MAIÚSCULO para o Google Sheets
                            finalizados = finalizados_raw[final_columns].copy()
                            finalizados.columns = [c.upper() for c in finalizados.columns]

                            # 2. Adicionar ao Histórico de Finalizados no Google Sheets
                            df_historico = pd.concat([st.session_state['bd_finalizados'], finalizados], ignore_index=True)
                            st.session_state['bd_finalizados'] = df_historico
                            salvar_bd(df_historico, PLANILHA_FINALIZADOS)
                            
                            # 3. Remover estes finalizados da tabela de Lotes Pendentes (Estoque)
                            lotes_para_remover = finalizados['N° LOTE'].astype(str).tolist()
                            df_lotes_restantes = st.session_state['bd_lotes'][~st.session_state['bd_lotes']['LOTE'].astype(str).isin(lotes_para_remover)]
                            
                            # Atualiza a memória e a planilha de lotes
                            st.session_state['bd_lotes'] = df_lotes_restantes[DEFAULT_HEADERS_LOTES]
                            salvar_bd(st.session_state['bd_lotes'], PLANILHA_LOTES)
                            
                            st.balloons()
                            st.rerun() 
                        else:
                            st.warning("⚠️ Nenhum pedido pronto para finalizar (NF 555 + NF 551 + Impresso) ou com Ticket preenchido.")
                    else:
                        st.warning("A tabela está vazia.")

            # Atualiza para o processamento de finalização fora do form
            df_fat_final = df_editavel
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
            # Remove colunas duplicadas e reseta o índice para evitar KeyError no Styler
            df_sobras = pd.DataFrame(lotes_sobrando_amarelo).loc[:, ~pd.DataFrame(lotes_sobrando_amarelo).columns.duplicated()]
            df_sobras = df_sobras.reset_index(drop=True)
            st.dataframe(df_sobras.style.set_properties(**{'background-color': '#ffeb9c', 'color': 'black'}), use_container_width=True, hide_index=True)
        else:
            st.info("Não há lotes pendentes fora da cubagem atual.")

    # --- TAB FINALIZADOS ---
    with tab_finalizados:
        st.subheader("Histórico de Pedidos Concluídos")
        if not st.session_state['bd_finalizados'].empty:
            df_hist = st.session_state['bd_finalizados']
            
            # Separando em dois tópicos
            finalizados_ok = df_hist[(df_hist['TICKET'].isna()) | (df_hist['TICKET'].astype(str).str.strip() == "")]
            finalizados_ticket = df_hist[(df_hist['TICKET'].notna()) & (df_hist['TICKET'].astype(str).str.strip() != "")]
            
            st.markdown("### ✅ Finalizados Corretamente")
            st.dataframe(finalizados_ok, use_container_width=True, hide_index=True)
            
            st.markdown("### ⚠️ Finalizados com Ticket (Problemas/TI)")
            st.dataframe(finalizados_ticket, use_container_width=True, hide_index=True)
        else:
            st.write("Nenhum pedido foi finalizado ainda.")

    # Adiciona botão na sidebar para salvar o estado atual dos lotes (mesmo sem finalizar faturamento)
    st.sidebar.button("💾 Atualizar Banco de Lotes (Estoque)", on_click=lambda: salvar_bd(st.session_state['bd_lotes'][DEFAULT_HEADERS_LOTES], PLANILHA_LOTES))

else: # This else block should remain at the very end of the script.
    st.info("👈 Por favor, faça o upload dos relatórios de Lotes Geral e Cubagem no menu lateral esquerdo para começar.")
