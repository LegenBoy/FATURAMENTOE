import streamlit as st
import pandas as pd
import os
import re
import gspread
import json

# ==========================================
# CONFIGURAÇÃO DA PÁGINA E MÓDULOS
# ==========================================
st.set_page_config(page_title="Portal de Faturamento Multi-Módulo", layout="wide")

st.sidebar.header("⚙️ Módulo de Trabalho")
modulo = st.sidebar.radio("Selecione a Operação:", ["🛍️ Ecommerce", "🏭 Transferência Matriz"])

st.title(f"📦 Portal de Faturamento - {modulo.split(' ')[1]}")

# ==========================================
# FUNÇÕES DE APOIO E LIMPEZA
# ==========================================
def clean_id(val):
    """Limpa zeros flutuantes e espaços para garantir cruzamentos perfeitos (ex: '123.0' -> '123')"""
    if pd.isna(val) or val == "": return ""
    v = str(val).strip().upper()
    if v.endswith('.0'): v = v[:-2]
    return v

def get_cod_filial(texto):
    """Extrai apenas os números da filial/fantasia para cruzamento (ex: 'AX 137 A' -> '137')"""
    return re.sub(r'\D', '', str(texto)).strip().lstrip('0')

# ==========================================
# CONFIGURAÇÃO DAS PLANILHAS GOOGLE SHEETS
# ==========================================
PLANILHA_LOTES_ECO = "lotes_pendentes_ecommerce"
PLANILHA_FINALIZADOS_ECO = "finalizados_ecommerce"

DEFAULT_HEADERS_LOTES_ECO = [
    "ROTA", "FILIAL", "CIDADE", "LOTE", "PEDIDO_ECOMMERCE",
    "PEDIDO_SITE", "PRODUTO", "DESCRICAO", "QUANTIDADE",
    "CUBTOTAL_PRODUTO", "CLIENTE", "DATA_PAGAMENTO"
]

DEFAULT_HEADERS_FINALIZADOS_ECO = [
    "N° LOTE", "ROTA", "AX - CIDADE", "PEDIDO CLIENTE ECOMMERCE",
    "CLIENTE", "NÚMERO NF 555", "NÚMERO NF 551", "CÓD PRODUTO",
    "DATA PLANILHA DE CUBAGEM", "TICKET"
]

PLANILHA_LOTES_MATRIZ = "lotes_pendentes_matriz"
PLANILHA_FINALIZADOS_MATRIZ = "finalizados_matriz"

DEFAULT_HEADERS_LOTES_MATRIZ = [
    "PRACA", "NRO_LOTE", "CIDADE", "FANTASIA", "NRO_PEDIDOS", 
    "NRO_ITENS", "PESO_TOTAL", "CUBAGEM_TOTAL", "VLR_TOTAL"
]

DEFAULT_HEADERS_FINALIZADOS_MATRIZ = [
    "DATA CARREGAMENTO", "ROTA", "AX - CIDADE", "NRO_LOTE", "QTD PEDIDOS", "VALOR TOTAL", "NÚMERO NF 55", "TICKET"
]

@st.cache_resource(ttl=3600)
def get_gsheet_client():
    try:
        creds_data = st.secrets["gcp_service_account"]
        if isinstance(creds_data, str): creds = json.loads(creds_data)
        else: creds = dict(creds_data)
        gc = gspread.service_account_from_dict(creds)
        return gc
    except Exception as e:
        st.error(f"Erro ao conectar ao Google Sheets. Verifique o st.secrets: {e}")
        st.stop()

gc = get_gsheet_client()

def carregar_bd(caminho):
    default_headers = []
    if caminho == PLANILHA_LOTES_ECO: default_headers = DEFAULT_HEADERS_LOTES_ECO
    elif caminho == PLANILHA_FINALIZADOS_ECO: default_headers = DEFAULT_HEADERS_FINALIZADOS_ECO
    elif caminho == PLANILHA_LOTES_MATRIZ: default_headers = DEFAULT_HEADERS_LOTES_MATRIZ
    elif caminho == PLANILHA_FINALIZADOS_MATRIZ: default_headers = DEFAULT_HEADERS_FINALIZADOS_MATRIZ

    try:
        sh = gc.open(caminho)
        worksheet = sh.get_worksheet(0)
        records = worksheet.get_all_records()
        
        if not records:
            if default_headers:
                worksheet.update([default_headers])
                return pd.DataFrame(columns=[h.upper() for h in default_headers])
            return pd.DataFrame()
        
        df = pd.DataFrame(records)
        if not df.empty:
            df.columns = df.columns.astype(str).str.strip().str.upper()
            for header in default_headers:
                if header.upper() not in df.columns: df[header.upper()] = ""
        return df
    except gspread.exceptions.SpreadsheetNotFound:
        try:
            sh = gc.create(caminho)
            sh.share(st.secrets["gcp_service_account"]["client_email"], perm_type='user', role='writer')
            worksheet = sh.get_worksheet(0)
            if default_headers:
                worksheet.update([default_headers])
                return pd.DataFrame(columns=[h.upper() for h in default_headers])
            return pd.DataFrame()
        except Exception: return pd.DataFrame()
    except Exception: return pd.DataFrame()

def salvar_bd(df, caminho):
    try:
        sh = gc.open(caminho)
        worksheet = sh.get_worksheet(0)
        header = [str(c) if pd.notna(c) and str(c).lower() not in ['nan', 'inf', '-inf'] else "" for c in df.columns.tolist()]
        values = [[str(val) if pd.notna(val) and str(val).lower() not in ['nan', 'inf', '-inf'] else "" for val in row] for row in df.values.tolist()]
        worksheet.clear()
        worksheet.update([header] + values)
    except Exception as e:
        st.error(f"Erro ao salvar dados na planilha '{caminho}': {e}")

# Inicializando Bases no Cache
for db_key, sheet_name in [
    ('bd_lotes_eco', PLANILHA_LOTES_ECO), ('bd_finalizados_eco', PLANILHA_FINALIZADOS_ECO),
    ('bd_lotes_matriz', PLANILHA_LOTES_MATRIZ), ('bd_finalizados_matriz', PLANILHA_FINALIZADOS_MATRIZ)
]:
    if db_key not in st.session_state: st.session_state[db_key] = carregar_bd(sheet_name)

if 'checks_persistentes_eco' not in st.session_state: st.session_state['checks_persistentes_eco'] = {}
if 'checks_persistentes_matriz' not in st.session_state: st.session_state['checks_persistentes_matriz'] = {}

# ==========================================
# IDENTIFICAÇÃO AUTOMÁTICA DOS ARQUIVOS
# ==========================================
def identificar_tipo_arquivo(df):
    colunas = df.columns.tolist()
    if 'DOCAS' in colunas and 'BATIDA' in colunas: return 'cubagem'
    elif 'LOTE' in colunas and 'PEDIDO_ECOMMERCE' in colunas: return 'lotes_geral_eco'
    elif 'NRO_LOTE' in colunas and 'FANTASIA' in colunas and 'PRACA' in colunas: return 'lotes_matriz'
    elif 'FILIAL' in colunas and 'N.F. DE SAIDA' in colunas and 'TIPO' in colunas:
        if not df.empty:
            tipo_nota = str(df['TIPO'].iloc[0]).split('.')[0].strip()
            if tipo_nota == '555': return 'faturamento_555'
            elif tipo_nota == '551': return 'faturamento_551'
            elif tipo_nota == '55': return 'faturamento_55'
    return 'desconhecido'

def extrair_ax_cidade(texto):
    try:
        if '-' in texto:
            ax = texto.split('-')[0].strip().lstrip('0')
            cidade = texto.split('-')[1].split('/')[0].strip()
            return f"{ax} - {cidade}"
        return texto
    except: return texto

# ==========================================
# ÁREA DE UPLOAD E LEITURA
# ==========================================
st.sidebar.divider()
st.sidebar.header("📤 Upload de Relatórios")
st.sidebar.write(f"Envie os arquivos do módulo: **{modulo.split(' ')[1]}**")

arquivos_upados = st.sidebar.file_uploader("Arraste os ficheiros CSV ou Excel", accept_multiple_files=True)

dados = {
    'cubagem': pd.DataFrame(), 'lotes_geral_eco': pd.DataFrame(),
    'faturamento_555': pd.DataFrame(), 'faturamento_551': pd.DataFrame(),
    'lotes_matriz': pd.DataFrame(), 'faturamento_55': pd.DataFrame()
}

if arquivos_upados:
    for arquivo in arquivos_upados:
        try:
            if arquivo.name.endswith('.csv'): df = pd.read_csv(arquivo, encoding='utf-8') 
            else: df = pd.read_excel(arquivo)
            df.columns = df.columns.astype(str).str.strip().str.upper()
            tipo = identificar_tipo_arquivo(df)
            if tipo != 'desconhecido':
                dados[tipo] = df
                st.sidebar.success(f"✅ {tipo.replace('_', ' ').upper()} carregado!")
            else:
                st.sidebar.warning(f"⚠️ Não identifiquei: {arquivo.name}")
        except Exception as e:
            st.sidebar.error(f"Erro ao ler {arquivo.name}: {e}")

st.divider()

# ==========================================
# EXTRAÇÃO DA CUBAGEM (O CORAÇÃO DO SISTEMA)
# ==========================================
filiais_info = {}
if not dados['cubagem'].empty:
    df_cubagem = dados['cubagem']
    data_cubagem = str(df_cubagem['DATA'].iloc[0]) if 'DATA' in df_cubagem.columns else "N/D"
    col_rota = 'ROTAS' if 'ROTAS' in df_cubagem.columns else ('ROTA' if 'ROTA' in df_cubagem.columns else None)
    
    for idx, row in df_cubagem.iterrows():
        rota_nome = str(row.get(col_rota, 'N/D'))
        for col in df_cubagem.columns:
            if 'filial' in col.lower() and 'cubagem' in col.lower():
                celula = str(row[col])
                if '-' in celula:
                    cod_cruzamento = celula.split('-')[0].strip().lstrip('0')
                    display_filial = extrair_ax_cidade(celula)
                    ordem_filial = col.split('/')[0].strip().replace('filial', 'Filial ')
                    filiais_info[cod_cruzamento] = {
                        "Codigo": cod_cruzamento, "Display": display_filial,
                        "Data": data_cubagem, "Rota/Ordem": f"{rota_nome} ({ordem_filial})"
                    }

# ==========================================
# 🛍️ LÓGICA DO MÓDULO ECOMMERCE
# ==========================================
if modulo == "🛍️ Ecommerce":
    if not dados['cubagem'].empty and not dados['lotes_geral_eco'].empty:
        df_lotes_hoje = dados['lotes_geral_eco'].copy()
        
        # Limpeza para Merge Seguro
        df_lotes_hoje['CHAVE_LOTE'] = df_lotes_hoje['LOTE'].apply(clean_id)
        df_lotes_hoje['CHAVE_PEDIDO'] = df_lotes_hoje['PEDIDO_ECOMMERCE'].apply(clean_id)
        
        df_hist = st.session_state['bd_lotes_eco'].copy()
        if not df_hist.empty:
            df_hist['CHAVE_LOTE'] = df_hist['LOTE'].apply(clean_id)
            df_hist['CHAVE_PEDIDO'] = df_hist['PEDIDO_ECOMMERCE'].apply(clean_id)
        else:
            df_hist['CHAVE_LOTE'] = pd.Series(dtype=str)
            df_hist['CHAVE_PEDIDO'] = pd.Series(dtype=str)

        tamanho_antes = len(df_hist)
        df_lotes_combinado = pd.concat([df_hist, df_lotes_hoje]).drop_duplicates(subset=['CHAVE_LOTE', 'CHAVE_PEDIDO'], keep='last')
        
        if len(df_lotes_combinado) > tamanho_antes:
            salvar_bd(df_lotes_combinado[DEFAULT_HEADERS_LOTES_ECO], PLANILHA_LOTES_ECO)
            st.toast("✅ Base de Lotes E-commerce atualizada!")
            st.session_state['bd_lotes_eco'] = df_lotes_combinado[DEFAULT_HEADERS_LOTES_ECO]
        
        # FILTRO DOS FINALIZADOS
        if not st.session_state['bd_finalizados_eco'].empty:
            pedidos_finalizados = set(st.session_state['bd_finalizados_eco']['PEDIDO CLIENTE ECOMMERCE'].apply(clean_id).tolist())
            df_lotes_combinado = df_lotes_combinado[~df_lotes_combinado['CHAVE_PEDIDO'].isin(pedidos_finalizados)]
        
        faturamento_view = []
        lotes_sobrando_amarelo = []
        
        # CRUZAMENTO COM A CUBAGEM
        for idx, row in df_lotes_combinado.iterrows():
            filial_lote = clean_id(row.get('FILIAL', '')).lstrip('0')
            if filial_lote in filiais_info:
                pedido = clean_id(row.get('PEDIDO_ECOMMERCE', ''))
                lote_num = clean_id(row.get('LOTE', ''))
                cod_produto = str(row.get('PRODUTO', 'N/D')).strip()
                
                status_555, status_551 = "NÃO FATURADO", "BLOQUEADO"
                st_nf_555, st_nf_551 = None, None
                
                if not dados['faturamento_555'].empty:
                    match_555 = dados['faturamento_555'][dados['faturamento_555']['LOTE'].apply(clean_id) == lote_num]
                    if not match_555.empty:
                        status_555 = clean_id(match_555['N.F. DE SAIDA'].iloc[0])
                        status_551 = "PRONTO P/ FATURAR"
                        if 'STATUS' in match_555.columns: st_nf_555 = match_555['STATUS'].iloc[0]
                
                if not dados['faturamento_551'].empty and status_555 != "NÃO FATURADO":
                    mask = dados['faturamento_551'].astype(str).apply(lambda col: col.str.contains(pedido, na=False, flags=re.IGNORECASE)).any(axis=1)
                    match_551 = dados['faturamento_551'][mask]
                    if not match_551.empty:
                        status_551 = clean_id(match_551['N.F. DE SAIDA'].iloc[0])
                        if 'STATUS' in match_551.columns: st_nf_551 = match_551['STATUS'].iloc[0]

                is_551_faturado = False
                if status_551 not in ["NÃO FATURADO", "BLOQUEADO", "PRONTO P/ FATURAR", "NAN", "", "N/D"]:
                    try: is_551_faturado = float(status_551) > 0
                    except: pass

                existing_entrada = row.get('ENTRADA', False)
                existing_impresso = row.get('IMPRESSO', False)
                existing_ticket = row.get('TICKET', "")

                chave_memoria = (lote_num, pedido)
                if chave_memoria in st.session_state['checks_persistentes_eco']:
                    mem_edits = st.session_state['checks_persistentes_eco'][chave_memoria]
                    if 'Entrada' in mem_edits: existing_entrada = mem_edits['Entrada']
                    if 'Impresso' in mem_edits: existing_impresso = mem_edits['Impresso']
                    if 'Ticket' in mem_edits: existing_ticket = mem_edits['Ticket']

                final_entrada = (str(existing_entrada).strip().upper() == 'TRUE' or existing_entrada is True) or is_551_faturado
                final_impresso = (str(existing_impresso).strip().upper() == 'TRUE' or existing_impresso is True)
                final_ticket = str(existing_ticket) if pd.notna(existing_ticket) and str(existing_ticket).strip() != "" else ""

                faturamento_view.append({
                    "Data Planilha de Cubagem": filiais_info[filial_lote]["Data"],
                    "Rota": filiais_info[filial_lote]["Rota/Ordem"],
                    "AX - Cidade": filiais_info[filial_lote]["Display"],
                    "N° Lote": lote_num,
                    "Pedido Cliente Ecommerce": pedido,
                    "Cód Produto": cod_produto,
                    "Cliente": str(row.get('CLIENTE', '')),
                    "Número NF 555": status_555,
                    "ST 555": st_nf_555,
                    "Entrada": final_entrada,
                    "Número NF 551": status_551,
                    "ST 551": st_nf_551,
                    "Impresso": final_impresso,
                    "Ticket": final_ticket,
                }) 
            else:
                # SE NÃO ESTÁ NA CUBAGEM, VAI PRO ESTOQUE PENDENTE
                lotes_sobrando_amarelo.append(row)

        df_fat_final = pd.DataFrame(faturamento_view)
        
        tab_pendentes, tab_cubagem, tab_lotes_geral, tab_finalizados = st.tabs([
            "📋 Faturamentos Pendentes", "🚛 Planilha de Cubagem", "📦 Lotes (Estoque)", "✅ Histórico"
        ])

        def colorir_texto_status(row):
            def is_nf_val(val):
                v = str(val).strip().upper()
                if v in ["NÃO FATURADO", "BLOQUEADO", "PRONTO P/ FATURAR", "NAN", "", "N/D"]: return False
                try: return float(v) > 0
                except: return False
            styles = [''] * len(row)
            fundo_verde = 'background-color: #d4edda; color: #155724; font-weight: bold;'
            fundo_amarelo = 'background-color: #fff3cd; color: #856404; font-weight: bold;'
            fundo_vermelho = 'background-color: #f8d7da; color: #721c24; font-weight: bold;'

            for col in ['Número NF 555', 'Número NF 551']:
                if col in row.index:
                    val = str(row[col]).strip().upper()
                    idx = row.index.get_loc(col)
                    if is_nf_val(val):
                        styles[idx] = fundo_verde 
                        st_col = 'ST 555' if col == 'Número NF 555' else 'ST 551'
                        if st_col in row.index and str(row.get(st_col, '')).strip() in ['6', '6.0']:
                            styles[idx] = fundo_vermelho
                    elif val == "PRONTO P/ FATURAR": styles[idx] = fundo_amarelo 
                    elif val in ["NÃO FATURADO", "BLOQUEADO"]: styles[idx] = fundo_vermelho 
            return styles

        with tab_pendentes:
            if not df_fat_final.empty:
                config_col = {
                    "Data Planilha de Cubagem": st.column_config.Column(width="small"),
                    "Rota": st.column_config.Column(width="small"),
                    "AX - Cidade": st.column_config.Column(width="medium"),
                    "N° Lote": st.column_config.Column("N° Lote", disabled=True),
                    "Número NF 555": st.column_config.Column("Número NF 555", disabled=True),
                    "Número NF 551": st.column_config.Column("Número NF 551", disabled=True),
                    "ST 555": None, "ST 551": None,
                    "Entrada": st.column_config.CheckboxColumn("Entrada"),
                    "Impresso": st.column_config.CheckboxColumn("Impresso"),
                    "Ticket": st.column_config.TextColumn("N° Ticket"),
                }
                
                with st.form("form_faturamento_eco"):
                    df_editavel = st.data_editor(df_fat_final.style.apply(colorir_texto_status, axis=1), use_container_width=True, column_config=config_col, hide_index=True, key="editor_faturamento_eco")
                    col_sync, col_finalize = st.columns([1, 1])
                    with col_sync: sync_btn = st.form_submit_button("🔄 Sincronizar Marcações", use_container_width=True)
                    with col_finalize: finalize_btn = st.form_submit_button("🚀 Finalizar Faturamentos", type="primary", use_container_width=True)

                if sync_btn:
                    edits = st.session_state.get("editor_faturamento_eco", {}).get("edited_rows", {})
                    for row_idx_str, row_changes in edits.items():
                        lote_ref = df_editavel.iloc[int(row_idx_str)]['N° Lote']
                        ped_ref = df_editavel.iloc[int(row_idx_str)]['Pedido Cliente Ecommerce']
                        if (lote_ref, ped_ref) not in st.session_state['checks_persistentes_eco']: st.session_state['checks_persistentes_eco'][(lote_ref, ped_ref)] = {}
                        for col_name, novo_valor in row_changes.items(): st.session_state['checks_persistentes_eco'][(lote_ref, ped_ref)][col_name] = novo_valor
                    st.rerun()

                if finalize_btn:
                    def is_numeric_nf(val):
                        v = str(val).split('.')[0].strip()
                        return v.isdigit() and int(v) > 0
                    finalizados_raw = df_editavel[
                        ((df_editavel['Número NF 555'].apply(is_numeric_nf)) & 
                         (df_editavel['Número NF 551'].apply(is_numeric_nf)) &
                         (df_editavel['Impresso'] == True)) |
                        (df_editavel['Ticket'].astype(str).str.strip() != "")
                    ]
                    if not finalizados_raw.empty:
                        final_columns = ["N° Lote", "Rota", "AX - Cidade", "Pedido Cliente Ecommerce", "Cliente", "Número NF 555", "Número NF 551", "Cód Produto", "Data Planilha de Cubagem", "Ticket"]
                        finalizados = finalizados_raw[final_columns].copy()
                        finalizados.columns = finalizados.columns.str.upper()

                        df_historico = pd.concat([st.session_state['bd_finalizados_eco'], finalizados], ignore_index=True)
                        st.session_state['bd_finalizados_eco'] = df_historico
                        salvar_bd(df_historico, PLANILHA_FINALIZADOS_ECO)
                        
                        lotes_remover = finalizados['N° LOTE'].apply(clean_id).tolist()
                        df_lotes_restantes = st.session_state['bd_lotes_eco'][~st.session_state['bd_lotes_eco']['LOTE'].apply(clean_id).isin(lotes_remover)]
                        st.session_state['bd_lotes_eco'] = df_lotes_restantes[DEFAULT_HEADERS_LOTES_ECO]
                        salvar_bd(st.session_state['bd_lotes_eco'], PLANILHA_LOTES_ECO)
                        st.balloons()
                        st.rerun()
                    else: st.warning("Nenhum pedido atende os requisitos para finalizar.")

        with tab_cubagem:
            # Colorir Cubagem Ecommerce
            codigos_pendentes_eco = set(df_lotes_combinado['FILIAL'].apply(get_cod_filial))
            codigos_finalizados_eco = set()
            if not st.session_state['bd_finalizados_eco'].empty:
                for ax_cid in st.session_state['bd_finalizados_eco']['AX - CIDADE']:
                    ax_part = str(ax_cid).split('-')[0]
                    num = get_cod_filial(ax_part)
                    if num: codigos_finalizados_eco.add(num)

            def colorir_cubagem_eco(df_c):
                st_df = pd.DataFrame('', index=df_c.index, columns=df_c.columns)
                for col in df_c.columns:
                    if 'FILIAL' in col and 'CUBAGEM' in col:
                        for row_idx in df_c.index:
                            val = str(df_c.at[row_idx, col])
                            if '-' in val:
                                code = val.split('-')[0].strip().lstrip('0')
                                if code in codigos_pendentes_eco: continue
                                elif code in codigos_finalizados_eco: st_df.at[row_idx, col] = 'background-color: #c6efce; color: #006100;'
                                else: st_df.at[row_idx, col] = 'background-color: #ffeb9c; color: #9c5700;'
                return st_df
            st.dataframe(dados['cubagem'].style.apply(colorir_cubagem_eco, axis=None), use_container_width=True, hide_index=True)
            
        with tab_lotes_geral:
            if lotes_sobrando_amarelo: st.dataframe(pd.DataFrame(lotes_sobrando_amarelo)[DEFAULT_HEADERS_LOTES_ECO], use_container_width=True, hide_index=True)
        with tab_finalizados:
            st.dataframe(st.session_state['bd_finalizados_eco'], use_container_width=True, hide_index=True)

    else: st.info("👈 Faça o upload dos relatórios de Lotes e Cubagem do E-commerce para começar.")


# ==========================================
# 🏭 LÓGICA DO MÓDULO TRANSFERÊNCIA MATRIZ
# ==========================================
elif modulo == "🏭 Transferência Matriz":
    st.info("💡 **Atenção**: Este painel cruza os Lotes diretamente com a planilha de **Cubagem do Dia**.")

    if not dados['cubagem'].empty and not dados['lotes_matriz'].empty:
        df_lotes_hoje = dados['lotes_matriz'].copy()
        
        # Limpeza rigorosa
        df_lotes_hoje['CHAVE_LOTE'] = df_lotes_hoje['NRO_LOTE'].apply(clean_id)
        
        df_hist = st.session_state['bd_lotes_matriz'].copy()
        if not df_hist.empty: df_hist['CHAVE_LOTE'] = df_hist['NRO_LOTE'].apply(clean_id)
        else: df_hist['CHAVE_LOTE'] = pd.Series(dtype=str)

        tamanho_antes = len(df_hist)
        df_lotes_combinado = pd.concat([df_hist, df_lotes_hoje]).drop_duplicates(subset=['CHAVE_LOTE'], keep='last')
        
        if len(df_lotes_combinado) > tamanho_antes:
            salvar_bd(df_lotes_combinado[DEFAULT_HEADERS_LOTES_MATRIZ], PLANILHA_LOTES_MATRIZ)
            st.toast("✅ Base de Lotes Matriz atualizada!")
            st.session_state['bd_lotes_matriz'] = df_lotes_combinado[DEFAULT_HEADERS_LOTES_MATRIZ]
            
        # FILTRO DOS FINALIZADOS
        if not st.session_state['bd_finalizados_matriz'].empty:
            lotes_finalizados = set(st.session_state['bd_finalizados_matriz']['NRO_LOTE'].apply(clean_id).tolist())
            df_lotes_combinado = df_lotes_combinado[~df_lotes_combinado['CHAVE_LOTE'].isin(lotes_finalizados)]

        faturamento_view = []
        lotes_sobrando_amarelo = []
        
        # CRUZAMENTO COM A CUBAGEM DA MATRIZ
        for idx, row in df_lotes_combinado.iterrows():
            fantasia_lote = str(row.get('FANTASIA', '')).strip()
            cidade_lote = str(row.get('CIDADE', '')).strip()
            
            # Extrai apenas os números da FANTASIA (ex: "AX 137 A" -> "137") para bater com a cubagem
            filial_lote_match = get_cod_filial(fantasia_lote)
            
            if filial_lote_match in filiais_info:
                rota_ordem_full = filiais_info[filial_lote_match]["Rota/Ordem"]
                data_carregamento = filiais_info[filial_lote_match]["Data"]
                
                lote_num = clean_id(row.get('NRO_LOTE', ''))
                status_55 = "NÃO FATURADO"
                st_nf_55 = None
                
                if not dados['faturamento_55'].empty:
                    match_55 = dados['faturamento_55'][dados['faturamento_55']['LOTE'].apply(clean_id) == lote_num]
                    if not match_55.empty:
                        status_55 = clean_id(match_55['N.F. DE SAIDA'].iloc[0])
                        if 'STATUS' in match_55.columns: st_nf_55 = match_55['STATUS'].iloc[0]
                            
                existing_impresso = row.get('IMPRESSO', False)
                existing_ticket = row.get('TICKET', "")

                if lote_num in st.session_state['checks_persistentes_matriz']:
                    mem_edits = st.session_state['checks_persistentes_matriz'][lote_num]
                    if 'Impresso' in mem_edits: existing_impresso = mem_edits['Impresso']
                    if 'Ticket' in mem_edits: existing_ticket = mem_edits['Ticket']

                final_impresso = (str(existing_impresso).strip().upper() == 'TRUE' or existing_impresso is True)
                final_ticket = str(existing_ticket) if pd.notna(existing_ticket) and str(existing_ticket).strip() != "" else ""

                ax_cidade = f"{fantasia_lote} - {cidade_lote}"

                faturamento_view.append({
                    "Data Carregamento": data_carregamento,
                    "Rota": rota_ordem_full,
                    "AX - Cidade": ax_cidade,
                    "Nro_Lote": lote_num,
                    "Qtd Pedidos": str(row.get('NRO_PEDIDOS', '')),
                    "Valor Total": str(row.get('VLR_TOTAL', '')),
                    "Número NF 55": status_55,
                    "ST 55": st_nf_55,
                    "Impresso": final_impresso,
                    "Ticket": final_ticket,
                })
            else:
                # SE NÃO ESTÁ NA CUBAGEM, VAI PRO ESTOQUE PENDENTE
                lotes_sobrando_amarelo.append(row)
            
        df_fat_final = pd.DataFrame(faturamento_view)
        
        tab_pendentes_matriz, tab_cubagem_matriz, tab_lotes_matriz, tab_finalizados_matriz = st.tabs([
            "📋 Faturamentos MOV 55", "🚛 Planilha de Cubagem", "📦 Lotes Matriz (Estoque)", "✅ Histórico"
        ])
        
        def colorir_texto_status_matriz(row):
            def is_nf_val(val):
                v = str(val).strip().upper()
                if v in ["NÃO FATURADO", "BLOQUEADO", "NAN", "", "N/D"]: return False
                try: return float(v) > 0
                except: return False
            styles = [''] * len(row)
            fundo_verde = 'background-color: #d4edda; color: #155724; font-weight: bold;'
            fundo_vermelho = 'background-color: #f8d7da; color: #721c24; font-weight: bold;'

            if 'Número NF 55' in row.index:
                val = str(row['Número NF 55']).strip().upper()
                idx = row.index.get_loc('Número NF 55')
                if is_nf_val(val):
                    styles[idx] = fundo_verde
                    if 'ST 55' in row.index and str(row.get('ST 55', '')).strip() in ['6', '6.0']: styles[idx] = fundo_vermelho
                elif val == "NÃO FATURADO": styles[idx] = fundo_vermelho
            return styles

        with tab_pendentes_matriz:
            if not df_fat_final.empty:
                # Expansor de Copiar Notas
                notas_para_imprimir = []
                for _, r in df_fat_final.iterrows():
                    nf_limpa = clean_id(r['Número NF 55'])
                    st_55 = str(r.get('ST 55', '')).strip()
                    if nf_limpa.isdigit() and st_55 not in ['6', '6.0'] and not r['Impresso']:
                        notas_para_imprimir.append({'NF': nf_limpa, 'Rota': str(r['Rota']), 'Cidade': str(r['AX - Cidade']), 'Lote': r['Nro_Lote']})
                                
                if notas_para_imprimir:
                    df_print = pd.DataFrame(notas_para_imprimir)
                    df_print['Rota_Base'] = df_print['Rota'].apply(lambda x: str(x).split('(')[0].strip())
                    with st.expander("🖨️ Copiar Notas MOV 55 para Impressão", expanded=True):
                        col_f1, col_f2 = st.columns(2)
                        rotas_sel = col_f1.multiselect("Filtrar por Rota:", sorted(df_print['Rota_Base'].unique()))
                        cidades_sel = col_f2.multiselect("Filtrar por Cidade:", sorted(df_print['Cidade'].unique()))

                        if rotas_sel: df_print = df_print[df_print['Rota_Base'].isin(rotas_sel)]
                        if cidades_sel: df_print = df_print[df_print['Cidade'].isin(cidades_sel)]

                        nfs_filtradas = list(dict.fromkeys(df_print['NF'].tolist()))
                        if len(nfs_filtradas) > 0:
                            if st.button("✅ Marcar notas MOV 55 como Impressas"):
                                for _, r_print in df_print.iterrows():
                                    if r_print['Lote'] not in st.session_state['checks_persistentes_matriz']: st.session_state['checks_persistentes_matriz'][r_print['Lote']] = {}
                                    st.session_state['checks_persistentes_matriz'][r_print['Lote']]['Impresso'] = True
                                st.rerun()
                            st.code("\n".join(nfs_filtradas), language="text")
                    
                config_col_matriz = {
                    "Data Carregamento": st.column_config.Column(width="small", disabled=True),
                    "Rota": st.column_config.Column(width="small", disabled=True),
                    "AX - Cidade": st.column_config.Column(disabled=True),
                    "Nro_Lote": st.column_config.Column(disabled=True),
                    "Qtd Pedidos": st.column_config.Column(disabled=True),
                    "Valor Total": st.column_config.Column(disabled=True),
                    "Número NF 55": st.column_config.Column(disabled=True),
                    "ST 55": None,
                    "Impresso": st.column_config.CheckboxColumn("Impresso"),
                    "Ticket": st.column_config.TextColumn("N° Ticket"),
                }
                
                with st.form("form_faturamento_matriz"):
                    df_editavel_matriz = st.data_editor(
                        df_fat_final.style.apply(colorir_texto_status_matriz, axis=1), 
                        use_container_width=True, column_config=config_col_matriz, hide_index=True, key="editor_faturamento_matriz"
                    )
                    col_sync, col_finalize = st.columns([1, 1])
                    with col_sync: sync_btn = st.form_submit_button("🔄 Sincronizar Marcações", use_container_width=True)
                    with col_finalize: finalize_btn = st.form_submit_button("🚀 Finalizar Faturamentos", type="primary", use_container_width=True)

                if sync_btn:
                    edits = st.session_state.get("editor_faturamento_matriz", {}).get("edited_rows", {})
                    for row_idx_str, row_changes in edits.items():
                        chave = df_editavel_matriz.iloc[int(row_idx_str)]['Nro_Lote']
                        if chave not in st.session_state['checks_persistentes_matriz']: st.session_state['checks_persistentes_matriz'][chave] = {}
                        for col_name, novo_valor in row_changes.items(): st.session_state['checks_persistentes_matriz'][chave][col_name] = novo_valor
                    st.rerun()

                if finalize_btn:
                    def is_numeric_nf(val):
                        v = str(val).split('.')[0].strip()
                        return v.isdigit() and int(v) > 0

                    finalizados_raw = df_editavel_matriz[
                        ((df_editavel_matriz['Número NF 55'].apply(is_numeric_nf)) & (df_editavel_matriz['Impresso'] == True)) |
                        (df_editavel_matriz['Ticket'].astype(str).str.strip() != "")
                    ]

                    if not finalizados_raw.empty:
                        final_cols = ["Data Carregamento", "Rota", "AX - Cidade", "Nro_Lote", "Qtd Pedidos", "Valor Total", "Número NF 55", "Ticket"]
                        finalizados = finalizados_raw[final_cols].copy()
                        finalizados.columns = finalizados.columns.str.upper().str.replace('NRO_LOTE', 'NRO_LOTE')

                        df_historico = pd.concat([st.session_state['bd_finalizados_matriz'], finalizados], ignore_index=True)
                        st.session_state['bd_finalizados_matriz'] = df_historico
                        salvar_bd(df_historico, PLANILHA_FINALIZADOS_MATRIZ)
                        
                        lotes_remover = finalizados['NRO_LOTE'].apply(clean_id).tolist()
                        df_lotes_restantes = st.session_state['bd_lotes_matriz'][~st.session_state['bd_lotes_matriz']['NRO_LOTE'].apply(clean_id).isin(lotes_remover)]
                        st.session_state['bd_lotes_matriz'] = df_lotes_restantes[DEFAULT_HEADERS_LOTES_MATRIZ]
                        salvar_bd(st.session_state['bd_lotes_matriz'], PLANILHA_LOTES_MATRIZ)
                        st.balloons()
                        st.rerun()
                    else: st.warning("⚠️ Nada a finalizar! NF 55 precisa estar faturada e a caixa 'Impresso' marcada.")
        
        with tab_cubagem_matriz:
            # Colorir Cubagem Matriz - O que falta vs O que foi faturado
            codigos_pendentes_matriz = set(df_lotes_combinado['FANTASIA'].apply(get_cod_filial))
            codigos_finalizados_matriz = set()
            if not st.session_state['bd_finalizados_matriz'].empty:
                for ax_cid in st.session_state['bd_finalizados_matriz']['AX - CIDADE']:
                    ax_part = str(ax_cid).split('-')[0]
                    num = get_cod_filial(ax_part)
                    if num: codigos_finalizados_matriz.add(num)

            def colorir_cubagem_matriz(df_c):
                st_df = pd.DataFrame('', index=df_c.index, columns=df_c.columns)
                for col in df_c.columns:
                    if 'FILIAL' in col and 'CUBAGEM' in col:
                        for row_idx in df_c.index:
                            val = str(df_c.at[row_idx, col])
                            if '-' in val:
                                code = val.split('-')[0].strip().lstrip('0')
                                if code in codigos_pendentes_matriz: continue
                                elif code in codigos_finalizados_matriz: st_df.at[row_idx, col] = 'background-color: #c6efce; color: #006100;'
                                else: st_df.at[row_idx, col] = 'background-color: #ffeb9c; color: #9c5700;'
                return st_df

            st.dataframe(dados['cubagem'].style.apply(colorir_cubagem_matriz, axis=None), use_container_width=True, hide_index=True)

        with tab_lotes_matriz:
            if lotes_sobrando_amarelo: st.dataframe(pd.DataFrame(lotes_sobrando_amarelo)[DEFAULT_HEADERS_LOTES_MATRIZ], use_container_width=True, hide_index=True)
            else: st.info("Não há lotes pendentes fora da cubagem atual.")
            
        with tab_finalizados_matriz:
            st.dataframe(st.session_state['bd_finalizados_matriz'], use_container_width=True, hide_index=True)

    else: st.info("👈 Por favor, faça o upload da Planilha de Cubagem e dos Lotes da Matriz (Previa) no menu lateral.")

# ==========================================
# BOTÕES DE UTILIDADES NO MENU LATERAL
# ==========================================
st.sidebar.divider()
st.sidebar.caption("Como várias pessoas operam, clique para puxar dados em tempo real:")
if st.sidebar.button("🔄 Sincronizar com Planilhas Google"):
    for k in ['bd_lotes_eco', 'bd_finalizados_eco', 'bd_lotes_matriz', 'bd_finalizados_matriz']:
        if k in st.session_state: del st.session_state[k]
    st.rerun()
