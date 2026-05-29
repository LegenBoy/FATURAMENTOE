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
# CONFIGURAÇÃO DAS PLANILHAS GOOGLE SHEETS
# ==========================================
# ECOMMERCE
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

# MATRIZ
PLANILHA_LOTES_MATRIZ = "lotes_pendentes_matriz"
PLANILHA_FINALIZADOS_MATRIZ = "finalizados_matriz"

DEFAULT_HEADERS_LOTES_MATRIZ = [
    "PRACA", "NRO_LOTE", "CIDADE", "FANTASIA", "NRO_PEDIDOS", 
    "NRO_ITENS", "PESO_TOTAL", "CUBAGEM_TOTAL", "VLR_TOTAL"
]

DEFAULT_HEADERS_FINALIZADOS_MATRIZ = [
    "NRO_LOTE", "PRACA", "CIDADE", "FANTASIA", "NÚMERO NF 55", "TICKET"
]

# Configuração de Conexão GSheets
@st.cache_resource(ttl=3600)
def get_gsheet_client():
    try:
        creds_data = st.secrets["gcp_service_account"]
        if isinstance(creds_data, str):
            creds = json.loads(creds_data)
        else:
            creds = dict(creds_data)
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
                if header.upper() not in df.columns:
                    df[header.upper()] = ""
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
        st.success(f"Dados salvos na planilha '{caminho}' com sucesso!")
    except Exception as e:
        st.error(f"Erro ao salvar dados na planilha '{caminho}': {e}")

# Inicializando e Separando as Bases de Dados no Cache
if 'bd_lotes_eco' not in st.session_state: st.session_state['bd_lotes_eco'] = carregar_bd(PLANILHA_LOTES_ECO)
if 'bd_finalizados_eco' not in st.session_state: st.session_state['bd_finalizados_eco'] = carregar_bd(PLANILHA_FINALIZADOS_ECO)
if 'checks_persistentes_eco' not in st.session_state: st.session_state['checks_persistentes_eco'] = {}

if 'bd_lotes_matriz' not in st.session_state: st.session_state['bd_lotes_matriz'] = carregar_bd(PLANILHA_LOTES_MATRIZ)
if 'bd_finalizados_matriz' not in st.session_state: st.session_state['bd_finalizados_matriz'] = carregar_bd(PLANILHA_FINALIZADOS_MATRIZ)
if 'checks_persistentes_matriz' not in st.session_state: st.session_state['checks_persistentes_matriz'] = {}

# ==========================================
# IDENTIFICAÇÃO AUTOMÁTICA DOS ARQUIVOS
# ==========================================
def identificar_tipo_arquivo(df):
    colunas = df.columns.tolist()
    # Padrões Ecommerce
    if 'DOCAS' in colunas and 'BATIDA' in colunas:
        return 'cubagem_eco'
    elif 'LOTE' in colunas and 'PEDIDO_ECOMMERCE' in colunas:
        return 'lotes_geral_eco'
    # Padrão Matriz (Lotes)
    elif 'NRO_LOTE' in colunas and 'FANTASIA' in colunas and 'PRACA' in colunas:
        return 'lotes_matriz'
    # Padrões de Faturamento (Notas Fiscais)
    elif 'FILIAL' in colunas and 'N.F. DE SAIDA' in colunas and 'TIPO' in colunas:
        if not df.empty:
            tipo_nota = str(df['TIPO'].iloc[0]).split('.')[0].strip()
            if tipo_nota == '555': return 'faturamento_555'
            elif tipo_nota == '551': return 'faturamento_551'
            elif tipo_nota == '55': return 'faturamento_55'
    return 'desconhecido'

# ==========================================
# ÁREA DE UPLOAD E LEITURA
# ==========================================
st.sidebar.divider()
st.sidebar.header("📤 Upload de Relatórios")
st.sidebar.write(f"Envie os arquivos referentes ao módulo: **{modulo.split(' ')[1]}**")

arquivos_upados = st.sidebar.file_uploader("Arraste os ficheiros CSV ou Excel", accept_multiple_files=True)

dados = {
    'cubagem_eco': pd.DataFrame(), 'lotes_geral_eco': pd.DataFrame(),
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
# 🛒 LÓGICA DO MÓDULO ECOMMERCE
# ==========================================
if modulo == "🛍️ Ecommerce":
    if not dados['cubagem_eco'].empty and not dados['lotes_geral_eco'].empty:
        df_lotes_hoje = dados['lotes_geral_eco']
        df_lotes_historico = st.session_state['bd_lotes_eco']
        
        df_lotes_combinado = pd.concat([df_lotes_historico, df_lotes_hoje]).drop_duplicates(subset=['LOTE', 'PEDIDO_ECOMMERCE'])
        st.session_state['bd_lotes_eco'] = df_lotes_combinado
        
        if not st.session_state['bd_finalizados_eco'].empty:
            pedidos_finalizados = set(st.session_state['bd_finalizados_eco']['PEDIDO CLIENTE ECOMMERCE'].astype(str).tolist())
            df_lotes_combinado = df_lotes_combinado[~df_lotes_combinado['PEDIDO_ECOMMERCE'].astype(str).isin(pedidos_finalizados)]

        def extrair_ax_cidade(texto):
            try:
                if '-' in texto:
                    ax = texto.split('-')[0].strip().lstrip('0')
                    cidade = texto.split('-')[1].split('/')[0].strip()
                    return f"{ax} - {cidade}"
                return texto
            except: return texto

        filiais_info = {}
        df_cubagem = dados['cubagem_eco']
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
                
                status_555, status_551 = "NÃO FATURADO", "BLOQUEADO"
                st_nf_555, st_nf_551 = None, None
                
                if not dados['faturamento_555'].empty:
                    match_555 = dados['faturamento_555'][dados['faturamento_555']['LOTE'].astype(str).str.strip() == lote_num]
                    if not match_555.empty:
                        nf_555 = str(match_555['N.F. DE SAIDA'].iloc[0])
                        status_555 = nf_555
                        status_551 = "PRONTO P/ FATURAR"
                        if 'STATUS' in match_555.columns: st_nf_555 = match_555['STATUS'].iloc[0]
                
                if not dados['faturamento_551'].empty and status_555 != "NÃO FATURADO":
                    mask = dados['faturamento_551'].astype(str).apply(lambda col: col.str.contains(pedido, na=False, flags=re.IGNORECASE)).any(axis=1)
                    match_551 = dados['faturamento_551'][mask]
                    if not match_551.empty:
                        status_551 = str(match_551['N.F. DE SAIDA'].iloc[0])
                        if 'STATUS' in match_551.columns: st_nf_551 = match_551['STATUS'].iloc[0]

                is_551_faturado = False
                v551 = str(status_551).strip().upper()
                if v551 not in ["NÃO FATURADO", "BLOQUEADO", "PRONTO P/ FATURAR", "NAN", "NONE", "N/D", ""]:
                    try: is_551_faturado = float(v551) > 0
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

                rota_ordem_full = filiais_info[filial_lote]["Rota/Ordem"]
                display_filial_full = filiais_info[filial_lote]["Display"]
                ax_val = display_filial_full.split('-')[0].strip() if '-' in display_filial_full else ""
                cidade_val = display_filial_full.split('-')[1].strip() if '-' in display_filial_full else display_filial_full
                
                faturamento_view.append({
                    "Data Planilha de Cubagem": filiais_info[filial_lote]["Data"],
                    "Rota": rota_ordem_full,
                    "AX - Cidade": f"{ax_val} - {cidade_val}",
                    "N° Lote": lote_num,
                    "Pedido Cliente Ecommerce": pedido,
                    "Cód Produto": cod_produto,
                    "Cliente": row.get('CLIENTE', ''),
                    "Número NF 555": status_555,
                    "ST 555": st_nf_555,
                    "Entrada": final_entrada,
                    "Número NF 551": status_551,
                    "ST 551": st_nf_551,
                    "Impresso": final_impresso,
                    "Ticket": final_ticket,
                }) 
            else: lotes_sobrando_amarelo.append(row)

        df_fat_final = pd.DataFrame(faturamento_view)
        
        codigos_pendentes = set(df_lotes_combinado['FILIAL'].astype(str).str.strip().str.lstrip('0').unique())
        codigos_finalizados = set()
        if not st.session_state['bd_finalizados_eco'].empty:
            codigos_finalizados = set(st.session_state['bd_finalizados_eco']['AX - CIDADE'].astype(str).str.split('-').str[0].str.strip().str.lstrip('0').unique())

        def colorir_cubagem(df_c):
            st_df = pd.DataFrame('', index=df_c.index, columns=df_c.columns)
            for col in df_c.columns:
                if 'FILIAL' in col and 'CUBAGEM' in col:
                    for row_idx in df_c.index:
                        val = str(df_c.at[row_idx, col])
                        if '-' in val:
                            code = val.split('-')[0].strip().lstrip('0')
                            if code in codigos_pendentes: continue
                            elif code in codigos_finalizados: st_df.at[row_idx, col] = 'background-color: #c6efce; color: #006100;'
                            else: st_df.at[row_idx, col] = 'background-color: #ffeb9c; color: #9c5700;'
            return st_df

        tab_pendentes, tab_cubagem, tab_lotes_geral, tab_finalizados = st.tabs([
            "📋 Faturamentos Pendentes", "🚛 Planilha de Cubagem", "📦 Lotes Geral (Estoque)", "✅ Histórico"
        ])

        def colorir_texto_status(row):
            def is_nf_val(val):
                v = str(val).strip().upper()
                if v in ["NÃO FATURADO", "BLOQUEADO", "PRONTO P/ FATURAR", "NAN", "NONE", "N/D", "", "NULL"]: return False
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
            st.subheader("Painel de Faturamento do Dia")
            if not df_fat_final.empty:
                def is_nf_val_static(val):
                    v = str(val).strip().upper()
                    if v in ["NÃO FATURADO", "BLOQUEADO", "PRONTO P/ FATURAR", "NAN", "NONE", "N/D", ""]: return False
                    try: return float(v) > 0
                    except: return False

                notas_para_imprimir = []
                nfs_bloqueadas_status6 = []
                
                for _, r in df_fat_final.iterrows():
                    st_551 = str(r.get('ST 551', '')).strip()
                    is_err_6 = st_551 in ['6', '6.0']
                    if is_nf_val_static(r['Número NF 551']):
                        nf_limpa = str(r['Número NF 551']).split('.')[0].strip()
                        if is_err_6: nfs_bloqueadas_status6.append(nf_limpa)
                        elif not r['Impresso']:
                            if nf_limpa.isdigit():
                                notas_para_imprimir.append({
                                    'NF': nf_limpa, 'Rota': str(r['Rota']), 'Cidade': str(r['AX - Cidade']),
                                    'Lote': r['N° Lote'], 'Pedido': r['Pedido Cliente Ecommerce']
                                })

                if notas_para_imprimir:
                    df_print_base = pd.DataFrame(notas_para_imprimir)
                    df_print_base['Rota_Base'] = df_print_base['Rota'].apply(lambda x: str(x).split('(')[0].strip())
                    
                    with st.expander("🖨️ Copiar Notas para Impressão", expanded=True):
                        col_f1, col_f2 = st.columns(2)
                        with col_f1:
                            rotas_disp = sorted(df_print_base['Rota_Base'].unique())
                            rotas_selecionadas = st.multiselect("Filtrar por Rota principal:", rotas_disp)
                        with col_f2:
                            cidades_disp = sorted(df_print_base['Cidade'].unique())
                            cidades_selecionadas = st.multiselect("Filtrar por Cidade:", cidades_disp)
                        
                        df_print_filtrado = df_print_base
                        if rotas_selecionadas: df_print_filtrado = df_print_filtrado[df_print_filtrado['Rota_Base'].isin(rotas_selecionadas)]
                        if cidades_selecionadas: df_print_filtrado = df_print_filtrado[df_print_filtrado['Cidade'].isin(cidades_selecionadas)]
                        
                        nfs_filtradas_lista = list(dict.fromkeys(df_print_filtrado['NF'].tolist()))
                        st.info(f"Mostrando **{len(nfs_filtradas_lista)}** de notas exclusivas para impressão.")
                        
                        if len(nfs_filtradas_lista) > 0:
                            if st.button("✅ Marcar notas filtradas como Impressas"):
                                for _, r_print in df_print_filtrado.iterrows():
                                    chave = (r_print['Lote'], r_print['Pedido'])
                                    if chave not in st.session_state['checks_persistentes_eco']: st.session_state['checks_persistentes_eco'][chave] = {}
                                    st.session_state['checks_persistentes_eco'][chave]['Impresso'] = True
                                st.success("Notas marcadas! Clique no botão de Sincronizar.")
                                st.rerun()

                            st.code("\n".join(nfs_filtradas_lista), language="text")

                if nfs_bloqueadas_status6: st.error(f"🚨 ALERTA: Notas canceladas: {', '.join(nfs_bloqueadas_status6)}")

                config_col = {
                    "Data Planilha de Cubagem": st.column_config.Column(width="small"),
                    "Rota": st.column_config.Column(width="small"),
                    "AX - Cidade": st.column_config.Column(width="medium"),
                    "N° Lote": st.column_config.Column("N° Lote"),
                    "Número NF 555": st.column_config.Column("Número NF 555"),
                    "Número NF 551": st.column_config.Column("Número NF 551"),
                    "ST 555": None, "ST 551": None,
                    "Entrada": st.column_config.CheckboxColumn("Entrada"),
                    "Impresso": st.column_config.CheckboxColumn("Impresso"),
                    "Ticket": st.column_config.TextColumn("N° Ticket"),
                }
                
                with st.form("form_faturamento_eco"):
                    df_editavel = st.data_editor(df_fat_final.style.apply(colorir_texto_status, axis=1), use_container_width=True, column_config=config_col, hide_index=True, key="editor_faturamento_eco")
                    col_sync, col_finalize = st.columns([1, 1])
                    with col_sync: sync_btn = st.form_submit_button("🔄 Sincronizar e Salvar", use_container_width=True)
                    with col_finalize: finalize_btn = st.form_submit_button("🚀 Finalizar Faturamentos", type="primary", use_container_width=True)

                if sync_btn:
                    edits = st.session_state.get("editor_faturamento_eco", {}).get("edited_rows", {})
                    if edits:
                        for row_idx_str, row_changes in edits.items():
                            row_idx = int(row_idx_str)
                            lote_ref = df_editavel.iloc[row_idx]['N° Lote']
                            ped_ref = df_editavel.iloc[row_idx]['Pedido Cliente Ecommerce']
                            chave_origem = (lote_ref, ped_ref)

                            if chave_origem not in st.session_state['checks_persistentes_eco']: st.session_state['checks_persistentes_eco'][chave_origem] = {}
                            for col_name, novo_valor in row_changes.items(): st.session_state['checks_persistentes_eco'][chave_origem][col_name] = novo_valor
                        st.rerun()

                if finalize_btn:
                    if not df_editavel.empty:
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
                            final_columns_for_gsheet = ["N° Lote", "Rota", "AX - Cidade", "Pedido Cliente Ecommerce", "Cliente", "Número NF 555", "Número NF 551", "Cód Produto", "Data Planilha de Cubagem", "Ticket"]
                            finalizados = finalizados_raw[final_columns_for_gsheet].copy()
                            finalizados.columns = finalizados.columns.str.upper()

                            df_historico = pd.concat([st.session_state['bd_finalizados_eco'], finalizados], ignore_index=True)
                            st.session_state['bd_finalizados_eco'] = df_historico
                            salvar_bd(df_historico, PLANILHA_FINALIZADOS_ECO)
                            
                            lotes_para_remover = finalizados['N° LOTE'].astype(str).tolist()
                            df_lotes_restantes = st.session_state['bd_lotes_eco'][~st.session_state['bd_lotes_eco']['LOTE'].astype(str).isin(lotes_para_remover)]
                            st.session_state['bd_lotes_eco'] = df_lotes_restantes[DEFAULT_HEADERS_LOTES_ECO]
                            salvar_bd(st.session_state['bd_lotes_eco'], PLANILHA_LOTES_ECO)
                            st.balloons()
                            st.rerun()
                        else: st.warning("Nenhum pedido atende os requisitos para finalizar (NFs + Impresso ou Ticket).")
        with tab_cubagem:
            if not dados['cubagem_eco'].empty: st.dataframe(dados['cubagem_eco'].style.apply(colorir_cubagem, axis=None), use_container_width=True, hide_index=True)
        with tab_lotes_geral:
            if lotes_sobrando_amarelo:
                df_sobras = pd.DataFrame(lotes_sobrando_amarelo).loc[:, ~pd.DataFrame(lotes_sobrando_amarelo).columns.duplicated()].reset_index(drop=True)
                st.dataframe(df_sobras.style.set_properties(**{'background-color': '#ffeb9c', 'color': 'black'}), use_container_width=True, hide_index=True)
        with tab_finalizados:
            if not st.session_state['bd_finalizados_eco'].empty:
                df_hist = st.session_state['bd_finalizados_eco']
                c_ticket = df_hist['TICKET'].isna() | df_hist['TICKET'].astype(str).str.strip().isin(["", "nan", "None", "NaN"])
                st.markdown("### ✅ Finalizados Corretamente"); st.dataframe(df_hist[c_ticket], use_container_width=True, hide_index=True)
                st.markdown("### ⚠️ Finalizados com Ticket"); st.dataframe(df_hist[~c_ticket], use_container_width=True, hide_index=True)
    else: st.info("👈 Faça o upload dos relatórios de Lotes e Cubagem do E-commerce para começar.")


# ==========================================
# 🏭 LÓGICA DO MÓDULO TRANSFERÊNCIA MATRIZ
# ==========================================
elif modulo == "🏭 Transferência Matriz":
    st.info("💡 **Atenção (Multiusuários)**: Como mais de uma pessoa acessa este módulo da Matriz, lembre-se sempre de clicar em 'Sincronizar com Planilhas Google' no menu lateral caso queira puxar dados feitos por outro operador.")

    if not dados['lotes_matriz'].empty or not st.session_state['bd_lotes_matriz'].empty:
        df_lotes_hoje = dados['lotes_matriz']
        df_lotes_historico = st.session_state['bd_lotes_matriz']
        
        if not df_lotes_hoje.empty:
            df_lotes_combinado = pd.concat([df_lotes_historico, df_lotes_hoje]).drop_duplicates(subset=['NRO_LOTE'])
            st.session_state['bd_lotes_matriz'] = df_lotes_combinado
        else:
            df_lotes_combinado = df_lotes_historico.copy()
            
        if not st.session_state['bd_finalizados_matriz'].empty:
            lotes_finalizados = set(st.session_state['bd_finalizados_matriz']['NRO_LOTE'].astype(str).tolist())
            df_lotes_combinado = df_lotes_combinado[~df_lotes_combinado['NRO_LOTE'].astype(str).isin(lotes_finalizados)]

        faturamento_view = []
        for idx, row in df_lotes_combinado.iterrows():
            lote_num = str(row.get('NRO_LOTE', '')).strip()
            status_55 = "NÃO FATURADO"
            st_nf_55 = None
            
            if not dados['faturamento_55'].empty:
                match_55 = dados['faturamento_55'][dados['faturamento_55']['LOTE'].astype(str).str.strip() == lote_num]
                if not match_55.empty:
                    status_55 = str(match_55['N.F. DE SAIDA'].iloc[0])
                    if 'STATUS' in match_55.columns:
                        st_nf_55 = match_55['STATUS'].iloc[0]
                        
            existing_impresso = row.get('IMPRESSO', False)
            existing_ticket = row.get('TICKET', "")

            chave_memoria = lote_num
            if chave_memoria in st.session_state['checks_persistentes_matriz']:
                mem_edits = st.session_state['checks_persistentes_matriz'][chave_memoria]
                if 'Impresso' in mem_edits: existing_impresso = mem_edits['Impresso']
                if 'Ticket' in mem_edits: existing_ticket = mem_edits['Ticket']

            final_impresso = (str(existing_impresso).strip().upper() == 'TRUE' or existing_impresso is True)
            final_ticket = str(existing_ticket) if pd.notna(existing_ticket) and str(existing_ticket).strip() != "" else ""

            faturamento_view.append({
                "N° Lote": lote_num,
                "Praça": str(row.get('PRACA', '')),
                "Cidade": str(row.get('CIDADE', '')),
                "Fantasia": str(row.get('FANTASIA', '')),
                "Qtd Pedidos": str(row.get('NRO_PEDIDOS', '')),
                "Valor Total": str(row.get('VLR_TOTAL', '')),
                "Número NF 55": status_55,
                "ST 55": st_nf_55,
                "Impresso": final_impresso,
                "Ticket": final_ticket,
            })
            
        df_fat_final = pd.DataFrame(faturamento_view)
        
        tab_pendentes_matriz, tab_lotes_matriz, tab_finalizados_matriz = st.tabs([
            "📋 Faturamentos MOV 55", "📦 Base de Lotes (Matriz)", "✅ Histórico"
        ])
        
        def colorir_texto_status_matriz(row):
            def is_nf_val(val):
                v = str(val).strip().upper()
                if v in ["NÃO FATURADO", "BLOQUEADO", "NAN", "NONE", "N/D", "", "NULL"]: return False
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
                def is_nf_val_static(val):
                    v = str(val).strip().upper()
                    if v in ["NÃO FATURADO", "BLOQUEADO", "NAN", "NONE", "N/D", ""]: return False
                    try: return float(v) > 0
                    except: return False

                notas_para_imprimir = []
                nfs_bloqueadas_status6 = []

                for _, r in df_fat_final.iterrows():
                    st_55 = str(r.get('ST 55', '')).strip()
                    is_err_6 = st_55 in ['6', '6.0']

                    if is_nf_val_static(r['Número NF 55']):
                        nf_limpa = str(r['Número NF 55']).split('.')[0].strip()
                        if is_err_6: nfs_bloqueadas_status6.append(nf_limpa)
                        elif not r['Impresso']:
                            if nf_limpa.isdigit():
                                notas_para_imprimir.append({
                                    'NF': nf_limpa, 'Fantasia': str(r['Fantasia']), 'Cidade': str(r['Cidade']), 'Lote': r['N° Lote']
                                })
                                
                if notas_para_imprimir:
                    df_print_base = pd.DataFrame(notas_para_imprimir)
                    df_print_base['Fantasia_Base'] = df_print_base['Fantasia'].apply(lambda x: str(x).split('-')[0].strip() if '-' in str(x) else str(x))

                    with st.expander("🖨️ Copiar Notas para Impressão", expanded=True):
                        col_f1, col_f2 = st.columns(2)
                        with col_f1:
                            fantasias_disp = sorted(df_print_base['Fantasia_Base'].unique())
                            fantasias_selecionadas = st.multiselect("Filtrar por Filial/Fantasia:", fantasias_disp)
                        with col_f2:
                            cidades_disp = sorted(df_print_base['Cidade'].unique())
                            cidades_selecionadas = st.multiselect("Filtrar por Cidade:", cidades_disp)

                        df_print_filtrado = df_print_base
                        if fantasias_selecionadas: df_print_filtrado = df_print_filtrado[df_print_filtrado['Fantasia_Base'].isin(fantasias_selecionadas)]
                        if cidades_selecionadas: df_print_filtrado = df_print_filtrado[df_print_filtrado['Cidade'].isin(cidades_selecionadas)]

                        nfs_filtradas_lista = list(dict.fromkeys(df_print_filtrado['NF'].tolist()))
                        st.info(f"Mostrando **{len(nfs_filtradas_lista)}** de notas MOV 55 prontas.")

                        if len(nfs_filtradas_lista) > 0:
                            if st.button("✅ Marcar notas MOV 55 como Impressas"):
                                for _, r_print in df_print_filtrado.iterrows():
                                    chave = r_print['Lote']
                                    if chave not in st.session_state['checks_persistentes_matriz']: st.session_state['checks_persistentes_matriz'][chave] = {}
                                    st.session_state['checks_persistentes_matriz'][chave]['Impresso'] = True
                                st.success("Notas marcadas! Clique no botão Sincronizar embaixo.")
                                st.rerun()

                            st.code("\n".join(nfs_filtradas_lista), language="text")

                if nfs_bloqueadas_status6: st.error(f"🚨 ALERTA: As notas {', '.join(nfs_bloqueadas_status6)} estão canceladas.")
                    
                config_col_matriz = {
                    "N° Lote": st.column_config.Column(width="small", disabled=True),
                    "Praça": st.column_config.Column(disabled=True),
                    "Cidade": st.column_config.Column(disabled=True),
                    "Fantasia": st.column_config.Column(disabled=True),
                    "Qtd Pedidos": st.column_config.Column(disabled=True),
                    "Valor Total": st.column_config.Column(disabled=True),
                    "Número NF 55": st.column_config.Column("Número NF 55", disabled=True),
                    "ST 55": None,
                    "Impresso": st.column_config.CheckboxColumn("Impresso"),
                    "Ticket": st.column_config.TextColumn("N° Ticket", help="Digite se houver problema TI"),
                }
                
                with st.form("form_faturamento_matriz"):
                    df_editavel_matriz = st.data_editor(
                        df_fat_final.style.apply(colorir_texto_status_matriz, axis=1), 
                        use_container_width=True, column_config=config_col_matriz, hide_index=True, key="editor_faturamento_matriz"
                    )

                    col_sync, col_finalize = st.columns([1, 1])
                    with col_sync: sync_btn = st.form_submit_button("🔄 Sincronizar e Salvar Marcações", use_container_width=True)
                    with col_finalize: finalize_btn = st.form_submit_button("🚀 Finalizar Faturamentos (NF 55)", type="primary", use_container_width=True)

                if sync_btn:
                    edits = st.session_state.get("editor_faturamento_matriz", {}).get("edited_rows", {})
                    if edits:
                        for row_idx_str, row_changes in edits.items():
                            row_idx = int(row_idx_str)
                            chave_origem = df_editavel_matriz.iloc[row_idx]['N° Lote']
                            if chave_origem not in st.session_state['checks_persistentes_matriz']: st.session_state['checks_persistentes_matriz'][chave_origem] = {}
                            for col_name, novo_valor in row_changes.items(): st.session_state['checks_persistentes_matriz'][chave_origem][col_name] = novo_valor
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
                        final_columns_for_gsheet = ["N° Lote", "Praça", "Cidade", "Fantasia", "Número NF 55", "Ticket"]
                        finalizados = finalizados_raw[final_columns_for_gsheet].copy()
                        finalizados.columns = finalizados.columns.str.upper().str.replace('N° LOTE', 'NRO_LOTE').str.replace('PRAÇA', 'PRACA')

                        df_historico = pd.concat([st.session_state['bd_finalizados_matriz'], finalizados], ignore_index=True)
                        st.session_state['bd_finalizados_matriz'] = df_historico
                        salvar_bd(df_historico, PLANILHA_FINALIZADOS_MATRIZ)
                        
                        lotes_para_remover = finalizados['NRO_LOTE'].astype(str).tolist()
                        df_lotes_restantes = st.session_state['bd_lotes_matriz'][~st.session_state['bd_lotes_matriz']['NRO_LOTE'].astype(str).isin(lotes_para_remover)]
                        st.session_state['bd_lotes_matriz'] = df_lotes_restantes[DEFAULT_HEADERS_LOTES_MATRIZ]
                        salvar_bd(st.session_state['bd_lotes_matriz'], PLANILHA_LOTES_MATRIZ)
                        st.balloons()
                        st.rerun()
                    else:
                        st.warning("⚠️ Nada a finalizar! A NF 55 precisa estar faturada e a caixa 'Impresso' precisa estar marcada (Ou um ticket deve ser preenchido).")
        
        with tab_lotes_matriz:
            st.dataframe(df_lotes_combinado, use_container_width=True, hide_index=True)

        with tab_finalizados_matriz:
            if not st.session_state['bd_finalizados_matriz'].empty:
                df_hist = st.session_state['bd_finalizados_matriz']
                c_ticket = df_hist['TICKET'].isna() | df_hist['TICKET'].astype(str).str.strip().isin(["", "nan", "None", "NaN"])
                st.markdown("### ✅ Finalizados Corretamente"); st.dataframe(df_hist[c_ticket], use_container_width=True, hide_index=True)
                st.markdown("### ⚠️ Finalizados com Ticket"); st.dataframe(df_hist[~c_ticket], use_container_width=True, hide_index=True)
    else: st.info("👈 Por favor, faça o upload da Previa_Cubagens (Lotes Matriz) no menu lateral esquerdo.")

# ==========================================
# BOTÕES DE UTILIDADES E ATUALIZAÇÃO NO MENU LATERAL
# ==========================================
st.sidebar.divider()
if modulo == "🛍️ Ecommerce":
    st.sidebar.button("💾 Salvar Banco Lotes (Estoque Eco)", on_click=lambda: salvar_bd(st.session_state['bd_lotes_eco'][DEFAULT_HEADERS_LOTES_ECO], PLANILHA_LOTES_ECO))
else:
    st.sidebar.button("💾 Salvar Banco Lotes (Estoque Matriz)", on_click=lambda: salvar_bd(st.session_state['bd_lotes_matriz'][DEFAULT_HEADERS_LOTES_MATRIZ], PLANILHA_LOTES_MATRIZ))

st.sidebar.markdown("---")
st.sidebar.caption("Como várias pessoas operam, use este botão para puxar atualizações dos colegas em tempo real:")
if st.sidebar.button("🔄 Sincronizar com Planilhas Google"):
    for k in ['bd_lotes_eco', 'bd_finalizados_eco', 'bd_lotes_matriz', 'bd_finalizados_matriz']:
        if k in st.session_state: del st.session_state[k]
    st.rerun()
