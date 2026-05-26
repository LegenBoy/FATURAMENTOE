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
        return pd.read_excel(caminho)
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
    colunas = df.columns.astype(str).str.lower().tolist()
    if 'docas' in colunas and 'batida' in colunas:
        return 'cubagem'
    elif 'lote' in colunas and 'pedido_ecommerce' in colunas:
        return 'lotes_geral'
    elif 'filial ' in colunas and 'n.f. de saida ' in colunas and 'tipo ' in colunas:
        # Verifica a primeira linha para descobrir se é a nota 555 ou 551
        if not df.empty:
            tipo_nota = str(df['Tipo '].iloc[0]).strip()
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
    
    # 2. Extrair Filiais ativas na Cubagem de hoje
    filiais_na_cubagem = []
    df_cubagem = dados['cubagem']
    
    for col in df_cubagem.columns[4:16]: # Colunas das filiais/lojas na sua planilha
        filiais_encontradas = df_cubagem[col].dropna().astype(str)
        for f in filiais_encontradas:
            if '-' in f:
                cod_filial = f.split('-')[0].strip().lstrip('0')
                filiais_na_cubagem.append(cod_filial)
                
    filiais_na_cubagem = set(filiais_na_cubagem)
    
    # 3. Cruzamento de Dados (Lotes vs Cubagem vs Notas Fiscais)
    faturamento_view = []
    lotes_sobrando_amarelo = []
    
    for idx, row in df_lotes_combinado.iterrows():
        filial_lote = str(row.get('FILIAL', '')).strip().lstrip('0')
        pedido = str(row.get('PEDIDO_ECOMMERCE', '')).strip()
        lote_num = str(row.get('LOTE', '')).strip()
        
        status_555 = "NÃO FATURADO"
        status_551 = "BLOQUEADO"
        
        # Procura a Nota 555 (pelo Lote)
        if not dados['faturamento_555'].empty:
            tem_555 = lote_num in dados['faturamento_555']['Lote '].astype(str).str.strip().values
            if tem_555:
                status_555 = "✅ FATURADO (555)"
                status_551 = "PRONTO P/ FATURAR" # Libera para a nota 551
        
        # Procura a Nota 551 (pelo Pedido Ecommerce)
        if not dados['faturamento_551'].empty and status_555.startswith("✅"):
            # A nota 551 tem o pedido no meio da string, então procuramos se a string do pedido existe em alguma coluna
            tem_551 = dados['faturamento_551'].astype(str).apply(lambda col: col.str.contains(pedido, na=False, flags=re.IGNORECASE)).any().any()
            if tem_551:
                status_551 = "✅ FATURADO (551)"

        if filial_lote in filiais_na_cubagem:
            faturamento_view.append({
                "Lote": lote_num,
                "Filial": filial_lote,
                "Pedido": pedido,
                "Cliente": row.get('CLIENTE', ''),
                "Status 555": status_555,
                "Status 551": status_551
            })
        else:
            # Lote existe, mas a filial não está na cubagem (Amarelo)
            lotes_sobrando_amarelo.append(row)

    df_fat_final = pd.DataFrame(faturamento_view)
    
    # ==========================================
    # APRESENTAÇÃO NO ECRÃ (Tabelas e Cores)
    # ==========================================
    
    def colorir_faturamento(val):
        if "✅" in str(val):
            return 'background-color: #c6efce; color: #006100; font-weight: bold;' # Verde
        elif val == "PRONTO P/ FATURAR":
            return 'background-color: #ffeb9c; color: #9c5700; font-weight: bold;' # Amarelo
        elif val == "NÃO FATURADO" or val == "BLOQUEADO":
            return 'background-color: #ffc7ce; color: #9c0006; font-weight: bold;' # Vermelho
        return ''

    st.subheader("📋 Painel de Faturamento (Lotes vs Cubagem)")
    if not df_fat_final.empty:
        st.dataframe(df_fat_final.style.applymap(colorir_faturamento, subset=['Status 555', 'Status 551']), use_container_width=True)
    else:
        st.info("Nenhum pedido da tabela de Lotes corresponde à Cubagem de hoje.")

    # Mostrar os Lotes Pendentes (Amarelos)
    if lotes_sobrando_amarelo:
        st.subheader("⚠️ Lotes Pendentes (Fora da Cubagem)")
        st.write("Estes lotes estão ativos, mas a filial não está na Cubagem de hoje. Eles continuarão guardados na base de dados para os próximos dias.")
        df_amarelos = pd.DataFrame(lotes_sobrando_amarelo)
        st.dataframe(df_amarelos.style.set_properties(**{'background-color': '#ffeb9c', 'color': 'black'}), use_container_width=True)

    # ==========================================
    # FUNÇÃO DE FINALIZAR E GUARDAR NO HISTÓRICO
    # ==========================================
    st.divider()
    col1, col2 = st.columns([1, 3])
    with col1:
        if st.button("🚀 Finalizar Faturamentos e Limpar", type="primary", use_container_width=True):
            if not df_fat_final.empty:
                # 1. Isolar quem está 100% faturado (555 e 551 concluídos)
                finalizados = df_fat_final[
                    (df_fat_final['Status 555'].str.contains("✅")) & 
                    (df_fat_final['Status 551'].str.contains("✅"))
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
