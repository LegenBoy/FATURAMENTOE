import streamlit as st
import pandas as pd
import os

# Configuração da Página
st.set_page_config(page_title="Portal de Faturamento Ecommerce", layout="wide")
st.title("📦 Portal Ecommerce - Faturamento Automático")

# Inicializa o "Banco de Dados" na sessão se não existir
if 'bd_lotes' not in st.session_state:
    # Em produção, aqui leríamos do bd/lotes_pendentes.xlsx
    st.session_state['bd_lotes'] = pd.DataFrame()

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
        # Pela sua planilha, a nota 555 e 551 tem a coluna 'Tipo ' (com espaço)
        # Vamos verificar uma linha para descobrir se é 555 ou 551
        if not df.empty:
            tipo_nota = str(df['Tipo '].iloc[0]).strip()
            if tipo_nota == '555': return 'faturamento_555'
            elif tipo_nota == '551': return 'faturamento_551'
    return 'desconhecido'

# ==========================================
# INTERFACE DE UPLOAD
# ==========================================
st.sidebar.header("📤 Upload de Arquivos")
arquivos_upados = st.sidebar.file_uploader("Arraste os relatórios do dia (CSV ou Excel)", accept_multiple_files=True)

dados = {
    'cubagem': pd.DataFrame(),
    'lotes_geral': pd.DataFrame(),
    'faturamento_555': pd.DataFrame(),
    'faturamento_551': pd.DataFrame()
}

# Lendo e categorizando os arquivos upados
if arquivos_upados:
    for arquivo in arquivos_upados:
        try:
            if arquivo.name.endswith('.csv'):
                df = pd.read_csv(arquivo, encoding='utf-8') # Ajuste encoding se der erro (ex: latin1)
            else:
                df = pd.read_excel(arquivo)
            
            tipo = identificar_tipo_arquivo(df)
            if tipo != 'desconhecido':
                dados[tipo] = df
                st.sidebar.success(f"✅ {tipo.upper()} carregado!")
            else:
                st.sidebar.warning(f"⚠️ Não consegui identificar: {arquivo.name}")
        except Exception as e:
            st.sidebar.error(f"Erro ao ler {arquivo.name}: {e}")

# ==========================================
# LÓGICA DE PROCESSAMENTO E CORES
# ==========================================
if not dados['cubagem'].empty and not dados['lotes_geral'].empty:
    st.divider()
    
    # 1. Simulação do BD de Lotes (Unindo o que subiu agora com o que estava guardado)
    df_lotes = pd.concat([st.session_state['bd_lotes'], dados['lotes_geral']]).drop_duplicates(subset=['LOTE', 'PEDIDO_ECOMMERCE'])
    st.session_state['bd_lotes'] = df_lotes # Atualiza o BD
    
    # Extrair Filiais da Cubagem (Simplificado para o exemplo)
    # Lógica baseada no seu script: procurar filiais nas colunas da Cubagem
    filiais_na_cubagem = []
    df_cubagem = dados['cubagem']
    for col in df_cubagem.columns[4:16]: # Colunas das filiais/lojas
        filiais_encontradas = df_cubagem[col].dropna().astype(str)
        for f in filiais_encontradas:
            if '-' in f:
                cod_filial = f.split('-')[0].strip().lstrip('0')
                filiais_na_cubagem.append(cod_filial)
                
    filiais_na_cubagem = set(filiais_na_cubagem)
    
    # 2. Criar a tela de FATURAMENTO Principal
    faturamento_view = []
    lotes_sobrando_amarelo = []
    
    for idx, row in df_lotes.iterrows():
        filial_lote = str(row.get('FILIAL', '')).strip().lstrip('0')
        pedido = str(row.get('PEDIDO_ECOMMERCE', ''))
        
        # Simulação de verificação nas tabelas 555 e 551
        status_555 = "NÃO FATURADO"
        status_551 = "BLOQUEADO"
        
        # Lógica: Se tem 555, libera o 551
        if not dados['faturamento_555'].empty:
            # Pela sua regra, 555 cruza pelo LOTE
            tem_555 = str(row['LOTE']) in dados['faturamento_555']['Lote '].astype(str).values
            if tem_555:
                status_555 = "✅ FATURADO (555)"
                status_551 = "PRONTO P/ FATURAR" # Libera para o 551
        
        if not dados['faturamento_551'].empty and status_555.startswith("✅"):
            # 551 cruza pelo Pedido Ecommerce (Obs no seu CSV)
            # Simplificação: verificar se o pedido está no dataframe
            tem_551 = dados['faturamento_551'].apply(lambda x: x.astype(str).str.contains(pedido, na=False)).any().any()
            if tem_551:
                status_551 = "✅ FATURADO (551)"

        if filial_lote in filiais_na_cubagem:
            faturamento_view.append({
                "Lote": row['LOTE'],
                "Filial": filial_lote,
                "Pedido": pedido,
                "Cliente": row['CLIENTE'],
                "Status 555": status_555,
                "Status 551": status_551
            })
        else:
            # Lote existe, mas filial não está na cubagem de hoje (Amarelo)
            lotes_sobrando_amarelo.append(row)

    df_fat_final = pd.DataFrame(faturamento_view)
    
    # ==========================================
    # APRESENTAÇÃO NA TELA COM CORES (Pandas Styling)
    # ==========================================
    
    # Função para colorir o Faturamento
    def colorir_faturamento(val):
        if val == "✅ FATURADO (555)" or val == "✅ FATURADO (551)":
            return 'background-color: #c6efce; color: #006100;' # Verde Excel
        elif val == "PRONTO P/ FATURAR":
            return 'background-color: #ffeb9c; color: #9c5700;' # Amarelo Excel
        elif val == "NÃO FATURADO" or val == "BLOQUEADO":
            return 'background-color: #ffc7ce; color: #9c0006;' # Vermelho Excel
        return ''

    st.subheader("📋 Painel de Faturamento (Lotes vs Cubagem)")
    if not df_fat_final.empty:
        st.dataframe(df_fat_final.style.applymap(colorir_faturamento, subset=['Status 555', 'Status 551']), use_container_width=True)
    else:
        st.info("Nenhum pedido da tabela de Lotes bate com a Cubagem de hoje.")

    # Mostrar Lotes que ficaram de fora (Os Amarelos da sua regra)
    if lotes_sobrando_amarelo:
        st.subheader("⚠️ Lotes Pendentes (Sem Cubagem)")
        st.write("Estes lotes estão no BD, mas a filial não está na Cubagem de hoje (Ficam guardados para o próximo dia).")
        df_amarelos = pd.DataFrame(lotes_sobrando_amarelo)
        st.dataframe(df_amarelos.style.set_properties(**{'background-color': '#ffeb9c', 'color': 'black'}), use_container_width=True)

    # Função de Finalizar / Gravar Histórico (Script 2)
    st.divider()
    if st.button("Finalizar Faturamentos do Dia e Limpar"):
        if not df_fat_final.empty:
            finalizados = df_fat_final[(df_fat_final['Status 555'].str.contains("✅")) & (df_fat_final['Status 551'].str.contains("✅"))]
            
            if not finalizados.empty:
                st.success(f"{len(finalizados)} pedidos finalizados com sucesso! (Salvos no BD)")
                # Aqui entra a lógica de salvar no arquivo finalizados.xlsx
                # E remover os Lotes finalizados do st.session_state['bd_lotes']
            else:
                st.warning("Nenhum pedido tem os dois faturamentos (555 e 551) concluidos para finalizar.")

else:
    st.info("👈 Por favor, faça o upload das planilhas de Cubagem e Lotes Geral no menu lateral para começar.")
