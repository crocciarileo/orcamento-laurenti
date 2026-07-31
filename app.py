import streamlit as st
import sqlite3
import pandas as pd
from datetime import datetime, timedelta

# Configuração da página
st.set_page_config(page_title="Orçamentos - Laurenti Móveis", page_icon="📝", layout="wide")

# -----------------------------------------------------------------------------
# 1. INICIALIZAÇÃO DO BANCO DE DADOS (SQLite)
# -----------------------------------------------------------------------------
def get_connection():
    conn = sqlite3.connect('orcamentos.db', check_same_thread=False)
    return conn

def init_db():
    conn = get_connection()
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS orcamentos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            cliente TEXT NOT NULL,
            contato TEXT,
            tipo_contato TEXT,
            telefone TEXT,
            email TEXT,
            data TEXT NOT NULL,
            dias_validade INTEGER NOT NULL,
            validade TEXT NOT NULL,
            prazo_entrega TEXT,
            condicoes_pagamento TEXT,
            observacoes TEXT,
            total_liquido REAL DEFAULT 0,
            total_com_opcionais REAL DEFAULT 0
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS ambientes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            orcamento_id INTEGER,
            ordem INTEGER,
            nome_ambiente TEXT NOT NULL,
            especificacoes TEXT,
            total_ambiente REAL DEFAULT 0,
            FOREIGN KEY(orcamento_id) REFERENCES orcamentos(id) ON DELETE CASCADE
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS itens (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ambiente_id INTEGER,
            ordem INTEGER,
            descricao TEXT NOT NULL,
            valor REAL DEFAULT 0,
            eh_opcional INTEGER DEFAULT 0,
            FOREIGN KEY(ambiente_id) REFERENCES ambientes(id) ON DELETE CASCADE
        )
    """)
    conn.commit()

init_db()

# -----------------------------------------------------------------------------
# 2. GERENCIAMENTO DE ESTADO (SESSION STATE)
# -----------------------------------------------------------------------------
if 'ambientes' not in st.session_state:
    st.session_state.ambientes = []  # Estrutura na memória antes de salvar

def recalcular_totais():
    total_liquido = 0.0
    total_opcionais = 0.0
    
    for amb in st.session_state.ambientes:
        total_amb = 0.0
        for item in amb['itens']:
            val = float(item['valor'])
            if not item['eh_opcional']:
                total_amb += val
            else:
                total_opcionais += val
        amb['total_ambiente'] = total_amb
        total_liquido += total_amb
        
    return total_liquido, (total_liquido + total_opcionais)

# -----------------------------------------------------------------------------
# 3. INTERFACE PRINCIPAL
# -----------------------------------------------------------------------------
st.title("🏭 Laurenti Móveis — Sistema de Orçamentos")

menu = st.sidebar.radio("Navegação", ["➕ Novo Orçamento", "📋 Orçamentos Salvos", "⚙️ Configurações"])

if menu == "➕ Novo Orçamento":
    st.subheader("Novo Orçamento")
    
    # Bloco Dados do Cliente
    with st.expander("👤 Dados do Cliente & Condições", expanded=True):
        col1, col2, col3 = st.columns(3)
        with col1:
            cliente = st.text_input("Cliente *", placeholder="Nome da pessoa ou empresa")
            contato = st.text_input("Contato", placeholder="Nome do responsável")
            tipo_contato = st.selectbox("Tipo de Contato", ["Arquitetura", "Cliente Final", "Indicado", "Outros"])
        
        with col2:
            telefone = st.text_input("Telefone", placeholder="(00) 00000-0000")
            email = st.text_input("E-mail", placeholder="cliente@email.com")
            data_atual = st.date_input("Data do Orçamento", value=datetime.now().date())
            
        with col3:
            dias_validade = st.radio("Validade em Dias", [7, 10, 15, 30], index=2, horizontal=True)
            data_validade = data_atual + timedelta(days=dias_validade)
            st.info(f"📅 **Validade calculada:** {data_validade.strftime('%d/%m/%Y')}")
            
        col_cond1, col_cond2 = st.columns(2)
        with col_cond1:
            prazo_entrega = st.text_input("Prazo de Entrega", value="45 dias úteis após aprovação do projeto")
        with col_cond2:
            condicoes_pagamento = st.text_input("Condições de Pagamento", value="50% na entrada + 50% na entrega")
            
        observacoes = st.text_area("Observações Gerais", placeholder="Detalhes adicionais do projeto, frete, montagem...")

    st.markdown("---")
    st.subheader("🛋️ Ambientes e Itens")

    # Form para Adicionar Ambiente
    with st.form("form_novo_ambiente", clear_on_submit=True):
        col_amb1, col_amb2 = st.columns([2, 3])
        with col_amb1:
            nome_amb = st.text_input("Nome do Ambiente *", placeholder="Ex: Cozinha, Suíte Master")
        with col_amb2:
            espec_amb = st.text_input("Especificações Técnicas", placeholder="Ex: MDF Laca J155, frentes molduradas, LED")
        
        btn_add_amb = st.form_submit_button("➕ Adicionar Ambiente")
        if btn_add_amb:
            if nome_amb:
                st.session_state.ambientes.append({
                    'nome': nome_amb,
                    'especificacoes': espec_amb,
                    'total_ambiente': 0.0,
                    'itens': []
                })
                st.rerun()
            else:
                st.error("Digite o nome do ambiente.")

    # Exibição dos Ambientes e seus Subitens
    if st.session_state.ambientes:
        for idx_amb, amb in enumerate(st.session_state.ambientes):
            st.markdown(f"#### **Ambiente {idx_amb + 1}: {amb['nome']}**")
            if amb['especificacoes']:
                st.caption(f"🔧 *{amb['especificacoes']}*")
            
            # Adicionar Item dentro deste Ambiente
            with st.expander(f"➕ Adicionar Item ao Ambiente {idx_amb + 1} ({amb['nome']})", expanded=True):
                col_i1, col_i2, col_i3 = st.columns([3, 1.5, 1])
                with col_i1:
                    desc_item = st.text_input(f"Descrição do Item", key=f"desc_{idx_amb}")
                with col_i2:
                    valor_item = st.number_input(f"Valor (R$)", min_value=0.0, step=100.0, format="%.2f", key=f"val_{idx_amb}")
                with col_i3:
                    eh_opc = st.checkbox("Item Opcional?", key=f"opc_{idx_amb}")
                
                if st.button("Adicionar Item", key=f"btn_add_item_{idx_amb}"):
                    if desc_item and valor_item > 0:
                        amb['itens'].append({
                            'descricao': desc_item,
                            'valor': valor_item,
                            'eh_opcional': eh_opc
                        })
                        st.rerun()
                    else:
                        st.warning("Preencha a descrição e um valor válido.")

            # Tabela de itens do ambiente
            if amb['itens']:
                df_itens = pd.DataFrame(amb['itens'])
                df_itens['eh_opcional_str'] = df_itens['eh_opcional'].apply(lambda x: "Sim (Opcional)" if x else "Não")
                df_itens['valor_str'] = df_itens['valor'].apply(lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                
                st.dataframe(
                    df_itens[['descricao', 'valor_str', 'eh_opcional_str']].rename(
                        columns={'descricao': 'Descrição', 'valor_str': 'Valor', 'eh_opcional_str': 'Opcional?'}
                    ),
                    use_container_width=True
                )
            
            st.markdown(f"**Total do Ambiente {idx_amb + 1}: R$ {amb['total_ambiente']:,.2f}**".replace(",", "X").replace(".", ",").replace("X", "."))
            st.markdown("---")

    # Resumo Final e Salvamento
    tot_liquido, tot_com_opcionais = recalcular_totais()
    
    st.markdown("### 📊 Totais do Orçamento")
    c_tot1, c_tot2 = st.columns(2)
    c_tot1.metric("Total Líquido (Base)", f"R$ {tot_liquido:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    c_tot2.metric("Total c/ Opcionais", f"R$ {tot_com_opcionais:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

    if st.button("💾 Salvar Orçamento no Banco de Dados", type="primary", use_container_width=True):
        if not cliente:
            st.error("Por favor, preencha o nome do Cliente.")
        elif not st.session_state.ambientes:
            st.error("Adicione pelo menos um ambiente antes de salvar.")
        else:
            conn = get_connection()
            c = conn.cursor()
            
            # Insere Orçamento
            c.execute("""
                INSERT INTO orcamentos (cliente, contato, tipo_contato, telefone, email, data, dias_validade, validade, prazo_entrega, condicoes_pagamento, observacoes, total_liquido, total_com_opcionais)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (cliente, contato, tipo_contato, telefone, email, str(data_atual), dias_validade, str(data_validade), prazo_entrega, condicoes_pagamento, observacoes, tot_liquido, tot_com_opcionais))
            
            orc_id = c.lastrowid
            
            # Insere Ambientes e Itens
            for idx_a, amb in enumerate(st.session_state.ambientes):
                c.execute("""
                    INSERT INTO ambientes (orcamento_id, ordem, nome_ambiente, especificacoes, total_ambiente)
                    VALUES (?, ?, ?, ?, ?)
                """, (orc_id, idx_a + 1, amb['nome'], amb['especificacoes'], amb['total_ambiente']))
                
                amb_id = c.lastrowid
                
                for idx_i, item in enumerate(amb['itens']):
                    c.execute("""
                        INSERT INTO itens (ambiente_id, ordem, descricao, valor, eh_opcional)
                        VALUES (?, ?, ?, ?, ?)
                    """, (amb_id, idx_i + 1, item['descricao'], item['valor'], 1 if item['eh_opcional'] else 0))
            
            conn.commit()
            st.success(f"✅ Orçamento #{orc_id} salvo com sucesso no banco de dados SQLite!")
            st.session_state.ambientes = []

elif menu == "📋 Orçamentos Salvos":
    st.subheader("Orçamentos Salvos no Banco")
    conn = get_connection()
    df_orc = pd.read_sql_query("SELECT id, cliente, data, validade, total_liquido, total_com_opcionais FROM orcamentos ORDER BY id DESC", conn)
    
    if not df_orc.empty:
        st.dataframe(df_orc, use_container_width=True)
    else:
        st.info("Nenhum orçamento salvo ainda.")

elif menu == "⚙️ Configurações":
    st.subheader("Configurações da Empresa")
    st.write("**Laurenti Móveis**")
    st.write("CNPJ: 00.000.000/0001-00")
    st.write("Rua Henrique Villa, 59 - Ariranha/SP — Tel: (17) 3576-1464")
