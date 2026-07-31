import streamlit as st
import sqlite3
import pandas as pd
from datetime import datetime, timedelta
import io
import os
from PIL import Image

# ReportLab para PDF institucional Laurenti Móveis
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas

# Configuração da página Streamlit
st.set_page_config(page_title="Orçamentos - Laurenti Móveis", page_icon="📝", layout="wide")

# -----------------------------------------------------------------------------
# 1. BANCO DE DADOS E MIGRAÇÕES (SQLite)
# -----------------------------------------------------------------------------
def get_connection():
    return sqlite3.connect('orcamentos.db', check_same_thread=False)

def init_db():
    conn = get_connection()
    c = conn.cursor()
    
    # Configurações da Empresa
    c.execute("""
        CREATE TABLE IF NOT EXISTS config_empresa (
            id INTEGER PRIMARY KEY,
            nome_empresa TEXT,
            cnpj TEXT,
            ie TEXT,
            endereco TEXT,
            telefone TEXT,
            email TEXT,
            logo_path TEXT
        )
    """)
    
    c.execute("SELECT COUNT(*) FROM config_empresa")
    if c.fetchone()[0] == 0:
        c.execute("""
            INSERT INTO config_empresa (id, nome_empresa, cnpj, ie, endereco, telefone, email, logo_path)
            VALUES (1, 'Fábrica de Móveis Laurenti Ltda', '44.331.015/0001-08', '186000158114', 
                    'Rua Henrique Villa, 59- Jardim Maria Emília. CEP: 15960000, ARIRANHA-SP', 
                    '(17) 3576-1464', 'contato@laurentimoveis.com.br', '')
        """)

    # Consultores
    c.execute("""
        CREATE TABLE IF NOT EXISTS consultores (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL
        )
    """)
    
    c.execute("SELECT COUNT(*) FROM consultores")
    if c.fetchone()[0] == 0:
        c.executemany("INSERT INTO consultores (nome) VALUES (?)", [("KATIA LUCIA LOURENCO",), ("Sem Consultor",)])

    # Orçamentos
    c.execute("""
        CREATE TABLE IF NOT EXISTS orcamentos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            proposta_num INTEGER,
            cliente TEXT NOT NULL,
            contato TEXT,
            tipo_contato TEXT,
            telefone TEXT,
            email TEXT,
            consultor TEXT,
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

    # Ambientes
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

    # Itens / Subitens
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

# Funções Auxiliares protegidas
def get_config():
    conn = get_connection()
    df = pd.read_sql_query("SELECT * FROM config_empresa WHERE id = 1", conn)
    if not df.empty:
        d = df.iloc[0].to_dict()
        d.setdefault('nome_empresa', 'Fábrica de Móveis Laurenti Ltda')
        d.setdefault('cnpj', '44.331.015/0001-08')
        d.setdefault('ie', '186000158114')
        d.setdefault('endereco', 'Rua Henrique Villa, 59- Jardim Maria Emília. CEP: 15960000, ARIRANHA-SP')
        d.setdefault('telefone', '(17) 3576-1464')
        d.setdefault('email', 'contato@laurentimoveis.com.br')
        d.setdefault('logo_path', '')
        return d
    return {'nome_empresa': 'Fábrica de Móveis Laurenti Ltda', 'cnpj': '44.331.015/0001-08', 'ie': '186000158114', 'endereco': '', 'telefone': '', 'email': '', 'logo_path': ''}

def get_consultores():
    conn = get_connection()
    return pd.read_sql_query("SELECT * FROM consultores ORDER BY nome ASC", conn)

def get_proxima_proposta():
    conn = get_connection()
    c = conn.cursor()
    c.execute("SELECT proposta_num FROM orcamentos ORDER BY id DESC LIMIT 1")
    row = c.fetchone()
    if row and row[0] is not None:
        try:
            val = int(row[0])
            return val + 1
        except ValueError:
            return 2358
    return 2358

def get_ultimas_condicoes():
    conn = get_connection()
    c = conn.cursor()
    c.execute("SELECT consultor, prazo_entrega, condicoes_pagamento, observacoes FROM orcamentos ORDER BY id DESC LIMIT 1")
    row = c.fetchone()
    if row:
        return {'consultor': row[0] or 'Sem Consultor', 'prazo_entrega': row[1] or '120 dias após medições finais.', 'condicoes_pagamento': row[2] or '8 PARCELAS', 'observacoes': row[3] or 'Especificações Gerais: MDF externo cores a definir 18mm / MDF interno Branco TX 18mm'}
    return {'consultor': 'Sem Consultor', 'prazo_entrega': '120 dias após medições finais.', 'condicoes_pagamento': '8 PARCELAS', 'observacoes': 'Especificações Gerais: MDF externo cores a definir 18mm / MDF interno Branco TX 18mm'}

# Canvas Dinâmico com Numeração de Páginas (Página X de Y)
class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_number(num_pages)
            super().showPage()
        super().save()

    def draw_page_number(self, page_count):
        config = get_config()
        self.saveState()
        self.setFont("Helvetica", 7)
        self.setFillColor(colors.HexColor("#333333"))
        
        self.setStrokeColor(colors.HexColor("#BDC3C7"))
        self.setLineWidth(0.5)
        self.line(1.2*cm, 1.8*cm, A4[0] - 1.2*cm, 1.8*cm)
        
        line1 = f"{config.get('nome_empresa', '')} CNPJ: {config.get('cnpj', '')} IE: {config.get('ie', '')}"
        line2 = f"{config.get('endereco', '')} / Fone: {config.get('telefone', '')}"
        
        self.drawString(1.2*cm, 1.3*cm, line1)
        self.drawString(1.2*cm, 0.9*cm, line2)
        
        data_hora_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        page_str = f"Página {self._pageNumber} de {page_count}"
        
        self.drawRightString(A4[0] - 1.2*cm, 1.3*cm, page_str)
        self.drawRightString(A4[0] - 1.2*cm, 0.9*cm, data_hora_str)
        self.restoreState()

# -----------------------------------------------------------------------------
# 2. GERADOR DE PDF
# -----------------------------------------------------------------------------
def gerar_pdf_orcamento(cliente_info, ambientes_list):
    config = get_config()
    buffer = io.BytesIO()
    
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=1.2*cm,
        leftMargin=1.2*cm,
        topMargin=1.2*cm,
        bottomMargin=2.2*cm
    )
    
    story = []
    styles = getSampleStyleSheet()
    
    body_style = ParagraphStyle('BodyStyle', parent=styles['Normal'], fontSize=8, leading=11, textColor=colors.HexColor('#222222'))
    body_bold = ParagraphStyle('BodyBold', parent=body_style, fontName='Helvetica-Bold')
    right_bold = ParagraphStyle('RightBold', parent=body_bold, alignment=2)
    header_title = ParagraphStyle('HeaderTitle', parent=body_bold, fontSize=9, textColor=colors.HexColor('#1E293B'))
    amb_title_style = ParagraphStyle('AmbTitleStyle', parent=body_bold, fontSize=9, textColor=colors.HexColor('#0F172A'))

    col_logo = Paragraph("<b>LAURENTI MÓVEIS</b>", body_bold)
    logo_p = config.get('logo_path', '')
    if logo_p and os.path.exists(logo_p):
        try:
            col_logo = RLImage(logo_p, width=3.8*cm, height=1.4*cm)
        except Exception:
            pass

    cli_text = f"""
    <b>Cliente:</b> {cliente_info['cliente']}<br/>
    <b>Contato:</b> {cliente_info['contato']}<br/>
    <b>Consultor:</b> {cliente_info.get('consultor', 'N/A')}<br/>
    <b>Tipo de Contato:</b> {cliente_info['tipo_contato']}<br/>
    <b>Telefone:</b> {cliente_info['telefone']} &nbsp;&nbsp;&nbsp; <b>E-mail:</b> {cliente_info['email']}
    """
    
    prop_num = cliente_info.get('proposta_num', '2358')
    prop_text = f"""
    <b>Proposta:</b> {prop_num}<br/><br/>
    <b>Data:</b> {cliente_info['data']}<br/><br/>
    <b>Validade:</b> {cliente_info['validade']}
    """
    
    t_head = Table([
        [col_logo, Paragraph(cli_text, body_style), Paragraph(prop_text, body_style)]
    ], colWidths=[4*cm, 10.2*cm, 4.4*cm])
    
    t_head.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 5),
        ('BACKGROUND', (2,0), (2,0), colors.HexColor('#F8FAFC')),
    ]))
    story.append(t_head)
    story.append(Spacer(1, 10))

    table_data = [
        [Paragraph("<b>Item</b>", header_title), Paragraph("<b>Valor (R$)</b>", ParagraphStyle('HRight', parent=header_title, alignment=2))]
    ]
    
    tot_liquido = 0.0
    tot_opcionais = 0.0
    row_styles = []

    for idx_a, amb in enumerate(ambientes_list):
        ordem_amb = idx_a + 1
        nome_amb_upper = amb['nome'].upper()
        
        itens_normais = [i for i in amb['itens'] if not i['eh_opcional']]
        itens_opcionais = [i for i in amb['itens'] if i['eh_opcional']]
        
        tot_amb = sum(float(i['valor']) for i in itens_normais)
        tot_liquido += tot_amb
        tot_amb_str = f"R$ {tot_amb:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

        current_row = len(table_data)
        amb_header_text = f"<b>{ordem_amb}- {nome_amb_upper}</b>"
        table_data.append([Paragraph(amb_header_text, amb_title_style), Paragraph(f"<b>{tot_amb_str}</b>", right_bold)])
        row_styles.append(('BACKGROUND', (0, current_row), (-1, current_row), colors.HexColor('#F1F5F9')))
        
        desc_amb_text = f"<i>{amb.get('especificacoes', '')}</i>" if amb.get('especificacoes') else ""
        
        for item in itens_normais:
            val_f = f"{float(item['valor']):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            desc_amb_text += f"<br/>- {item['descricao']} - R$ {val_f}"
            
        if itens_opcionais:
            desc_amb_text += "<br/><br/><b>• Acessórios opcionais a serem acrescidos:</b>"
            for item in itens_opcionais:
                val_opc = float(item['valor'])
                tot_opcionais += val_opc
                val_f = f"{val_opc:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                desc_amb_text += f"<br/>- {item['descricao']}, acréscimo - R$ {val_f}"

        table_data.append([Paragraph(desc_amb_text, body_style), Paragraph("", body_style)])

    t_itens = Table(table_data, colWidths=[14.6*cm, 4*cm])
    
    t_style = [
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#E2E8F0')),
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#94A3B8')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
        ('PADDING', (0,0), (-1,-1), 6),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
    ] + row_styles
    
    t_itens.setStyle(TableStyle(t_style))
    story.append(t_itens)
    story.append(Spacer(1, 10))

    tot_com_opc = tot_liquido + tot_opcionais
    tot_liquido_str = f"R$ {tot_liquido:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    tot_com_opc_str = f"R$ {tot_com_opc:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    cond_text = f"""
    <b>Prazo de Entrega:</b> {cliente_info['prazo_entrega']}<br/>
    <b>Condição de Pagamento:</b> {cliente_info['condicoes_pagamento']}<br/>
    <b>Observações:</b> {cliente_info['observacoes']}
    """
    
    totais_box_text = f"""
    <font size=9 color='#64748B'><b>Total Líquido:</b></font><br/>
    <font size=11 color='#0F172A'><b>{tot_liquido_str}</b></font><br/><br/>
    <font size=8 color='#64748B'>Total c/ Opcionais:</font><br/>
    <font size=10 color='#D97706'><b>{tot_com_opc_str}</b></font>
    """
    
    t_cond = Table([
        [Paragraph(cond_text, body_style), Paragraph(totais_box_text, right_bold)]
    ], colWidths=[13.6*cm, 5*cm])
    
    t_cond.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
        ('BACKGROUND', (0,0), (0,0), colors.HexColor('#F8FAFC')),
        ('BACKGROUND', (1,0), (1,0), colors.HexColor('#F1F5F9')),
        ('PADDING', (0,0), (-1,-1), 8),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    story.append(t_cond)

    doc.build(story, canvasmaker=NumberedCanvas)
    buffer.seek(0)
    return buffer

# -----------------------------------------------------------------------------
# 3. INTERFACE STREAMLIT
# -----------------------------------------------------------------------------
if 'ambientes' not in st.session_state:
    st.session_state.ambientes = []

if 'edit_index' not in st.session_state:
    st.session_state.edit_index = None

if 'confirm_del' not in st.session_state:
    st.session_state.confirm_del = None

if 'active_tab' not in st.session_state:
    st.session_state.active_tab = "➕ Novo / Editar Orçamento"

def recalcular_totais():
    total_liquido = 0.0
    total_opcionais = 0.0
    for amb in st.session_state.ambientes:
        tot_amb = 0.0
        for item in amb['itens']:
            val = float(item['valor'])
            if not item['eh_opcional']:
                tot_amb += val
            else:
                total_opcionais += val
        amb['total_ambiente'] = tot_amb
        total_liquido += tot_amb
    return total_liquido, (total_liquido + total_opcionais)

st.title("🏭 Laurenti Móveis — Gestão de Orçamentos")

menu = st.sidebar.radio("Navegação", ["➕ Novo / Editar Orçamento", "📋 Orçamentos Salvos", "⚙️ Configurações"], key="main_menu")

# Sync caso botão de edição peça para mudar a aba
if 'navigate_to' in st.session_state and st.session_state.navigate_to:
    menu = st.session_state.navigate_to
    st.session_state.navigate_to = None

# --- ABA 1: NOVO / EDITAR ORÇAMENTO ---
if menu == "➕ Novo / Editar Orçamento":
    st.subheader("Formulário de Orçamento")
    
    ultimas_cond = get_ultimas_condicoes()
    df_cons = get_consultores()
    consultores_opts = df_cons['nome'].tolist() if not df_cons.empty else ["Sem Consultor"]

    if st.button("✨ Criar Novo Orçamento Limpo"):
        st.session_state.ambientes = []
        st.session_state.edit_index = None
        st.session_state.cli_nome = ""
        st.session_state.cli_contato = ""
        st.session_state.cli_tel = ""
        st.session_state.cli_email = ""
        st.session_state.cli_tipo = "Residencial"
        st.session_state.cli_consultor = ultimas_cond['consultor']
        st.session_state.cli_prazo = ultimas_cond['prazo_entrega']
        st.session_state.cli_cond = ultimas_cond['condicoes_pagamento']
        st.session_state.cli_obs = ultimas_cond['observacoes']
        st.session_state.cli_prop = get_proxima_proposta()
        st.rerun()

    prop_num_atual = st.session_state.get('cli_prop', get_proxima_proposta())

    with st.expander("👤 Dados do Cliente e Proposta", expanded=True):
        col1, col2, col3 = st.columns(3)
        with col1:
            cliente = st.text_input("Cliente *", value=st.session_state.get('cli_nome', ''), placeholder="Nome da pessoa ou empresa")
            contato = st.text_input("Contato", value=st.session_state.get('cli_contato', ''), placeholder="Nome do responsável")
            
            tipo_opts = ["Residencial", "Comercial", "Arquitetura", "Outros"]
            idx_tipo = tipo_opts.index(st.session_state.get('cli_tipo', 'Residencial')) if st.session_state.get('cli_tipo') in tipo_opts else 0
            tipo_contato = st.selectbox("Tipo de Contato", tipo_opts, index=idx_tipo)
            
            cons_val = st.session_state.get('cli_consultor', ultimas_cond['consultor'])
            idx_cons_padrao = consultores_opts.index(cons_val) if cons_val in consultores_opts else 0
            consultor = st.selectbox("Consultor *", consultores_opts, index=idx_cons_padrao)
        
        with col2:
            telefone = st.text_input("Telefone", value=st.session_state.get('cli_tel', ''), placeholder="(00) 00000-0000")
            email = st.text_input("E-mail", value=st.session_state.get('cli_email', ''), placeholder="cliente@email.com")
            data_atual = st.date_input("Data da Proposta", value=datetime.now().date())
            
        with col3:
            st.text_input("Proposta Nº (Automático)", value=str(prop_num_atual), disabled=True)
            dias_validade = st.radio("Validade em Dias", [7, 10, 15, 30], index=3, horizontal=True)
            data_validade = data_atual + timedelta(days=dias_validade)
            st.info(f"📅 **Validade:** {data_validade.strftime('%d/%m/%Y')}")
            
        col_cond1, col_cond2 = st.columns(2)
        with col_cond1:
            prazo_entrega = st.text_input("Prazo de Entrega", value=st.session_state.get('cli_prazo', ultimas_cond['prazo_entrega']))
        with col_cond2:
            condicoes_pagamento = st.text_input("Condições de Pagamento", value=st.session_state.get('cli_cond', ultimas_cond['condicoes_pagamento']))
            
        observacoes = st.text_area("Observações Gerais", value=st.session_state.get('cli_obs', ultimas_cond['observacoes']))

    st.markdown("---")
    st.subheader("🛋️ Ambientes e Subitens")

    with st.form("form_novo_ambiente", clear_on_submit=True):
        col_amb1, col_amb2 = st.columns([2, 3])
        with col_amb1:
            nome_amb = st.text_input("Nome do Ambiente *", placeholder="Ex: COZINHA, GOURMET, SALA")
        with col_amb2:
            espec_amb = st.text_input("Especificações do Ambiente", placeholder="Ex: MDF Nogueira Ambar / MDF Nude Vel")
        
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

    if st.session_state.ambientes:
        for idx_amb, amb in enumerate(st.session_state.ambientes):
            ordem_amb = idx_amb + 1
            amb['itens'] = sorted(amb['itens'], key=lambda x: x['eh_opcional'])

            with st.expander(f"🛋️ **Item {ordem_amb}: {amb['nome'].upper()}** — Total: R$ {amb['total_ambiente']:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), expanded=True):
                
                col_e1, col_e2, col_e3 = st.columns([2, 3, 1])
                with col_e1:
                    amb['nome'] = st.text_input(f"Nome #{ordem_amb}", value=amb['nome'], key=f"edit_nome_{idx_amb}")
                with col_e2:
                    amb['especificacoes'] = st.text_input(f"Especificações #{ordem_amb}", value=amb['especificacoes'], key=f"edit_espec_{idx_amb}")
                with col_e3:
                    st.write(" ")
                    if st.button("🗑️ Remover Ambiente", key=f"del_amb_btn_{idx_amb}"):
                        st.session_state.confirm_del = f"amb_{idx_amb}"

                if st.session_state.confirm_del == f"amb_{idx_amb}":
                    st.warning("⚠️ Confirma a exclusão deste ambiente e todos seus itens?")
                    c_del1, c_del2 = st.columns(2)
                    if c_del1.button("✅ Confirmar Exclusão", key=f"conf_del_amb_{idx_amb}"):
                        st.session_state.ambientes.pop(idx_amb)
                        st.session_state.confirm_del = None
                        st.rerun()
                    if c_del2.button("❌ Cancelar", key=f"canc_del_amb_{idx_amb}"):
                        st.session_state.confirm_del = None
                        st.rerun()

                st.markdown(f"##### Subitens do Ambiente {ordem_amb}")
                
                with st.form(f"form_add_subitem_{idx_amb}", clear_on_submit=True):
                    col_i1, col_i2, col_i3, col_i4 = st.columns([3, 1.5, 1, 1])
                    with col_i1:
                        desc_sub = st.text_input("Descrição do Subitem", placeholder="Ex: Armário aéreo 3,67m, com 8 portas")
                    with col_i2:
                        val_sub = st.number_input("Valor (R$)", min_value=0.0, step=100.0, format="%.2f")
                    with col_i3:
                        st.write(" ")
                        opc_sub = st.checkbox("Opcional?")
                    with col_i4:
                        st.write(" ")
                        btn_sub = st.form_submit_button("➕ Incluir Subitem")
                        
                    if btn_sub:
                        if desc_sub and val_sub > 0:
                            amb['itens'].append({
                                'descricao': desc_sub,
                                'valor': val_sub,
                                'eh_opcional': opc_sub
                            })
                            st.rerun()

                if amb['itens']:
                    for idx_item, item in enumerate(amb['itens']):
                        tag_opc = " (OPCIONAL)" if item['eh_opcional'] else ""
                        col_it1, col_it2, col_it3, col_it4 = st.columns([3, 1.5, 1, 0.5])
                        with col_it1:
                            item['descricao'] = st.text_input(f"Subitem {ordem_amb}.{idx_item+1}{tag_opc}", value=item['descricao'], key=f"item_desc_{idx_amb}_{idx_item}")
                        with col_it2:
                            item['valor'] = st.number_input("Valor R$", value=float(item['valor']), step=50.0, format="%.2f", key=f"item_val_{idx_amb}_{idx_item}")
                        with col_it3:
                            st.write(" ")
                            item['eh_opcional'] = st.checkbox("Opcional?", value=bool(item['eh_opcional']), key=f"item_opc_{idx_amb}_{idx_item}")
                        with col_it4:
                            st.write(" ")
                            if st.button("❌", key=f"del_item_{idx_amb}_{idx_item}"):
                                amb['itens'].pop(idx_item)
                                st.rerun()

    tot_liquido, tot_com_opcionais = recalcular_totais()
    
    st.markdown("---")
    st.markdown("### 📊 Totais e Emissão")
    c_tot1, c_tot2 = st.columns(2)
    c_tot1.metric("Total Líquido (Sem Opcionais)", f"R$ {tot_liquido:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    c_tot2.metric("Total Com Opcionais", f"R$ {tot_com_opcionais:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

    col_btn1, col_btn2 = st.columns(2)
    
    with col_btn1:
        if st.button("💾 Salvar Orçamento no Banco", type="primary", use_container_width=True):
            if not cliente:
                st.error("Por favor, preencha o nome do Cliente.")
            elif not st.session_state.ambientes:
                st.error("Adicione pelo menos um ambiente antes de salvar.")
            else:
                conn = get_connection()
                c = conn.cursor()
                
                if st.session_state.edit_index:
                    orc_id = st.session_state.edit_index
                    c.execute("DELETE FROM ambientes WHERE orcamento_id = ?", (orc_id,))
                    c.execute("""
                        UPDATE orcamentos SET proposta_num=?, cliente=?, contato=?, tipo_contato=?, telefone=?, email=?, consultor=?, data=?, dias_validade=?, validade=?, prazo_entrega=?, condicoes_pagamento=?, observacoes=?, total_liquido=?, total_com_opcionais=?
                        WHERE id=?
                    """, (int(prop_num_atual), cliente, contato, tipo_contato, telefone, email, consultor, str(data_atual), dias_validade, str(data_validade.strftime('%d/%m/%Y')), prazo_entrega, condicoes_pagamento, observacoes, tot_liquido, tot_com_opcionais, orc_id))
                else:
                    c.execute("""
                        INSERT INTO orcamentos (proposta_num, cliente, contato, tipo_contato, telefone, email, consultor, data, dias_validade, validade, prazo_entrega, condicoes_pagamento, observacoes, total_liquido, total_com_opcionais)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (int(prop_num_atual), cliente, contato, tipo_contato, telefone, email, consultor, str(data_atual), dias_validade, str(data_validade.strftime('%d/%m/%Y')), prazo_entrega, condicoes_pagamento, observacoes, tot_liquido, tot_com_opcionais))
                    orc_id = c.lastrowid
                
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
                st.success(f"✅ Orçamento Proposta Nº {prop_num_atual} salvo com sucesso!")

    with col_btn2:
        if cliente and st.session_state.ambientes:
            cli_info = {
                'proposta_num': prop_num_atual,
                'cliente': cliente,
                'contato': contato,
                'tipo_contato': tipo_contato,
                'telefone': telefone,
                'email': email,
                'consultor': consultor,
                'data': data_atual.strftime('%d/%m/%Y'),
                'dias_validade': dias_validade,
                'validade': data_validade.strftime('%d/%m/%Y'),
                'prazo_entrega': prazo_entrega,
                'condicoes_pagamento': condicoes_pagamento,
                'observacoes': observacoes
            }
            pdf_bytes = gerar_pdf_orcamento(cli_info, st.session_state.ambientes)
            st.download_button(
                label="📄 Exportar Orçamento em PDF",
                data=pdf_bytes,
                file_name=f"Orcamento_{prop_num_atual}_{cliente.replace(' ', '_')}.pdf",
                mime="application/pdf",
                use_container_width=True
            )

# --- ABA 2: ORÇAMENTOS SALVOS ---
elif menu == "📋 Orçamentos Salvos":
    st.subheader("📋 Orçamentos Salvos no Banco de Dados")
    conn = get_connection()
    c = conn.cursor()
    
    df_orc = pd.read_sql_query("SELECT id, proposta_num, cliente, consultor, data, total_liquido FROM orcamentos ORDER BY id DESC", conn)
    
    if not df_orc.empty:
        for idx, row in df_orc.iterrows():
            with st.container():
                col_a1, col_a2, col_a3, col_a4, col_a5 = st.columns([1.5, 3, 2, 1.2, 1.2])
                col_a1.write(f"**Proposta: #{row['proposta_num']}**")
                col_a2.write(f"**{row['cliente']}** ({row['consultor']})")
                col_a3.write(f"R$ {row['total_liquido']:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                
                # BOTÃO EDITAR: Carrega absolutamente TODOS os dados e navega sozinho!
                if col_a4.button("✏️ Editar", key=f"btn_edit_orc_{row['id']}"):
                    st.session_state.edit_index = row['id']
                    
                    c.execute("SELECT * FROM orcamentos WHERE id = ?", (row['id'],))
                    o = c.fetchone()
                    
                    st.session_state.cli_prop = o[1]
                    st.session_state.cli_nome = o[2]
                    st.session_state.cli_contato = o[3]
                    st.session_state.cli_tipo = o[4]
                    st.session_state.cli_tel = o[5]
                    st.session_state.cli_email = o[6]
                    st.session_state.cli_consultor = o[7]
                    st.session_state.cli_prazo = o[11]
                    st.session_state.cli_cond = o[12]
                    st.session_state.cli_obs = o[13]
                    
                    ambs_db = []
                    c.execute("SELECT id, nome_ambiente, especificacoes, total_ambiente FROM ambientes WHERE orcamento_id = ? ORDER BY ordem", (row['id'],))
                    for amb_row in c.fetchall():
                        itens_db = []
                        c.execute("SELECT descricao, valor, eh_opcional FROM itens WHERE ambiente_id = ? ORDER BY ordem", (amb_row[0],))
                        for item_row in c.fetchall():
                            itens_db.append({
                                'descricao': item_row[0],
                                'valor': item_row[1],
                                'eh_opcional': bool(item_row[2])
                            })
                        ambs_db.append({
                            'nome': amb_row[1],
                            'especificacoes': amb_row[2],
                            'total_ambiente': amb_row[3],
                            'itens': itens_db
                        })
                    st.session_state.ambientes = ambs_db
                    st.session_state.navigate_to = "➕ Novo / Editar Orçamento"
                    st.rerun()

                if col_a5.button("🗑️ Excluir", key=f"btn_del_orc_{row['id']}"):
                    st.session_state.confirm_del = f"orc_{row['id']}"

                if st.session_state.confirm_del == f"orc_{row['id']}":
                    st.warning(f"⚠️ Confirma a exclusão permanente do Orçamento Proposta #{row['proposta_num']}?")
                    c_del1, c_del2 = st.columns(2)
                    if c_del1.button("✅ Confirmar", key=f"conf_del_orc_{row['id']}"):
                        c.execute("DELETE FROM orcamentos WHERE id = ?", (row['id'],))
                        conn.commit()
                        st.session_state.confirm_del = None
                        st.rerun()
                    if c_del2.button("❌ Cancelar", key=f"canc_del_orc_{row['id']}"):
                        st.session_state.confirm_del = None
                        st.rerun()

                st.markdown("---")
    else:
        st.info("Nenhum orçamento cadastrado.")

# --- ABA 3: CONFIGURAÇÕES E LOGO ---
elif menu == "⚙️ Configurações":
    st.subheader("⚙️ Configurações da Empresa e Consultores")
    conn = get_connection()
    c = conn.cursor()
    
    st.markdown("#### 🏢 Logo e Dados da Fábrica")
    config = get_config()
    
    uploaded_logo = st.file_uploader("Enviar Logo da Empresa (PNG / JPG)", type=["png", "jpg", "jpeg"])
    if uploaded_logo:
        logo_path = os.path.join("logo_empresa.png")
        with open(logo_path, "wb") as f:
            f.write(uploaded_logo.getbuffer())
        c.execute("UPDATE config_empresa SET logo_path = ? WHERE id = 1", (logo_path,))
        conn.commit()
        st.success("Logo atualizada com sucesso!")

    logo_actual = config.get('logo_path', '')
    if logo_actual and os.path.exists(logo_actual):
        st.image(logo_actual, width=200, caption="Logo Atual")

    with st.form("form_config_empresa"):
        nome_empresa = st.text_input("Razão Social", value=config.get('nome_empresa', 'Fábrica de Móveis Laurenti Ltda'))
        cnpj = st.text_input("CNPJ", value=config.get('cnpj', '44.331.015/0001-08'))
        ie = st.text_input("Inscrição Estadual (IE)", value=config.get('ie', '186000158114'))
        endereco = st.text_input("Endereço Completo", value=config.get('endereco', 'Rua Henrique Villa, 59- Jardim Maria Emília. CEP: 15960000, ARIRANHA-SP'))
        telefone = st.text_input("Telefone", value=config.get('telefone', '(17) 3576-1464'))
        email = st.text_input("E-mail", value=config.get('email', 'contato@laurentimoveis.com.br'))
        
        if st.form_submit_button("💾 Salvar Dados da Empresa"):
            c.execute("""
                UPDATE config_empresa
                SET nome_empresa = ?, cnpj = ?, ie = ?, endereco = ?, telefone = ?, email = ?
                WHERE id = 1
            """, (nome_empresa, cnpj, ie, endereco, telefone, email))
            conn.commit()
            st.success("Dados da empresa salvos!")
            st.rerun()

    st.markdown("---")
    st.markdown("#### 👤 Gestão de Consultores")
    
    with st.form("form_add_consultor", clear_on_submit=True):
        novo_c = st.text_input("Nome do Novo Consultor")
        if st.form_submit_button("➕ Cadastrar Consultor"):
            if novo_c:
                c.execute("INSERT INTO consultores (nome) VALUES (?)", (novo_c,))
                conn.commit()
                st.rerun()

    df_c = get_consultores()
    if not df_c.empty:
        for _, r in df_c.iterrows():
            col_cons1, col_cons2 = st.columns([4, 1])
            col_cons1.write(f"👤 **{r['nome']}**")
            
            if col_cons2.button("🗑️ Excluir", key=f"del_cons_btn_{r['id']}"):
                st.session_state.confirm_del = f"cons_{r['id']}"

            if st.session_state.confirm_del == f"cons_{r['id']}":
                st.warning(f"⚠️ Remover o consultor '{r['nome']}'?")
                c_del1, c_del2 = st.columns(2)
                if c_del1.button("✅ Confirmar", key=f"conf_del_cons_{r['id']}"):
                    c.execute("DELETE FROM consultores WHERE id = ?", (r['id'],))
                    conn.commit()
                    st.session_state.confirm_del = None
                    st.rerun()
                if c_del2.button("❌ Cancelar", key=f"canc_del_cons_{r['id']}"):
                    st.session_state.confirm_del = None
                    st.rerun()