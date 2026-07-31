import streamlit as st
import sqlite3
import pandas as pd
from datetime import datetime, timedelta
import io

# ReportLab para geração de PDF de alta qualidade
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm

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
# 2. GERADOR DE PDF (ReportLab)
# -----------------------------------------------------------------------------
def gerar_pdf_orcamento(cliente_info, ambientes_list):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=1.5*cm,
        leftMargin=1.5*cm,
        topMargin=1.5*cm,
        bottomMargin=2.0*cm
    )
    
    story = []
    styles = getSampleStyleSheet()
    
    # Estilos customizados
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=18, textColor=colors.HexColor('#2C3E50'), spaceAfter=4)
    subtitle_style = ParagraphStyle('SubTitleStyle', parent=styles['Normal'], fontSize=9, textColor=colors.HexColor('#7F8C8D'))
    h2_style = ParagraphStyle('H2Style', parent=styles['Heading2'], fontSize=12, textColor=colors.HexColor('#2C3E50'), spaceBefore=8, spaceAfter=4)
    body_style = ParagraphStyle('BodyStyle', parent=styles['Normal'], fontSize=9, leading=11, textColor=colors.HexColor('#333333'))
    body_bold = ParagraphStyle('BodyBold', parent=body_style, fontName='Helvetica-Bold')
    right_bold = ParagraphStyle('RightBold', parent=body_bold, alignment=2)
    right_normal = ParagraphStyle('RightNormal', parent=body_style, alignment=2)
    center_bold = ParagraphStyle('CenterBold', parent=body_bold, alignment=1)

    # Cabeçalho da Empresa
    story.append(Paragraph("<b>LAURENTI MÓVEIS</b>", title_style))
    story.append(Paragraph("Móveis Planejados & Marcenaria de Alto Padrão", subtitle_style))
    story.append(Spacer(1, 10))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#BDC3C7'), spaceBefore=0, spaceAfter=10))

    # Tabela de Informações do Cliente e Orçamento
    data_cli = [
        [
            Paragraph(f"<b>Cliente:</b> {cliente_info['cliente']}", body_style),
            Paragraph(f"<b>Data:</b> {cliente_info['data']}", body_style)
        ],
        [
            Paragraph(f"<b>Contato:</b> {cliente_info['contato']} ({cliente_info['tipo_contato']})", body_style),
            Paragraph(f"<b>Validade:</b> {cliente_info['validade']} ({cliente_info['dias_validade']} dias)", body_style)
        ],
        [
            Paragraph(f"<b>Telefone:</b> {cliente_info['telefone']}", body_style),
            Paragraph(f"<b>E-mail:</b> {cliente_info['email']}", body_style)
        ]
    ]
    t_cli = Table(data_cli, colWidths=[11*cm, 7*cm])
    t_cli.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8F9FA')),
        ('PADDING', (0,0), (-1,-1), 6),
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    story.append(t_cli)
    story.append(Spacer(1, 15))

    # Tabela Principal de Ambientes e Itens
    table_data = [
        [Paragraph("<b>Item</b>", center_bold), Paragraph("<b>Descrição do Ambiente / Subitem</b>", body_bold), Paragraph("<b>Parcial (R$)</b>", right_bold), Paragraph("<b>Total (R$)</b>", right_bold)]
    ]
    
    tot_liquido = 0.0
    tot_opcionais = 0.0

    for idx_a, amb in enumerate(ambientes_list):
        ordem_amb = idx_a + 1
        spec_text = f"<br/><font color='#666666'><i>Espec: {amb['especificacoes']}</i></font>" if amb.get('especificacoes') else ""
        amb_title = f"<b>{amb['nome']}</b>{spec_text}"
        
        # Calcular total do ambiente
        tot_amb = sum(float(item['valor']) for item in amb['itens'] if not item['eh_opcional'])
        tot_liquido += tot_amb
        
        table_data.append([
            Paragraph(f"<b>{ordem_amb}</b>", center_bold),
            Paragraph(amb_title, body_style),
            Paragraph("", body_style),
            Paragraph(f"<b>{tot_amb:,.2f}</b>".replace(",", "X").replace(".", ",").replace("X", "."), right_bold)
        ])
        
        # Subitens
        for idx_i, item in enumerate(amb['itens']):
            ordem_item = f"{ordem_amb}.{idx_i + 1}"
            opc_tag = " <font color='#D97706'><b>(OPCIONAL)</b></font>" if item['eh_opcional'] else ""
            desc_item = f"&nbsp;&nbsp;&nbsp;&nbsp;• {item['descricao']}{opc_tag}"
            val_f = f"{float(item['valor']):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            
            if item['eh_opcional']:
                tot_opcionais += float(item['valor'])

            table_data.append([
                Paragraph(ordem_item, ParagraphStyle('SubItemCode', parent=body_style, alignment=1, fontSize=8)),
                Paragraph(desc_item, body_style),
                Paragraph(val_f, right_normal),
                Paragraph("", body_style)
            ])

    t_itens = Table(table_data, colWidths=[1.8*cm, 10.2*cm, 3*cm, 3*cm])
    t_itens.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2C3E50')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('PADDING', (0,0), (-1,-1), 5),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    # Ajustar cor das linhas dos títulos da tabela
    for i in range(len(table_data)):
        if i == 0:
            continue
        # Aplica fundo sutil para cabeçalhos de ambiente
        if table_data[i][3].text != "":
            t_itens.setStyle(TableStyle([('BACKGROUND', (0, i), (-1, i), colors.HexColor('#F1F5F9'))]))

    story.append(t_itens)
    story.append(Spacer(1, 15))

    # Tabela de Condições e Totais
    tot_com_opc = tot_liquido + tot_opcionais
    
    cond_text = f"""
    <b>Prazo de Entrega:</b> {cliente_info['prazo_entrega']}<br/>
    <b>Condições de Pagamento:</b> {cliente_info['condicoes_pagamento']}<br/><br/>
    <b>Observações:</b><br/>{cliente_info['observacoes']}
    """
    
    totais_text = f"""
    <font size=10><b>Total Líquido:</b></font><br/>
    <font size=13 color='#2C3E50'><b>R$ {tot_liquido:,.2f}</b></font><br/><br/>
    <font size=9 color='#666666'>Total c/ Opcionais:</font><br/>
    <font size=10 color='#D97706'><b>R$ {tot_com_opc:,.2f}</b></font>
    """.replace(",", "X").replace(".", ",").replace("X", ".")
    
    t_totais = Table([
        [Paragraph(cond_text, body_style), Paragraph(totais_text, right_bold)]
    ], colWidths=[12*cm, 6*cm])
    
    t_totais.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8F9FA')),
        ('PADDING', (0,0), (-1,-1), 8),
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
    ]))
    story.append(t_totais)

    # Rodapé Nativo em Todas as Páginas
    def add_footer(canvas, doc):
        canvas.saveState()
        footer_line1 = "LAURENTI MÓVEIS — CNPJ: 00.000.000/0001-00"
        footer_line2 = "Rua Henrique Villa, 59 - Ariranha/SP — Tel: (17) 3576-1464"
        
        canvas.setFont('Helvetica', 8)
        canvas.setFillColor(colors.HexColor('#64748B'))
        canvas.drawCentredString(A4[0]/2, 1.2*cm, footer_line1)
        canvas.drawCentredString(A4[0]/2, 0.8*cm, footer_line2)
        
        page_num = canvas.getPageNumber()
        canvas.drawRightString(A4[0] - 1.5*cm, 0.8*cm, f"Página {page_num}")
        canvas.restoreState()

    doc.build(story, onFirstPage=add_footer, onLaterPages=add_footer)
    buffer.seek(0)
    return buffer

# -----------------------------------------------------------------------------
# 3. GERENCIAMENTO DE ESTADO (SESSION STATE)
# -----------------------------------------------------------------------------
if 'ambientes' not in st.session_state:
    st.session_state.ambientes = []

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
# 4. INTERFACE PRINCIPAL
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
        
        btn_add_amb = st.form_submit_button("➕ Criar Ambiente")
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

    # Exibição, Edição e Exclusão de Ambientes e seus Subitens
    if st.session_state.ambientes:
        for idx_amb, amb in enumerate(st.session_state.ambientes):
            with st.expander(f"🛋️ **Ambiente {idx_amb + 1}: {amb['nome']}** — Total: R$ {amb['total_ambiente']:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), expanded=True):
                
                # Edição do próprio Ambiente
                col_e1, col_e2, col_e3 = st.columns([2, 3, 1])
                with col_e1:
                    novo_nome = st.text_input(f"Nome do Ambiente #{idx_amb+1}", value=amb['nome'], key=f"edit_nome_{idx_amb}")
                with col_e2:
                    nova_espec = st.text_input(f"Especificações #{idx_amb+1}", value=amb['especificacoes'], key=f"edit_espec_{idx_amb}")
                with col_e3:
                    st.write(" ")
                    st.write(" ")
                    if st.button("🗑️ Apagar Ambiente", key=f"del_amb_{idx_amb}", type="secondary"):
                        st.session_state.ambientes.pop(idx_amb)
                        st.rerun()
                
                # Atualizar dados do ambiente se editados
                amb['nome'] = novo_nome
                amb['especificacoes'] = nova_espec

                st.markdown("##### **Subitens do Ambiente**")
                
                # Form para Adicionar Item
                col_i1, col_i2, col_i3, col_i4 = st.columns([3, 1.5, 1, 1])
                with col_i1:
                    desc_item = st.text_input("Descrição do Item", key=f"desc_add_{idx_amb}", placeholder="Ex: Armário Superior 2 Portas")
                with col_i2:
                    valor_item = st.number_input("Valor (R$)", min_value=0.0, step=50.0, format="%.2f", key=f"val_add_{idx_amb}")
                with col_i3:
                    st.write(" ")
                    eh_opc = st.checkbox("Opcional?", key=f"opc_add_{idx_amb}")
                with col_i4:
                    st.write(" ")
                    if st.button("➕ Adicionar Item", key=f"btn_add_item_{idx_amb}"):
                        if desc_item and valor_item > 0:
                            amb['itens'].append({
                                'descricao': desc_item,
                                'valor': valor_item,
                                'eh_opcional': eh_opc
                            })
                            st.rerun()
                        else:
                            st.warning("Informe a descrição e valor.")

                # Lista e Edição dos Itens Existentes
                if amb['itens']:
                    for idx_item, item in enumerate(amb['itens']):
                        col_it1, col_it2, col_it3, col_it4 = st.columns([3, 1.5, 1, 0.5])
                        with col_it1:
                            item['descricao'] = st.text_input(f"Item {idx_amb+1}.{idx_item+1}", value=item['descricao'], key=f"item_desc_{idx_amb}_{idx_item}")
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

    # Resumo Final, PDF e Salvamento
    tot_liquido, tot_com_opcionais = recalcular_totais()
    
    st.markdown("---")
    st.markdown("### 📊 Resumo de Valores e Emissão")
    c_tot1, c_tot2 = st.columns(2)
    c_tot1.metric("Total Líquido (Base)", f"R$ {tot_liquido:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    c_tot2.metric("Total c/ Opcionais", f"R$ {tot_com_opcionais:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

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
                c.execute("""
                    INSERT INTO orcamentos (cliente, contato, tipo_contato, telefone, email, data, dias_validade, validade, prazo_entrega, condicoes_pagamento, observacoes, total_liquido, total_com_opcionais)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (cliente, contato, tipo_contato, telefone, email, str(data_atual), dias_validade, str(data_validade.strftime('%d/%m/%Y')), prazo_entrega, condicoes_pagamento, observacoes, tot_liquido, tot_com_opcionais))
                
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
                st.success(f"✅ Orçamento #{orc_id} salvo com sucesso no banco de dados!")

    with col_btn2:
        if cliente and st.session_state.ambientes:
            cli_info = {
                'cliente': cliente,
                'contato': contato,
                'tipo_contato': tipo_contato,
                'telefone': telefone,
                'email': email,
                'data': data_atual.strftime('%d/%m/%Y'),
                'dias_validade': dias_validade,
                'validade': data_validade.strftime('%d/%m/%Y'),
                'prazo_entrega': prazo_entrega,
                'condicoes_pagamento': condicoes_pagamento,
                'observacoes': observacoes
            }
            pdf_bytes = gerar_pdf_orcamento(cli_info, st.session_state.ambientes)
            st.download_button(
                label="📄 Gerar e Baixar Orçamento em PDF",
                data=pdf_bytes,
                file_name=f"Orcamento_{cliente.replace(' ', '_')}.pdf",
                mime="application/pdf",
                use_container_width=True
            )

elif menu == "📋 Orçamentos Salvos":
    st.subheader("Orçamentos Salvos no Banco de Dados")
    conn = get_connection()
    df_orc = pd.read_sql_query("SELECT id, cliente, data, validade, total_liquido, total_com_opcionais FROM orcamentos ORDER BY id DESC", conn)
    
    if not df_orc.empty:
        st.dataframe(df_orc, use_container_width=True)
    else:
        st.info("Nenhum orçamento salvo ainda.")

elif menu == "⚙️ Configurações":
    st.subheader("Configurações da Empresa")
    st.write("**LAURENTI MÓVEIS**")
    st.write("CNPJ: 00.000.000/0001-00")
    st.write("Rua Henrique Villa, 59 - Ariranha/SP — Tel: (17) 3576-1464")
