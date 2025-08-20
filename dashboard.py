import streamlit as st
import sqlite3
import pandas as pd
from datetime import datetime, timedelta
import matplotlib.pyplot as plt
import seaborn as sns
import os
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
import pywhatkit

# Configurar o Streamlit
st.set_page_config(page_title="Dashboard Hospedagem MRV", layout="wide")

# Função para criar o banco de dados
def create_db():
    conn = sqlite3.connect('hospedagem.db')
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS hospedagens (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT,
            empresa TEXT,
            apt INTEGER,
            data_inicio TEXT,
            data_fim TEXT,
            diarias INTEGER,
            valor_diaria REAL,
            projeto TEXT
        )
    ''')
    conn.commit()
    conn.close()

# Função para inserir hospedagem
def insert_hospedagem(nome, empresa, apt, data_inicio, data_fim, diarias, valor_diaria, projeto):
    conn = sqlite3.connect('hospedagem.db')
    c = conn.cursor()
    c.execute('''
        INSERT INTO hospedagens (nome, empresa, apt, data_inicio, data_fim, diarias, valor_diaria, projeto)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    ''', (nome, empresa, apt, data_inicio, data_fim, diarias, valor_diaria, projeto))
    conn.commit()
    conn.close()

# Função para obter hospedagens com filtros
def get_hospedagens(projeto=None, data_min=None, data_max=None, empresa=None, apt=None):
    conn = sqlite3.connect('hospedagem.db')
    query = 'SELECT * FROM hospedagens'
    params = []
    where = []
    if projeto:
        where.append('projeto = ?')
        params.append(projeto)
    if data_min:
        where.append('data_inicio >= ?')
        params.append(data_min)
    if data_max:
        where.append('data_fim <= ?')
        params.append(data_max)
    if empresa:
        where.append('empresa LIKE ?')
        params.append(f'%{empresa}%')
    if apt:
        where.append('apt = ?')
        params.append(apt)
    if where:
        query += ' WHERE ' + ' AND '.join(where)
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    if not df.empty:
        df['total'] = df['diarias'] * df['valor_diaria']
    return df

# Função para importar planilha
def import_from_excel(file_path):
    xl = pd.ExcelFile(file_path)
    for sheet_name in xl.sheet_names:
        if sheet_name == 'TOTAL GERAL':
            continue  # Pular, pois é um resumo
        df = pd.read_excel(xl, sheet_name=sheet_name, header=None)
        data_start_row = None
        for i, row in df.iterrows():
            if 'NOME' in str(row[0]) or 'APTOS' in str(row[0]):
                data_start_row = i
                break
        if data_start_row is None:
            continue
        df_data = pd.read_excel(xl, sheet_name=sheet_name, skiprows=data_start_row)
        
        for index, row in df_data.iterrows():
            nome = row.get('NOME') or row.get('APTOS') or ''
            if pd.isna(nome) or not nome.strip():
                continue
            empresa = row.get('EMPRESA') or row.get('Funcionários') or ''
            apt = int(row.get('APT', 0)) if pd.notna(row.get('APT')) else 0
            def excel_date(serial):
                if pd.isna(serial) or not isinstance(serial, (int, float)):
                    return datetime.now().strftime('%Y-%m-%d')
                return (datetime(1899, 12, 30) + timedelta(days=serial)).strftime('%Y-%m-%d')
            diarias = 0
            data_inicio = None
            data_fim = None
            for col in df_data.columns:
                if str(col).startswith('45'):  # Colunas de datas seriais
                    if row[col] == 1:
                        diarias += 1
                        serial = int(col)
                        date = (datetime(1899, 12, 30) + timedelta(days=serial)).strftime('%Y-%m-%d')
                        if data_inicio is None or date < data_inicio:
                            data_inicio = date
                        if data_fim is None or date > data_fim:
                            data_fim = date
            if diarias == 0:
                diarias = int(row.get('TOTAL DE DIÁRIAS:', 0)) or 0
            valor_diaria = row.get('Valor por diária', 42.0)
            if isinstance(valor_diaria, str) and 'R$' in valor_diaria:
                valor_diaria = float(valor_diaria.replace('R$', '').strip()) or 42.0
            insert_hospedagem(
                nome,
                empresa,
                apt,
                data_inicio or datetime.now().strftime('%Y-%m-%d'),
                data_fim or datetime.now().strftime('%Y-%m-%d'),
                diarias,
                valor_diaria,
                projeto=sheet_name
            )

# Função para gerar PDF com reportlab
def generate_pdf(data, filename='relatorio.pdf'):
    if data.empty:
        return None
    doc = SimpleDocTemplate(filename, pagesize=A4)
    styles = getSampleStyleSheet()
    elements = []
    elements.append(Paragraph("Relatório de Hospedagem", styles['Title']))
    data_list = [data.columns.tolist()] + data.values.tolist()
    table = Table(data_list)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    elements.append(table)
    doc.build(elements)
    return filename

# Interface Streamlit
st.title("🏨 Dashboard Hospedagem MRV")
create_db()

# Sidebar para navegação
st.sidebar.header("Navegação")
page = st.sidebar.selectbox("Selecione a Página", ["Resumo Geral", "Ocupação", "Hospedagem Externa", "Detalhes por Projeto", "Lançar Demanda"])

# Filtros globais
st.sidebar.header("Filtros")
projeto_filter = st.sidebar.selectbox("Projeto", ["Todos", "Beach Plaza", "Norte Plaza", "Torre dos Seridós"])
empresa_filter = st.sidebar.text_input("Empresa (ou parte do nome)")
apt_filter = st.sidebar.number_input("Apartamento", min_value=0, value=0, step=1)
data_min = st.sidebar.date_input("Data Início")
data_max = st.sidebar.date_input("Data Fim")

# Aplicar filtros
filtros = {}
if projeto_filter != "Todos":
    filtros['projeto'] = projeto_filter
if empresa_filter:
    filtros['empresa'] = empresa_filter
if apt_filter > 0:
    filtros['apt'] = apt_filter
if data_min:
    filtros['data_min'] = data_min.strftime('%Y-%m-%d')
if data_max:
    filtros['data_max'] = data_max.strftime('%Y-%m-%d')

# Página: Resumo Geral
if page == "Resumo Geral":
    st.header("Resumo Geral")
    df = get_hospedagens(**filtros)
    if not df.empty:
        st.subheader("Totais por Projeto")
        totais = df.groupby('projeto').agg({'diarias': 'sum', 'total': 'sum'}).reset_index()
        st.dataframe(totais)
        
        # Gráfico de barras
        fig, ax = plt.subplots()
        sns.barplot(data=totais, x='projeto', y='diarias', ax=ax)
        plt.title("Diárias por Projeto")
        plt.xticks(rotation=45)
        st.pyplot(fig)
        
        # Interessante: Maior projeto
        max_projeto = totais.loc[totais['diarias'].idxmax()]['projeto']
        st.write(f"**Fato Interessante**: O projeto {max_projeto} tem o maior número de diárias ({totais['diarias'].max()}).")

# Página: Ocupação
if page == "Ocupação":
    st.header("Ocupação de Apartamentos")
    df = get_hospedagens(**filtros)
    if not df.empty:
        ocupacao = df.groupby(['apt', 'projeto']).size().reset_index(name='ocupantes')
        st.dataframe(ocupacao)
        
        # Gráfico de ocupação
        fig, ax = plt.subplots()
        sns.barplot(data=ocupacao, x='apt', y='ocupantes', hue='projeto', ax=ax)
        plt.title("Ocupantes por Apartamento")
        st.pyplot(fig)

# Página: Hospedagem Externa
if page == "Hospedagem Externa":
    st.header("Hospedagem Externa")
    df = get_hospedagens(**filtros)
    if not df.empty:
        externa = df[df['apt'].isin([8, 14])][['apt', 'diarias', 'total']]
        st.dataframe(externa)
        
        # Gráfico
        fig, ax = plt.subplots()
        sns.barplot(data=externa, x='apt', y='total', ax=ax)
        plt.title("Custo Total de Hospedagem Externa")
        st.pyplot(fig)

# Página: Detalhes por Projeto
if page == "Detalhes por Projeto":
    st.header("Detalhes por Projeto")
    df = get_hospedagens(**filtros)
    if not df.empty:
        st.dataframe(df)
        
        # Gráfico de diárias por empresa
        fig, ax = plt.subplots()
        sns.barplot(data=df, x='empresa', y='diarias', hue='projeto', ax=ax)
        plt.title("Diárias por Empresa e Projeto")
        plt.xticks(rotation=45)
        st.pyplot(fig)

# Página: Lançar Demanda
if page == "Lançar Demanda":
    st.header("Lançar Nova Demanda")
    with st.form("form_demanda"):
        nome = st.text_input("Nome")
        empresa = st.text_input("Empresa")
        apt = st.number_input("Apartamento", min_value=0, step=1)
        data_inicio = st.date_input("Data Início")
        data_fim = st.date_input("Data Fim")
        diarias = st.number_input("Diárias", min_value=0, step=1)
        valor_diaria = st.number_input("Valor Diária", min_value=0.0, value=42.0)
        projeto = st.selectbox("Projeto", ["Beach Plaza", "Norte Plaza", "Torre dos Seridós"])
        submitted = st.form_submit_button("Salvar")
        if submitted:
            insert_hospedagem(
                nome,
                empresa,
                apt,
                data_inicio.strftime('%Y-%m-%d'),
                data_fim.strftime('%Y-%m-%d'),
                diarias,
                valor_diaria,
                projeto
            )
            st.success("Demanda salva com sucesso!")

# Botão para importar planilha
st.sidebar.header("Importar Planilha")
uploaded_file = st.sidebar.file_uploader("Carregar HOSPEDAGEM MRV-14.xlsx", type=["xlsx"])
if uploaded_file:
    with open("temp.xlsx", "wb") as f:
        f.write(uploaded_file.getbuffer())
    import_from_excel("temp.xlsx")
    st.sidebar.success("Planilha importada com sucesso!")

# Botão para gerar e compartilhar relatório
if st.sidebar.button("Gerar e Compartilhar Relatório"):
    df = get_hospedagens(**filtros)
    pdf_file = generate_pdf(df)
    if pdf_file:
        with open(pdf_file, "rb") as f:
            st.sidebar.download_button("Baixar PDF", f, file_name="relatorio.pdf")
        # Para WhatsApp (envio manual ou automatizado)
        st.sidebar.write("PDF gerado. Envie manualmente via WhatsApp ou configure um número:")
        numero_whatsapp = st.sidebar.text_input("Número WhatsApp (+55DDDNUMERO)")
        if numero_whatsapp and st.sidebar.button("Enviar via WhatsApp"):
            pywhatkit.sendwhats_image(numero_whatsapp, pdf_file, "Relatório de Hospedagem")
            st.sidebar.success("Enviado para o WhatsApp!")