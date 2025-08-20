import sqlite3
import pandas as pd
from datetime import datetime, timedelta
import matplotlib.pyplot as plt
import os
import tempfile
import webbrowser

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
def insert_hospedagem(nome, empresa, apt, data_inicio, data_fim, diarias, valor_diaria=42.0, projeto='Beach Plaza'):
    conn = sqlite3.connect('hospedagem.db')
    c = conn.cursor()
    c.execute('''
        INSERT INTO hospedagens (nome, empresa, apt, data_inicio, data_fim, diarias, valor_diaria, projeto)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    ''', (nome, empresa, apt, data_inicio, data_fim, diarias, valor_diaria, projeto))
    conn.commit()
    conn.close()

# Função para obter hospedagens com filtros
def get_hospedagens(projeto=None, data_min=None, data_max=None):
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
    if where:
        query += ' WHERE ' + ' AND '.join(where)
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    if not df.empty:
        df['total'] = df['diarias'] * df['valor_diaria']
    return df

# Função para gerar PDF usando matplotlib
def generate_pdf(data, filename='relatorio.pdf'):
    if data.empty:
        print("Nenhum dado para gerar relatório.")
        return
    fig, ax = plt.subplots(figsize=(12, len(data)*0.5 + 2))  # Ajuste o tamanho
    ax.axis('tight')
    ax.axis('off')
    col_labels = list(data.columns)
    table = ax.table(cellText=data.values, colLabels=col_labels, cellLoc='center', loc='center')
    table.auto_set_font_size(False)
    table.set_fontsize(8)
    table.scale(1.2, 1.2)
    plt.title('Relatório de Hospedagem')
    plt.savefig(filename, format='pdf', bbox_inches='tight')
    plt.close()

# Função melhorada para importar de Excel, adaptada à sua planilha
def import_from_excel(file_path):
    xl = pd.ExcelFile(file_path)
    for sheet_name in xl.sheet_names:
        if sheet_name in ['TOTAL GERAL', 'OCUPACAO', 'HOSPEDAGEM EXTERNA']:  # Sheets auxiliares, pule ou processe separadamente
            continue
        df = pd.read_excel(xl, sheet_name=sheet_name, header=None)  # Sem header fixo, pois irregular
        # Encontre a linha de dados reais (ex: após cabeçalhos como "Obra:", "NOME,EMPRESA")
        data_start_row = None
        for i, row in df.iterrows():
            if 'NOME' in str(row[0]) or 'APTOS' in str(row[0]):
                data_start_row = i
                break
        if data_start_row is None:
            continue
        df_data = pd.read_excel(xl, sheet_name=sheet_name, skiprows=data_start_row)
        
        # Limpeza e parsing
        for index, row in df_data.iterrows():
            nome = row.get('NOME') or row.get('APTOS') or ''
            if pd.isna(nome) or not nome.strip():
                continue
            empresa = row.get('EMPRESA') or row.get('Funcionários') or ''
            apt = int(row.get('APT', 0)) if pd.notna(row.get('APT')) else 0
            
            # Calcular diárias somando colunas de datas (ex: 45859, 45860...) onde valor=1
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
                diarias = int(row.get('TOTAL DE DIÁRIAS:', 0)) or 0  # Fallback para totais
            
            insert_hospedagem(
                nome,
                empresa,
                apt,
                data_inicio or datetime.now().strftime('%Y-%m-%d'),
                data_fim or datetime.now().strftime('%Y-%m-%d'),
                diarias,
                projeto=sheet_name
            )
    print("Importação concluída! Dados parseados das sheets principais como 'Beach Plaza'.")

# Exemplo de uso CLI
def main():
    create_db()
    while True:
        print("\n1. Lançar Demanda")
        print("2. Gerar Relatório")
        print("3. Importar Excel")
        print("4. Sair")
        choice = input("Escolha: ")
        if choice == '1':
            nome = input("Nome: ")
            empresa = input("Empresa: ")
            apt = int(input("APT: "))
            data_inicio = input("Data Início (YYYY-MM-DD): ")
            data_fim = input("Data Fim (YYYY-MM-DD): ")
            diarias = int(input("Diárias: "))
            projeto = input("Projeto (default Beach Plaza): ") or 'Beach Plaza'
            insert_hospedagem(nome, empresa, apt, data_inicio, data_fim, diarias, projeto=projeto)
        elif choice == '2':
            projeto = input("Filtro Projeto (ou vazio): ")
            data_min = input("Data Min (YYYY-MM-DD ou vazio): ")
            data_max = input("Data Max (YYYY-MM-DD ou vazio): ")
            df = get_hospedagens(projeto if projeto else None, data_min if data_min else None, data_max if data_max else None)
            print(df)
            generate_pdf(df)
            print("PDF gerado: relatorio.pdf")
            # Para imprimir: os.startfile('relatorio.pdf', 'print')  # Windows; ajuste para seu OS
            # Para WhatsApp: instale pywhatkit e use: import pywhatkit; pywhatkit.sendwhatmsg('+numero', 'Relatório', 0, True, attachment='relatorio.pdf')
        elif choice == '3':
            file_path = input("Caminho do Excel (ex: HOSPEDAGEM MRV.xlsx): ")
            import_from_excel(file_path)
        elif choice == '4':
            break

if __name__ == '__main__':
    main()
