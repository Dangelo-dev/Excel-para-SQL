import pandas as pd
from sqlalchemy import create_engine, exc, text

def abrir_planilha(loja_id, arquivo):
    df = pd.read_excel('7087.xlsx', skiprows=10, usecols=[1,4,5,6,7])
    df = df[:-1] # Excluindo a ultima linha que contém o total acumulado
    df.columns = ['Data', 'Boletos', 'Itens', 'Venda_Liquida', 'Desconto']
    df.insert(0, 'Loja', loja_id)
    return df

# lista de lojas e arquivos correspondentes
lojas = {
    '7087': '7087.xlsx',
    '7162': '7162.xlsx',
    '8333': '8333.xlsx',
    '11276': '11276.xlsx',
    '11996': '11996.xlsx',
    '12268': '12268.xlsx',
    '12396': '12396.xlsx',
    '12674': '12674.xlsx',
    '13003': '13003.xlsx',
    '13391': '13391.xlsx',
    '13557': '13557.xlsx',
    '17868': '17868.xlsx',
    '19386': '19386.xlsx',
    '19530': '19530.xlsx',
    '19964': '19964.xlsx',
    '21372': '21372.xlsx'
}

engine = create_engine('sqlite:///vendas_lojas.db')

criar_tabela = """
    CREATE TABLE IF NOT EXISTS vendas_lojas (
        Loja TEXT NOT NULL,
        Data DATE NOT NULL,
        Boletos INTEGER NOT NULL,
        Itens INTEGER NOT NULL,
        Venda_Liquida REAL NOT NULL,
        Desconto REAL,
        UNIQUE (Loja, Data, Boletos)
    )
"""

with engine.connect() as conn:
    conn.execute(text(criar_tabela))
    
for loja_id, arquivo in lojas.items():
    df_result = abrir_planilha(loja_id, arquivo)
    try:
        df_result.to_sql('vendas_lojas', con=engine, if_exists='append', index=False)
    except exc.IntegrityError as e:
        print(f"Erro ao inserir dados para a loja {loja_id}: {e.orig}")

print('Dados inseridos com sucesso!')
