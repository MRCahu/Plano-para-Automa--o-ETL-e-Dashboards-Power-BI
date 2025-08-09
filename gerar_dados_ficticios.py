import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import random

# Configurar seed para reprodutibilidade
np.random.seed(42)
random.seed(42)

# Definir parâmetros
num_registros = 1000

# Listas de dados fictícios
departamentos = [
    'Recursos Humanos', 'Financeiro', 'Tecnologia da Informação', 
    'Compras e Licitações', 'Planejamento', 'Jurídico', 
    'Comunicação', 'Infraestrutura', 'Saúde', 'Educação'
]

categorias = [
    'Material de Escritório', 'Equipamentos', 'Serviços', 
    'Manutenção', 'Capacitação', 'Combustível', 
    'Alimentação', 'Limpeza', 'Segurança', 'Consultoria'
]

tipos_transacao = ['Despesa', 'Receita', 'Transferência']

status_opcoes = ['Aprovado', 'Pendente', 'Rejeitado', 'Em Análise']

fornecedores = [
    'Empresa Alpha Ltda', 'Beta Serviços SA', 'Gamma Tecnologia',
    'Delta Suprimentos', 'Epsilon Consultoria', 'Zeta Equipamentos',
    'Eta Manutenção', 'Theta Sistemas', 'Iota Materiais', 'Kappa Serviços'
]

# Gerar dados
dados = []

for i in range(num_registros):
    # Data aleatória nos últimos 2 anos
    data_inicio = datetime.now() - timedelta(days=730)
    data_aleatoria = data_inicio + timedelta(days=random.randint(0, 730))
    
    # Valores aleatórios com distribuição realista
    if random.choice(tipos_transacao) == 'Receita':
        valor = round(random.uniform(5000, 50000), 2)
    else:
        valor = round(random.uniform(100, 15000), 2)
    
    registro = {
        'ID': f'TXN{str(i+1).zfill(4)}',
        'Data': data_aleatoria.strftime('%Y-%m-%d'),
        'Departamento': random.choice(departamentos),
        'Categoria': random.choice(categorias),
        'Tipo_Transacao': random.choice(tipos_transacao),
        'Valor': valor,
        'Fornecedor': random.choice(fornecedores),
        'Status': random.choice(status_opcoes),
        'Descricao': f'Transação {i+1} - {random.choice(categorias)}',
        'Mes': data_aleatoria.month,
        'Ano': data_aleatoria.year,
        'Trimestre': f'Q{((data_aleatoria.month-1)//3)+1}',
        'Responsavel': f'Funcionário {random.randint(1, 50)}',
        'Centro_Custo': f'CC{random.randint(1000, 9999)}',
        'Prioridade': random.choice(['Alta', 'Média', 'Baixa'])
    }
    
    dados.append(registro)

# Criar DataFrame
df = pd.DataFrame(dados)

# Adicionar algumas colunas calculadas
df['Valor_Acumulado'] = df.groupby('Departamento')['Valor'].cumsum()
df['Media_Mensal_Dept'] = df.groupby(['Departamento', 'Mes'])['Valor'].transform('mean')

# Salvar em Excel
nome_arquivo = '/home/ubuntu/dados_ficticios_1000_linhas.xlsx'
with pd.ExcelWriter(nome_arquivo, engine='openpyxl') as writer:
    # Planilha principal
    df.to_excel(writer, sheet_name='Dados_Principais', index=False)
    
    # Planilha de resumo por departamento
    resumo_dept = df.groupby('Departamento').agg({
        'Valor': ['sum', 'mean', 'count'],
        'ID': 'count'
    }).round(2)
    resumo_dept.columns = ['Total_Valor', 'Media_Valor', 'Count_Valor', 'Total_Transacoes']
    resumo_dept.to_excel(writer, sheet_name='Resumo_Departamento')
    
    # Planilha de resumo mensal
    resumo_mensal = df.groupby(['Ano', 'Mes']).agg({
        'Valor': 'sum',
        'ID': 'count'
    }).round(2)
    resumo_mensal.columns = ['Total_Valor', 'Total_Transacoes']
    resumo_mensal.to_excel(writer, sheet_name='Resumo_Mensal')

print(f"Arquivo Excel criado com sucesso: {nome_arquivo}")
print(f"Total de registros: {len(df)}")
print(f"Colunas: {list(df.columns)}")
print("\nPrimeiras 5 linhas:")
print(df.head())

# Estatísticas básicas
print("\nEstatísticas básicas:")
print(df.describe())

