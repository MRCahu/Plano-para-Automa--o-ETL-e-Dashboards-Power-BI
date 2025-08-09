"""
Script para criar visualizações impressionantes para o README do GitHub
Gera gráficos, diagramas e imagens para documentação
"""

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from datetime import datetime, timedelta
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch
import warnings
warnings.filterwarnings('ignore')

# Configurar estilo
plt.style.use('default')
sns.set_palette("husl")

def criar_banner_projeto():
    """Cria banner principal do projeto"""
    fig, ax = plt.subplots(figsize=(16, 6))
    
    # Fundo gradiente
    gradient = np.linspace(0, 1, 256).reshape(1, -1)
    gradient = np.vstack((gradient, gradient))
    
    # Cores do gradiente (azul para roxo)
    ax.imshow(gradient, aspect='auto', cmap='viridis', alpha=0.8, extent=[0, 16, 0, 6])
    
    # Título principal
    ax.text(8, 4.2, '🚀 AUTOMAÇÃO ETL & DASHBOARDS POWER BI', 
            fontsize=28, fontweight='bold', ha='center', va='center', 
            color='white', family='sans-serif')
    
    # Subtítulo
    ax.text(8, 3.2, 'Pipeline Completo de Dados com Python, Excel e Power BI', 
            fontsize=16, ha='center', va='center', 
            color='white', alpha=0.9, family='sans-serif')
    
    # Ícones e tecnologias
    tecnologias = ['🐍 Python', '📊 Excel', '📈 Power BI', '🔄 ETL', '📋 VBA', '☁️ Cloud']
    for i, tech in enumerate(tecnologias):
        x_pos = 1.5 + (i * 2.3)
        ax.text(x_pos, 1.5, tech, fontsize=12, ha='center', va='center',
                color='white', fontweight='bold',
                bbox=dict(boxstyle="round,pad=0.3", facecolor=(1, 1, 1, 0.2), 
                         edgecolor='white', linewidth=1))
    
    ax.set_xlim(0, 16)
    ax.set_ylim(0, 6)
    ax.axis('off')
    
    plt.tight_layout()
    plt.savefig('/home/ubuntu/banner_projeto.png', dpi=300, bbox_inches='tight', 
                facecolor='none', edgecolor='none')
    plt.close()

def criar_arquitetura_etl():
    """Cria diagrama de arquitetura ETL"""
    fig, ax = plt.subplots(figsize=(14, 10))
    
    # Cores
    cor_extract = '#FF6B6B'
    cor_transform = '#4ECDC4'
    cor_load = '#45B7D1'
    cor_arrow = '#2C3E50'
    
    # Função para criar caixas
    def criar_caixa(ax, x, y, width, height, text, color, text_color='white'):
        box = FancyBboxPatch((x, y), width, height,
                           boxstyle="round,pad=0.1",
                           facecolor=color, edgecolor='white', linewidth=2)
        ax.add_patch(box)
        ax.text(x + width/2, y + height/2, text, ha='center', va='center',
                fontsize=11, fontweight='bold', color=text_color, wrap=True)
    
    # Função para criar setas
    def criar_seta(ax, x1, y1, x2, y2):
        ax.annotate('', xy=(x2, y2), xytext=(x1, y1),
                   arrowprops=dict(arrowstyle='->', lw=3, color=cor_arrow))
    
    # EXTRACT (Fontes de Dados)
    ax.text(2, 9, 'EXTRACT', fontsize=16, fontweight='bold', ha='center', color=cor_extract)
    
    fontes = [
        ('📊 Excel Files', 0.5, 8),
        ('🗄️ Databases', 0.5, 7),
        ('📄 CSV Files', 0.5, 6),
        ('🌐 APIs', 3.5, 8),
        ('☁️ Cloud Storage', 3.5, 7),
        ('📋 ERP Systems', 3.5, 6)
    ]
    
    for fonte, x, y in fontes:
        criar_caixa(ax, x, y, 2.5, 0.8, fonte, cor_extract)
    
    # Setas para Transform
    for _, x, y in fontes:
        criar_seta(ax, x + 2.5, y + 0.4, 6.5, 5.5)
    
    # TRANSFORM (Processamento)
    ax.text(7.5, 9, 'TRANSFORM', fontsize=16, fontweight='bold', ha='center', color=cor_transform)
    
    transforms = [
        ('🧹 Data Cleaning', 6, 7.5),
        ('🔄 Data Validation', 6, 6.5),
        ('📊 Aggregations', 6, 5.5),
        ('🏷️ Categorization', 9, 7.5),
        ('📈 Calculations', 9, 6.5),
        ('🔍 Quality Checks', 9, 5.5)
    ]
    
    for transform, x, y in transforms:
        criar_caixa(ax, x, y, 2.5, 0.8, transform, cor_transform)
    
    # Setas para Load
    for _, x, y in transforms:
        criar_seta(ax, x + 2.5, y + 0.4, 13, 4)
    
    # LOAD (Destinos)
    ax.text(13.5, 9, 'LOAD', fontsize=16, fontweight='bold', ha='center', color=cor_load)
    
    destinos = [
        ('📊 Power BI', 12, 4.5),
        ('📈 Dashboards', 12, 3.5),
        ('📋 Excel Reports', 12, 2.5),
        ('📧 Email Reports', 12, 1.5)
    ]
    
    for destino, x, y in destinos:
        criar_caixa(ax, x, y, 2.5, 0.8, destino, cor_load)
    
    # Título
    ax.text(7.5, 10.5, 'ARQUITETURA ETL - PIPELINE DE DADOS', 
            fontsize=18, fontweight='bold', ha='center', color='#2C3E50')
    
    # Legenda
    legend_elements = [
        mpatches.Patch(color=cor_extract, label='Extração de Dados'),
        mpatches.Patch(color=cor_transform, label='Transformação'),
        mpatches.Patch(color=cor_load, label='Carga e Visualização')
    ]
    ax.legend(handles=legend_elements, loc='lower center', ncol=3, 
              bbox_to_anchor=(0.5, -0.05), fontsize=12)
    
    ax.set_xlim(0, 15)
    ax.set_ylim(0, 11)
    ax.axis('off')
    
    plt.tight_layout()
    plt.savefig('/home/ubuntu/arquitetura_etl.png', dpi=300, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    plt.close()

def criar_dashboard_kpis():
    """Cria visualização de dashboard de KPIs"""
    # Carregar dados
    try:
        df = pd.read_excel('/home/ubuntu/dados_ficticios_1000_linhas.xlsx')
    except:
        # Criar dados fictícios se não existir
        np.random.seed(42)
        df = pd.DataFrame({
            'Valor': np.random.normal(10000, 5000, 1000),
            'Departamento': np.random.choice(['RH', 'TI', 'Financeiro', 'Compras'], 1000),
            'Data': pd.date_range('2023-01-01', periods=1000, freq='D')
        })
    
    fig = plt.figure(figsize=(16, 12))
    
    # Layout do dashboard
    gs = fig.add_gridspec(4, 4, hspace=0.3, wspace=0.3)
    
    # Cores do tema
    cores = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7', '#DDA0DD']
    
    # KPI 1: Total de Transações
    ax1 = fig.add_subplot(gs[0, 0])
    ax1.text(0.5, 0.7, f'{len(df):,}', ha='center', va='center', 
             fontsize=32, fontweight='bold', color=cores[0])
    ax1.text(0.5, 0.3, 'Total de\nTransações', ha='center', va='center', 
             fontsize=14, color='#2C3E50')
    ax1.set_xlim(0, 1)
    ax1.set_ylim(0, 1)
    ax1.axis('off')
    ax1.add_patch(plt.Rectangle((0.05, 0.05), 0.9, 0.9, fill=False, 
                               edgecolor=cores[0], linewidth=3))
    
    # KPI 2: Valor Total
    ax2 = fig.add_subplot(gs[0, 1])
    valor_total = df['Valor'].sum()
    ax2.text(0.5, 0.7, f'R$ {valor_total/1000000:.1f}M', ha='center', va='center', 
             fontsize=28, fontweight='bold', color=cores[1])
    ax2.text(0.5, 0.3, 'Valor Total\nProcessado', ha='center', va='center', 
             fontsize=14, color='#2C3E50')
    ax2.set_xlim(0, 1)
    ax2.set_ylim(0, 1)
    ax2.axis('off')
    ax2.add_patch(plt.Rectangle((0.05, 0.05), 0.9, 0.9, fill=False, 
                               edgecolor=cores[1], linewidth=3))
    
    # KPI 3: Departamentos
    ax3 = fig.add_subplot(gs[0, 2])
    n_dept = df['Departamento'].nunique()
    ax3.text(0.5, 0.7, f'{n_dept}', ha='center', va='center', 
             fontsize=32, fontweight='bold', color=cores[2])
    ax3.text(0.5, 0.3, 'Departamentos\nAtivos', ha='center', va='center', 
             fontsize=14, color='#2C3E50')
    ax3.set_xlim(0, 1)
    ax3.set_ylim(0, 1)
    ax3.axis('off')
    ax3.add_patch(plt.Rectangle((0.05, 0.05), 0.9, 0.9, fill=False, 
                               edgecolor=cores[2], linewidth=3))
    
    # KPI 4: Valor Médio
    ax4 = fig.add_subplot(gs[0, 3])
    valor_medio = df['Valor'].mean()
    ax4.text(0.5, 0.7, f'R$ {valor_medio:,.0f}', ha='center', va='center', 
             fontsize=24, fontweight='bold', color=cores[3])
    ax4.text(0.5, 0.3, 'Valor Médio\npor Transação', ha='center', va='center', 
             fontsize=14, color='#2C3E50')
    ax4.set_xlim(0, 1)
    ax4.set_ylim(0, 1)
    ax4.axis('off')
    ax4.add_patch(plt.Rectangle((0.05, 0.05), 0.9, 0.9, fill=False, 
                               edgecolor=cores[3], linewidth=3))
    
    # Gráfico 1: Distribuição por Departamento
    ax5 = fig.add_subplot(gs[1, :2])
    dept_data = df.groupby('Departamento')['Valor'].sum().sort_values(ascending=True)
    bars = ax5.barh(dept_data.index, dept_data.values, color=cores[:len(dept_data)])
    ax5.set_title('Distribuição de Valores por Departamento', fontsize=14, fontweight='bold')
    ax5.set_xlabel('Valor Total (R$)')
    
    # Adicionar valores nas barras
    for i, bar in enumerate(bars):
        width = bar.get_width()
        ax5.text(width + width*0.01, bar.get_y() + bar.get_height()/2, 
                f'R$ {width/1000:.0f}K', ha='left', va='center', fontsize=10)
    
    # Gráfico 2: Tendência Temporal
    ax6 = fig.add_subplot(gs[1, 2:])
    df['Data'] = pd.to_datetime(df['Data'])
    temporal = df.groupby(df['Data'].dt.to_period('M'))['Valor'].sum()
    ax6.plot(range(len(temporal)), temporal.values, marker='o', linewidth=3, 
             markersize=8, color=cores[4])
    ax6.set_title('Tendência Temporal de Valores', fontsize=14, fontweight='bold')
    ax6.set_xlabel('Período (Meses)')
    ax6.set_ylabel('Valor Total (R$)')
    ax6.grid(True, alpha=0.3)
    
    # Gráfico 3: Distribuição de Valores
    ax7 = fig.add_subplot(gs[2, :2])
    ax7.hist(df['Valor'], bins=30, alpha=0.7, color=cores[5], edgecolor='black')
    ax7.set_title('Distribuição de Valores das Transações', fontsize=14, fontweight='bold')
    ax7.set_xlabel('Valor (R$)')
    ax7.set_ylabel('Frequência')
    ax7.grid(True, alpha=0.3)
    
    # Gráfico 4: Pizza por Departamento
    ax8 = fig.add_subplot(gs[2, 2:])
    dept_counts = df['Departamento'].value_counts()
    wedges, texts, autotexts = ax8.pie(dept_counts.values, labels=dept_counts.index, 
                                      autopct='%1.1f%%', colors=cores[:len(dept_counts)])
    ax8.set_title('Distribuição de Transações por Departamento', fontsize=14, fontweight='bold')
    
    # Estatísticas resumo
    ax9 = fig.add_subplot(gs[3, :])
    stats_text = f"""
    📊 ESTATÍSTICAS RESUMO DO DATASET
    
    • Período dos Dados: {df['Data'].min().strftime('%d/%m/%Y')} a {df['Data'].max().strftime('%d/%m/%Y')}
    • Valor Mínimo: R$ {df['Valor'].min():,.2f}  |  Valor Máximo: R$ {df['Valor'].max():,.2f}
    • Desvio Padrão: R$ {df['Valor'].std():,.2f}  |  Mediana: R$ {df['Valor'].median():,.2f}
    • Departamento com Maior Volume: {dept_data.index[-1]} (R$ {dept_data.iloc[-1]:,.0f})
    """
    
    ax9.text(0.5, 0.5, stats_text, ha='center', va='center', fontsize=12,
             bbox=dict(boxstyle="round,pad=0.5", facecolor='lightgray', alpha=0.8))
    ax9.set_xlim(0, 1)
    ax9.set_ylim(0, 1)
    ax9.axis('off')
    
    # Título principal
    fig.suptitle('🎯 DASHBOARD DE ANÁLISE DE DADOS - VISÃO EXECUTIVA', 
                fontsize=20, fontweight='bold', y=0.98)
    
    plt.savefig('/home/ubuntu/dashboard_kpis.png', dpi=300, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    plt.close()

def criar_fluxograma_processo():
    """Cria fluxograma do processo de automação"""
    fig, ax = plt.subplots(figsize=(14, 10))
    
    # Cores
    cores = {
        'inicio': '#2ECC71',
        'processo': '#3498DB', 
        'decisao': '#F39C12',
        'fim': '#E74C3C',
        'dados': '#9B59B6'
    }
    
    def criar_elemento(ax, x, y, width, height, text, tipo, color):
        if tipo == 'decisao':
            # Losango para decisão
            diamond = mpatches.FancyBboxPatch((x, y), width, height,
                                            boxstyle="round,pad=0.02",
                                            facecolor=color, edgecolor='white', linewidth=2)
        else:
            # Retângulo para outros
            diamond = mpatches.FancyBboxPatch((x, y), width, height,
                                            boxstyle="round,pad=0.05",
                                            facecolor=color, edgecolor='white', linewidth=2)
        ax.add_patch(diamond)
        ax.text(x + width/2, y + height/2, text, ha='center', va='center',
                fontsize=10, fontweight='bold', color='white', wrap=True)
    
    def criar_seta(ax, x1, y1, x2, y2, texto=''):
        ax.annotate('', xy=(x2, y2), xytext=(x1, y1),
                   arrowprops=dict(arrowstyle='->', lw=2, color='#2C3E50'))
        if texto:
            mid_x, mid_y = (x1 + x2) / 2, (y1 + y2) / 2
            ax.text(mid_x + 0.2, mid_y, texto, fontsize=9, color='#2C3E50', fontweight='bold')
    
    # Elementos do fluxograma
    elementos = [
        (2, 9, 2.5, 0.8, '🚀 INÍCIO\nCarregar Dados', 'inicio', cores['inicio']),
        (2, 7.5, 2.5, 0.8, '📊 EXTRACT\nFontes de Dados', 'processo', cores['processo']),
        (2, 6, 2.5, 0.8, '✅ Dados\nVálidos?', 'decisao', cores['decisao']),
        (6, 6, 2.5, 0.8, '🧹 TRANSFORM\nLimpar Dados', 'processo', cores['processo']),
        (10, 6, 2.5, 0.8, '📈 LOAD\nGerar Relatórios', 'processo', cores['processo']),
        (6, 4, 2.5, 0.8, '📊 Criar\nVisualizações', 'processo', cores['dados']),
        (10, 4, 2.5, 0.8, '📧 Enviar\nRelatórios', 'processo', cores['processo']),
        (6, 2, 2.5, 0.8, '🎯 FIM\nProcesso Concluído', 'fim', cores['fim']),
        (2, 4, 2.5, 0.8, '⚠️ Corrigir\nErros', 'processo', cores['fim'])
    ]
    
    for x, y, w, h, text, tipo, color in elementos:
        criar_elemento(ax, x, y, w, h, text, tipo, color)
    
    # Setas do fluxo
    setas = [
        (3.25, 9, 3.25, 8.3),  # Início -> Extract
        (3.25, 7.5, 3.25, 6.8),  # Extract -> Validação
        (4.5, 6.4, 6, 6.4, 'SIM'),  # Validação -> Transform
        (3.25, 6, 3.25, 4.8, 'NÃO'),  # Validação -> Corrigir
        (2, 4.4, 2, 5.6),  # Corrigir -> Validação (volta)
        (8.5, 6.4, 10, 6.4),  # Transform -> Load
        (7.25, 6, 7.25, 4.8),  # Transform -> Visualizações
        (11.25, 6, 11.25, 4.8),  # Load -> Enviar
        (8.5, 4.4, 10, 4.4),  # Visualizações -> Enviar
        (11.25, 4, 7.25, 2.8),  # Enviar -> Fim
    ]
    
    for seta in setas:
        if len(seta) == 5:
            criar_seta(ax, seta[0], seta[1], seta[2], seta[3], seta[4])
        else:
            criar_seta(ax, seta[0], seta[1], seta[2], seta[3])
    
    # Título
    ax.text(7, 10.5, '🔄 FLUXOGRAMA DO PROCESSO DE AUTOMAÇÃO ETL', 
            fontsize=16, fontweight='bold', ha='center', color='#2C3E50')
    
    # Legenda
    legend_elements = [
        mpatches.Patch(color=cores['inicio'], label='Início/Fim'),
        mpatches.Patch(color=cores['processo'], label='Processo'),
        mpatches.Patch(color=cores['decisao'], label='Decisão'),
        mpatches.Patch(color=cores['dados'], label='Dados/Visualização')
    ]
    ax.legend(handles=legend_elements, loc='upper right', fontsize=10)
    
    ax.set_xlim(0, 14)
    ax.set_ylim(1, 11)
    ax.axis('off')
    
    plt.tight_layout()
    plt.savefig('/home/ubuntu/fluxograma_processo.png', dpi=300, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    plt.close()

def criar_comparativo_ferramentas():
    """Cria gráfico comparativo de ferramentas"""
    fig, ax = plt.subplots(figsize=(12, 8))
    
    # Dados das ferramentas
    ferramentas = ['Excel + VBA', 'Python + Pandas', 'Power BI', 'SQL + ETL', 'Power Query']
    categorias = ['Facilidade de Uso', 'Performance', 'Escalabilidade', 'Flexibilidade', 'Custo-Benefício']
    
    # Scores (0-10)
    scores = {
        'Excel + VBA': [9, 6, 4, 7, 9],
        'Python + Pandas': [6, 9, 9, 10, 8],
        'Power BI': [8, 7, 7, 6, 7],
        'SQL + ETL': [5, 10, 10, 8, 6],
        'Power Query': [9, 7, 6, 7, 9]
    }
    
    # Configurar gráfico radar
    angles = np.linspace(0, 2 * np.pi, len(categorias), endpoint=False).tolist()
    angles += angles[:1]  # Fechar o círculo
    
    cores_ferramentas = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7']
    
    for i, (ferramenta, cor) in enumerate(zip(ferramentas, cores_ferramentas)):
        valores = scores[ferramenta]
        valores += valores[:1]  # Fechar o círculo
        
        ax.plot(angles, valores, 'o-', linewidth=2, label=ferramenta, color=cor)
        ax.fill(angles, valores, alpha=0.25, color=cor)
    
    # Configurar eixos
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(categorias, fontsize=12)
    ax.set_ylim(0, 10)
    ax.set_yticks(range(0, 11, 2))
    ax.grid(True)
    
    # Título e legenda
    ax.set_title('🔧 COMPARATIVO DE FERRAMENTAS DE AUTOMAÇÃO', 
                fontsize=16, fontweight='bold', pad=20)
    ax.legend(loc='upper right', bbox_to_anchor=(1.3, 1.0), fontsize=10)
    
    plt.tight_layout()
    plt.savefig('/home/ubuntu/comparativo_ferramentas.png', dpi=300, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    plt.close()

def criar_timeline_implementacao():
    """Cria timeline de implementação do projeto"""
    fig, ax = plt.subplots(figsize=(14, 8))
    
    # Dados do timeline
    fases = [
        ('Semana 1-2', 'Análise e Planejamento', '📋', '#3498DB'),
        ('Semana 3-4', 'Setup do Ambiente', '⚙️', '#2ECC71'),
        ('Semana 5-6', 'Desenvolvimento ETL', '🔄', '#E74C3C'),
        ('Semana 7-8', 'Criação de Dashboards', '📊', '#F39C12'),
        ('Semana 9-10', 'Testes e Validação', '✅', '#9B59B6'),
        ('Semana 11-12', 'Deploy e Treinamento', '🚀', '#1ABC9C')
    ]
    
    y_positions = range(len(fases))
    
    # Criar barras horizontais
    for i, (periodo, atividade, icone, cor) in enumerate(fases):
        # Barra principal
        ax.barh(i, 2, left=i*2, height=0.6, color=cor, alpha=0.8, edgecolor='white', linewidth=2)
        
        # Texto da atividade
        ax.text(i*2 + 1, i, f'{icone} {atividade}', ha='center', va='center',
                fontsize=11, fontweight='bold', color='white')
        
        # Período
        ax.text(i*2 + 1, i - 0.4, periodo, ha='center', va='center',
                fontsize=9, color='#2C3E50', fontweight='bold')
    
    # Setas conectoras
    for i in range(len(fases) - 1):
        ax.annotate('', xy=(i*2 + 2.1, i + 1), xytext=(i*2 + 1.9, i),
                   arrowprops=dict(arrowstyle='->', lw=2, color='#2C3E50'))
    
    # Configurações do gráfico
    ax.set_xlim(-0.5, len(fases) * 2 - 0.5)
    ax.set_ylim(-0.8, len(fases) - 0.2)
    ax.set_yticks([])
    ax.set_xticks([])
    
    # Título
    ax.text(len(fases), len(fases) - 0.5, '📅 TIMELINE DE IMPLEMENTAÇÃO DO PROJETO', 
            fontsize=16, fontweight='bold', ha='center', color='#2C3E50')
    
    # Remover bordas
    for spine in ax.spines.values():
        spine.set_visible(False)
    
    plt.tight_layout()
    plt.savefig('/home/ubuntu/timeline_implementacao.png', dpi=300, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    plt.close()

def main():
    """Executa todas as funções de criação de visualizações"""
    print("🎨 Criando visualizações para o README...")
    
    criar_banner_projeto()
    print("✅ Banner do projeto criado")
    
    criar_arquitetura_etl()
    print("✅ Diagrama de arquitetura ETL criado")
    
    criar_dashboard_kpis()
    print("✅ Dashboard de KPIs criado")
    
    criar_fluxograma_processo()
    print("✅ Fluxograma do processo criado")
    
    criar_comparativo_ferramentas()
    print("✅ Comparativo de ferramentas criado")
    
    criar_timeline_implementacao()
    print("✅ Timeline de implementação criado")
    
    print("\n🎉 Todas as visualizações foram criadas com sucesso!")
    print("📁 Arquivos salvos:")
    print("   • banner_projeto.png")
    print("   • arquitetura_etl.png") 
    print("   • dashboard_kpis.png")
    print("   • fluxograma_processo.png")
    print("   • comparativo_ferramentas.png")
    print("   • timeline_implementacao.png")

if __name__ == "__main__":
    main()

