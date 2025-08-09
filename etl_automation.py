"""
Script de Automação ETL (Extract, Transform, Load)
Demonstra as melhores práticas para processamento de dados
Autor: Sistema de Automação de Dados
Data: 2025
"""

import pandas as pd
import numpy as np
import logging
from datetime import datetime, timedelta
import os
import json
from pathlib import Path

# Configuração de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('etl_log.log'),
        logging.StreamHandler()
    ]
)

class ETLProcessor:
    """Classe principal para processamento ETL"""
    
    def __init__(self, config_file=None):
        self.config = self.load_config(config_file)
        self.data = None
        self.processed_data = None
        
    def load_config(self, config_file):
        """Carrega configurações do ETL"""
        default_config = {
            "input_file": "dados_ficticios_1000_linhas.xlsx",
            "output_file": "dados_processados.xlsx",
            "sheet_name": "Dados_Principais",
            "date_columns": ["Data"],
            "numeric_columns": ["Valor", "Valor_Acumulado", "Media_Mensal_Dept"],
            "categorical_columns": ["Departamento", "Categoria", "Status"],
            "validation_rules": {
                "valor_min": 0,
                "valor_max": 100000,
                "required_columns": ["ID", "Data", "Departamento", "Valor"]
            }
        }
        
        if config_file and os.path.exists(config_file):
            with open(config_file, 'r') as f:
                user_config = json.load(f)
                default_config.update(user_config)
        
        return default_config
    
    def extract(self):
        """Fase de Extração - Carrega dados de diferentes fontes"""
        try:
            logging.info("Iniciando fase de EXTRAÇÃO...")
            
            # Verificar se arquivo existe
            if not os.path.exists(self.config["input_file"]):
                raise FileNotFoundError(f"Arquivo não encontrado: {self.config['input_file']}")
            
            # Carregar dados do Excel
            self.data = pd.read_excel(
                self.config["input_file"], 
                sheet_name=self.config["sheet_name"]
            )
            
            logging.info(f"Dados extraídos com sucesso: {len(self.data)} registros")
            logging.info(f"Colunas encontradas: {list(self.data.columns)}")
            
            return True
            
        except Exception as e:
            logging.error(f"Erro na extração: {str(e)}")
            return False
    
    def validate_data(self):
        """Validação inicial dos dados"""
        try:
            logging.info("Iniciando VALIDAÇÃO dos dados...")
            
            # Verificar colunas obrigatórias
            required_cols = self.config["validation_rules"]["required_columns"]
            missing_cols = [col for col in required_cols if col not in self.data.columns]
            
            if missing_cols:
                raise ValueError(f"Colunas obrigatórias ausentes: {missing_cols}")
            
            # Verificar dados vazios
            null_counts = self.data.isnull().sum()
            if null_counts.sum() > 0:
                logging.warning(f"Valores nulos encontrados:\n{null_counts[null_counts > 0]}")
            
            # Validar valores numéricos
            if "Valor" in self.data.columns:
                valor_min = self.config["validation_rules"]["valor_min"]
                valor_max = self.config["validation_rules"]["valor_max"]
                
                invalid_values = self.data[
                    (self.data["Valor"] < valor_min) | 
                    (self.data["Valor"] > valor_max)
                ]
                
                if len(invalid_values) > 0:
                    logging.warning(f"Valores fora do intervalo válido: {len(invalid_values)} registros")
            
            logging.info("Validação concluída com sucesso")
            return True
            
        except Exception as e:
            logging.error(f"Erro na validação: {str(e)}")
            return False
    
    def transform(self):
        """Fase de Transformação - Limpa e processa os dados"""
        try:
            logging.info("Iniciando fase de TRANSFORMAÇÃO...")
            
            # Criar cópia dos dados para transformação
            self.processed_data = self.data.copy()
            
            # 1. Limpeza de dados
            self._clean_data()
            
            # 2. Conversão de tipos
            self._convert_data_types()
            
            # 3. Criação de colunas derivadas
            self._create_derived_columns()
            
            # 4. Agregações e cálculos
            self._calculate_aggregations()
            
            # 5. Padronização de categorias
            self._standardize_categories()
            
            logging.info("Transformação concluída com sucesso")
            return True
            
        except Exception as e:
            logging.error(f"Erro na transformação: {str(e)}")
            return False
    
    def _clean_data(self):
        """Limpeza dos dados"""
        logging.info("Executando limpeza de dados...")
        
        # Remover duplicatas
        initial_count = len(self.processed_data)
        self.processed_data = self.processed_data.drop_duplicates()
        removed_duplicates = initial_count - len(self.processed_data)
        
        if removed_duplicates > 0:
            logging.info(f"Removidas {removed_duplicates} duplicatas")
        
        # Tratar valores nulos
        for col in self.processed_data.columns:
            if self.processed_data[col].dtype == 'object':
                self.processed_data[col] = self.processed_data[col].fillna('Não Informado')
            elif self.processed_data[col].dtype in ['int64', 'float64']:
                self.processed_data[col] = self.processed_data[col].fillna(0)
        
        # Limpar strings
        string_columns = self.processed_data.select_dtypes(include=['object']).columns
        for col in string_columns:
            self.processed_data[col] = self.processed_data[col].astype(str).str.strip()
    
    def _convert_data_types(self):
        """Conversão de tipos de dados"""
        logging.info("Convertendo tipos de dados...")
        
        # Converter datas
        for col in self.config["date_columns"]:
            if col in self.processed_data.columns:
                self.processed_data[col] = pd.to_datetime(self.processed_data[col])
        
        # Converter numéricos
        for col in self.config["numeric_columns"]:
            if col in self.processed_data.columns:
                self.processed_data[col] = pd.to_numeric(self.processed_data[col], errors='coerce')
        
        # Converter categóricos
        for col in self.config["categorical_columns"]:
            if col in self.processed_data.columns:
                self.processed_data[col] = self.processed_data[col].astype('category')
    
    def _create_derived_columns(self):
        """Criação de colunas derivadas"""
        logging.info("Criando colunas derivadas...")
        
        if "Data" in self.processed_data.columns:
            # Extrair componentes de data
            self.processed_data['Dia_Semana'] = self.processed_data['Data'].dt.day_name()
            self.processed_data['Semana_Ano'] = self.processed_data['Data'].dt.isocalendar().week
            self.processed_data['Dias_Desde_Hoje'] = (datetime.now() - self.processed_data['Data']).dt.days
        
        if "Valor" in self.processed_data.columns:
            # Categorizar valores
            self.processed_data['Faixa_Valor'] = pd.cut(
                self.processed_data['Valor'],
                bins=[0, 1000, 5000, 15000, float('inf')],
                labels=['Baixo', 'Médio', 'Alto', 'Muito Alto']
            )
            
            # Calcular percentis
            self.processed_data['Percentil_Valor'] = self.processed_data['Valor'].rank(pct=True)
    
    def _calculate_aggregations(self):
        """Cálculos e agregações"""
        logging.info("Calculando agregações...")
        
        # Estatísticas por departamento
        dept_stats = self.processed_data.groupby('Departamento')['Valor'].agg([
            'count', 'sum', 'mean', 'std', 'min', 'max'
        ]).round(2)
        
        # Merge com dados principais
        dept_stats.columns = [f'Dept_{col}' for col in dept_stats.columns]
        self.processed_data = self.processed_data.merge(
            dept_stats, 
            left_on='Departamento', 
            right_index=True, 
            how='left'
        )
        
        # Ranking por valor
        self.processed_data['Ranking_Valor'] = self.processed_data['Valor'].rank(
            method='dense', ascending=False
        )
    
    def _standardize_categories(self):
        """Padronização de categorias"""
        logging.info("Padronizando categorias...")
        
        # Padronizar status
        status_mapping = {
            'aprovado': 'Aprovado',
            'APROVADO': 'Aprovado',
            'pendente': 'Pendente',
            'PENDENTE': 'Pendente',
            'rejeitado': 'Rejeitado',
            'REJEITADO': 'Rejeitado'
        }
        
        if 'Status' in self.processed_data.columns:
            self.processed_data['Status'] = self.processed_data['Status'].replace(status_mapping)
    
    def load(self):
        """Fase de Carga - Salva dados processados"""
        try:
            logging.info("Iniciando fase de CARGA...")
            
            # Criar diretório de saída se não existir
            output_path = Path(self.config["output_file"])
            output_path.parent.mkdir(parents=True, exist_ok=True)
            
            # Salvar em múltiplas abas
            with pd.ExcelWriter(self.config["output_file"], engine='openpyxl') as writer:
                # Dados principais processados
                self.processed_data.to_excel(writer, sheet_name='Dados_Processados', index=False)
                
                # Resumo executivo
                self._create_executive_summary().to_excel(writer, sheet_name='Resumo_Executivo')
                
                # Dados por departamento
                dept_summary = self.processed_data.groupby('Departamento').agg({
                    'Valor': ['count', 'sum', 'mean'],
                    'ID': 'count'
                }).round(2)
                dept_summary.to_excel(writer, sheet_name='Resumo_Departamento')
                
                # Dados mensais
                if 'Data' in self.processed_data.columns:
                    monthly_summary = self.processed_data.groupby([
                        self.processed_data['Data'].dt.year,
                        self.processed_data['Data'].dt.month
                    ]).agg({
                        'Valor': ['count', 'sum', 'mean'],
                        'ID': 'count'
                    }).round(2)
                    monthly_summary.to_excel(writer, sheet_name='Resumo_Mensal')
            
            # Salvar também em CSV para compatibilidade
            csv_file = self.config["output_file"].replace('.xlsx', '.csv')
            self.processed_data.to_csv(csv_file, index=False, encoding='utf-8-sig')
            
            logging.info(f"Dados salvos com sucesso em: {self.config['output_file']}")
            logging.info(f"Arquivo CSV criado: {csv_file}")
            
            return True
            
        except Exception as e:
            logging.error(f"Erro na carga: {str(e)}")
            return False
    
    def _create_executive_summary(self):
        """Cria resumo executivo dos dados"""
        summary_data = {
            'Métrica': [
                'Total de Registros',
                'Período dos Dados',
                'Total de Departamentos',
                'Valor Total',
                'Valor Médio',
                'Maior Transação',
                'Menor Transação',
                'Departamento com Maior Volume',
                'Status Mais Comum'
            ],
            'Valor': [
                len(self.processed_data),
                f"{self.processed_data['Data'].min().strftime('%Y-%m-%d')} a {self.processed_data['Data'].max().strftime('%Y-%m-%d')}",
                self.processed_data['Departamento'].nunique(),
                f"R$ {self.processed_data['Valor'].sum():,.2f}",
                f"R$ {self.processed_data['Valor'].mean():,.2f}",
                f"R$ {self.processed_data['Valor'].max():,.2f}",
                f"R$ {self.processed_data['Valor'].min():,.2f}",
                self.processed_data.groupby('Departamento')['Valor'].sum().idxmax(),
                self.processed_data['Status'].mode().iloc[0]
            ]
        }
        
        return pd.DataFrame(summary_data)
    
    def run_etl(self):
        """Executa o pipeline ETL completo"""
        logging.info("=== INICIANDO PIPELINE ETL ===")
        
        # Extração
        if not self.extract():
            return False
        
        # Validação
        if not self.validate_data():
            return False
        
        # Transformação
        if not self.transform():
            return False
        
        # Carga
        if not self.load():
            return False
        
        logging.info("=== PIPELINE ETL CONCLUÍDO COM SUCESSO ===")
        return True
    
    def generate_data_quality_report(self):
        """Gera relatório de qualidade dos dados"""
        if self.processed_data is None:
            logging.error("Dados não processados. Execute o ETL primeiro.")
            return None
        
        quality_report = {
            'total_records': len(self.processed_data),
            'columns': list(self.processed_data.columns),
            'data_types': self.processed_data.dtypes.to_dict(),
            'null_values': self.processed_data.isnull().sum().to_dict(),
            'duplicate_records': self.processed_data.duplicated().sum(),
            'memory_usage': self.processed_data.memory_usage(deep=True).sum(),
            'numeric_summary': self.processed_data.describe().to_dict()
        }
        
        # Salvar relatório
        with open('data_quality_report.json', 'w', encoding='utf-8') as f:
            json.dump(quality_report, f, indent=2, ensure_ascii=False, default=str)
        
        logging.info("Relatório de qualidade salvo em: data_quality_report.json")
        return quality_report

# Exemplo de uso
if __name__ == "__main__":
    # Criar instância do processador ETL
    etl = ETLProcessor()
    
    # Executar pipeline completo
    success = etl.run_etl()
    
    if success:
        # Gerar relatório de qualidade
        etl.generate_data_quality_report()
        
        print("\n=== RESUMO DO PROCESSAMENTO ===")
        print(f"Registros processados: {len(etl.processed_data)}")
        print(f"Colunas finais: {len(etl.processed_data.columns)}")
        print(f"Arquivo de saída: {etl.config['output_file']}")
        
        # Mostrar primeiras linhas dos dados processados
        print("\nPrimeiras 5 linhas dos dados processados:")
        print(etl.processed_data.head())
    else:
        print("Erro no processamento ETL. Verifique os logs.")

