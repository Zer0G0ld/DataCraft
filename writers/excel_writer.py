"""
Escritor para arquivos Excel
"""

import pandas as pd
from openpyxl.utils import get_column_letter
from .base_writer import BaseWriter

class ExcelWriter(BaseWriter):
    """Escritor especializado para Excel"""
    
    def write(self, df: pd.DataFrame, output_path: str, **kwargs) -> str:
        sheet_name = kwargs.get('sheet_name', 'Dados_Completos')
        create_summary = kwargs.get('create_summary', True)
        auto_width = kwargs.get('auto_width', True)
        
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Aba principal
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # Aba resumida
            if create_summary and len(df) > 0:
                summary_df = self._create_summary(df)
                summary_df.to_excel(writer, sheet_name='Resumo', index=False)
            
            # Ajusta largura das colunas
            if auto_width:
                self._auto_adjust_width(writer, sheet_name)
        
        return output_path
    
    def _create_summary(self, df: pd.DataFrame) -> pd.DataFrame:
        """Cria uma aba resumida com estatísticas"""
        summary_data = {
            'Métrica': [],
            'Valor': []
        }
        
        # Informações básicas
        summary_data['Métrica'].append('Total de Registros')
        summary_data['Valor'].append(len(df))
        
        summary_data['Métrica'].append('Total de Colunas')
        summary_data['Valor'].append(len(df.columns))
        
        # Estatísticas para colunas numéricas
        numeric_cols = df.select_dtypes(include=['number']).columns
        for col in numeric_cols[:5]:  # Limita a 5 colunas
            summary_data['Métrica'].append(f'Média - {col}')
            summary_data['Valor'].append(round(df[col].mean(), 2))
            
            summary_data['Métrica'].append(f'Total - {col}')
            summary_data['Valor'].append(df[col].sum())
        
        # Contagem de valores nulos
        null_counts = df.isnull().sum()
        for col in null_counts[null_counts > 0].index[:5]:
            summary_data['Métrica'].append(f'Nulos em {col}')
            summary_data['Valor'].append(null_counts[col])
        
        return pd.DataFrame(summary_data)
    
    def _auto_adjust_width(self, writer, sheet_name):
        """Ajusta automaticamente a largura das colunas"""
        worksheet = writer.sheets[sheet_name]
        for column in worksheet.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width
    
    def get_supported_formats(self) -> list:
        return ['excel', 'xlsx', 'xls']