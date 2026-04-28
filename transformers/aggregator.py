"""
Transformador para agregação de dados
"""

import pandas as pd
from typing import Dict, List, Any

class DataAggregator:
    """Gerencia operações de agregação em DataFrames"""
    
    def aggregate(self, df: pd.DataFrame, aggregations: List[Dict[str, Any]]) -> pd.DataFrame:
        """
        Aplica agregações ao DataFrame
        
        Args:
            df: DataFrame de entrada
            aggregations: Lista de agregações no formato:
                [
                    {
                        'group_by': ['col1', 'col2'],
                        'aggregations': {'col3': 'sum', 'col4': 'mean'}
                    }
                ]
        
        Returns:
            DataFrame agregado
        """
        result = df.copy()
        
        for agg_config in aggregations:
            group_by = agg_config.get('group_by', [])
            agg_dict = agg_config.get('aggregations', {})
            
            if group_by and agg_dict:
                result = result.groupby(group_by).agg(agg_dict).reset_index()
        
        return result
    
    def get_summary_stats(self, df: pd.DataFrame) -> pd.DataFrame:
        """Retorna estatísticas resumidas do DataFrame"""
        summary = []
        
        for col in df.columns:
            col_data = {
                'Coluna': col,
                'Tipo': str(df[col].dtype),
                'Valores Únicos': df[col].nunique(),
                'Nulos': df[col].isnull().sum(),
                'Nulos %': round(df[col].isnull().sum() / len(df) * 100, 2)
            }
            
            if pd.api.types.is_numeric_dtype(df[col]):
                col_data['Mínimo'] = df[col].min()
                col_data['Máximo'] = df[col].max()
                col_data['Média'] = round(df[col].mean(), 2)
            else:
                col_data['Mínimo'] = '-'
                col_data['Máximo'] = '-'
                col_data['Média'] = '-'
            
            summary.append(col_data)
        
        return pd.DataFrame(summary)