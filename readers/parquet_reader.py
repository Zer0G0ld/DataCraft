"""
Leitor de arquivos Parquet
"""

import pandas as pd
from pathlib import Path
from .base_reader import BaseReader

class ParquetReader(BaseReader):
    """Leitor para arquivos Parquet (formato columnar de alta performance)"""
    
    def read(self, file_path: str, **kwargs) -> pd.DataFrame:
        """
        Lê arquivo Parquet
        
        Args:
            file_path: Caminho do arquivo .parquet
            **kwargs:
                - columns: Lista de colunas específicas para carregar
                - filters: Filtros para aplicação (formato pyarrow)
                - nrows: Número máximo de linhas para ler
        """
        columns = kwargs.get('columns')
        filters = kwargs.get('filters')
        nrows = kwargs.get('nrows')
        
        # Usa pyarrow engine para melhor performance
        try:
            # Tenta com engine='pyarrow' (mais rápido)
            df = pd.read_parquet(
                file_path,
                engine='pyarrow',
                columns=columns,
                filters=filters
            )
        except:
            # Fallback para engine='fastparquet'
            df = pd.read_parquet(
                file_path,
                engine='fastparquet',
                columns=columns
            )
        
        # Limita número de linhas se especificado
        if nrows and nrows < len(df):
            df = df.head(nrows)
        
        return df
    
    def read_partitioned(self, directory_path: str, **kwargs) -> pd.DataFrame:
        """
        Lê um dataset Parquet particionado em diretório
        
        Args:
            directory_path: Diretório raiz com partições Parquet
            **kwargs: Mesmas opções do método read
        """
        import pyarrow.parquet as pq
        
        dataset = pq.ParquetDataset(directory_path)
        table = dataset.read()
        df = table.to_pandas()
        
        # Aplica filtros adicionais
        if kwargs.get('columns'):
            df = df[kwargs['columns']]
        
        if kwargs.get('nrows') and kwargs['nrows'] < len(df):
            df = df.head(kwargs['nrows'])
        
        return df
    
    def get_metadata(self, file_path: str) -> dict:
        """Retorna metadados do arquivo Parquet"""
        import pyarrow.parquet as pq
        
        parquet_file = pq.ParquetFile(file_path)
        metadata = parquet_file.metadata
        
        return {
            'num_rows': metadata.num_rows,
            'num_columns': metadata.num_columns,
            'num_row_groups': metadata.num_row_groups,
            'serialized_size': metadata.serialized_size,
            'schema': str(metadata.schema),
            'row_groups': [
                {
                    'num_rows': rg.num_rows,
                    'total_byte_size': rg.total_byte_size,
                    'num_columns': rg.num_columns
                }
                for rg in parquet_file.row_groups
            ]
        }
    
    def get_schema(self, file_path: str) -> pd.DataFrame:
        """Retorna o schema do arquivo Parquet"""
        df = self.read(file_path, nrows=0)
        schema_df = pd.DataFrame({
            'column': df.columns,
            'dtype': df.dtypes.astype(str),
            'nullable': df.isnull().any()
        })
        return schema_df
    
    def get_supported_extensions(self) -> list:
        return ['.parquet', '.pq']