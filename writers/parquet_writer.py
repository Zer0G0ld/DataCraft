"""
Escritor para arquivos Parquet
"""

import pandas as pd
from pathlib import Path
from .base_writer import BaseWriter

class ParquetWriter(BaseWriter):
    """Escritor para formato Parquet (alta performance, compressão)"""
    
    def write(self, df: pd.DataFrame, output_path: str, **kwargs) -> str:
        """
        Escreve DataFrame para Parquet
        
        Args:
            df: DataFrame a ser escrito
            output_path: Caminho do arquivo .parquet
            **kwargs:
                - compression: 'snappy', 'gzip', 'brotli', 'lz4', 'zstd' (padrão: 'snappy')
                - index: Incluir índice (padrão: False)
                - partition_cols: Colunas para particionamento (opcional)
                - row_group_size: Tamanho do row group em linhas
        """
        compression = kwargs.get('compression', 'snappy')
        index = kwargs.get('index', False)
        partition_cols = kwargs.get('partition_cols')
        row_group_size = kwargs.get('row_group_size')
        
        # Garante extensão .parquet
        if not output_path.endswith('.parquet'):
            output_path = f"{output_path}.parquet"
        
        # Se tem particionamento, escreve como dataset particionado
        if partition_cols:
            # Escreve particionado
            df.to_parquet(
                output_path,
                engine='pyarrow',
                compression=compression,
                index=index,
                partition_cols=partition_cols
            )
        else:
            # Escreve arquivo único
            df.to_parquet(
                output_path,
                engine='pyarrow',
                compression=compression,
                index=index,
                row_group_size=row_group_size
            )
        
        return output_path
    
    def write_partitioned(self, df: pd.DataFrame, directory_path: str, partition_cols: list, **kwargs) -> str:
        """
        Escreve DataFrame como dataset particionado
        
        Args:
            df: DataFrame
            directory_path: Diretório raiz para o dataset
            partition_cols: Colunas para particionamento
        """
        import pyarrow as pa
        import pyarrow.parquet as pq
        
        # Converte pandas para pyarrow
        table = pa.Table.from_pandas(df)
        
        # Escreve particionado
        pq.write_to_dataset(
            table,
            root_path=directory_path,
            partition_cols=partition_cols,
            use_legacy_dataset=False
        )
        
        return directory_path
    
    def get_compression_options(self) -> dict:
        """Retorna opções de compressão disponíveis"""
        return {
            'snappy': 'Boa compressão, rápido (padrão)',
            'gzip': 'Melhor compressão, mais lento',
            'brotli': 'Alta compressão, bom para texto',
            'lz4': 'Muito rápido, compressão média',
            'zstd': 'Bom equilíbrio compressão/velocidade',
            'none': 'Sem compressão'
        }
    
    def get_supported_formats(self) -> list:
        return ['parquet', 'pq']