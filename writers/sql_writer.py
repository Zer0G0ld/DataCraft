"""
Escritor para bancos de dados SQL
"""

import pandas as pd
from sqlalchemy import create_engine, text
from pathlib import Path
from .base_writer import BaseWriter

class SQLWriter(BaseWriter):
    """Escritor para bancos de dados SQL"""
    
    def write(self, df: pd.DataFrame, output_path: str, **kwargs) -> str:
        """
        Escreve DataFrame para SQL
        
        Args:
            df: DataFrame a ser escrito
            output_path: Caminho para arquivo SQLite ou string de conexão
            **kwargs:
                - table_name: Nome da tabela (padrão: 'data_table')
                - if_exists: 'replace', 'append', 'fail' (padrão: 'replace')
                - index: Incluir índice como coluna (padrão: False)
                - chunksize: Linhas por lote (padrão: None)
                - create_sql_file: Criar arquivo .sql complementar (padrão: True)
        """
        table_name = kwargs.get('table_name', 'data_table')
        if_exists = kwargs.get('if_exists', 'replace')
        index = kwargs.get('index', False)
        chunksize = kwargs.get('chunksize')
        create_sql_file = kwargs.get('create_sql_file', True)
        
        # Determina tipo de conexão
        if output_path.endswith(('.db', '.sqlite', '.sqlite3')):
            connection_string = f'sqlite:///{output_path}'
        elif output_path.startswith(('postgresql://', 'mysql://', 'sqlite://', 'mssql://')):
            connection_string = output_path
        else:
            # Assume SQLite com extensão .db
            if not output_path.endswith('.db'):
                output_path = f"{output_path}.db"
            connection_string = f'sqlite:///{output_path}'
        
        # Cria engine
        engine = create_engine(connection_string)
        
        try:
            # Escreve tabela
            df.to_sql(
                table_name,
                engine,
                if_exists=if_exists,
                index=index,
                chunksize=chunksize
            )
            
            # Cria arquivo .sql complementar
            if create_sql_file:
                sql_file = output_path.replace('.db', '.sql').replace('.sqlite', '.sql')
                self._create_sql_script(df, table_name, sql_file)
            
            # Se for SQLite, retorna o caminho do arquivo
            if 'sqlite' in connection_string:
                return output_path
            else:
                return connection_string
                
        finally:
            engine.dispose()
    
    def _create_sql_script(self, df: pd.DataFrame, table_name: str, sql_file: str):
        """Cria um arquivo .sql com comandos INSERT"""
        with open(sql_file, 'w', encoding='utf-8') as f:
            f.write(f"-- DataCraft SQL Export\n")
            f.write(f"-- Gerado em: {pd.Timestamp.now()}\n")
            f.write(f"-- Tabela: {table_name}\n")
            f.write(f"-- Registros: {len(df)}\n\n")
            
            # CREATE TABLE (simplificado)
            columns_def = []
            for col in df.columns:
                dtype = df[col].dtype
                if 'int' in str(dtype):
                    sql_type = 'INTEGER'
                elif 'float' in str(dtype):
                    sql_type = 'REAL'
                elif 'datetime' in str(dtype):
                    sql_type = 'TIMESTAMP'
                else:
                    sql_type = 'TEXT'
                columns_def.append(f'    "{col}" {sql_type}')
            
            f.write(f"CREATE TABLE IF NOT EXISTS {table_name} (\n")
            f.write(",\n".join(columns_def))
            f.write("\n);\n\n")
            
            # INSERTs (limitado a 1000 para não estourar)
            for idx, row in df.head(1000).iterrows():
                columns = ', '.join([f'"{col}"' for col in df.columns])
                values = []
                for col in df.columns:
                    val = row[col]
                    if pd.isna(val):
                        values.append('NULL')
                    elif isinstance(val, str):
                        # Escapa aspas simples
                        val_escaped = val.replace("'", "''")
                        values.append(f"'{val_escaped}'")
                    elif isinstance(val, (pd.Timestamp, datetime)):
                        values.append(f"'{val}'")
                    else:
                        values.append(str(val))
                
                values_str = ', '.join(values)
                f.write(f"INSERT INTO {table_name} ({columns}) VALUES ({values_str});\n")
            
            if len(df) > 1000:
                f.write(f"\n-- Nota: Apenas os primeiros 1000 registros foram incluídos neste script SQL\n")
                f.write(f"-- Para os {len(df) - 1000} registros restantes, use o banco de dados diretamente.\n")
    
    def write_multiple_tables(self, tables: dict, output_path: str, **kwargs) -> list:
        """
        Escreve múltiplas tabelas no mesmo banco
        
        Args:
            tables: Dicionário {table_name: dataframe}
            output_path: Caminho do banco
        """
        written = []
        
        for table_name, df in tables.items():
            result = self.write(df, output_path, table_name=table_name, **kwargs)
            written.append(result)
        
        return written
    
    def get_supported_formats(self) -> list:
        return ['sql', 'db', 'sqlite', 'sqlite3']