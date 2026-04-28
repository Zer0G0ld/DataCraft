"""
Leitor de arquivos SQL e bancos de dados
"""

import pandas as pd
from sqlalchemy import create_engine, text, inspect
from pathlib import Path
from .base_reader import BaseReader

class SQLReader(BaseReader):
    """Leitor para bancos de dados SQL (SQLite, PostgreSQL, MySQL, etc)"""
    
    def read(self, file_path: str, **kwargs) -> pd.DataFrame:
        """
        Lê dados de um banco SQL
        
        Suporta:
        - Arquivos SQLite (.db, .sqlite, .sqlite3)
        - Strings de conexão para outros bancos
        
        Args:
            file_path: Caminho para arquivo SQLite ou string de conexão
            **kwargs: 
                - table_name: Nome da tabela a ser lida (obrigatório para alguns casos)
                - query: Query SQL customizada (opcional)
                - schema: Esquema do banco (opcional)
        """
        table_name = kwargs.get('table_name')
        query = kwargs.get('query')
        schema = kwargs.get('schema')
        
        # Determina se é arquivo SQLite ou string de conexão
        if Path(file_path).exists() and file_path.endswith(('.db', '.sqlite', '.sqlite3')):
            connection_string = f'sqlite:///{file_path}'
        elif file_path.startswith(('postgresql://', 'mysql://', 'sqlite://', 'mssql://')):
            connection_string = file_path
        else:
            # Tenta como SQLite mesmo sem extensão
            connection_string = f'sqlite:///{file_path}'
        
        # Cria engine de conexão
        engine = create_engine(connection_string)
        
        try:
            with engine.connect() as conn:
                # Se tem query customizada, executa diretamente
                if query:
                    df = pd.read_sql_query(query, conn)
                    return df
                
                # Se tem nome de tabela, lê ela
                if table_name:
                    df = pd.read_sql_table(table_name, conn, schema=schema)
                    return df
                
                # Tenta encontrar a primeira tabela com dados
                inspector = inspect(engine)
                tables = inspector.get_table_names(schema=schema)
                
                if not tables:
                    raise ValueError("Nenhuma tabela encontrada no banco de dados!")
                
                # Pega a primeira tabela não-vazia
                for table in tables:
                    try:
                        df = pd.read_sql_table(table, conn, schema=schema)
                        if len(df) > 0:
                            self._last_table = table
                            return df
                    except:
                        continue
                
                # Se todas estavam vazias, pega a primeira
                df = pd.read_sql_table(tables[0], conn, schema=schema)
                return df
                
        finally:
            engine.dispose()
    
    def get_tables(self, file_path: str, schema: str = None) -> list:
        """Retorna lista de tabelas no banco"""
        if Path(file_path).exists() and file_path.endswith(('.db', '.sqlite', '.sqlite3')):
            connection_string = f'sqlite:///{file_path}'
        else:
            connection_string = file_path
        
        engine = create_engine(connection_string)
        try:
            inspector = inspect(engine)
            return inspector.get_table_names(schema=schema)
        finally:
            engine.dispose()
    
    def get_table_info(self, file_path: str, table_name: str, schema: str = None) -> dict:
        """Retorna informações sobre uma tabela específica"""
        if Path(file_path).exists() and file_path.endswith(('.db', '.sqlite', '.sqlite3')):
            connection_string = f'sqlite:///{file_path}'
        else:
            connection_string = file_path
        
        engine = create_engine(connection_string)
        try:
            inspector = inspect(engine)
            columns = inspector.get_columns(table_name, schema=schema)
            
            # Conta registros
            with engine.connect() as conn:
                result = conn.execute(text(f'SELECT COUNT(*) FROM {table_name}'))
                row_count = result.scalar()
            
            return {
                'name': table_name,
                'columns': len(columns),
                'rows': row_count,
                'column_details': columns
            }
        finally:
            engine.dispose()
    
    def execute_query(self, file_path: str, query: str) -> pd.DataFrame:
        """Executa uma query SQL customizada"""
        if Path(file_path).exists() and file_path.endswith(('.db', '.sqlite', '.sqlite3')):
            connection_string = f'sqlite:///{file_path}'
        else:
            connection_string = file_path
        
        engine = create_engine(connection_string)
        try:
            with engine.connect() as conn:
                return pd.read_sql_query(query, conn)
        finally:
            engine.dispose()
    
    def get_supported_extensions(self) -> list:
        return ['.db', '.sqlite', '.sqlite3', '.sql']