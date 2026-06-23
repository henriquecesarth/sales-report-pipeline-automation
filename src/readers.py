import os
import glob
import pandas as pd
from typing import List, Dict, Any

# Mapeamento de termos comuns para colunas canônicas (todas em minúsculo e sem acentos/caracteres especiais)
CANONICAL_MAPPING = {
    'id_venda': ['id_venda', 'id venda', 'idvenda', 'venda_id', 'venda id', 'sale_id', 'sale id', 'codigo_venda', 'cod_venda', 'codigo venda', 'cod venda'],
    'data': ['data', 'data_venda', 'data venda', 'date', 'sale_date', 'data_do_pedido', 'data do pedido', 'data_venda', 'data_pedido'],
    'canal': ['canal', 'canal_venda', 'canal de venda', 'canal de vendas', 'channel', 'origem', 'tipo_venda', 'tipo venda', 'canal_de_venda'],
    'produto': ['produto', 'nome_produto', 'nome produto', 'product', 'item', 'descrição', 'descricao', 'descriçao_produto', 'descricao_produto'],
    'quantidade': ['quantidade', 'qtd', 'quantity', 'quant', 'quant.', 'quantidades', 'unidades'],
    'valor_unitario': ['valor_unitario', 'valor unitario', 'valor_unitário', 'valor unitário', 'preco_unitario', 'preco unitario', 'preço unitário', 'preço_unitário', 'unit_price', 'price', 'valor', 'valor_uni', 'valor uni', 'preço_uni', 'preco_uni'],
    'id_cliente': ['id_cliente', 'id cliente', 'idcliente', 'customer_id', 'customer id', 'cliente_id', 'cliente id', 'cpf_cliente', 'cod_cliente', 'codigo_cliente', 'codigo cliente']
}

def normalize_text(text: str) -> str:
    """Remove espaços extras, converte para minúsculas e remove acentos básicos para mapeamento."""
    if not isinstance(text, str):
        return ""
    
    text = text.strip().lower()
    
    # Substituições simples de acentuação comum em cabeçalhos
    replacements = {
        'á': 'a', 'à': 'a', 'ã': 'a', 'â': 'a',
        'é': 'e', 'ê': 'e',
        'í': 'i',
        'ó': 'o', 'ô': 'o', 'õ': 'o',
        'ú': 'u', 'ü': 'u',
        'ç': 'c',
        ' ': '_', '-': '_', '.': ''
    }
    for char, replacement in replacements.items():
        text = text.replace(char, replacement)
    
    # Remove múltiplos underscores consecutivos
    while '__' in text:
        text = text.replace('__', '_')
        
    return text.strip('_')

def map_column_names(df: pd.DataFrame) -> pd.DataFrame:
    """Normaliza e mapeia os nomes de colunas de um DataFrame para os nomes canônicos."""
    rename_dict = {}
    
    for col in df.columns:
        norm_col = normalize_text(str(col))
        mapped = False
        for canonical, aliases in CANONICAL_MAPPING.items():
            # Verifica se o nome normalizado coincide com o nome canônico ou seus aliases normalizados
            normalized_aliases = [normalize_text(alias) for alias in aliases]
            if norm_col == canonical or norm_col in normalized_aliases:
                rename_dict[col] = canonical
                mapped = True
                break
        if not mapped:
            # Se não mapear, mantém o nome normalizado básico
            rename_dict[col] = norm_col
            
    return df.rename(columns=rename_dict)

def read_file(file_path: str) -> pd.DataFrame:
    """
    Lê um arquivo de dados (CSV ou Excel) de forma inteligente e resiliente.
    Trata diferentes encodings e delimitadores em arquivos CSV.
    """
    ext = os.path.splitext(file_path)[1].lower()
    df = pd.DataFrame()
    
    if ext == '.csv':
        # Tenta ler com detecção automática de separador e tratamento de encoding
        encodings = ['utf-8', 'utf-8-sig', 'latin-1', 'cp1252']
        for encoding in encodings:
            try:
                # Primeiro tenta ler com sep=None para detectar delimitadores automaticamente (como , ou ;)
                # O motor 'python' é necessário para detecção automática, mas se falhar tentamos o padrão
                df = pd.read_csv(file_path, sep=None, encoding=encoding, engine='python')
                
                # Se o DataFrame tiver apenas uma coluna, pode indicar que a detecção falhou
                if len(df.columns) <= 1:
                    # Tenta forçar os separadores mais comuns
                    for sep in [',', ';', '\t']:
                        df_temp = pd.read_csv(file_path, sep=sep, encoding=encoding)
                        if len(df_temp.columns) > len(df.columns):
                            df = df_temp
                break # Se leu sem erro de encoding, interrompe o loop
            except (UnicodeDecodeError, pd.errors.ParserError):
                continue
        else:
            raise ValueError(f"Não foi possível ler o arquivo CSV '{file_path}' com os encodings padrão.")
            
    elif ext in ['.xlsx', '.xls']:
        # Lê arquivos Excel (lê apenas a primeira aba por padrão)
        try:
            df = pd.read_excel(file_path, sheet_name=0)
        except Exception as e:
            raise ValueError(f"Erro ao ler o arquivo Excel '{file_path}': {e}")
    else:
        raise ValueError(f"Extensão de arquivo não suportada: {ext}")
        
    if df.empty:
        raise ValueError(f"O arquivo '{file_path}' está vazio ou não pôde ser lido.")
        
    return df

def scan_raw_folder(raw_dir: str) -> List[str]:
    """Escaneia a pasta raw em busca de arquivos CSV e Excel."""
    file_patterns = ['*.csv', '*.xlsx', '*.xls']
    files = []
    for pattern in file_patterns:
        files.extend(glob.glob(os.path.join(raw_dir, pattern)))
    return sorted(files)

def consolidate_data(raw_dir: str) -> pd.DataFrame:
    """
    Escaneia a pasta data/raw/, lê todos os arquivos de forma automatizada,
    padroniza as colunas de cada um e os consolida em um único DataFrame estruturado.
    """
    files = scan_raw_folder(raw_dir)
    if not files:
        raise FileNotFoundError(f"Nenhum arquivo de dados (CSV ou Excel) foi encontrado na pasta '{raw_dir}'.")
        
    print(f"Encontrados {len(files)} arquivos para processamento na pasta '{raw_dir}'.")
    
    dfs_to_concat = []
    
    for file_path in files:
        file_name = os.path.basename(file_path)
        print(f"-> Lendo arquivo: {file_name}")
        try:
            # Lê o arquivo
            df = read_file(file_path)
            # Mapeia colunas para os nomes canônicos
            df_normalized = map_column_names(df)
            # Adiciona metadados de origem da venda
            df_normalized['arquivo_origem'] = file_name
            dfs_to_concat.append(df_normalized)
        except Exception as e:
            print(f"   [AVISO] Falha ao processar arquivo '{file_name}': {e}. Pulando arquivo.")
            
    if not dfs_to_concat:
        raise ValueError("Nenhum arquivo pôde ser processado com sucesso para a consolidação.")
        
    # Consolida os DataFrames em um único
    # O Pandas cuida de alinhar as colunas com base nos nomes mapeados
    consolidated_df = pd.concat(dfs_to_concat, ignore_index=True)
    
    print(f"Consolidação concluída. Base consolidada possui {consolidated_df.shape[0]} linhas e {consolidated_df.shape[1]} colunas.")
    return consolidated_df
