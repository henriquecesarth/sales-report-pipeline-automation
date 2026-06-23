import pandas as pd
from typing import Dict, Any

def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    """
    Realiza a limpeza de dados da base consolidada:
    - Remove linhas completamente duplicadas.
    - Remove duplicadas com base no ID da venda.
    - Remove registros sem identificação crítica (vendas fantasmas).
    """
    df_clean = df.copy()
    
    # 1. Remove linhas completamente idênticas
    initial_rows = df_clean.shape[0]
    df_clean.drop_duplicates(inplace=True)
    duplicate_rows = initial_rows - df_clean.shape[0]
    if duplicate_rows > 0:
        print(f"   [Limpeza] Removidas {duplicate_rows} linhas duplicadas idênticas.")
        
    # 2. Remove registros com IDs de venda duplicados (garante unicidade da venda no relatório)
    # Se uma venda possui o mesmo ID, mantemos apenas o primeiro registro
    initial_rows = df_clean.shape[0]
    df_clean.drop_duplicates(subset=['id_venda'], keep='first', inplace=True)
    duplicate_ids = initial_rows - df_clean.shape[0]
    if duplicate_ids > 0:
        print(f"   [Limpeza] Removidas {duplicate_ids} vendas com ID duplicado.")

    # 3. Remove "vendas fantasmas" - registros sem os dados essenciais
    # (id_venda, id_cliente, produto, quantidade, valor_unitario)
    critical_columns = ['id_venda', 'id_cliente', 'produto', 'quantidade', 'valor_unitario']
    initial_rows = df_clean.shape[0]
    # Remove linhas onde qualquer uma das colunas críticas é nula
    df_clean.dropna(subset=critical_columns, how='any', inplace=True)
    ghost_sales = initial_rows - df_clean.shape[0]
    if ghost_sales > 0:
        print(f"   [Limpeza] Removidas {ghost_sales} 'vendas fantasmas' (registros com campos cruciais nulos).")
        
    # Castings e sanitizações básicas
    df_clean['id_venda'] = df_clean['id_venda'].astype(int)
    df_clean['quantidade'] = pd.to_numeric(df_clean['quantidade']).astype(int)
    df_clean['valor_unitario'] = pd.to_numeric(df_clean['valor_unitario']).astype(float)
    
    return df_clean

def standardize_channel(channel: Any) -> str:
    """Padroniza a string do canal de venda para 'Balcão' ou 'E-commerce'."""
    if not isinstance(channel, str):
        return "Outros"
    
    val = channel.strip().lower()
    
    # Mapeamento dinâmico
    if 'balc' in val or 'balcao' in val:
        return 'Balcão'
    elif 'online' in val or 'e-commerce' in val or 'ecommerce' in val or 'net' in val:
        return 'E-commerce'
    else:
        return channel.strip().title()

def standardize_data(df: pd.DataFrame) -> pd.DataFrame:
    """
    Padroniza formatos de data e campos de texto:
    - Converte a coluna 'data' para datetime (formato YYYY-MM-DD).
    - Padroniza canais de venda (Balcão / E-commerce).
    - Capitaliza nomes de produtos de forma uniforme.
    """
    df_std = df.copy()
    
    # 1. Padronizar Datas
    # format='mixed' ajuda a tratar múltiplos formatos de data sem warnings
    df_std['data'] = pd.to_datetime(df_std['data'], format='mixed', errors='coerce')
    
    # Se alguma data ficou nula após a conversão, removemos
    initial_rows = df_std.shape[0]
    df_std.dropna(subset=['data'], inplace=True)
    invalid_dates = initial_rows - df_std.shape[0]
    if invalid_dates > 0:
        print(f"   [Padronização] Removidas {invalid_dates} linhas devido a formato de data inválido.")
        
    # 2. Padronizar Canais de Venda
    df_std['canal'] = df_std['canal'].apply(standardize_channel)
    
    # 3. Padronizar nomes de Produtos (remover espaços extras e capitalizar)
    df_std['produto'] = df_std['produto'].astype(str).str.strip().str.title()
    
    # 4. Calcular Faturamento Total da linha (Quantidade * Valor Unitário)
    df_std['valor_total'] = df_std['quantidade'] * df_std['valor_unitario']
    
    # Ordenar dados por data para apresentação lógica
    df_std.sort_values(by='data', ascending=True, inplace=True)
    df_std.reset_index(drop=True, inplace=True)
    
    return df_std

def calculate_metrics(df: pd.DataFrame) -> Dict[str, Any]:
    """
    Calcula métricas de negócio:
    - Faturamento Total.
    - Ticket Médio por Cliente (Faturamento Total / Total de Clientes Únicos).
    - Produto mais vendido (em quantidade total de unidades).
    """
    if df.empty:
        return {
            "faturamento_total": 0.0,
            "ticket_medio_cliente": 0.0,
            "produto_mais_vendido": "N/A",
            "qtd_produto_mais_vendido": 0
        }
        
    # 1. Faturamento Total
    faturamento_total = float(df['valor_total'].sum())
    
    # 2. Ticket Médio por cliente
    # Calculado como: Faturamento Total / Quantidade de Clientes Únicos
    unique_clients = df['id_cliente'].nunique()
    ticket_medio = faturamento_total / unique_clients if unique_clients > 0 else 0.0
    
    # 3. Produto mais vendido (por quantidade de unidades vendidas)
    prod_sales = df.groupby('produto')['quantidade'].sum()
    if not prod_sales.empty:
        produto_mais_vendido = prod_sales.idxmax()
        qtd_produto_mais_vendido = int(prod_sales.max())
    else:
        produto_mais_vendido = "N/A"
        qtd_produto_mais_vendido = 0
        
    return {
        "faturamento_total": faturamento_total,
        "ticket_medio_cliente": ticket_medio,
        "produto_mais_vendido": produto_mais_vendido,
        "qtd_produto_mais_vendido": qtd_produto_mais_vendido,
        "total_clientes_unicos": unique_clients,
        "total_vendas_processadas": df.shape[0]
    }

def process_sales_data(df: pd.DataFrame) -> tuple[pd.DataFrame, Dict[str, Any]]:
    """
    Orquestra a limpeza, padronização e cálculo de métricas para a base consolidada.
    Retorna o DataFrame final limpo/padronizado e um dicionário com as métricas obtidas.
    """
    print("\nIniciando processamento e limpeza dos dados com Pandas...")
    
    # Etapa 1: Limpeza de duplicatas e nulos
    df_clean = clean_data(df)
    
    # Etapa 2: Padronização de datas, canais e cálculo do valor total
    df_final = standardize_data(df_clean)
    
    # Etapa 3: Cálculo das métricas de negócio
    metrics = calculate_metrics(df_final)
    
    print("Processamento concluído com sucesso!")
    print(f"-> Base final limpa contém {df_final.shape[0]} registros únicos.")
    print(f"-> Faturamento Total: R$ {metrics['faturamento_total']:,.2f}")
    print(f"-> Ticket Médio por Cliente: R$ {metrics['ticket_medio_cliente']:,.2f}")
    print(f"-> Produto mais vendido: {metrics['produto_mais_vendido']} ({metrics['qtd_produto_mais_vendido']} unidades)")
    
    return df_final, metrics
