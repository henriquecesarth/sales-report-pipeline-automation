import os
import sys
import time

# Garante que o diretório raiz está no path para evitar problemas de importação
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.readers import consolidate_data
from src.processors import process_sales_data
from src.reporters import generate_excel_report

# Caminhos padrão do projeto
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RAW_DIR = os.path.join(BASE_DIR, "data", "raw")
OUTPUT_DIR = os.path.join(BASE_DIR, "data", "output")
OUTPUT_PATH = os.path.join(OUTPUT_DIR, "relatorio_executivo.xlsx")

def main():
    print("=" * 60)
    print("  PIPELINE DE AUTOMAÇÃO DE RELATÓRIOS EXECUTIVOS DE VENDAS  ")
    print("=" * 60)
    start_time = time.time()
    
    # Validação inicial de pastas
    if not os.path.exists(RAW_DIR):
        print(f"[ERRO] Pasta de planilhas de entrada '{RAW_DIR}' não encontrada.")
        print("Crie a pasta e adicione os arquivos de vendas nela antes de executar.")
        sys.exit(1)
        
    try:
        # PASSO 1: Consolidação Inteligente
        print("\n[Etapa 1/3] Consolidando arquivos de dados de 'data/raw'...")
        df_consolidated = consolidate_data(RAW_DIR)
        
        # PASSO 2: Limpeza e Cruzamento com Pandas
        print("\n[Etapa 2/3] Executando regras de negócio e cálculo de métricas...")
        df_clean, metrics = process_sales_data(df_consolidated)
        
        # PASSO 3: Geração do Excel Estilizado
        print("\n[Etapa 3/3] Formatando e gerando o relatório final estilizado...")
        generate_excel_report(df_clean, OUTPUT_PATH)
        
        # Exibição de Resumo de Execução no Console
        elapsed_time = time.time() - start_time
        print("\n" + "=" * 60)
        print("               RESUMO DO RELATÓRIO PROCESSADO               ")
        print("=" * 60)
        print(f"  Faturamento Total:          R$ {metrics['faturamento_total']:,.2f}")
        print(f"  Ticket Médio por Cliente:   R$ {metrics['ticket_medio_cliente']:,.2f}")
        print(f"  Produto Mais Vendido:       {metrics['produto_mais_vendido']} ({metrics['qtd_produto_mais_vendido']} unidades)")
        print(f"  Clientes Únicos Atendidos:  {metrics['total_clientes_unicos']}")
        print(f"  Total de Vendas Processadas: {metrics['total_vendas_processadas']}")
        print("-" * 60)
        print(f"  Tempo de processamento:     {elapsed_time:.2f} segundos")
        print(f"  Relatório Executivo Salvo em: {OUTPUT_PATH}")
        print("=" * 60 + "\n")
        
    except FileNotFoundError as fnf:
        print(f"\n[ERRO DE ARQUIVO] {fnf}")
    except Exception as e:
        print(f"\n[ERRO DURANTE O PIPELINE] Ocorreu uma falha na execução: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
