"""
main.py — Ponto de entrada do gerador de relatório gerencial de obras.

Uso:
    python main.py
    python main.py --dados meu_projeto.json --saida RelatorioMarco26.xlsx
"""

import argparse
import os
import sys
from datetime import datetime

def parse_args():
    parser = argparse.ArgumentParser(
        description="Gerador de Relatório Gerencial Mensal - Obras"
    )
    parser.add_argument(
        "--dados",
        default="dados_relatorio.json",
        help="Caminho para o arquivo JSON com os dados do projeto "
             "(padrão: dados_relatorio.json)",
    )
    parser.add_argument(
        "--saida",
        default=None,
        help="Caminho/nome do arquivo Excel de saída. "
             "Se omitido, gera automaticamente com data/hora.",
    )
    return parser.parse_args()


def main():
    args = parse_args()

    # Resolve caminhos relativos ao diretório do script
    base_dir = os.path.dirname(os.path.abspath(__file__))
    caminho_dados = os.path.join(base_dir, args.dados)

    if not os.path.exists(caminho_dados):
        print(f"[ERRO] Arquivo de dados não encontrado: {caminho_dados}")
        sys.exit(1)

    if args.saida:
        caminho_saida = args.saida
        if not os.path.isabs(caminho_saida):
            caminho_saida = os.path.join(base_dir, caminho_saida)
    else:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        caminho_saida = os.path.join(base_dir, f"Relatorio_Gerencial_{timestamp}.xlsx")

    # Importa após resolver caminhos para evitar erro de import em execução direta
    try:
        from gerador_relatorio import GeradorRelatorio
    except ImportError:
        # Permite rodar de qualquer diretório
        sys.path.insert(0, base_dir)
        from gerador_relatorio import GeradorRelatorio

    gerador = GeradorRelatorio(caminho_dados)
    gerador.gerar(caminho_saida)


if __name__ == "__main__":
    main()
