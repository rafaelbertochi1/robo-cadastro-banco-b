"""
Monta um Excel com UMA proposta só, pra testar com segurança o que a
Central de Gestão faz ao receber uma proposta que JÁ EXISTE e já foi
FINALIZADA (ou seja, que não está mais em "Em Aberto").

Por que isso existe: a deduplicação do lado do sistema ("reenviar não
duplica") foi confirmada em teste real apenas com propostas que ainda
estavam em Em Aberto. Ninguém verificou o que acontece com uma proposta
já finalizada - as duas possibilidades são bem diferentes:

  a) o sistema reconhece o número e ignora  -> reenviar é inofensivo;
  b) o sistema não acha em Em Aberto e RECRIA a proposta lá -> reenviar
     um lote grande ressuscitaria centenas de propostas já finalizadas,
     desfazendo o trabalho da Etapa 5.

Com uma proposta só, o teste custa nada e a resposta é definitiva.

Uso:
    python preparar_teste_reenvio.py 10000013
    python preparar_teste_reenvio.py 10000013 caminho\\para\\export.xlsx

Depois, envie SÓ esse arquivo e confira na tela:
    python cadastrar_central_gestao.py data\\envios\\teste_<numero>.xlsx
"""

import os
import sys

from openpyxl import load_workbook

from comum import achar_coluna_por_cabecalho, achar_excel_mais_recente

PASTA_SCRIPT = os.path.dirname(os.path.abspath(__file__))
EXPORTS_DIR = os.path.join(PASTA_SCRIPT, "data", "exports")
ENVIOS_DIR = os.path.join(PASTA_SCRIPT, "data", "envios")
NOME_COLUNA_PROPOSTA = "proposta"


def montar_excel_de_uma_proposta(caminho_origem, num_proposta):
    """Devolve o caminho de uma cópia do export contendo só a linha dessa
    proposta (o original não é tocado), ou None se não achar."""
    wb = load_workbook(caminho_origem)
    achou = False

    for nome_aba in wb.sheetnames:
        ws = wb[nome_aba]
        coluna = achar_coluna_por_cabecalho(ws, NOME_COLUNA_PROPOSTA)
        if coluna is None:
            continue
        manter = []
        for linha in range(2, ws.max_row + 1):
            valor = str(ws.cell(row=linha, column=coluna).value or "").strip()
            if valor == num_proposta:
                manter.append(linha)
        # apaga de baixo pra cima pra não bagunçar os índices
        for linha in range(ws.max_row, 1, -1):
            if linha not in manter:
                ws.delete_rows(linha)
        if manter:
            achou = True

    if not achou:
        return None

    os.makedirs(ENVIOS_DIR, exist_ok=True)
    caminho = os.path.join(ENVIOS_DIR, f"teste_{num_proposta}.xlsx")
    wb.save(caminho)
    return caminho


def main():
    if len(sys.argv) < 2:
        print("Falta o número da proposta.")
        print(f"  python {os.path.basename(__file__)} 10000013")
        return

    num_proposta = sys.argv[1].strip()
    caminho_origem = sys.argv[2] if len(sys.argv) > 2 else achar_excel_mais_recente(EXPORTS_DIR)

    if not caminho_origem or not os.path.exists(caminho_origem):
        print("[ERRO] Não achei nenhum Excel exportado em data/exports.")
        return

    print(f"Export usado: {caminho_origem}")
    caminho = montar_excel_de_uma_proposta(caminho_origem, num_proposta)

    if caminho is None:
        print(f"[ERRO] A proposta {num_proposta} não está nesse Excel.")
        print("Escolha uma que apareça no export E que você saiba que já foi finalizada.")
        return

    print(f"\nPronto: {caminho}")
    print("\nAgora envie SÓ esse arquivo (IGNORAR_LOG_LOCAL=1 é obrigatório: sem")
    print("isso o log local filtra a proposta e nada sobe - o teste não acontece):")
    print(f"  $env:HEADLESS=\"false\"; $env:IGNORAR_LOG_LOCAL=\"1\"; python cadastrar_central_gestao.py {caminho}")
    print("\nDepois abra Vistorias > Banco B > Em Aberto e procure essa proposta:")
    print("  - NÃO apareceu  -> o sistema ignorou. Reenviar o lote todo é seguro.")
    print("  - APARECEU      -> o sistema recria as finalizadas. NÃO envie o lote;")
    print("                     me avise que eu monto a deduplicação antes.")


if __name__ == "__main__":
    main()
