"""
Faz login no sistema de gestão (Central de Gestão), vai até a tela Cadastrar
de "Vistorias > Banco B > Em Aberto" e envia o Excel exportado da
Plataforma na seção "PLATAFORMA" - cria as propostas novas.

Adaptado do robô irmão que já faz esse mesmo cadastro pro Banco A
(robo-cadastro-banco-a) - mesma mecânica de login/sessão/envio.

Os seletores são os do BANCO B, confirmados no HTML real da tela
(mapear_central_gestao.py, 09/09/2026) - e vários são diferentes dos do
Banco A, apesar de ser o mesmo sistema:

    campo de arquivo     Banco A #excelFilePlataforma  -> Banco B #excelFile
    campo oculto (JSON)  Banco A #dadosExcelPlataforma -> Banco B #dadosExcel
    formulário           Banco A #excelFormPlataforma  -> Banco B action /storeExcel
    aba Cadastrar        #tab2-tab (igual nos dois)

Usar os nomes do Banco A fazia o robô esperar 5 minutos por um elemento
inexistente e desistir em silêncio, sem enviar nada.

Duas camadas de deduplicação:
1. A própria Central de Gestão identifica pelo número da proposta e ignora
   as que já existem ao importar (confirmado em teste real no fluxo do
   Banco A - reenviar não duplica). Essa é a garantia de verdade, sempre
   vale mesmo sem o item 2 abaixo.
2. Antes de enviar, este script também remove do Excel as propostas que
   ELE MESMO já enviou com sucesso antes (log local, ver
   PROPOSTAS_ENVIADAS_FILE) - não evita duplicata (o item 1 já evita),
   só evita reenviar à toa e mostra quantas propostas são realmente
   novas antes de pedir confirmação.

Antes de enviar de verdade, mostra quantas linhas tem no Excel (já
descontadas as já enviadas antes) e pede confirmação explícita (digitar
CONFIRMAR) - trava de segurança enquanto este fluxo (recém adaptado pro
Banco B) ainda está em teste.

Uso: python cadastrar_central_gestao.py [caminho\\para\\arquivo.xlsx]
Se não passar o caminho, usa o Excel mais recente em data/exports/.

Login: usa sessão salva (mesmo arquivo do mapeador - se você já logou
rodando mapear_central_gestao.py, essa sessão já vale aqui). Se expirar,
pausa e pede login manual uma vez (precisa rodar com HEADLESS = False
nesse caso).
"""

import os
import sys
from datetime import datetime

from openpyxl import load_workbook
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout

from comum import (
    abrir_pagina,
    Tee,
    modo_automatico,
    pausar_para_usuario,
    pedir_confirmacao,
    headless_configurado,
    contar_linhas_excel,
    achar_coluna_por_cabecalho,
    achar_excel_mais_recente,
    espera_carregamento_gestao,
    bloquear_recursos_visuais,
    abrir_gestao,
    credenciais_configuradas,
    login_automatico,
)

LOGIN_URL = "https://central-gestao.example.com/login"
# confirmado em execução real: a URL segue o mesmo padrão do Banco A
# (/gestao-banco-a -> /gestao-banco-b).
GESTAO_BANCO_B_URL = "https://central-gestao.example.com/gestao-banco-b"
# se a URL acima não responder, o robô procura no menu do sistema um
# link cujo href contenha esta palavra (ver comum.abrir_gestao) - assim
# uma eventual mudança de URL não trava a etapa.
PALAVRA_CHAVE_CLIENTE = "banco_b"
PASTA_SCRIPT = os.path.dirname(os.path.abspath(__file__))
EXPORTS_DIR = os.path.join(PASTA_SCRIPT, "data", "exports")
# cópia filtrada que é enviada pro sistema, gravada FORA de data/exports
# de propósito: o export da Etapa 1 nunca pode ser alterado (as Etapas 3
# e 4 leem o mesmo arquivo depois), e um .xlsx dentro de data/exports
# viraria "o Excel mais recente" pra elas. Ver filtrar_propostas_ja_enviadas
# (o robô irmão do Banco A sobrescrevia o próprio export aqui até 17/09/2026
# e isso fazia as etapas seguintes enxergarem só as propostas da última
# rodada - mesmo risco existia aqui, corrigido antes de acontecer).
ENVIOS_DIR = os.path.join(PASTA_SCRIPT, "data", "envios")
# mesmo arquivo de sessão do mapeador - login feito lá já vale aqui.
SESSION_FILE = os.path.join(PASTA_SCRIPT, "sessao_central_gestao.json")
LOG_DIR = os.path.join(PASTA_SCRIPT, "logs")
# log local (fora do git) das propostas que este robô já mandou com
# sucesso pra Em Aberto - usado só pra não reenviar à toa, não é a
# garantia de duplicidade (essa é do próprio sistema, ver docstring).
PROPOSTAS_ENVIADAS_FILE = os.path.join(PASTA_SCRIPT, "propostas_ja_enviadas_em_aberto.txt")
NOME_COLUNA_PROPOSTA = "proposta"

# Seletores da tela Cadastrar, confirmados no HTML real do Banco B
# (mapear_central_gestao.py, 09/09/2026). NÃO são os mesmos do Banco A: lá o
# campo é #excelFilePlataforma e o oculto é #dadosExcelPlataforma; aqui
# são #excelFile e #dadosExcel. Usar os nomes do Banco A fazia o robô
# esperar 5 minutos por um elemento que não existe e desistir em
# silêncio, sem enviar nada.
ABA_CADASTRAR = "#tab2-tab"
CAMPO_ARQUIVO_PLATAFORMA = "#excelFile"
CAMPO_DADOS_PLATAFORMA = "#dadosExcel"
# a página tem DOIS formulários com id="excelForm" (Plataforma e Cetip) -
# id duplicado é HTML inválido, mas é o que está no ar. Por isso o
# formulário é escolhido pela action, que é única: /storeExcel (o do
# Cetip é /storeExcelCetip, não casa com o $=). Escolher por #excelForm
# poderia mandar o Excel da Plataforma pro formulário do Cetip.
FORM_PLATAFORMA = "form[action$='/storeExcel']"
# a página de Em Aberto "parece" carregada bem antes de estar pronta de
# verdade (JS/jQuery ainda inicializando por trás) - confirmado no
# fluxo irmão do Banco A (09/09/2026): clicar cedo demais trava. Pausa
# fixa, não um timeout de espera por elemento (esse já passa antes).
# Ajustável sem editar o código, ver comum.espera_carregamento_gestao.
ESPERA_CARREGAMENTO_GESTAO = espera_carregamento_gestao()

# pra ver a janela e logar manualmente, rode com a variável de ambiente
# HEADLESS=false - não edite esta linha (ver comum.headless_configurado).
HEADLESS = headless_configurado()


def novo_contexto_pagina(browser, com_sessao):
    context = browser.new_context(
        accept_downloads=True,
        viewport={"width": 1920, "height": 1080} if HEADLESS else None,
        no_viewport=None if HEADLESS else True,
        storage_state=SESSION_FILE if com_sessao else None,
    )
    if HEADLESS:
        bloquear_recursos_visuais(context)
    return context, context.new_page()


def ler_propostas_ja_enviadas():
    """Devolve o conjunto de números de proposta que este robô já
    mandou com sucesso em execuções anteriores (conjunto vazio se o
    arquivo ainda não existe - primeira execução)."""
    if not os.path.exists(PROPOSTAS_ENVIADAS_FILE):
        return set()
    with open(PROPOSTAS_ENVIADAS_FILE, "r", encoding="utf-8") as f:
        return {linha.strip() for linha in f if linha.strip()}


def salvar_propostas_enviadas(propostas):
    """Acrescenta números de proposta ao log local (nunca sobrescreve o
    que já tinha)."""
    if not propostas:
        return
    with open(PROPOSTAS_ENVIADAS_FILE, "a", encoding="utf-8") as f:
        for proposta in sorted(propostas):
            f.write(proposta + "\n")


def extrair_numeros_proposta(caminho_arquivo):
    """Lê todos os números de proposta presentes no Excel (todas as
    abas), achando a coluna pelo cabeçalho - usado tanto pra filtrar
    quanto, depois de um envio confirmado, pra atualizar o log local."""
    if os.path.splitext(caminho_arquivo)[1].lower() != ".xlsx":
        return set()
    wb = load_workbook(caminho_arquivo, read_only=True, data_only=True)
    propostas = set()
    for nome_aba in wb.sheetnames:
        ws = wb[nome_aba]
        coluna = achar_coluna_por_cabecalho(ws, NOME_COLUNA_PROPOSTA)
        if coluna is None:
            continue
        for linha in ws.iter_rows(min_row=2):
            valor = linha[coluna - 1].value
            if valor not in (None, ""):
                propostas.add(str(valor).strip())
    return propostas


def filtrar_propostas_ja_enviadas(caminho_arquivo, propostas_ja_enviadas):
    """Gera uma CÓPIA do Excel sem as linhas cuja coluna de Nº de Proposta
    já esteja em `propostas_ja_enviadas`. Devolve
    (caminho_da_copia, total_antes, total_removidas), ou None se não deu
    pra filtrar (sem propostas conhecidas ainda, ou arquivo não é .xlsx).

    IMPORTANTE - o arquivo original NUNCA é tocado. A cópia vai pra
    data/envios (fora de data/exports) justamente pra não ser confundida
    com "o Excel mais recente" pelas etapas seguintes (que leem o mesmo
    export original depois)."""
    if not propostas_ja_enviadas or os.path.splitext(caminho_arquivo)[1].lower() != ".xlsx":
        return None

    wb = load_workbook(caminho_arquivo)
    total_antes = 0
    total_removidas = 0
    achou_coluna = False
    for nome_aba in wb.sheetnames:
        ws = wb[nome_aba]
        coluna = achar_coluna_por_cabecalho(ws, NOME_COLUNA_PROPOSTA)
        if coluna is None:
            continue
        achou_coluna = True
        linhas_para_remover = []
        for linha in range(2, ws.max_row + 1):
            valor = str(ws.cell(row=linha, column=coluna).value or "").strip()
            if not valor:
                continue
            total_antes += 1
            if valor in propostas_ja_enviadas:
                linhas_para_remover.append(linha)
        for linha in reversed(linhas_para_remover):
            ws.delete_rows(linha)
            total_removidas += 1
    if not achou_coluna:
        print(f"      [AVISO] Não achei uma coluna com '{NOME_COLUNA_PROPOSTA}' no cabeçalho - "
              "pulei a deduplicação local (o próprio sistema ainda vai ignorar repetidas ao importar).")
        return None

    os.makedirs(ENVIOS_DIR, exist_ok=True)
    nome = os.path.splitext(os.path.basename(caminho_arquivo))[0]
    caminho_envio = os.path.join(ENVIOS_DIR, f"envio_{nome}_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
    wb.save(caminho_envio)
    return caminho_envio, total_antes, total_removidas


def abas_que_nao_serao_enviadas(caminho_arquivo):
    """Devolve as abas do Excel que a página vai ignorar (todas menos a
    primeira) - lista vazia quando só existe uma aba.

    O JS da própria tela Cadastrar faz:
        let firstSheet = workbook.SheetNames[0];
        XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet]);
    ou seja, só a PRIMEIRA aba vira JSON e sobe. As demais somem sem
    nenhum aviso na tela. No export do Banco B isso não aparece (vem
    uma aba só, depois que a Etapa 1 remove a de metadados), mas o
    export do Banco A já veio com mais de uma aba de dados - se um dia
    acontecer aqui, as linhas das outras abas seriam perdidas em
    silêncio. Melhor avisar antes de enviar do que descobrir depois."""
    if os.path.splitext(caminho_arquivo)[1].lower() != ".xlsx":
        return []
    try:
        wb = load_workbook(caminho_arquivo, read_only=True, data_only=True)
    except Exception:
        return []
    return wb.sheetnames[1:]


def enviar_excel_plataforma(page, caminho_arquivo):
    """Na tela Cadastrar (aba #tab2-tab) da gestão Banco B, seleciona o
    Excel na seção Plataforma e clica em Enviar."""
    page.locator(ABA_CADASTRAR).click()
    # a aba Cadastrar (e a tela de Em Aberto em geral) já demorou muito
    # pra ficar pronta em execuções reais no fluxo irmão do Banco A (mesmo
    # motivo do timeout=300000 no goto pra /gestao-banco-b, abaixo) -
    # por isso timeout bem folgado aqui em vez de um valor curto que
    # trava à toa num dia mais lento.
    page.wait_for_selector(CAMPO_ARQUIVO_PLATAFORMA, state="visible", timeout=300000)

    page.locator(CAMPO_ARQUIVO_PLATAFORMA).set_input_files(caminho_arquivo)
    print(f"      Arquivo selecionado: {caminho_arquivo}")

    # o JS da própria página lê o arquivo, converte pra JSON e joga no
    # campo oculto do formulário - o Enviar só faz sentido depois disso.
    try:
        page.wait_for_function(
            f"document.querySelector('{CAMPO_DADOS_PLATAFORMA}').value.length > 0",
            timeout=60000,
        )
    except PlaywrightTimeout:
        raise RuntimeError(
            "Selecionei o arquivo mas a página não terminou de processá-lo "
            f"(campo '{CAMPO_DADOS_PLATAFORMA}' continuou vazio). Rode com "
            "HEADLESS=false pra ver se apareceu algum erro na tela."
        )

    with page.expect_navigation(timeout=120000):
        page.locator(f"{FORM_PLATAFORMA} button[type=submit]").click()


def _tentar_login_automatico(page):
    """Tenta logar sozinho com as credenciais de ambiente (CENTRAL_USUARIO /
    CENTRAL_SENHA). Só devolve True se realmente SAIU da tela de login -
    nunca confia em "o clique não deu erro". Sem credenciais
    configuradas, nem tenta."""
    usuario, senha = credenciais_configuradas("CENTRAL")
    if not usuario:
        return False
    print("      Sessão expirada - tentando login automático...")
    if not login_automatico(page, usuario, senha):
        print("      [AVISO] Não achei o formulário de login pra preencher sozinho.")
        return False
    try:
        page.wait_for_url(lambda url: "/login" not in url, timeout=30000)
    except PlaywrightTimeout:
        pass
    if "/login" in page.url:
        print("      [AVISO] Preenchi o login mas continuo na tela de login "
              "(credencial errada, ou o sistema pediu verificação extra).")
        return False
    return True


def main():
    os.makedirs(LOG_DIR, exist_ok=True)

    log_path = os.path.join(LOG_DIR, f"cadastro_central_{datetime.now():%Y%m%d_%H%M%S}.txt")
    log_file = open(log_path, "w", encoding="utf-8")
    stdout_original = sys.stdout
    sys.stdout = Tee(stdout_original, log_file)

    try:
        print("=" * 60)
        print(" CADASTRO NO SISTEMA DE GESTÃO - CENTRAL DE GESTÃO (BANCO B)")
        print("=" * 60)
        print(f"Log desta execução: {log_path}")

        caminho_arquivo = sys.argv[1] if len(sys.argv) > 1 else achar_excel_mais_recente(EXPORTS_DIR)
        if not caminho_arquivo or not os.path.exists(caminho_arquivo):
            print("\n[ERRO] Não achei nenhum Excel pra enviar.")
            print("Exporte primeiro com exportar_status_banco_b.py, ou passe o caminho:")
            print(f"  python {os.path.basename(__file__)} caminho\\para\\arquivo.xlsx")
            return
        print(f"Arquivo a cadastrar: {caminho_arquivo}")

        propostas_ja_enviadas = ler_propostas_ja_enviadas()
        # IGNORAR_LOG_LOCAL=1 pula a deduplicação local de propósito -
        # usado pelo teste de reenvio (preparar_teste_reenvio.py), que
        # PRECISA mandar uma proposta que o log já conhece pra ver o que
        # o sistema faz com ela. Fora desse teste, não use.
        if os.environ.get("IGNORAR_LOG_LOCAL", "").strip() == "1":
            print("      [IGNORAR_LOG_LOCAL=1] Deduplicação local desligada nesta execução.")
            propostas_ja_enviadas = set()
        # o que sobe pro sistema é a cópia filtrada; o export original
        # fica intacto porque as Etapas 3, 4 e 5 vão ler ele depois.
        caminho_envio = caminho_arquivo
        resultado_filtro = filtrar_propostas_ja_enviadas(caminho_arquivo, propostas_ja_enviadas)
        if resultado_filtro is not None:
            caminho_envio, total_antes, total_removidas = resultado_filtro
            print(f"      {total_antes} proposta(s) no Excel, {total_removidas} já enviada(s) antes por "
                  f"este robô (removida(s) do envio), {total_antes - total_removidas} nova(s) a enviar.")
            if total_antes - total_removidas == 0:
                print("\nNada de novo pra enviar - todas as propostas do Excel já foram enviadas antes.")
                print("(O export continua intacto - as Etapas 3, 4 e 5 seguem normalmente.)")
                return

        total_linhas = contar_linhas_excel(caminho_envio)
        aviso_linhas = f"{total_linhas} linha(s) de dados no Excel." if total_linhas is not None else "Não consegui contar as linhas (arquivo não é .xlsx?)."

        avisos_extras = []
        abas_ignoradas = abas_que_nao_serao_enviadas(caminho_envio)
        if abas_ignoradas:
            avisos_extras.append(
                f"[ATENÇÃO] A página só importa a PRIMEIRA aba do Excel - estas NÃO vão subir: "
                f"{', '.join(abas_ignoradas)}"
            )

        confirmado = pedir_confirmacao(
            f"Arquivo: {caminho_envio}",
            aviso_linhas,
            *avisos_extras,
            "Isso vai subir esse Excel DE VERDADE pro sistema de gestão",
            "(Vistorias > Banco B > Em Aberto).",
        )
        if not confirmado:
            print("\nEnvio cancelado pelo usuário - nada foi enviado.")
            return

        concluido = False

        with sync_playwright() as p:
            browser = p.chromium.launch(
                headless=HEADLESS,
                args=[] if HEADLESS else ["--start-maximized"],
            )
            sessao_existente = os.path.exists(SESSION_FILE)
            context, page = novo_contexto_pagina(browser, com_sessao=sessao_existente)

            def salvar_sessao():
                context.storage_state(path=SESSION_FILE)

            try:
                print("\n[1/3] Verificando sessão salva...")
                # importante: sempre passar pela tela de LOGIN primeiro, nunca
                # ir direto pra /gestao-banco-b sem sessão - esse sistema não
                # redireciona pra login quando não autenticado, ele quebra
                # com um erro de servidor (Auth::user() nulo).
                abrir_pagina(page, LOGIN_URL)
                if sessao_existente and "/login" not in page.url:
                    print("      Sessão válida, painel carregado direto.")
                elif _tentar_login_automatico(page):
                    salvar_sessao()
                    print("      Login automático OK - sessão renovada e salva.")
                elif HEADLESS or modo_automatico():
                    print("\n[ERRO] A sessão salva expirou e não consegui logar sozinho.")
                    print("Pra rodar sem ninguém na frente, configure as credenciais uma vez:")
                    print('  [Environment]::SetEnvironmentVariable("CENTRAL_USUARIO", "voce@exemplo.com", "User")')
                    print('  [Environment]::SetEnvironmentVariable("CENTRAL_SENHA", "suasenha", "User")')
                    print("Ou rode com a janela visível e logue na mão:")
                    print(f'  $env:HEADLESS="false"; python {os.path.basename(__file__)}')
                    return
                else:
                    pausar_para_usuario(
                        "A sessão salva não está mais válida (ou é a primeira vez).",
                        "1. Faça o login manualmente na janela do Chrome (e-mail e senha).",
                        "2. Espere a página carregar depois do login (a tela do Dashboard",
                        "   aparecer de vez) ANTES de voltar aqui e apertar ENTER.",
                    )
                    # se o ENTER veio um pouco cedo (redirect ainda em curso), dá mais
                    # alguns segundos de chance antes de desistir.
                    if "/login" in page.url:
                        try:
                            page.wait_for_url(lambda url: "/login" not in url, timeout=15000)
                        except PlaywrightTimeout:
                            pass
                    if "/login" in page.url:
                        print("\n[ERRO] Ainda estou vendo a tela de login.")
                        print("Encerrando esta execução.")
                        return
                    salvar_sessao()
                    print("      Sessão salva para as próximas execuções.")

                url_gestao = abrir_gestao(
                    page, GESTAO_BANCO_B_URL, PALAVRA_CHAVE_CLIENTE, timeout=300000
                )
                if url_gestao is None:
                    print("\n[ERRO] Não consegui chegar na tela de gestão do Banco B.")
                    print("Encerrando esta execução sem enviar nada.")
                    return
                if "/login" in page.url:
                    print("\n[ERRO] Cheguei na tela de gestão mas voltei pra tela de login.")
                    print("Encerrando esta execução.")
                    return

                # a página "aparenta" carregada bem antes de estar pronta de
                # verdade (o JS/jQuery que faz a tela funcionar ainda está
                # inicializando por trás) - confirmado no fluxo do Banco A:
                # clicar cedo demais na aba Cadastrar trava. Isso não é a
                # mesma coisa que "elemento ainda não apareceu" (timeout não
                # resolve) - por isso essa pausa fixa aqui, não um timeout.
                print(f"      Aguardando {ESPERA_CARREGAMENTO_GESTAO // 1000}s pra página terminar de carregar de verdade...")
                page.wait_for_timeout(ESPERA_CARREGAMENTO_GESTAO)

                print("[2/3] Abrindo aba Cadastrar > seção PLATAFORMA...")
                print("[3/3] Selecionando e enviando o arquivo...")
                enviar_excel_plataforma(page, caminho_envio)
                concluido = True

                print(f"\n      Página depois do envio: {page.url}")
                try:
                    resultado_pagina = os.path.join(LOG_DIR, f"resultado_cadastro_{datetime.now():%Y%m%d_%H%M%S}.png")
                    page.screenshot(path=resultado_pagina, full_page=True)
                    print(f"      Print da página de resultado salvo em: {resultado_pagina}")
                except Exception:
                    pass

            except PlaywrightTimeout as e:
                print("\n[ERRO] O robô travou esperando um elemento aparecer na tela.")
                print("Copie a mensagem abaixo e me envie, junto com o que estava")
                print("aparecendo na janela do Chrome nesse momento:")
                print(str(e))
            except Exception as e:
                print("\n[ERRO INESPERADO]")
                print(str(e))
            finally:
                try:
                    salvar_sessao()
                except Exception:
                    pass
                if not HEADLESS and not modo_automatico():
                    print("\n" + "#" * 60)
                    print("#  A AÇÃO É SUA AGORA - O ROBÔ ESTÁ PAUSADO")
                    print("#" * 60)
                    input(">>> Pressione ENTER aqui para fechar o navegador... ")
                # salva de novo aqui - se o login foi concluído durante a
                # pausa acima (ex.: depois de um erro anterior), essa é a
                # sessão mais recente que temos chance de capturar.
                try:
                    salvar_sessao()
                except Exception:
                    pass
                browser.close()

        if concluido:
            print("\nEnviado! Confira o print salvo (ou a tela, se rodou com")
            print("HEADLESS = False) pra ver se o sistema confirmou o cadastro.")
            novas_propostas = extrair_numeros_proposta(caminho_envio)
            if novas_propostas:
                salvar_propostas_enviadas(novas_propostas)
                print(f"      {len(novas_propostas)} proposta(s) registrada(s) no log local "
                      f"({os.path.basename(PROPOSTAS_ENVIADAS_FILE)}) pra não reenviar à toa nas próximas execuções.")
        else:
            print("\nNão consegui concluir o envio - veja o erro acima.")

    finally:
        sys.stdout = stdout_original
        log_file.close()


if __name__ == "__main__":
    main()
