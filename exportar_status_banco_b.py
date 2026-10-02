"""
Entra na Plataforma (cliente Banco B), filtra o Relatório Analítico por
Tipo de Inspeção (todos menos AVM - ver TIPOS_INSPECAO) e Status de
Inspeção (ver STATUS_INSPECAO),
a partir de uma data desejada até hoje, exporta pra Excel e, antes de
terminar, abre cada inspeção Cancelada pra conferir se ela teve
vistoria de verdade (aba Laudos com pelo menos uma versão publicada) -
Cancelada sem nenhum laudo é removida do Excel final (não é erro, é
esperado: proposta cancelada antes de qualquer vistoria acontecer).

Esta é a base do fluxo de atualização de status - o Excel gerado aqui
alimenta o cadastro em "Em Aberto" feito por cadastrar_central_gestao.py.

Adaptado do robô irmão que já faz esse mesmo fluxo pro Banco A
(robo-cadastro-banco-a) - mesma base de login/sessão/navegação/filtros,
trocando o cliente selecionado na Plataforma e o nome dos arquivos
gerados. Ver "Pontos que podem precisar de ajuste fino" no README: o que
veio confirmado do robô do Banco A foi mantido igual; o que é específico
da tela do Banco B (ainda não testado contra o sistema real) está
marcado lá.

Uso: pip install -r requirements.txt && playwright install chromium
     && python exportar_status_banco_b.py

Login: usa sessão salva (sem pedir e-mail/senha). Se expirar, pausa e
pede login manual uma vez (precisa rodar com HEADLESS = False nesse caso).
"""

import os
import re
import sys
import time
from datetime import datetime, timedelta

from openpyxl import load_workbook
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout

from comum import (
    abrir_pagina,
    Tee,
    credenciais_configuradas,
    login_automatico,
    pedir_dado,
    pausar_para_usuario,
    headless_configurado,
    contar_linhas_excel,
    achar_coluna_por_cabecalho,
    modo_automatico,
    bloquear_recursos_visuais,
    chave_autenticador_configurada,
    gerar_codigo_totp,
)

# quantos dias pra trás cobrir por padrão no modo automático (sem
# ninguém pra digitar a data). JANELA_DIAS (variável de ambiente)
# sobrepõe esse padrão.
#
# ATENÇÃO - este número NÃO é "de quanto em quanto tempo eu rodo".
# O filtro da Plataforma é por DATA DE SOLICITAÇÃO da inspeção, mas a
# proposta só aparece no export depois que o laudo é aceito, o que
# demora dias. Então uma proposta solicitada no dia D só é capturada
# se ainda estiver dentro da janela no dia em que o laudo fica pronto:
# é preciso JANELA_DIAS > (dias entre solicitação e laudo aceito).
# O que sobrar fora da janela nunca mais é exportado - some de vez.
#
# Medido no Banco B (30/09/2026, export real de 01/09 a 29/09, 497
# laudos sem Cancelada, "Data Criação" x "Data Status"): 20 dias pegavam
# só 96,6% - 17 laudos ficavam de fora pra sempre - e 30 dias pegam
# 100% (maior atraso visto: 27 dias). Confirmado na prática em
# 29/09/2026: uma execução com 30 dias achou 11 propostas que nunca
# tinham sido cadastradas. O fluxo pula o que já foi feito, então a
# janela maior custa quase só o tamanho do export. (Antes era 20, o
# valor medido no Banco A; e antes disso 4, que perdia quase metade.)
DIAS_JANELA_AUTOMATICA = 30

# piso: nunca puxar nada anterior a esta data, mesmo que a janela (ou
# uma varredura de recuperação com JANELA_DIAS grande) alcance mais pra
# trás. Regra de negócio, não técnica - confirmado com o usuário
# (17/09/2026) que é o mesmo piso usado no Banco A: 01/09/2026, data em
# que o trabalho começou. Sobrepõe/desliga com a variável de ambiente
# DATA_MINIMA (dd/mm/aaaa); DATA_MINIMA="" tira o piso.
DATA_MINIMA_PADRAO = "01/09/2026"

# aba extra que a Plataforma inclui no export só com metadados do
# relatório (usuário, período, filtros...) - não é dado de vistoria e
# atrapalha o cadastro na Central de Gestão depois, por isso é removida.
NOME_ABA_INFORMACOES = "Informações"

# Tipos de inspeção a manter marcados no filtro. Confirmado na tela real
# do Banco B (09/09/2026) que a lista aqui é DIFERENTE da do Banco A (que
# tinha "Laudo Eletrônico"/"Laudo Físico"): o Banco B oferece AVM,
# Laudão, Urbano Alto Padrão e Urbano Padrão. Queremos todos menos AVM.
TIPOS_INSPECAO = ["Laudão", "Urbano Alto Padrão", "Urbano Padrão"]

# Status de inspeção a manter marcados no filtro. Confirmado na lista
# real do Banco B (09/09/2026): além de "Laudo Aceito" e "Cancelada"
# (os dois usados no Banco A), aqui existe também "Laudo Aceito com
# Ressalvas" - o laudo foi aceito do mesmo jeito, a vistoria aconteceu,
# então entra no cadastro igual aos outros.
#
# "Laudo Indeferido" e "Laudo Devolvido" entraram em 29/09/2026, a pedido
# do usuário: nesses status a vistoria foi feita e o engenheiro tem de
# ser pago. Rótulos copiados do print real do filtro do Banco B
# (29/09/2026). A etiqueta na lista diz "Laudo Devolvido (Prest)", mas
# no filtro a opção é só "Laudo Devolvido" (não existe uma com
# "(Prest)"). Entram direto, como as aceitas (decisão do usuário) - só
# a Cancelada é conferida inspeção por inspeção (STATUS_CANCELADA).
STATUS_INSPECAO = [
    "Laudo Aceito",
    "Laudo Aceito com Ressalvas",
    "Laudo Indeferido",
    "Laudo Devolvido",
    "Cancelada",
]

# único status que precisa ser conferido inspeção por inspeção (tem
# vistoria de verdade?). Os demais status do filtro são laudos aceitos -
# não seriam aceitos sem vistoria, então abrir cada um seria checar o
# óbvio pagando o preço de abrir/fechar um modal por linha.
#
# Atenção: o texto do status no EXCEL não é o mesmo do rótulo na TELA.
# Confirmado no export real (09/09/2026): a tela chama "Laudo Aceito com
# Ressalvas", o Excel escreve "Laudo Aceito c/ Ressalva". Por isso
# STATUS_INSPECAO (acima) usa os rótulos da tela - são pra clicar nos
# checkboxes - e a leitura do Excel usa a palavra "Cancelada", que é
# igual nos dois lugares.
STATUS_CANCELADA = "Cancelada"

# colunas do Excel exportado usadas pra saber quais linhas conferir.
# Confirmado no export real: coluna 1 = "Identificador" (o mesmo código
# TAUxxxx que aparece na lista da tela) e coluna 44 = "Status". A busca
# do Status é EXATA porque existe também uma coluna "Data Status".
NOME_COLUNA_IDENTIFICADOR = "Identificador"
NOME_COLUNA_STATUS = "Status"
# mesma coluna usada na Etapa 4 (ver NOME_COLUNA_DATA_VISTORIA em
# agendar_vistoria_central_gestao.py) - mesmo relatório da Plataforma, nome
# de coluna idêntico. Usada aqui só pra decidir se uma Cancelada que a
# tela não mostrou pode ficar mesmo assim (ver identificar_canceladas_sem_vistoria).
NOME_COLUNA_DATA_VISTORIA = "Data Vistoria"

LOGIN_URL = "https://plataforma-laudos.example.com/sistema/index.html#/home"
PASTA_SCRIPT = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(PASTA_SCRIPT, "data", "exports")
# só pra guardar print/HTML de uma falha real, pra próxima vez ter evidência
# em vez de chutar de novo - nunca sobe pro git (ver .gitignore).
DEBUG_DIR = os.path.join(PASTA_SCRIPT, "data", "debug_laudos")
# cookie de sessão da Plataforma - o navegador descarta ao fechar, então
# salvamos/recarregamos na mão. Nunca sobe pro GitHub (.gitignore).
# Arquivo separado do robô do Banco A (mesmo login, contas/clientes
# diferentes na mesma conta) pra nunca misturar sessão de um cliente
# com o outro.
SESSION_FILE = os.path.join(PASTA_SCRIPT, "sessao_plataforma_banco_b.json")
LOG_DIR = os.path.join(PASTA_SCRIPT, "logs")

# índice do card do cliente na tela "escolha o cliente" (0 = primeiro).
# Confirmado no robô irmão que já baixa laudos da Plataforma pro
# Banco B: os cards vêm na ordem Banco A (0), Banco B (1) - por isso 1
# aqui (não é chute, é o mesmo valor já usado nesse outro robô). Se a
# Plataforma mudar a ordem, ajuste este número.
INDICE_CLIENTE_BANCO_B = 1

# login e seleção de cliente são automáticos, não precisa ver a janela.
# se a sessão expirar de verdade, o robô avisa e para (não dá pra fazer
# login manual headless) - pra ver a janela e logar manualmente, rode com
# a variável de ambiente HEADLESS=false (ver comum.headless_configurado),
# não edite esta linha - editar cria conflito com o próximo `git pull`.
HEADLESS = headless_configurado()

# cada linha da lista é uma div (não uma <table> de verdade). A tela tem
# abas escondidas com as mesmas linhas no HTML, por isso o :visible no
# final.
LINHA_SELECTOR = "div.insp360-mouse-link.insp360-tabela-relatorio.insp360-cor-tabela-rel:visible"

# a tela da Plataforma anda lenta pra carregar conteúdo dentro dos modais
# de detalhe - timeout mais folgado que o padrão pra essas esperas
# específicas (confirmado necessário no fluxo do Banco A).
TIMEOUT_LONGO = 45000


def novo_contexto_pagina(browser, com_sessao):
    # viewport grande - a tabela usa rolagem virtual e só desenha as
    # linhas que cabem na altura visível
    context = browser.new_context(
        accept_downloads=True,
        viewport={"width": 1920, "height": 1080} if HEADLESS else None,
        no_viewport=None if HEADLESS else True,
        storage_state=SESSION_FILE if com_sessao else None,
    )
    if HEADLESS:
        bloquear_recursos_visuais(context)
    return context, context.new_page()


def codigo_pagina_de_erro(page):
    """Devolve o código da tela de erro da Plataforma ("401", "403"...) se
    a página atual for uma delas, ou None.

    A Plataforma não redireciona pra tela de login quando a sessão perde
    a validade: ela manda o navegador pra uma rota de erro do próprio
    sistema (.../sistema/http-status.html#/401, "ACESSO NÃO AUTORIZADO").
    Confirmado no robô irmão do Banco A - mesma plataforma, mesmo
    comportamento aqui. Sem reconhecer isso, o robô só via "não achei
    GRID DE INSPEÇÃO", tentava logar sozinho, não achava formulário de
    login nenhum (porque não existe formulário nessa tela) e parava com
    uma mensagem que não explicava nada."""
    url = page.url or ""
    if "http-status.html" not in url:
        return None
    if "#/" in url:
        return url.split("#/")[-1].strip("/") or "?"
    return "?"


def descartar_sessao_salva():
    """Apaga o arquivo de sessão. Usado quando a Plataforma devolve a tela
    de erro: insistir com um cookie que o servidor já recusou só repete o
    mesmo 401 - é preciso recomeçar limpo pra chegar no formulário de
    login de verdade."""
    try:
        os.remove(SESSION_FILE)
        return True
    except OSError:
        return False


def pagina_pede_mfa(page):
    """True se a Plataforma está na tela de verificação em duas etapas
    (index.html#/verificacao-mfa). Descoberto em execução real
    (22/09/2026) pelo print/URL salvos na falha do login automático: a
    credencial estava certa, o cliente Banco B chegou a ser
    selecionado, e a Plataforma pediu um segundo fator antes de liberar o
    painel. Senha sozinha não passa por aqui - esse é exatamente o
    objetivo do MFA - então o robô nunca vai conseguir "logar sozinho"
    nesse estado: precisa de alguém completando a verificação uma vez
    com a janela visível, e daí a sessão salva volta a valer."""
    return "verificacao-mfa" in (page.url or "")


def aguardar_painel_ou_erro(page, timeout=10000):
    """Espera até o painel (GRID DE INSPEÇÃO) aparecer OU a Plataforma
    mandar pra tela de erro dela OU pedir MFA - o que vier primeiro.
    Devolve ("painel", None), ("erro", "401"), ("mfa", None) ou
    ("nada", None) se o tempo acabou sem nenhum dos três.

    Existe porque a tela 401 não vem na hora: a SPA carrega index.html
    normalmente, faz uma chamada de API por trás, e SÓ ENTÃO redireciona
    pra http-status.html#/401. Checar a URL uma única vez logo depois do
    goto passava batido - o robô concluía "não achei o painel", tentava
    logar, e o usuário encontrava o Chrome parado no 401 sem formulário
    nenhum pra preencher. Mesmo aprendizado do robô irmão do Banco A
    (17/09/2026), confirmado real no Banco B em 22/09/2026: uma sessão
    genuinamente válida (chegou a selecionar o cliente) foi tratada como
    morta porque o painel demorou mais que os 10s de espera antiga."""
    prazo = time.time() + timeout / 1000
    while time.time() < prazo:
        codigo = codigo_pagina_de_erro(page)
        if codigo:
            return "erro", codigo
        if pagina_pede_mfa(page):
            return "mfa", None
        if page.locator("text=GRID DE INSPEÇÃO").count() > 0:
            return "painel", None
        page.wait_for_timeout(500)
    return "nada", None


def chegou_no_painel(page, timeout=8000):
    try:
        page.wait_for_selector("text=GRID DE INSPEÇÃO", timeout=timeout)
        return True
    except PlaywrightTimeout:
        return False


def selecionar_cliente_banco_b(page, timeout=6000):
    # tela "escolha o cliente" aparece mesmo com sessão válida - o logo é
    # imagem, sem texto, então usamos a ordem dos cards.
    selecionar = page.locator("text=SELECIONAR")
    try:
        selecionar.first.wait_for(timeout=timeout)
    except PlaywrightTimeout:
        # pode já ter caído direto no painel (só 1 cliente na conta) - ok
        return False

    if selecionar.count() <= INDICE_CLIENTE_BANCO_B:
        print("      [AVISO] Tela de cliente com layout inesperado - selecione manualmente.")
        return False

    selecionar.nth(INDICE_CLIENTE_BANCO_B).click()
    page.wait_for_timeout(500)
    print("      Cliente Banco B selecionado automaticamente.")
    return True


def abrir_relatorio_analitico(page):
    """Abre o Relatório Analítico pelo menu. A Plataforma pode organizar o
    menu de forma diferente pra cada cliente, então tentamos alguns
    caminhos possíveis antes de desistir."""
    caminhos_possiveis = [
        ["Administrativo", "Relatórios", "Inspeções", "Analítico"],
        ["Relatórios", "Inspeções", "Analítico"],
        ["Relatórios", "Analítico"],
    ]
    ultimo_erro = None
    for caminho in caminhos_possiveis:
        try:
            page.locator("button.hamburger").click()
            for item in caminho:
                page.click(f"text={item}", timeout=4000)
            return
        except PlaywrightTimeout as e:
            ultimo_erro = e
            continue
    raise RuntimeError(
        "Não consegui achar o caminho de menu até o Relatório Analítico "
        f"(tentei: {caminhos_possiveis}). Rode com HEADLESS = False, abra o "
        "menu manualmente e me diga os nomes exatos dos itens pra eu "
        "ajustar o script."
    ) from ultimo_erro


def _checkbox_apos(ancora, texto_alvo, timeout=10000):
    """Acha o checkbox do primeiro elemento com esse texto EXATO que
    aparece DEPOIS de `ancora` no documento - a tela de filtros repete
    "Selecionar Todos" (e alguns rótulos) em várias seções (Ramo,
    Produto Base, Tipo de Inspeção, Status de Inspeção, Região...),
    então é essencial ancorar a busca na seção certa, não fazer busca
    global."""
    alvo = ancora.locator(f"xpath=following::*[normalize-space(text())='{texto_alvo}'][1]").first
    alvo.wait_for(timeout=timeout)
    alvo.scroll_into_view_if_needed()
    return alvo.locator(
        "xpath=preceding::input[@type='checkbox'][1] | ancestor::label//input[@type='checkbox']"
    ).first


def definir_filtro_checkboxes(page, titulo_secao, rotulos_manter_marcados):
    """Desmarca 'Selecionar Todos' da seção e marca só os rótulos
    pedidos - usado pra restringir Tipo de Inspeção e Status de
    Inspeção antes de exibir/exportar."""
    print(f"      Ajustando filtro '{titulo_secao}' -> {', '.join(rotulos_manter_marcados)}...")
    titulo = page.locator(f"text={titulo_secao}").first
    titulo.wait_for(timeout=10000)
    titulo.scroll_into_view_if_needed()

    checkbox_todos = _checkbox_apos(titulo, "Selecionar Todos")
    if checkbox_todos.is_checked():
        checkbox_todos.click()
        page.wait_for_timeout(300)

    for rotulo in rotulos_manter_marcados:
        checkbox = _checkbox_apos(titulo, rotulo)
        if not checkbox.is_checked():
            checkbox.click()
        page.wait_for_timeout(200)


def preencher_periodo_e_exibir(page, data_inicio, data_fim):
    page.wait_for_selector("text=GRID DE INSPEÇÃO", timeout=20000)
    abrir_relatorio_analitico(page)

    page.wait_for_selector("text=PERÍODO DE SOLICITAÇÃO DA INSPEÇÃO", timeout=15000)
    # busca os campos de data DEPOIS desse título - o painel FILTROS tem
    # outros campos com a mesma classe CSS
    titulo_periodo = page.locator("text=PERÍODO DE SOLICITAÇÃO DA INSPEÇÃO")
    campos_data = titulo_periodo.locator(
        "xpath=following::input[contains(@class,'insp360-filtro-data')]"
    )

    def preencher_data(campo, valor):
        # .fill() não funciona com a máscara desse campo
        campo.click()
        campo.press("Control+A")
        campo.type(valor, delay=50)
        page.keyboard.press("Escape")
        page.wait_for_timeout(300)

    preencher_data(campos_data.nth(0), data_inicio)
    preencher_data(campos_data.nth(1), data_fim)
    # os botões "30/60/90" ao lado são atalhos de período ("últimos N
    # dias"), não itens por página - não mexer, sobrescrevem as datas

    definir_filtro_checkboxes(page, "TIPO DE INSPEÇÃO", TIPOS_INSPECAO)
    definir_filtro_checkboxes(page, "STATUS DE INSPEÇÃO", STATUS_INSPECAO)

    page.get_by_role("button", name="Exibir").click()
    page.wait_for_selector(LINHA_SELECTOR, timeout=20000)


def clicar_exportar_excel(page):
    """Clica no botão de exportar da tela de relatório (alguns sistemas
    disparam o download direto, outros abrem um submenu com o formato
    Excel/PDF/CSV antes) e devolve o download do Playwright."""
    gatilhos = [
        "button:has-text('Exportar')",
        "a:has-text('Exportar')",
        "[title*='Exportar' i]",
        "i.fa-file-excel, i.fa-file-excel-o",
        "text=Exportar",
    ]
    gatilho = None
    for seletor in gatilhos:
        loc = page.locator(seletor)
        if loc.count() > 0:
            gatilho = loc.first
            break

    if gatilho is None:
        raise RuntimeError(
            "Não encontrei o botão 'Exportar' na tela do relatório. Rode com "
            "HEADLESS = False, veja como o botão/ícone aparece na tela "
            "(texto, ícone, tooltip) e me diga pra eu ajustar o seletor."
        )

    try:
        with page.expect_download(timeout=8000) as download_info:
            gatilho.click()
        return download_info.value
    except PlaywrightTimeout:
        pass  # não baixou direto - deve ter aberto um submenu de formato

    # BUG real (22/09/2026, execução automática, mesma função copiada do
    # Banco A): depois do clique acima, o código antes conferia
    # `opcao_excel.count() == 0` NA HORA - sem esperar o submenu terminar
    # de renderizar. O clique tinha acontecido certo, só que a checagem
    # rodava antes da animação do menu terminar, então às vezes o "Excel"
    # ainda não existia no DOM no instante exato da contagem - timing, não
    # seletor errado nem mudança de tela. `wait_for` (que já espera até um
    # prazo generoso, sem inventar um `sleep` fixo) resolve isso: só falha
    # de verdade se o texto não aparecer dentro do prazo.
    opcao_excel = page.locator("text=Excel")
    try:
        opcao_excel.first.wait_for(state="visible", timeout=15000)
    except PlaywrightTimeout:
        _salvar_evidencia_falha_exportar(page, "sem_opcao_excel")
        raise RuntimeError(
            "Cliquei em 'Exportar' mas não achei uma opção 'Excel' em "
            "seguida, nem o download começou sozinho. Print e HTML da tela "
            f"nesse momento salvos em {DEBUG_DIR} pra eu ver o que apareceu, "
            "sem precisar reproduzir a falha."
        )
    with page.expect_download(timeout=30000) as download_info:
        opcao_excel.first.click()
    return download_info.value


def _salvar_evidencia_falha_exportar(page, motivo):
    """Salva print + HTML da tela no momento da falha, pra diagnosticar sem
    depender de reproduzir o problema (a causa pode ser passageira -
    lentidão do servidor, ordem de carregamento). Nunca sobe pro git
    (.gitignore)."""
    try:
        os.makedirs(DEBUG_DIR, exist_ok=True)
        marca = datetime.now().strftime("%Y%m%d_%H%M%S")
        base = os.path.join(DEBUG_DIR, f"exportar_{motivo}_{marca}")
        page.screenshot(path=f"{base}.png", full_page=True)
        with open(f"{base}.html", "w", encoding="utf-8") as f:
            f.write(page.content())
        print(f"      [DEBUG] Evidência da falha salva em {base}.png / .html")
    except Exception as e:
        print(f"      [AVISO] Não consegui salvar evidência da falha ({e}).")


def _tem_laudo_publicado(page):
    """Com DETALHAR INSPEÇÃO já aberto: vai pra aba Laudos e diz se tem
    pelo menos uma versão de laudo publicada.

    A decisão é tomada pelo aviso "Nenhum laudo encontrado" estar
    VISÍVEL - esse é o sinal positivo e inequívoco de "não tem laudo".
    A versão anterior decidia ao contrário (tem engrenagem => tem
    laudo) e dava falso positivo em 100% dos casos: o locator procurava
    a engrenagem na PÁGINA INTEIRA, e a tela do relatório atrás do modal
    tem engrenagens próprias, então a contagem nunca era zero e nenhuma
    Cancelada sem vistoria era detectada (confirmado em teste real no
    Banco B, 09/09/2026: a inspeção TAU5142 mostrava "Nenhum laudo
    encontrado até o momento" na tela e mesmo assim passou como se
    tivesse laudo). Por isso aqui a busca é escopada no modal e usa
    visibilidade, não contagem no DOM."""
    page.click("text=Laudos")
    modal = page.locator(".modal.in").last
    aviso_sem_laudo = modal.locator("text=Nenhum laudo encontrado")
    icone_engrenagem = modal.locator("i.fa-cog, i.fa-gear, .fa-cog")
    # espera a aba terminar de carregar: ou aparece o aviso, ou aparece
    # a lista de laudos (com a engrenagem de ações de cada versão)
    aviso_sem_laudo.or_(icone_engrenagem).first.wait_for(timeout=TIMEOUT_LONGO)
    return not aviso_sem_laudo.is_visible()


def _fechar_todos_modais(page, max_tentativas=3):
    """Fecha modal(is) abertos na Plataforma, em ordem crescente de força
    (ESC -> clique no ícone -> clique forçado -> remoção via JS como
    último recurso). Devolve True se conseguiu limpar tudo. Seletor do
    ícone de fechar (ng-click="fecharModal(true)") confirmado via
    Inspecionar elemento no robô irmão que lê laudos dessa mesma tela."""
    for _ in range(max_tentativas):
        if page.locator(".modal.in").count() == 0:
            return True
        page.keyboard.press("Escape")
        page.wait_for_timeout(300)
        if page.locator(".modal.in").count() == 0:
            return True
        modais = page.locator(".modal.in")
        fechar = modais.last.locator('i[ng-click="fecharModal(true)"]')
        alvo = fechar if fechar.count() > 0 else modais.last.locator("i.insp360-icone-fechar")
        try:
            alvo.first.click(timeout=2000)
        except PlaywrightTimeout:
            try:
                alvo.first.click(timeout=2000, force=True)
            except Exception:
                pass
        page.wait_for_timeout(300)

    if page.locator(".modal.in").count() == 0:
        return True
    try:
        page.evaluate(
            "document.querySelectorAll('.modal.in, .modal-backdrop').forEach(el => el.remove());"
            "document.body.classList.remove('modal-open');"
        )
        page.wait_for_timeout(300)
    except Exception:
        pass
    return page.locator(".modal.in").count() == 0


def _recarregar_do_zero(page, data_inicio, data_fim):
    """Sai da tela de filtros e volta pelo caminho de login (mesmo
    caminho já comprovado confiável no início da execução) - usado só
    quando um modal preso não fecha de nenhum outro jeito."""
    abrir_pagina(page, LOGIN_URL, timeout=TIMEOUT_LONGO)
    selecionar_cliente_banco_b(page)
    chegou_no_painel(page, timeout=TIMEOUT_LONGO)
    preencher_periodo_e_exibir(page, data_inicio, data_fim)


def _garantir_lista_limpa(page, data_inicio, data_fim):
    """Garante que não sobrou modal preso antes de seguir pra próxima
    linha/página. Devolve True se seguiu sem precisar recarregar, False
    se precisou sair e voltar pelo login (o chamador recomeça a leitura
    da lista da página 1 nesse caso)."""
    try:
        if _fechar_todos_modais(page, max_tentativas=3):
            return True
        print("      [AVISO] Modal preso não fechou - saindo e voltando pelo login...")
        _recarregar_do_zero(page, data_inicio, data_fim)
        return False
    except Exception as e:
        print(f"      [AVISO] Recarregamento falhou ({e}) - seguindo mesmo assim.")
        return False


def ler_canceladas_do_excel(caminho_arquivo):
    """Lê do Excel recém-baixado quais Identificadores estão com status
    Cancelada - essa é a lista do que precisa ser conferido na tela.
    Devolve (ids_canceladas, ids_com_data_vistoria) - o segundo conjunto
    é o subconjunto do primeiro que já tem "Data Vistoria" preenchida
    (usado como exceção quando a tela não mostra a linha - ver
    identificar_canceladas_sem_vistoria). None se não deu pra ler.

    O Excel é a fonte confiável aqui, não a tela: a lista da Plataforma
    desenha só as linhas que cabem na altura visível, então percorrer a
    tela pode simplesmente não enxergar parte das inspeções (visto em
    teste real: 7 no Excel, 4 lidas na tela). O export, ao contrário,
    vem completo."""
    if os.path.splitext(caminho_arquivo)[1].lower() != ".xlsx":
        return None  # sem como ler - o chamador cai no modo antigo (varre a tela)

    wb = load_workbook(caminho_arquivo, read_only=True, data_only=True)
    ids_canceladas = set()
    ids_com_data_vistoria = set()
    status_vistos = set()
    achou_colunas = False
    for nome_aba in wb.sheetnames:
        ws = wb[nome_aba]
        coluna_id = achar_coluna_por_cabecalho(ws, NOME_COLUNA_IDENTIFICADOR, exato=True)
        coluna_status = achar_coluna_por_cabecalho(ws, NOME_COLUNA_STATUS, exato=True)
        coluna_data_vist = achar_coluna_por_cabecalho(ws, NOME_COLUNA_DATA_VISTORIA, exato=True)
        if coluna_id is None or coluna_status is None:
            continue
        achou_colunas = True
        for linha in ws.iter_rows(min_row=2):
            identificador = str(linha[coluna_id - 1].value or "").strip()
            status = str(linha[coluna_status - 1].value or "").strip()
            if not identificador:
                continue
            status_vistos.add(status)
            if STATUS_CANCELADA.lower() in status.lower():
                ids_canceladas.add(identificador)
                if coluna_data_vist is not None and str(linha[coluna_data_vist - 1].value or "").strip():
                    ids_com_data_vistoria.add(identificador)

    if not achou_colunas:
        print(f"      [AVISO] Não achei as colunas '{NOME_COLUNA_IDENTIFICADOR}' e '{NOME_COLUNA_STATUS}' "
              "no Excel - vou conferir varrendo a tela (menos confiável).")
        return None

    # mostra os status que realmente vieram no Excel: se um dia a Plataforma
    # escrever "Cancelada" de outro jeito, a busca acharia zero Canceladas e
    # a conferência seria pulada silenciosamente - com essa linha no log dá
    # pra perceber na hora. (O Excel já escreve diferente da tela num caso
    # conhecido: "Laudo Aceito c/ Ressalva" x "Laudo Aceito com Ressalvas".)
    print(f"      Status presentes no Excel: {', '.join(sorted(status_vistos)) or '(nenhum)'}")
    return ids_canceladas, ids_com_data_vistoria


def _carregar_todas_as_linhas(page):
    """Rola a lista até parar de aparecer linha nova e devolve quantas
    ficaram desenhadas - a Plataforma só desenha as linhas que cabem na
    altura visível, então sem rolar o robô não enxerga a lista toda."""
    qtd_anterior = -1
    for _ in range(30):
        linhas = page.locator(LINHA_SELECTOR)
        qtd_atual = linhas.count()
        if qtd_atual == qtd_anterior:
            break
        qtd_anterior = qtd_atual
        try:
            linhas.last.scroll_into_view_if_needed(timeout=5000)
        except Exception:
            break
        page.wait_for_timeout(400)
    return page.locator(LINHA_SELECTOR).count()


def identificar_canceladas_sem_vistoria(page, data_inicio, data_fim, ids_canceladas=None,
                                         ids_com_data_vistoria=None):
    """Percorre TODAS as páginas da lista já filtrada (ver
    TIPOS_INSPECAO e STATUS_INSPECAO) e abre APENAS as inspeções
    Canceladas pra checar, na aba Laudos, se existe pelo menos uma
    versão publicada. Devolve o conjunto de IDs que NÃO têm nenhum
    laudo - são Canceladas sem vistoria de verdade, que devem sair do
    Excel final.

    Só as Canceladas são abertas: um laudo aceito (com ou sem
    ressalvas) nunca seria aceito sem vistoria, então abrir essas
    linhas seria checar o óbvio e custa caro - cada uma abre um modal,
    espera carregar e fecha. Num período grande isso é a maior parte do
    tempo da Etapa 1.

    `ids_canceladas` vem do Excel (ver ler_canceladas_do_excel) e é a
    lista de quem PRECISA ser conferido. Quando ela é passada, qualquer
    Cancelada que a tela não tiver mostrado é removida por segurança no
    final - sem isso, uma Cancelada que ficasse fora da parte desenhada
    da lista passaria batida como se tivesse vistoria. Se vier None (não
    deu pra ler o Excel), cai no modo antigo: decide pelo texto da
    própria linha.

    `ids_com_data_vistoria` (também do Excel) é a exceção a essa regra:
    quando a Plataforma JÁ TEM a data da vistoria no Excel, a vistoria
    de fato aconteceu, então uma Cancelada nessa lista que a tela não
    mostrou fica mantida em vez de removida - a regra "não deu pra
    conferir -> remove por segurança" continua valendo pras que também
    não têm data (mesmo aprendizado do robô irmão do Banco A, confirmado
    lá em execução real: sem essa exceção, Cancelada com vistoria de
    verdade mas fora da parte desenhada da tela era descartada à toa).

    Se não conseguir checar uma Cancelada por erro técnico (não
    confundir com 'sem laudo', que é resultado normal), ela é removida
    por segurança - melhor conferir manualmente depois do que arriscar
    subir uma Cancelada sem vistoria de verdade."""
    ids_sem_vistoria = set()
    canceladas_conferidas = set()
    pagina_num = 1
    total_linhas_lidas = 0

    while True:
        qtd = _carregar_todas_as_linhas(page)
        linhas = page.locator(LINHA_SELECTOR)
        print(f"      Página {pagina_num}: {qtd} inspeção(ões) - checando vistoria das Canceladas...")

        codigos_vistos = set()
        precisou_recarregar = False
        for i in range(qtd):
            try:
                texto_linha = linhas.nth(i).inner_text(timeout=TIMEOUT_LONGO)
            except Exception as e:
                print(f"      [AVISO] Não consegui ler a linha {i + 1} ({e}) - recomeçando a leitura da lista.")
                precisou_recarregar = True
                break

            id_linha = texto_linha.split("\n")[0].strip()
            if not id_linha:
                print(f"        [AVISO] Linha {i + 1} sem ID legível (texto: '{texto_linha[:40]}...') - pulando.")
                continue
            if id_linha in codigos_vistos:
                continue  # mesma linha repetida (aba escondida com o mesmo HTML) - normal, sem aviso
            codigos_vistos.add(id_linha)
            total_linhas_lidas += 1

            if ids_canceladas is not None:
                if id_linha not in ids_canceladas:
                    continue  # laudo aceito (com ou sem ressalvas) sempre teve vistoria
            elif STATUS_CANCELADA.lower() not in texto_linha.lower():
                continue
            canceladas_conferidas.add(id_linha)

            linha = page.locator(LINHA_SELECTOR, has_text=id_linha)
            if linha.count() == 0:
                # a linha sumiu entre ler a lista e tentar clicar (ex.: a
                # tela recarregou sozinha) - não dá pra confundir isso com
                # "sem laudo" (resultado normal), mas também não podemos
                # deixar essa Cancelada passar sem checagem nenhuma -
                # removida por segurança, igual ao caso de erro abaixo.
                print(f"        [AVISO] [{id_linha}] não achei mais essa linha pra clicar - "
                      "removendo por segurança (não consegui checar a vistoria).")
                ids_sem_vistoria.add(id_linha)
                continue

            try:
                linha.first.click(timeout=15000)
                page.wait_for_selector("text=DETALHAR INSPEÇÃO", timeout=TIMEOUT_LONGO)
                if not _tem_laudo_publicado(page):
                    ids_sem_vistoria.add(id_linha)
                    print(f"        [{id_linha}] sem laudo publicado - será removida (Cancelada sem vistoria).")
            except Exception as e:
                print(f"        [AVISO] Não consegui checar {id_linha} ({e}) - removendo por segurança.")
                ids_sem_vistoria.add(id_linha)
            finally:
                if not _garantir_lista_limpa(page, data_inicio, data_fim):
                    precisou_recarregar = True
                    break

        if precisou_recarregar:
            pagina_num = 1
            continue

        botao_proxima = page.locator("a:visible", has_text="próxima")
        if botao_proxima.count() == 0 or botao_proxima.get_attribute("disabled") is not None:
            break
        botao_proxima.click(timeout=TIMEOUT_LONGO)
        pagina_num += 1
        page.wait_for_selector(LINHA_SELECTOR, timeout=TIMEOUT_LONGO)

    print(f"      {total_linhas_lidas} inspeção(ões) lida(s) na lista, "
          f"{len(canceladas_conferidas)} Cancelada(s) conferida(s), {len(ids_sem_vistoria)} sem vistoria.")

    if ids_canceladas is not None:
        nao_conferidas = ids_canceladas - canceladas_conferidas
        if nao_conferidas:
            com_data = ids_com_data_vistoria or set()
            removidas = sorted(i for i in nao_conferidas if i not in com_data)
            mantidas = sorted(i for i in nao_conferidas if i in com_data)
            print(f"      {len(nao_conferidas)} Cancelada(s) do Excel não apareceram na lista da tela:")
            if removidas:
                print(f"        {len(removidas)} sem data de vistoria na Plataforma -> removidas por segurança "
                      f"(sem como conferir, e sem vistoria não completam nunca): {', '.join(removidas)}")
                ids_sem_vistoria.update(removidas)
            if mantidas:
                print(f"        {len(mantidas)} com data de vistoria na Plataforma -> mantidas: {', '.join(mantidas)}")

    return ids_sem_vistoria


def remover_propostas_sem_vistoria(caminho_arquivo, ids_sem_vistoria):
    """Remove do Excel as linhas cujo ID (1ª coluna) esteja em
    ids_sem_vistoria - Canceladas sem nenhuma vistoria de verdade, que
    não devem entrar no cadastro em Em Aberto.

    Devolve True se está seguro confiar no Excel (nada a remover, ou
    removeu exatamente a quantidade esperada) e False se não dá pra
    confiar sem conferência manual - a contagem removida não bateu com
    o esperado (sinal de que a suposição "1ª coluna = mesmo
    Identificador da tela" pode estar errada nesse Excel) ou o arquivo
    nem é .xlsx. A checagem existe pra pegar esse caso automaticamente,
    em vez de confiar cegamente que a remoção funcionou."""
    if not ids_sem_vistoria:
        print("      Nenhuma Cancelada sem vistoria encontrada - Excel mantido como veio da Plataforma.")
        return True

    extensao = os.path.splitext(caminho_arquivo)[1].lower()
    if extensao != ".xlsx":
        print(f"      [AVISO] Arquivo é {extensao}, não .xlsx - não dá pra remover linhas automaticamente.")
        print(f"      Remova manualmente estas propostas antes de cadastrar: {', '.join(sorted(ids_sem_vistoria))}")
        return False

    wb = load_workbook(caminho_arquivo)
    total_removidas = 0
    for nome_aba in wb.sheetnames:
        ws = wb[nome_aba]
        linhas_para_remover = [
            linha for linha in range(2, ws.max_row + 1)  # linha 1 = cabeçalho
            if str(ws.cell(row=linha, column=1).value or "").strip() in ids_sem_vistoria
        ]
        for linha in reversed(linhas_para_remover):
            ws.delete_rows(linha)
            total_removidas += 1
    wb.save(caminho_arquivo)

    print(f"      {total_removidas} linha(s) removida(s) do Excel: {', '.join(sorted(ids_sem_vistoria))}")
    if total_removidas != len(ids_sem_vistoria):
        print(f"      [ERRO] Esperava remover {len(ids_sem_vistoria)} linha(s), removi {total_removidas} - "
              "a 1ª coluna do Excel pode não ser o mesmo Identificador mostrado na lista da Plataforma.")
        print(f"      O arquivo NÃO será usado automaticamente pelas próximas etapas - confira manualmente:")
        print(f"      {caminho_arquivo}")
        return False
    return True


def remover_aba_informacoes(caminho_arquivo):
    """Remove a aba 'Informações' (metadados do relatório) do Excel
    baixado, deixando só as abas com dados de vistoria."""
    extensao = os.path.splitext(caminho_arquivo)[1].lower()
    if extensao != ".xlsx":
        print(f"      [AVISO] Arquivo é {extensao}, não .xlsx - não dá pra editar as abas.")
        return

    wb = load_workbook(caminho_arquivo)
    abas_para_remover = [
        nome for nome in wb.sheetnames if nome.strip().lower() == NOME_ABA_INFORMACOES.lower()
    ]
    if not abas_para_remover:
        print(f"      Nenhuma aba '{NOME_ABA_INFORMACOES}' encontrada (nada a remover).")
        return

    for nome in abas_para_remover:
        if len(wb.sheetnames) <= 1:
            print(f"      [AVISO] '{nome}' é a única aba do arquivo - mantendo pra não ficar vazio.")
            continue
        del wb[nome]
        print(f"      Aba '{nome}' removida.")

    wb.save(caminho_arquivo)


def _salvar_evidencia_falha_login(page, motivo):
    try:
        os.makedirs(LOG_DIR, exist_ok=True)
        caminho_print = os.path.join(LOG_DIR, f"falha_login_plataforma_{motivo}_{datetime.now():%Y%m%d_%H%M%S}.png")
        page.screenshot(path=caminho_print, full_page=True)
        print(f"      Print da tela no momento da falha salvo em: {caminho_print}")
    except Exception as e:
        print(f"      [AVISO] Não consegui nem salvar o print da falha ({e}).")


def _avisar_mfa(page):
    """Mensagem única e honesta pro caso de MFA: não é credencial errada
    nem lentidão, é um segundo fator que senha nenhuma passa. Guarda o
    print pra registro, mas o que resolve é uma pessoa completar a
    verificação uma vez com a janela visível."""
    print("      [MFA] A Plataforma está pedindo verificação em duas etapas "
          f"(URL: {page.url}).")
    print("      Isso não dá pra passar por senha - é o objetivo do MFA. A credencial")
    print("      está certa (chegou até aqui autenticado); o que falta é o código da")
    print("      verificação, que só uma pessoa pode fornecer.")
    _salvar_evidencia_falha_login(page, "mfa")
    _descrever_tela_mfa(page)


JS_DESCREVER_TELA = """
() => {
  const vis = (el) => {
    const r = el.getBoundingClientRect();
    return r.width > 0 && r.height > 0;
  };
  const texto = (document.body.innerText || "").replace(/\\n{2,}/g, "\\n").trim();
  const campos = [];
  document.querySelectorAll("input, select, textarea, button, a.btn, [role='button']").forEach((el) => {
    if (!vis(el)) return;
    const rotulo = el.id ? document.querySelector(`label[for='${el.id}']`) : null;
    campos.push({
      tag: el.tagName.toLowerCase(),
      type: el.getAttribute("type") || "",
      id: el.id || "",
      name: el.getAttribute("name") || "",
      placeholder: el.getAttribute("placeholder") || "",
      texto: el.type === "checkbox" ? "" : (el.innerText || el.value || "").trim().slice(0, 60),
      rotulo: rotulo ? rotulo.innerText.trim().slice(0, 60) : "",
      checked: el.type === "checkbox" ? el.checked : null,
    });
  });
  return { texto: texto.slice(0, 1500), campos };
}
"""


def _descrever_tela_mfa(page):
    """Escreve no log o que a tela de MFA mostra (texto visível e
    campos/botões), pra decidir a solução em cima do que a Plataforma
    realmente pede - código por e-mail? app autenticador? SMS? tem
    'lembrar este dispositivo'? - em vez de adivinhar. Nunca digita
    nada; só lê."""
    try:
        dados = page.evaluate(JS_DESCREVER_TELA)
    except Exception as e:
        print(f"      (não consegui ler a tela de MFA pra descrever: {e})")
        return
    print("      ---- o que a tela de MFA mostra (texto visível) ----")
    for linha in (dados.get("texto") or "").splitlines():
        linha = linha.strip()
        if linha:
            print(f"      | {linha}")
    print("      ---- campos e botões visíveis ----")
    for c in dados.get("campos") or []:
        partes = [f"<{c['tag']}"]
        for chave in ("type", "id", "name", "placeholder"):
            if c.get(chave):
                partes.append(f"{chave}={c[chave]!r}")
        partes[-1] += ">"
        if c.get("rotulo"):
            partes.append(f"rótulo={c['rotulo']!r}")
        if c.get("texto"):
            partes.append(f"texto={c['texto']!r}")
        if c.get("checked") is not None:
            partes.append(f"marcado={c['checked']}")
        print("      | " + " ".join(partes))
    print("      ---------------------------------------------------")


# Tela real da verificação (print de 23/09/2026): "VERIFICAÇÃO DE
# SEGURANÇA - Passo 1 de 2: Escolha como quer receber o código", com os
# cards E-MAIL / SMS / App Autenticador e o botão "Avançar". Não existe
# "lembrar este dispositivo" - toda sessão nova passa por aqui. O passo 2
# (onde se digita o código) ainda não foi visto: por isso o campo é
# procurado de forma genérica e a tela é descrita no log sempre que o
# robô passa por ela, até confirmarmos o layout.
SELETOR_CAMPO_CODIGO = (
    "input:not([type=hidden]):not([type=checkbox]):not([type=radio])"
    ":not([type=submit]):not([type=button]):visible"
)
PADRAO_BOTAO_CONFIRMAR_CODIGO = (
    r"^\s*(validar|verificar|confirmar|avan[çc]ar|entrar|continuar|concluir|enviar)\b"
)


def _clicar_botao_por_texto(page, padrao):
    rx = re.compile(padrao, re.IGNORECASE)
    botoes = page.locator(
        "button:visible, input[type=submit]:visible, a.btn:visible, [role=button]:visible"
    )
    for botao in botoes.all():
        try:
            texto = (botao.inner_text() or botao.get_attribute("value") or "").strip()
        except Exception:
            continue
        if rx.search(texto):
            botao.click()
            return True
    return False


def _esperar_campos_codigo(page, timeout):
    prazo = time.time() + timeout / 1000
    while True:
        campos = page.locator(SELETOR_CAMPO_CODIGO).all()
        if campos or time.time() >= prazo:
            return campos
        page.wait_for_timeout(300)


def _responder_mfa_autenticador(page, chave):
    """Uma tentativa de passar pela verificação usando o App Autenticador.
    Devolve (True, None) se saiu da tela de MFA, ou (False, motivo)."""
    campos = _esperar_campos_codigo(page, 3000)
    if not campos:
        opcao = page.get_by_text("App Autenticador", exact=True)
        if opcao.count() == 0:
            return False, "sem_opcao_app_autenticador"
        opcao.first.click()
        page.wait_for_timeout(500)
        if not _clicar_botao_por_texto(page, r"^\s*Avan[çc]ar\s*$"):
            return False, "sem_botao_avancar"
        campos = _esperar_campos_codigo(page, 15000)
        if not campos:
            return False, "sem_campo_codigo"
        print("      Passo 2 da verificação - o que a tela mostra:")
        _descrever_tela_mfa(page)

    codigo = gerar_codigo_totp(chave)
    if len(campos) == 1:
        campos[0].fill(codigo)
    elif len(campos) == len(codigo):
        # uma caixinha por dígito
        for campo, digito in zip(campos, codigo):
            campo.fill(digito)
    else:
        return False, f"campos_inesperados_{len(campos)}"
    if not _clicar_botao_por_texto(page, PADRAO_BOTAO_CONFIRMAR_CODIGO):
        campos[-1].press("Enter")

    try:
        page.wait_for_url(lambda url: "verificacao-mfa" not in url, timeout=15000)
    except PlaywrightTimeout:
        return False, "codigo_recusado"
    return True, None


def _resolver_mfa(page):
    """Chamado quando a Plataforma está na tela de verificação em duas
    etapas. Com PLATAFORMA_CHAVE_AUTENTICADOR configurada, gera o código
    do App Autenticador (TOTP) e conclui a verificação sozinho; sem ela,
    explica o que falta. Devolve True só se o painel aparecer depois."""
    chave = chave_autenticador_configurada("PLATAFORMA")
    if not chave:
        _avisar_mfa(page)
        print("      Pra o robô responder essa verificação sozinho, configure o App")
        print("      Autenticador na Plataforma e a variável PLATAFORMA_CHAVE_AUTENTICADOR")
        print("      (passo a passo no README, seção Login).")
        return False
    print("      [MFA] Verificação em duas etapas - respondendo com o App Autenticador...")
    motivo = None
    for tentativa in (1, 2):
        ok, motivo = _responder_mfa_autenticador(page, chave)
        if ok:
            print("      [MFA] Verificação concluída.")
            selecionar_cliente_banco_b(page, timeout=8000)
            estado, codigo = aguardar_painel_ou_erro(page, timeout=TIMEOUT_LONGO)
            if estado == "painel":
                return True
            motivo = f"erro_{codigo}" if estado == "erro" else f"sem_painel_depois_do_mfa_{estado}"
            break
        if motivo != "codigo_recusado" or tentativa == 2:
            break
        # o código vale por janelas de 30s - se o relógio do Windows
        # estiver um pouco adiantado/atrasado, o próximo costuma passar
        espera = 30 - (time.time() % 30) + 1
        print(f"      Código recusado - esperando o próximo ({espera:.0f}s) e tentando de novo...")
        page.wait_for_timeout(espera * 1000)
    print(f"      [MFA] Não consegui concluir a verificação com o App Autenticador ({motivo}).")
    if motivo == "codigo_recusado":
        print("      Confira se a chave é a mesma cadastrada na Plataforma e se o relógio")
        print("      do Windows está certo (Configurações > Hora > Sincronizar agora).")
    _salvar_evidencia_falha_login(page, f"mfa_{motivo}")
    _descrever_tela_mfa(page)
    return False


def _tentar_login_automatico(page):
    """Tenta logar sozinho na Plataforma com PLATAFORMA_USUARIO /
    PLATAFORMA_SENHA. Só devolve True se o painel realmente aparecer -
    a Plataforma ainda seleciona cliente e pode pedir verificação de
    segurança, então confirmar de verdade é essencial."""
    usuario, senha = credenciais_configuradas("PLATAFORMA")
    if not usuario:
        print("      [AVISO] PLATAFORMA_USUARIO / PLATAFORMA_SENHA não estão definidas NESTE terminal -")
        print("      sem elas não dá pra logar sozinho. Se você já rodou o")
        print("      SetEnvironmentVariable, feche e reabra o VS Code inteiro.")
        return False
    print("      Sessão expirada - tentando login automático...")
    if pagina_pede_mfa(page):
        # já está na tela de MFA (a sessão salva foi aceita até a seleção
        # de cliente e a Plataforma pediu o segundo fator) - não adianta
        # procurar formulário de senha nem esperar painel.
        return _resolver_mfa(page)
    if codigo_pagina_de_erro(page):
        # a tela de erro não tem formulário; o link "plataforma-laudos.example.com" dela
        # leva pro login - é o mesmo clique que um humano daria.
        try:
            abrir_pagina(page, LOGIN_URL)
            page.wait_for_timeout(2000)
        except Exception:
            pass
    if not login_automatico(page, usuario, senha):
        # sem formulário de login E sem tela de erro: pode ser só a página
        # logada demorando pra desenhar o painel (mesmo caso real visto no
        # Banco B em 22/09/2026 - sessão válida, chegou a "selecionar o
        # cliente Banco B" antes de cair aqui, e "não achei o formulário"
        # só quer dizer que não tinha formulário mesmo, porque já estava
        # logado). Dá mais um tempo antes de declarar a sessão vencida.
        estado, _codigo = aguardar_painel_ou_erro(page, timeout=30000)
        if estado == "painel":
            print("      O painel apareceu (só estava lento) - a sessão salva ainda vale.")
            return True
        if estado == "mfa":
            # caso real de 22/09/2026 12:19: o print salvo mostrou a URL
            # index.html#/verificacao-mfa - não era formulário sumido, era
            # a Plataforma pedindo o segundo fator.
            return _resolver_mfa(page)
        print(f"      [AVISO] Não achei o formulário de login pra preencher sozinho (URL: {page.url}).")
        _salvar_evidencia_falha_login(page, "sem_formulario")
        return False
    page.wait_for_timeout(3000)
    selecionar_cliente_banco_b(page)
    estado, codigo = aguardar_painel_ou_erro(page, timeout=TIMEOUT_LONGO)
    if estado == "erro":
        print(f"      [AVISO] Preenchi o login e a Plataforma devolveu a tela de erro {codigo}.")
        _salvar_evidencia_falha_login(page, f"erro_{codigo}")
        return False
    if estado == "mfa":
        return _resolver_mfa(page)
    if estado != "painel":
        print("      [AVISO] Preenchi o login mas não cheguei no painel "
              "(credencial errada, ou a página não respondeu a tempo).")
        _salvar_evidencia_falha_login(page, "sem_painel")
        return False
    return True


def piso_configurado():
    """Devolve o piso de data (datetime) configurado, ou None se não
    houver piso / se o texto configurado não for uma data válida."""
    piso_texto = os.environ.get("DATA_MINIMA", DATA_MINIMA_PADRAO).strip()
    if not piso_texto:
        return None
    try:
        return datetime.strptime(piso_texto, "%d/%m/%Y")
    except ValueError:
        print(f"      [AVISO] DATA_MINIMA='{piso_texto}' não é uma data dd/mm/aaaa "
              f"válida - ignorando o piso nesta execução.")
        return None


def data_inicial_sugerida(dias):
    """Data de início de uma janela de N dias terminando hoje, já
    respeitando o piso - é o que a execução manual oferece no ENTER."""
    inicio = datetime.now() - timedelta(days=dias - 1)
    piso = piso_configurado()
    if piso and inicio < piso:
        inicio = piso
    return inicio.strftime("%d/%m/%Y")


def aplicar_data_minima(data_inicio):
    """Impede que o período comece antes do piso configurado. Devolve a
    data já corrigida (avisando na tela quando corrige), ou a original se
    não houver piso / se ela já for posterior a ele.

    Data digitada inválida: não mexe em nada - quem valida o formato é a
    própria tela da Plataforma, não vale a pena inventar erro aqui."""
    piso = piso_configurado()
    if piso is None:
        return data_inicio
    try:
        pedida = datetime.strptime(data_inicio, "%d/%m/%Y")
    except ValueError:
        return data_inicio
    if pedida >= piso:
        return data_inicio
    piso_texto = piso.strftime("%d/%m/%Y")
    print(f"      Data inicial {data_inicio} é anterior ao piso configurado "
          f"({piso_texto}) - usando o piso.")
    return piso_texto


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    os.makedirs(LOG_DIR, exist_ok=True)

    log_path = os.path.join(LOG_DIR, f"execucao_{datetime.now():%Y%m%d_%H%M%S}.txt")
    log_file = open(log_path, "w", encoding="utf-8")
    stdout_original = sys.stdout
    sys.stdout = Tee(stdout_original, log_file)

    try:
        print("=" * 60)
        print(" EXPORTADOR DE STATUS - BANCO B (PLATAFORMA)")
        print("=" * 60)
        print(f"Log desta execução: {log_path}")

        if modo_automatico():
            dias_texto = os.environ.get("JANELA_DIAS", "").strip()
            dias = int(dias_texto) if dias_texto.isdigit() and int(dias_texto) > 0 else DIAS_JANELA_AUTOMATICA
            data_inicio = data_inicial_sugerida(dias)
            print(f"[MODO_EXECUCAO=auto] Data inicial calculada sozinha: {data_inicio} (janela de {dias} dias)")
        else:
            # sugere a mesma janela do modo automático: assim a execução
            # manual não fica mais "curta" que a agendada por engano (uma
            # janela curta demais perde laudos pra sempre - ver o comentário
            # em DIAS_JANELA_AUTOMATICA).
            sugestao = data_inicial_sugerida(DIAS_JANELA_AUTOMATICA)
            data_inicio = pedir_dado(
                f"Data inicial (dd/mm/aaaa) - ENTER usa {sugestao} "
                f"(últimos {DIAS_JANELA_AUTOMATICA} dias, recomendado)"
            )
            if not data_inicio:
                data_inicio = sugestao
                print(f"      Usando a data sugerida: {data_inicio}")

        data_inicio = aplicar_data_minima(data_inicio)
        data_fim = datetime.now().strftime("%d/%m/%Y")
        print(f"Período usado: {data_inicio} até {data_fim} (hoje)")  # input() não vai pro log sozinho
        print(f"Filtros aplicados: Tipo de Inspeção = {', '.join(TIPOS_INSPECAO)} | "
              f"Status = {', '.join(STATUS_INSPECAO)}")

        caminho_saida = None

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
                print("\n[1/5] Verificando sessão salva...")
                abrir_pagina(page, LOGIN_URL)
                selecionar_cliente_banco_b(page)

                # a Plataforma manda pra uma tela de erro dela (401 ACESSO
                # NÃO AUTORIZADO) em vez de redirecionar pro login - mas ela
                # não vem na hora: a SPA carrega, faz uma chamada de API por
                # trás e só então redireciona. Por isso espera até 10s pelo
                # painel OU pelo erro, o que vier primeiro, em vez de checar
                # só uma vez logo depois do goto (isso já tratou uma sessão
                # válida como morta, 22/09/2026 - ver aguardar_painel_ou_erro).
                estado, codigo_erro = aguardar_painel_ou_erro(page, timeout=10000)
                if estado == "erro":
                    print(f"      A Plataforma respondeu com a tela de erro {codigo_erro} "
                          f"(ACESSO NÃO AUTORIZADO) - a sessão salva não vale mais.")
                    print("      Descartando a sessão salva e recomeçando do zero...")
                    descartar_sessao_salva()
                    # contexto NOVO, sem storage_state: a Plataforma é uma SPA
                    # e guarda o token no localStorage, não (só) em cookie -
                    # um clear_cookies() deixava o token velho lá e o reload
                    # devolvia o mesmo 401. Só um contexto limpo chega de
                    # fato no formulário de login.
                    try:
                        context.close()
                    except Exception:
                        pass
                    context, page = novo_contexto_pagina(browser, com_sessao=False)
                    sessao_existente = False
                    abrir_pagina(page, LOGIN_URL)
                    selecionar_cliente_banco_b(page)
                    estado, codigo_erro = aguardar_painel_ou_erro(page, timeout=10000)
                    if estado == "erro":
                        print(f"\n[ERRO] Mesmo começando do zero a Plataforma devolveu a tela "
                              f"de erro {codigo_erro}.")
                        print("Isso é do lado da Plataforma, não do robô - normalmente é o")
                        print("usuário sem permissão no cliente Banco B, ou o acesso suspenso.")
                        print("Abra plataforma-laudos.example.com no seu navegador normal e confira se você")
                        print("consegue entrar e ver o GRID DE INSPEÇÃO por lá.")
                        return

                if sessao_existente and estado == "painel":
                    print("      Sessão válida, painel carregado direto.")
                elif _tentar_login_automatico(page):
                    salvar_sessao()
                    print("      Login automático OK - sessão renovada e salva.")
                elif HEADLESS or modo_automatico():
                    if pagina_pede_mfa(page):
                        # não adianta sugerir credencial: ela já foi aceita.
                        print("\n[ERRO] A Plataforma pediu verificação em duas etapas (MFA) e o robô")
                        print("não conseguiu responder sozinho (veja as linhas [MFA] acima).")
                        print("Solução definitiva: App Autenticador + PLATAFORMA_CHAVE_AUTENTICADOR")
                        print("(README, seção Login). Pra destravar hoje, rode com a janela visível")
                        print("e digite o código na mão:")
                        print(f'  $env:HEADLESS="false"; $env:MODO_EXECUCAO=$null; python {os.path.basename(__file__)}')
                        return
                    print("\n[ERRO] A sessão salva expirou e não consegui logar sozinho.")
                    print("Pra rodar sem ninguém na frente, configure as credenciais uma vez:")
                    print('  [Environment]::SetEnvironmentVariable("PLATAFORMA_USUARIO", "voce@exemplo.com", "User")')
                    print('  [Environment]::SetEnvironmentVariable("PLATAFORMA_SENHA", "suasenha", "User")')
                    print("Ou rode com a janela visível e logue na mão:")
                    print(f'  $env:HEADLESS="false"; python {os.path.basename(__file__)}')
                    return
                else:
                    if pagina_pede_mfa(page):
                        print("      A Plataforma está pedindo verificação em duas etapas (MFA).")
                        print("      A senha já foi aceita - só falta o código, que precisa ser digitado")
                        print("      na janela do Chrome.")
                    if codigo_pagina_de_erro(page):
                        print(f"      O Chrome está na tela de erro {codigo_pagina_de_erro(page)} da Plataforma")
                        print("      (sem formulário). Clique no link 'plataforma-laudos.example.com' embaixo do erro")
                        print("      pra chegar no login - o resto é igual ao de sempre:")
                    pausar_para_usuario(
                        "A sessão salva não está mais válida (ou pediu login/",
                        "verificação de novo).",
                        "1. Faça o login manualmente na janela do Chrome (e-mail e senha).",
                        "2. Selecione o cliente Banco B, se for pedido (o robô já",
                        "   tenta fazer isso sozinho, mas confirme se ele conseguiu).",
                        "3. Complete a verificação de segurança, se pedir.",
                        "4. Espere até ver a tela com 'GRID DE INSPEÇÃO'.",
                    )
                    estado, codigo_erro = aguardar_painel_ou_erro(page, timeout=10000)
                    # Caso real de 23/09/2026 10:01: o usuário apertou ENTER no
                    # terminal com o Chrome ainda na tela de MFA (tinha escolhido
                    # o e-mail, o código ainda não tinha sido digitado) e o robô
                    # encerrou na primeira checagem. Enquanto for MFA ou "nada",
                    # a pessoa está no meio do caminho - dá mais chances em vez
                    # de jogar a execução fora.
                    tentativas_restantes = 5
                    while estado in ("mfa", "nada") and tentativas_restantes > 0:
                        tentativas_restantes -= 1
                        if estado == "mfa":
                            pausar_para_usuario(
                                "O Chrome AINDA está na tela de verificação em duas etapas (MFA).",
                                "Termine a verificação por lá (digite o código que chegou e",
                                "confirme) e só aperte ENTER aqui quando aparecer o",
                                "'GRID DE INSPEÇÃO'.",
                            )
                        else:
                            pausar_para_usuario(
                                "Ainda não vi o 'GRID DE INSPEÇÃO' no Chrome.",
                                "Espere ele aparecer e aperte ENTER de novo.",
                            )
                        estado, codigo_erro = aguardar_painel_ou_erro(page, timeout=10000)
                    if estado != "painel":
                        if estado == "erro":
                            print(f"\n[ERRO] Depois do login o Chrome continua na tela de erro {codigo_erro}.")
                        elif estado == "mfa":
                            print("\n[ERRO] O Chrome continua na tela de verificação em duas etapas (MFA) -")
                            print("a verificação não foi concluída.")
                        else:
                            print("\n[ERRO] Ainda não consegui ver o painel principal (GRID DE INSPEÇÃO).")
                        print("Encerrando esta execução.")
                        return
                    salvar_sessao()
                    print("      Sessão salva para as próximas execuções.")

                print("[2/5] Navegando até Relatório Analítico...")
                print("[3/5] Preenchendo filtros/período e clicando em Exibir...")
                preencher_periodo_e_exibir(page, data_inicio, data_fim)

                print("[4/5] Clicando em Exportar para Excel...")
                download = clicar_exportar_excel(page)

                extensao = os.path.splitext(download.suggested_filename or "")[1] or ".xlsx"
                data_inicio_arq = data_inicio.replace("/", "-")
                data_fim_arq = data_fim.replace("/", "-")
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                nome_arquivo = f"laudos_banco_b_{data_inicio_arq}_a_{data_fim_arq}_{timestamp}{extensao}"
                caminho_saida = os.path.join(OUTPUT_DIR, nome_arquivo)
                download.save_as(caminho_saida)

                print("      Removendo aba de metadados do Excel...")
                remover_aba_informacoes(caminho_saida)

                total_baixado = contar_linhas_excel(caminho_saida)
                if total_baixado is not None:
                    print(f"      Laudos baixados no Excel: {total_baixado}")

                print("[5/5] Conferindo Canceladas sem vistoria (abre cada uma e olha a aba Laudos)...")
                resultado_canceladas = ler_canceladas_do_excel(caminho_saida)
                if resultado_canceladas is not None:
                    ids_canceladas, ids_com_data_vistoria = resultado_canceladas
                    print(f"      {len(ids_canceladas)} Cancelada(s) no Excel a conferir"
                          f"{': ' + ', '.join(sorted(ids_canceladas)) if ids_canceladas else ''}.")
                else:
                    ids_canceladas, ids_com_data_vistoria = None, set()
                if ids_canceladas == set():
                    # nenhuma Cancelada no período: todo o resto é laudo aceito,
                    # que nunca seria aceito sem vistoria - nada a conferir.
                    print("      Nenhuma Cancelada no período - pulando a conferência (nada a checar).")
                    ids_sem_vistoria = set()
                else:
                    ids_sem_vistoria = identificar_canceladas_sem_vistoria(
                        page, data_inicio, data_fim, ids_canceladas=ids_canceladas,
                        ids_com_data_vistoria=ids_com_data_vistoria,
                    )
                remocao_confiavel = remover_propostas_sem_vistoria(caminho_saida, ids_sem_vistoria)

                total_final = contar_linhas_excel(caminho_saida)
                if total_final is not None:
                    print(f"      Laudos no Excel final (depois de remover Cancelada sem vistoria): {total_final}")

                if not remocao_confiavel:
                    # o arquivo existe em disco (pra conferência manual), mas
                    # main() não devolve o caminho - as próximas etapas
                    # (cadastro, atribuir engenheiro) nunca devem usar um
                    # Excel que pode ainda ter Cancelada sem vistoria dentro.
                    print("\n[ERRO] Não confiei na remoção de Cancelada sem vistoria - encerrando sem devolver o arquivo.")
                    caminho_saida = None

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
                browser.close()

        if caminho_saida:
            print("\nConcluído!")
            print("-" * 60)
            print(f"  Arquivo exportado: {caminho_saida}")
            total_final = contar_linhas_excel(caminho_saida)
            if total_final is not None:
                print(f"  Total de laudos no Excel final: {total_final}")
            print("-" * 60)
        else:
            print("\nNada foi exportado (ou o arquivo exportado não é seguro pra usar automaticamente - veja o [ERRO] acima).")

    finally:
        sys.stdout = stdout_original
        log_file.close()

    return caminho_saida


if __name__ == "__main__":
    main()
