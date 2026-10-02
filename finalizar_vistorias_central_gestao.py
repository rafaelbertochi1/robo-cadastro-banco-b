"""
Etapa 5: marca o checkbox de cada proposta que já está COMPLETA na tela
de "Em Aberto" (engenheiro de Gestão atribuído + Dt Ag. + Dt Vist.
preenchidas) e clica no botão verde "Finalizar Selecionados", mandando
todas elas de uma vez pra próxima etapa.

Adaptado do robô irmão que já faz essa mesma etapa pro Banco A
(robo-cadastro-banco-a). Os seletores abaixo foram conferidos um a um
contra o HTML real do Banco B (mapear_central_gestao.py, 17/09/2026):

    #table (tabela)                       Banco A usa #tabelaBancoA
    input.checkbox-proposta               idêntico ao Banco A
    #formFinalizarSelecionados            idêntico ao Banco A
    botão verde: title="Finalizar Selecionados"
                 onclick="return confirmarFinalizacao()"    idêntico

    ⚠️ action do formulário: https://central-gestao.example.com/gestao-banco_b/
    atualizar-status - REPARE O HÍFEN em "gestao-banco_b", diferente
    do resto das URLs do Banco B (que usam "gestao-banco-b", sem
    hífen) e diferente do Banco A (que usa "gestao-banco-a/atualizar-status",
    também sem hífen). Confirmado na captura real - não precisou ser
    escrito no código porque o robô nunca monta essa URL na mão: só
    clica no botão da própria tela, que já aponta pro lugar certo.

Diferente das outras etapas, essa NÃO é guiada pelo Excel: o critério é
lido da própria tela. Assim ela pega qualquer proposta que esteja
pronta, inclusive de lotes anteriores, e nunca finaliza uma que ainda
esteja incompleta - mesmo que o Excel desta rodada dissesse que estava
tudo certo.

Mecânica (confirmada no Banco A, seletores conferidos contra o HTML real
do Banco B):

1. Cada linha da tabela #table tem, na coluna "Ações", um
   input.checkbox-proposta cujo value é o id interno da linha (o mesmo
   id usado nos modais das Etapas 3 e 4).
2. O botão verde fica no formulário #formFinalizarSelecionados, com
   title="Finalizar Selecionados" e onclick="return confirmarFinalizacao()".
3. Esse confirmarFinalizacao() lê document.querySelectorAll(
   '.checkbox-proposta:checked'). Se nenhum estiver marcado, mostra um
   alert e bloqueia o envio - não existe envio vazio. Se tiver marcado,
   chama confirm('Tem certeza que deseja finalizar as propostas
   selecionadas?') - só monta os inputs hidden ids[] e deixa o
   formulário seguir se o usuário confirmar.

   ⚠️ Esse confirm() é um diálogo NATIVO do navegador (não é um modal
   HTML) - confirmado no JS real capturado da tela (18/09/2026). Sem um
   listener de 'dialog' registrado na página, o Playwright DESCARTA
   automaticamente qualquer diálogo desses (equivale a clicar
   Cancelar): o clique no botão não dá erro nenhum, mas o formulário
   nunca chega a ser enviado de verdade - foi exatamente o que aconteceu
   numa execução real (184 propostas confirmadas na prévia, 0
   finalizadas, 5 minutos de timeout esperando uma navegação que nunca
   ia vir). Corrigido registrando um handler que aceita automaticamente
   qualquer diálogo desta página (ver _aceitar_dialogo).

   Como confirmarFinalizacao() só olha o estado :checked (nenhum
   listener de evento), marcar os checkboxes direto no DOM funciona
   igual a clicar um por um - e é o que este script faz, porque a
   tabela mantém todas as linhas no HTML mas esconde a maioria
   (paginação do List.js), e o Playwright não clica em elemento
   invisível.
4. Não existe "marcar todos" no cabeçalho da tabela - por isso a marcação
   é feita linha a linha.

Critério de "completa" (as 3 condições, lidas da própria linha):
  - coluna "Eng" -> a parte "Gestão:" tem um nome de verdade (quando não
    tem engenheiro, o sistema mostra um traço "-" - confirmado na
    captura real do Banco B);
  - coluna "Dt Ag."  -> preenchida;
  - coluna "Dt Vist." -> preenchida.
Os índices dessas colunas são descobertos lendo o <thead> em tempo de
execução (não são fixos no código), pelo mesmo motivo que a leitura do
Excel procura a coluna pelo cabeçalho em vez de pela posição.

SEGURANÇA (finalizar tira a proposta de "Em Aberto" - é uma mudança de
estado de verdade, em lote):
- Prévia obrigatória: lista proposta por proposta o que seria finalizado
  (e um resumo do que ficou de fora e por quê) ANTES de marcar qualquer
  coisa, e só segue se você digitar CONFIRMAR.
- Ao marcar, desmarca explicitamente todo o resto da página - nada fora
  da lista aprovada pode ir junto por acidente.
- Antes de clicar no botão verde, confere que a quantidade de
  checkboxes realmente marcados na página bate com a quantidade
  aprovada; se não bater, aborta sem clicar.
- Depois de finalizar, relê a lista e confirma que as propostas sumiram
  de "Em Aberto" - as que continuarem aparecendo viram VERIFICAR.

SEM PARALELISMO, de propósito (mesma decisão do Banco A): aqui não existe o
custo por proposta que justifica as guias nas Etapas 3 e 4 (marcar
checkbox não recarrega a página). É um único envio em lote, e o estado
dos checkboxes vive numa página só - várias guias enviariam
finalizações parciais separadas, sem ganho nenhum de tempo.

Login automático: com CENTRAL_USUARIO/CENTRAL_SENHA configurados, tenta logar
sozinho quando a sessão salva expira (mesmo mecanismo das outras
etapas, ver comum.py).

⚠️ Ainda não rodou contra o sistema real do Banco B - os seletores
foram conferidos contra uma captura real da tela (não é chute), mas
precisa do primeiro teste real (com $env:HEADLESS="false") pra validar
antes de confiar nele com volume.

Uso: python finalizar_vistorias_central_gestao.py
"""

import os
import sys
from datetime import datetime

from openpyxl import Workbook
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout

from comum import (
    abrir_pagina,
    Tee,
    modo_automatico,
    pausar_para_usuario,
    pedir_confirmacao,
    headless_configurado,
    achar_excel_mais_recente,
    espera_carregamento_gestao,
    bloquear_recursos_visuais,
    url_sem_cache,
    abrir_gestao,
    credenciais_configuradas,
    login_automatico,
)

LOGIN_URL = "https://central-gestao.example.com/login"
GESTAO_BANCO_B_URL = "https://central-gestao.example.com/gestao-banco-b"
PALAVRA_CHAVE_CLIENTE = "banco_b"
PASTA_SCRIPT = os.path.dirname(os.path.abspath(__file__))
EXPORTS_DIR = os.path.join(PASTA_SCRIPT, "data", "exports")
# mesmo arquivo de sessão dos outros scripts da Central de Gestão.
SESSION_FILE = os.path.join(PASTA_SCRIPT, "sessao_central_gestao.json")
LOG_DIR = os.path.join(PASTA_SCRIPT, "logs")

# Caixa "Procurar..." da tela Consultar - mesma confirmada em
# atribuir_engenheiro_central_gestao.py (não tem name="search" no
# Banco B, diferente do Banco A).
CAMPO_BUSCA = "input.search[type='search']"

# como o sistema mostra a coluna Eng quando não há engenheiro de Gestão
# atribuído (confirmado na captura real do Banco B: "-").
MARCADORES_SEM_ENGENHEIRO = {"", "-", "--", "não definido", "nao definido"}

TIMEOUT_LONGO = 45000
TIMEOUT_PAGINA_EM_ABERTO = 300000
ESPERA_CARREGAMENTO_GESTAO = espera_carregamento_gestao()

HEADLESS = headless_configurado()

# Lê todas as linhas da tabela de Em Aberto de uma vez só.
#
# Usa textContent (não innerText) de propósito: innerText depende de
# layout e devolve string vazia pra elemento escondido - e a maioria das
# linhas fica com display:none por causa da paginação do List.js, o que
# faria toda linha escondida parecer "incompleta".
JS_LER_LINHAS = """
() => {
    const tabela = document.querySelector('#table');
    if (!tabela) return {erro: 'não achei a tabela #table na página'};

    const limpa = (el) => (el ? el.textContent : '').replace(/\\s+/g, ' ').trim();
    // ':scope >' é essencial: cada linha carrega, dentro do <td> de Ações,
    // os modais das Etapas 3 e 4 - e o modal de engenheiro tem uma tabela
    // própria com <thead>. Sem escopar, o cabeçalho lido vinha com
    // colunas dessas tabelas aninhadas junto (mesmo problema já
    // confirmado no Banco A rodando este JS contra a captura real).
    const cabecalhos = Array.from(tabela.querySelectorAll(':scope > thead th')).map(th => limpa(th));
    const indiceDe = (nome) => cabecalhos.findIndex(
        h => h.toLowerCase().replace(/\\.$/, '') === nome.toLowerCase().replace(/\\.$/, '')
    );

    const iEng = indiceDe('Eng');
    const iAg = indiceDe('Dt Ag.');
    const iVist = indiceDe('Dt Vist.');
    if (iEng < 0 || iAg < 0 || iVist < 0) {
        return {erro: 'não achei as colunas Eng / Dt Ag. / Dt Vist. no cabeçalho. Cabeçalho lido: '
                      + cabecalhos.join(' | ')};
    }

    const linhas = [];
    tabela.querySelectorAll(':scope > tbody.list > tr').forEach(tr => {
        const check = tr.querySelector('input.checkbox-proposta');
        if (!check) return;
        const tds = Array.from(tr.children).filter(el => el.tagName === 'TD');
        if (tds.length <= Math.max(iEng, iAg, iVist)) return;

        // o nome do engenheiro de Gestão vem no último <span> da célula
        // Eng (o primeiro texto é o id do Backoffice, sem span colorido)
        const spans = tds[iEng].querySelectorAll('span');
        const gestao = spans.length ? limpa(spans[spans.length - 1]) : '';

        const celulaProposta = tr.querySelector('td.num_proposta');
        linhas.push({
            id: check.value,
            num_proposta: limpa(celulaProposta).replace(/⭐️|⭐/g, '').trim(),
            engenheiro: gestao,
            dt_ag: limpa(tds[iAg]),
            dt_vist: limpa(tds[iVist]),
        });
    });
    return {linhas: linhas, cabecalhos: cabecalhos};
}
"""

# Marca só os ids aprovados e DESMARCA todo o resto - garante que nada
# fora da lista confirmada seja enviado junto no lote.
JS_MARCAR = """
(ids) => {
    const alvo = new Set(ids);
    let marcados = 0;
    document.querySelectorAll('input.checkbox-proposta').forEach(el => {
        if (alvo.has(el.value)) {
            el.checked = true;
            marcados++;
        } else {
            el.checked = false;
        }
    });
    return {
        marcados: marcados,
        conferencia: document.querySelectorAll('input.checkbox-proposta:checked').length
    };
}
"""


def _aceitar_dialogo(dialog):
    """Aceita automaticamente qualquer diálogo nativo do navegador
    (confirm()/alert()) nesta página.

    Sem um listener de 'dialog' registrado, o Playwright DESCARTA
    qualquer diálogo sozinho - equivale a clicar Cancelar. O botão
    'Finalizar Selecionados' chama confirm('Tem certeza que deseja
    finalizar as propostas selecionadas?') antes de montar e enviar o
    formulário (confirmado no JS real capturado da tela, 18/09/2026) -
    sem aceitar esse diálogo, o clique não dá erro nenhum no Playwright,
    mas o formulário nunca chega a ser enviado: o robô ficava esperando
    uma navegação que nunca ia acontecer até estourar os 5 minutos de
    timeout, e nenhuma proposta era finalizada de verdade (confirmado
    em execução real: 184 propostas confirmadas na prévia, 0
    finalizadas)."""
    print(f"      [diálogo do navegador] '{dialog.message}' - aceitando automaticamente.")
    dialog.accept()


def novo_contexto_pagina(browser, com_sessao):
    context = browser.new_context(
        accept_downloads=True,
        viewport={"width": 1920, "height": 1080} if HEADLESS else None,
        no_viewport=None if HEADLESS else True,
        storage_state=SESSION_FILE if com_sessao else None,
    )
    if HEADLESS:
        bloquear_recursos_visuais(context)
    page = context.new_page()
    page.on("dialog", _aceitar_dialogo)
    return context, page


def _aguardar_pagina_pronta(page):
    """Pausa fixa depois da caixa de busca já estar visível - a página
    "parece" pronta antes de estar (mesmo motivo documentado nos outros
    scripts da Central de Gestão)."""
    print(f"      Aguardando {ESPERA_CARREGAMENTO_GESTAO // 1000}s pra página terminar de carregar de verdade...")
    page.wait_for_timeout(ESPERA_CARREGAMENTO_GESTAO)


def ler_linhas_da_tela(page):
    """Lê todas as linhas da tabela de Em Aberto (inclusive as escondidas
    pela paginação do List.js, que continuam no HTML). Levanta erro se a
    estrutura da tela não for a esperada - nunca devolve lista vazia
    fingindo que "não tem nada pra fazer"."""
    resultado = page.evaluate(JS_LER_LINHAS)
    if resultado.get("erro"):
        raise RuntimeError(resultado["erro"])
    return resultado["linhas"]


def _tem_engenheiro(texto_engenheiro):
    return texto_engenheiro.strip().lower() not in MARCADORES_SEM_ENGENHEIRO


def classificar_linhas(linhas):
    """Separa as linhas entre as COMPLETAS (engenheiro de Gestão + Dt Ag.
    + Dt Vist.) e as que ficam de fora, com o motivo de cada uma."""
    completas = []
    de_fora = []
    for linha in linhas:
        faltando = []
        if not _tem_engenheiro(linha["engenheiro"]):
            faltando.append("engenheiro")
        if not linha["dt_ag"].strip():
            faltando.append("Dt Ag.")
        if not linha["dt_vist"].strip():
            faltando.append("Dt Vist.")

        if faltando:
            de_fora.append({**linha, "motivo": "sem " + " / sem ".join(faltando)})
        else:
            completas.append(linha)
    return completas, de_fora


def explicar_de_fora(de_fora, caminho_excel):
    """Diz, pra cada proposta que ficou de fora, POR QUE ela está
    incompleta - cruzando o que a tela mostra com o que o Excel da
    Plataforma tem. Sem isso o log só dizia "sem Dt Vist." e ninguém sabia
    se era a Etapa 4 que falhou ou se a Plataforma simplesmente ainda não
    tem a data (caso normal: a vistoria ainda não aconteceu).

    Devolve a mesma lista, com 'explicacao' em cada item."""
    com_data = set()
    presentes = set()
    if caminho_excel and os.path.exists(caminho_excel):
        try:
            from agendar_vistoria_central_gestao import ler_agendamentos
            from cadastrar_central_gestao import extrair_numeros_proposta
            com_data = {l["num_proposta"] for l in ler_agendamentos(caminho_excel)}
            presentes = extrair_numeros_proposta(caminho_excel)
        except Exception as e:
            print(f"      [AVISO] Não consegui cruzar com o Excel ({e}) - explicações ficam genéricas.")

    for linha in de_fora:
        p = linha["num_proposta"]
        if "engenheiro" in linha["motivo"]:
            linha["explicacao"] = "sem engenheiro na Gestão - processo interno, descartada"
        elif p in com_data:
            linha["explicacao"] = "ATENÇÃO: está no Excel COM data de vistoria - a Etapa 4 deveria ter agendado, confira"
        elif p in presentes:
            linha["explicacao"] = "Plataforma ainda sem data de vistoria - completa sozinha numa rodada futura"
        elif not presentes:
            linha["explicacao"] = "(sem Excel pra cruzar)"
        else:
            linha["explicacao"] = "não está no export (fora do período) - nada a fazer por aqui"
    return de_fora


def marcar_checkboxes(page, ids_aprovados):
    """Marca só os ids aprovados (e desmarca o resto). Devolve quantos
    ficaram realmente marcados na página, pra conferir antes de enviar."""
    resultado = page.evaluate(JS_MARCAR, ids_aprovados)
    return resultado["conferencia"]


def clicar_finalizar(page):
    """Clica no botão verde 'Finalizar Selecionados' e espera a página
    recarregar. Espera só 'domcontentloaded' (não o padrão 'load') pelo
    mesmo motivo já confirmado nas Etapas 3 e 4: nesta página pesada o
    'load' estoura o timeout mesmo quando o envio funcionou."""
    botao = page.locator("#formFinalizarSelecionados button[type='submit']")
    botao.wait_for(state="visible", timeout=TIMEOUT_LONGO)
    with page.expect_navigation(timeout=TIMEOUT_PAGINA_EM_ABERTO, wait_until="domcontentloaded"):
        botao.click(timeout=TIMEOUT_LONGO)


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
    os.makedirs(EXPORTS_DIR, exist_ok=True)

    log_path = os.path.join(LOG_DIR, f"finalizar_vistorias_{datetime.now():%Y%m%d_%H%M%S}.txt")
    log_file = open(log_path, "w", encoding="utf-8")
    stdout_original = sys.stdout
    sys.stdout = Tee(stdout_original, log_file)

    resultados = []

    try:
        print("=" * 60)
        print(" FINALIZAR VISTORIAS - CENTRAL DE GESTÃO (BANCO B)")
        print("=" * 60)
        print(f"Log desta execução: {log_path}")

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
                    if "/login" in page.url:
                        try:
                            page.wait_for_url(lambda url: "/login" not in url, timeout=15000)
                        except PlaywrightTimeout:
                            pass
                    if "/login" in page.url:
                        print("\n[ERRO] Ainda estou vendo a tela de login.")
                        return
                    salvar_sessao()
                    print("      Sessão salva para as próximas execuções.")

                print("[2/5] Abrindo Vistorias > Banco B > Em Aberto (Consultar)...")
                url_gestao = abrir_gestao(
                    page, GESTAO_BANCO_B_URL, PALAVRA_CHAVE_CLIENTE, timeout=TIMEOUT_PAGINA_EM_ABERTO
                )
                if url_gestao is None:
                    print("\n[ERRO] Não consegui chegar na tela de gestão do Banco B.")
                    return
                if "/login" in page.url:
                    print("\n[ERRO] Cheguei na tela de gestão mas voltei pra tela de login.")
                    return
                page.wait_for_selector(CAMPO_BUSCA, timeout=TIMEOUT_PAGINA_EM_ABERTO)
                _aguardar_pagina_pronta(page)

                print("[3/5] Lendo as linhas da tela e vendo quais estão completas...")
                linhas = ler_linhas_da_tela(page)
                print(f"      {len(linhas)} proposta(s) em Em Aberto.")

                completas, de_fora = classificar_linhas(linhas)

                print("\n" + "-" * 60)
                print(" PRÉVIA - nada foi marcado nem finalizado ainda")
                print("-" * 60)
                for linha in completas:
                    print(f"  [FINALIZAR] Proposta {linha['num_proposta']}: "
                          f"{linha['engenheiro']} | Ag. {linha['dt_ag']} | Vist. {linha['dt_vist']}")
                print("-" * 60)
                print(f"  {len(completas)} completa(s) seriam finalizadas, {len(de_fora)} ficariam de fora.")

                if de_fora:
                    de_fora = explicar_de_fora(de_fora, achar_excel_mais_recente(EXPORTS_DIR))
                    por_explicacao = {}
                    for linha in de_fora:
                        por_explicacao.setdefault(linha["explicacao"], []).append(linha)
                    print("\n  Quem fica de fora, e por quê (nenhuma delas é tocada):")
                    for explicacao, grupo in sorted(por_explicacao.items(), key=lambda x: -len(x[1])):
                        print(f"\n    {len(grupo):3d}x  {explicacao}")
                        for linha in grupo:
                            print(f"           {linha['num_proposta']}  ({linha['motivo']})")

                for linha in de_fora:
                    resultados.append({**linha, "resultado": "PULADO",
                                       "detalhe": f"{linha['motivo']} | {linha.get('explicacao', '')}"})

                if not completas:
                    print("\nNenhuma proposta completa - encerrando sem marcar nada.")
                    return

                confirmado = pedir_confirmacao(
                    f"{len(completas)} proposta(s) completas vão ser FINALIZADAS de verdade agora",
                    "(saem de 'Em Aberto' e vão pra próxima etapa).",
                    f"{len(de_fora)} incompletas não serão tocadas.",
                )
                if not confirmado:
                    print("\nCancelado pelo usuário - nada foi marcado nem finalizado.")
                    return

                print("\n[4/5] Marcando os checkboxes das propostas completas...")
                ids_aprovados = [linha["id"] for linha in completas]
                marcados = marcar_checkboxes(page, ids_aprovados)
                print(f"      {marcados} checkbox(es) marcados na página.")

                if marcados != len(ids_aprovados):
                    print(f"\n[ERRO] Eu esperava {len(ids_aprovados)} marcados, mas a página ficou com {marcados}.")
                    print("Não vou clicar em Finalizar com a seleção diferente do que você confirmou.")
                    print("Rode de novo - se repetir, me manda este log.")
                    for linha in completas:
                        resultados.append({**linha, "resultado": "ERRO",
                                           "detalhe": f"seleção não bateu ({marcados} de {len(ids_aprovados)}) - nada foi enviado"})
                    return

                print("[5/5] Clicando em 'Finalizar Selecionados' e reconferindo...")
                clicar_finalizar(page)

                # reconferência: as finalizadas devem sumir de Em Aberto
                page.wait_for_selector(CAMPO_BUSCA, timeout=TIMEOUT_PAGINA_EM_ABERTO)
                _aguardar_pagina_pronta(page)
                try:
                    linhas_depois = ler_linhas_da_tela(page)
                    ids_ainda_presentes = {linha["id"] for linha in linhas_depois}
                except Exception as e:
                    print(f"      [AVISO] Enviei, mas deu erro técnico relendo a tela ({e}) - confira manualmente.")
                    for linha in completas:
                        resultados.append({**linha, "resultado": "VERIFICAR",
                                           "detalhe": f"enviado, mas erro técnico ao reconferir: {e}"})
                    ids_ainda_presentes = None

                if ids_ainda_presentes is not None:
                    for linha in completas:
                        if linha["id"] in ids_ainda_presentes:
                            resultados.append({**linha, "resultado": "VERIFICAR",
                                               "detalhe": "continua aparecendo em Em Aberto depois de finalizar"})
                        else:
                            resultados.append({**linha, "resultado": "OK",
                                               "detalhe": "finalizada (saiu de Em Aberto)"})

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

        if resultados:
            caminho_resultado = os.path.join(
                EXPORTS_DIR, f"resultado_finalizacoes_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            )
            wb_resultado = Workbook()
            ws_resultado = wb_resultado.active
            ws_resultado.title = "Resultado"
            ws_resultado.append(["Nº Proposta", "Engenheiro (Gestão)", "Dt Ag.", "Dt Vist.", "Resultado", "Detalhe"])
            for r in resultados:
                ws_resultado.append([
                    r["num_proposta"],
                    r["engenheiro"],
                    r["dt_ag"],
                    r["dt_vist"],
                    r["resultado"],
                    r["detalhe"],
                ])
            wb_resultado.save(caminho_resultado)

            total_ok = sum(1 for r in resultados if r["resultado"] == "OK")
            total_verificar = sum(1 for r in resultados if r["resultado"] == "VERIFICAR")
            total_erro = sum(1 for r in resultados if r["resultado"] == "ERRO")
            total_pulado = sum(1 for r in resultados if r["resultado"] == "PULADO")
            print("\nConcluído!")
            print("-" * 60)
            print(f"  {total_ok} finalizada(s) e confirmada(s), {total_verificar} enviada(s) mas não confirmada(s) "
                  f"(confira na tela), {total_erro} com erro de verdade, "
                  f"{total_pulado} pulada(s) por estarem incompletas.")
            print(f"  Resultado detalhado: {caminho_resultado}")
            print("-" * 60)

    finally:
        sys.stdout = stdout_original
        log_file.close()


if __name__ == "__main__":
    main()
