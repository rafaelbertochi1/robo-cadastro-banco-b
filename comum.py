"""Funções compartilhadas entre os robôs deste projeto (Plataforma e
sistema de gestão Central de Gestão)."""

import base64
import glob
import hashlib
import hmac
import os
import re
import struct
import time
from urllib.parse import parse_qs, urlparse

from openpyxl import load_workbook


def headless_configurado():
    """Lê a variável de ambiente HEADLESS (true/false) em vez de exigir
    editar o .py na mão - assim dá pra rodar com o navegador visível sem
    criar uma alteração local que trava o próximo `git pull`.

    PowerShell (um único uso):   $env:HEADLESS="false"; python arquivo.py
    PowerShell (sessão inteira): $env:HEADLESS="false"
    Sem definir a variável, roda com HEADLESS = True (padrão)."""
    valor = os.environ.get("HEADLESS", "true").strip().lower()
    return valor not in ("false", "0", "nao", "não")


def modo_automatico():
    """Lê a variável de ambiente MODO_EXECUCAO (mesmo padrão do
    HEADLESS - variável de ambiente, nunca editar o .py na mão).

    'auto' faz o robô calcular tudo sozinho (data inicial da Etapa 1) e
    confirmar tudo sozinho (nenhum pedir_confirmacao pausa esperando
    resposta) - pensado pro fluxo agendado (ex.: Agendador de Tarefas
    do Windows rodando de 6 em 6h).

    Sem definir a variável (ou qualquer outro valor), roda no modo
    padrão: pede a data e exige digitar CONFIRMAR antes de qualquer
    ação real - use esse modo até confiar no fluxo inteiro.

    PowerShell (um único uso):   $env:MODO_EXECUCAO="auto"; python arquivo.py"""
    valor = os.environ.get("MODO_EXECUCAO", "confirmar").strip().lower()
    return valor == "auto"


def espera_carregamento_gestao(padrao_ms=10000):
    """Lê a variável de ambiente ESPERA_CARREGAMENTO_GESTAO_MS pra
    ajustar (sem editar o .py) a pausa fixa usada depois de carregar a
    tela de Em Aberto na Central de Gestão - ela existe porque a página
    "parece" pronta antes de estar de verdade (confirmado travando em
    teste real com um valor menor).

    Calibragem: 45s era um chute inicial a partir de "30s a 1 minuto"
    estimado pelo usuário, nunca medido. 25s rodou sem nenhuma falha em
    produção (18/09/2026) nas Etapas 4 e 5 - inclusive na Etapa 4 com 6
    guias em paralelo, que é o caso mais pesado. O padrão desceu pra
    20s (21/09/2026) porque 45s estava custando muito tempo de relógio:
    essa pausa roda a cada carregamento de página, ou seja, uma vez por
    proposta nas Etapas 3 e 4. Desceu de novo pra 10s (24/09/2026)
    depois de uma execução completa em produção com 10s (Etapas 2 a 5,
    6 guias em paralelo) rodar sem nenhuma falha - metade do tempo de
    espera por proposta.

    Se um dia a tela voltar a travar (erro de "não consegui localizar a
    proposta" ou timeout no Salvar aparecendo em série), suba de novo
    sem mexer no código:
    PowerShell (um único uso): $env:ESPERA_CARREGAMENTO_GESTAO_MS="45000"; python arquivo.py
    Sem definir a variável, usa `padrao_ms` (10000 = 10s)."""
    valor = os.environ.get("ESPERA_CARREGAMENTO_GESTAO_MS", "").strip()
    if valor.isdigit() and int(valor) > 0:
        return int(valor)
    return padrao_ms


def url_sem_cache(url):
    """Acrescenta um parâmetro que muda a cada chamada (timestamp), pra
    forçar o navegador a buscar a página no servidor em vez de servir
    uma cópia em cache.

    Isso não é preciosismo: já foi confirmado ao vivo neste mesmo
    sistema, no robô irmão do Banco A, que a Central de Gestão pode devolver
    página em cache mesmo quando o robô "recarrega" - lá isso fez a
    reconferência depois de salvar ler a versão ANTERIOR ao salvamento
    (atribuições bem-sucedidas sendo lidas como "não confirmadas").
    Aplicado por precaução nas mesmas navegações onde isso mordeu no
    Banco A, já que é o mesmo sistema por trás."""
    separador = "&" if "?" in url else "?"
    return f"{url}{separador}_={int(time.time() * 1000)}"


def paralelismo_configurado(variavel, padrao=6):
    """Lê uma variável de ambiente (o nome vem em `variavel`) pra ajustar
    (sem editar o .py) quantas guias trabalham em paralelo numa etapa que
    escreve de verdade na Central de Gestão - mesma estratégia validada no
    robô irmão do Banco A (lá, cada guia abre seu próprio navegador
    Playwright; é a forma segura de paralelizar com a API síncrona).

    Padrão 6 (24/09/2026; antes era 3, conservador): cada guia clica em
    Salvar de verdade, então o número só subiu depois de validado - 6
    guias rodaram sem falha em produção nas Etapas 3 e 4 do Banco B
    (18/09, 23/09 e 24/09/2026, inclusive com a espera de 10s). Mais que
    isso não é recomendado: no Banco A, 12 guias saturaram a tela real.

    PowerShell (um único uso): $env:PARALELISMO_ATRIBUICAO="6"; python arquivo.py
    Sem definir a variável, usa `padrao` (6)."""
    valor = os.environ.get(variavel, "").strip()
    if valor.isdigit() and int(valor) > 0:
        return int(valor)
    return padrao


def _inteiro_configurado(variavel, padrao):
    valor = os.environ.get(variavel, "").strip()
    if valor.isdigit():
        return int(valor)
    return padrao


def abrir_pagina(page, url, **kwargs):
    """page.goto com nova tentativa quando o site não responde a tempo.

    Caso real (25/09/2026 17:19, execução agendada): a primeira abertura
    da Plataforma estourou o tempo (Page.goto: Timeout 30000ms exceeded) e
    a execução inteira foi perdida - Etapa 1 sem Excel, Etapas 2 a 5
    puladas. Na execução seguinte a mesma página abriu normal: era
    lentidão passageira do site ou da rede, não falha do robô.

    Só tenta de novo em estouro de tempo ou erro de rede (net::ERR_...);
    qualquer outro erro sobe na hora, como antes. Na última tentativa o
    erro original sobe sem mudança, então quem chama continua tratando
    do mesmo jeito.

    TENTATIVAS_ABRIR_PAGINA (padrão 3) e PAUSA_NOVA_TENTATIVA_S (padrão
    30) ajustam sem editar o .py."""
    from playwright.sync_api import Error as PlaywrightError
    from playwright.sync_api import TimeoutError as PlaywrightTimeout

    tentativas = max(1, _inteiro_configurado("TENTATIVAS_ABRIR_PAGINA", 3))
    pausa_s = _inteiro_configurado("PAUSA_NOVA_TENTATIVA_S", 30)
    for tentativa in range(1, tentativas + 1):
        try:
            return page.goto(url, **kwargs)
        except PlaywrightError as e:
            transitorio = isinstance(e, PlaywrightTimeout) or "net::ERR_" in str(e)
            if not transitorio or tentativa == tentativas:
                raise
            motivo = str(e).splitlines()[0] if str(e) else type(e).__name__
            print(f"      [AVISO] A página não abriu (tentativa {tentativa} de {tentativas}): {motivo}")
            print(f"      Tentando de novo em {pausa_s}s...")
            time.sleep(pausa_s)


def bloquear_recursos_visuais(context):
    """Bloqueia o carregamento de imagens/fontes/vídeo nesse contexto -
    o robô nunca olha pra tela, só interage com o HTML, então baixar
    ícones/logos/fontes só consome tempo de rede à toa. Não bloqueia
    CSS nem JS (isso poderia mudar layout/comportamento e quebrar a
    automação) - só recursos puramente visuais.

    Só chame isso em HEADLESS=true - se tem alguém acompanhando a
    janela (HEADLESS=false), melhor deixar a página com a aparência
    normal."""
    def _filtrar(route):
        if route.request.resource_type in ("image", "font", "media"):
            route.abort()
        else:
            route.continue_()

    context.route("**/*", _filtrar)


def abrir_gestao(page, url_padrao, palavra_chave, timeout):
    """Abre a tela de gestão do cliente e devolve a URL que realmente
    funcionou (None se não conseguiu chegar em nenhuma).

    `url_padrao` é a URL esperada; se ela não responder (404, erro, ou
    cai de volta no login), procura no próprio menu do sistema um link
    que contenha `palavra_chave` no href. Existe porque a URL da gestão
    do Banco B foi deduzida do padrão do Banco A (/gestao-banco-a ->
    /gestao-banco-b) e nunca foi confirmada - e esta etapa GRAVA no
    sistema, então é melhor descobrir o caminho certo do que insistir
    num palpite e agir na tela errada.

    Nunca vai direto pra tela de gestão sem sessão: esse sistema não
    redireciona pro login quando não autenticado, ele quebra com erro de
    servidor (Auth::user() nulo) - quem chama já passou pelo /login."""
    from playwright.sync_api import TimeoutError as PlaywrightTimeout

    try:
        resposta = page.goto(url_padrao, timeout=timeout, wait_until="commit")
    except PlaywrightTimeout:
        print("      [AVISO] Navegação demorou demais pra 'começar' - tentando mesmo assim...")
        resposta = None

    status = resposta.status if resposta is not None else None
    if "/login" not in page.url and (status is None or status < 400):
        return url_padrao

    print(f"      [AVISO] {url_padrao} não respondeu como esperado"
          f"{f' (status {status})' if status else ''} - procurando o link no menu do sistema...")

    links = page.locator(f"a[href*='{palavra_chave}' i]")
    encontrados = []
    for i in range(links.count()):
        href = links.nth(i).get_attribute("href")
        if href and href not in encontrados:
            encontrados.append(href)

    if not encontrados:
        print(f"      [ERRO] Não achei nenhum link com '{palavra_chave}' no menu.")
        print("      Abra a tela de Em Aberto do Banco B no navegador e me diga a URL exata")
        print("      da barra de endereços pra eu corrigir o script.")
        return None

    print(f"      Link(s) encontrado(s) no menu: {', '.join(encontrados)}")
    for href in encontrados:
        try:
            resposta = page.goto(href, timeout=timeout, wait_until="commit")
        except PlaywrightTimeout:
            continue
        status = resposta.status if resposta is not None else None
        if "/login" not in page.url and (status is None or status < 400):
            print(f"      Usando: {page.url}")
            return page.url
    return None


class Tee:
    """Escreve simultaneamente no terminal e num arquivo de log.

    Dá flush a cada escrita (não só quando alguém chama .flush()
    explicitamente) - sem isso, o arquivo .txt no disco fica com um
    buffer de várias linhas atrasado em relação ao que já apareceu na
    tela, e um log copiado enquanto o robô ainda está rodando vem
    incompleto (foi exatamente o que aconteceu em vários logs
    enviados: sempre paravam bem antes do que realmente tinha
    acontecido na janela)."""

    def __init__(self, *streams):
        self.streams = streams

    def write(self, dado):
        for s in self.streams:
            s.write(dado)
        self.flush()

    def flush(self):
        for s in self.streams:
            s.flush()


def pedir_dado(pergunta):
    """Pede um dado ao usuário com um aviso visual bem claro."""
    print("\n" + "-" * 60)
    print(">>> PRECISO DE UMA INFORMAÇÃO SUA <<<")
    return input(f"{pergunta}: ").strip()


def pausar_para_usuario(*linhas_instrucao):
    """Pausa o robô e deixa bem claro que é a vez do usuário agir."""
    print("\n" + "#" * 60)
    print("#  A AÇÃO É SUA AGORA - O ROBÔ ESTÁ PAUSADO")
    print("#" * 60)
    for linha in linhas_instrucao:
        print(linha)
    input(">>> Quando terminar, clique aqui no terminal e pressione ENTER... ")
    print("#" * 60 + "\n")


def achar_excel_mais_recente(pasta_exports):
    """Devolve o caminho do .xlsx/.xls mais recente numa pasta (None se
    não tiver nenhum) - usado por qualquer script que precise do último
    Excel exportado, sem repetir a busca em cada arquivo.

    Ignora arquivos que começam com '~$' - são os arquivos de bloqueio
    temporário que o próprio Excel cria enquanto alguém tem a planilha
    de verdade aberta (pra visualizar/conferir, por exemplo); como
    ficam com data de modificação mais recente que o original, sem
    esse filtro a busca "mais recente" pegava o arquivo de bloqueio -
    que nem é um .xlsx de verdade - por engano.

    Ignora também os 'resultado_*.xlsx' que a Etapa 3 grava NESTA MESMA
    pasta ao terminar. Eles são relatório de saída, não lista de
    laudos - e como são sempre mais recentes que o export que os
    gerou, a próxima execução pegaria um relatório como se fosse a
    lista de propostas a processar."""
    arquivos = glob.glob(os.path.join(pasta_exports, "*.xlsx")) + glob.glob(
        os.path.join(pasta_exports, "*.xls")
    )
    arquivos = [
        a for a in arquivos
        if not os.path.basename(a).startswith("~$")
        and not os.path.basename(a).lower().startswith("resultado_")
    ]
    if not arquivos:
        return None
    return max(arquivos, key=os.path.getmtime)


def contar_linhas_excel(caminho_arquivo):
    """Conta quantas linhas de dados (todas as abas, sem contar
    cabeçalho) tem um Excel .xlsx - usado pra mostrar quantos laudos
    foram exportados/serão enviados. Devolve None se não conseguir
    contar (arquivo não é .xlsx, por exemplo)."""
    if os.path.splitext(caminho_arquivo)[1].lower() != ".xlsx":
        return None
    try:
        wb = load_workbook(caminho_arquivo, read_only=True, data_only=True)
        total = 0
        for nome_aba in wb.sheetnames:
            ws = wb[nome_aba]
            for linha in ws.iter_rows(min_row=2):
                if any(celula.value not in (None, "") for celula in linha):
                    total += 1
        return total
    except Exception as e:
        print(f"      [AVISO] Não consegui contar as linhas do Excel ({e}).")
        return None


def achar_coluna_por_cabecalho(ws, texto_procurado, exato=False):
    """Acha, pela linha 1 (cabeçalho), a 1ª coluna cujo texto contém
    `texto_procurado` (sem diferenciar maiúsc./minúsc.) - usa iter_rows
    em vez de ws.cell() pra funcionar tanto em workbook normal quanto
    read_only. Devolve None se não achar.

    `exato=True` exige o cabeçalho IGUAL ao procurado, não só contendo
    - necessário quando existe mais de uma coluna com o mesmo texto
    dentro do nome (ex.: 'Inspetor' e 'CPF Inspetor' são colunas
    diferentes, mas a busca por substring bateria nas duas - a que vier
    primeiro na planilha "ganharia" por engano)."""
    try:
        primeira_linha = next(ws.iter_rows(min_row=1, max_row=1))
    except StopIteration:
        return None
    alvo = texto_procurado.strip().lower()
    for indice, celula in enumerate(primeira_linha, start=1):
        cabecalho = str(celula.value or "").strip().lower()
        if (cabecalho == alvo) if exato else (alvo in cabecalho):
            return indice
    return None


def credenciais_configuradas(prefixo):
    """Lê usuário/senha das variáveis de ambiente `<PREFIXO>_USUARIO` e
    `<PREFIXO>_SENHA` (ex.: PLATAFORMA_USUARIO/PLATAFORMA_SENHA,
    CENTRAL_USUARIO/CENTRAL_SENHA) - nunca fica gravado no .py.

    Devolve (usuario, senha) ou (None, None) se qualquer uma faltar -
    nesse caso quem chama deve cair pra `pausar_para_usuario` (login
    manual), nunca travar o robô."""
    usuario = os.environ.get(f"{prefixo}_USUARIO", "").strip()
    senha = os.environ.get(f"{prefixo}_SENHA", "")
    if usuario and senha:
        return usuario, senha
    return None, None


def normalizar_chave_totp(texto):
    """Aceita a chave do App Autenticador do jeito que a tela mostrar:
    com espaços/hífens ("ABCD EFGH ..."), minúscula, ou a URL inteira do
    QR code ("otpauth://totp/...?secret=ABCD..."). Devolve só a chave em
    base32, maiúscula e sem separadores."""
    texto = (texto or "").strip()
    if texto.lower().startswith("otpauth://"):
        texto = parse_qs(urlparse(texto).query).get("secret", [""])[0]
    return texto.replace(" ", "").replace("-", "").upper()


def gerar_codigo_totp(chave, instante=None, digitos=6, periodo=30):
    """Código de 6 dígitos do App Autenticador (TOTP, RFC 6238 - o mesmo
    cálculo que Google/Microsoft Authenticator fazem no celular). Só
    biblioteca padrão, sem dependência nova. Validado contra o vetor de
    teste oficial da RFC (ver README)."""
    chave = normalizar_chave_totp(chave)
    chave_bytes = base64.b32decode(chave + "=" * (-len(chave) % 8), casefold=True)
    contador = int((time.time() if instante is None else instante) // periodo)
    mac = hmac.new(chave_bytes, struct.pack(">Q", contador), hashlib.sha1).digest()
    offset = mac[-1] & 0x0F
    valor = struct.unpack(">I", mac[offset:offset + 4])[0] & 0x7FFFFFFF
    return str(valor % (10 ** digitos)).zfill(digitos)


def chave_autenticador_configurada(prefixo):
    """Lê `<PREFIXO>_CHAVE_AUTENTICADOR` (ex.: PLATAFORMA_CHAVE_AUTENTICADOR).
    Devolve a chave normalizada, ou None se não estiver definida ou não
    for uma chave válida. Nunca imprime a chave."""
    chave = normalizar_chave_totp(os.environ.get(f"{prefixo}_CHAVE_AUTENTICADOR", ""))
    if not chave:
        return None
    try:
        gerar_codigo_totp(chave, instante=0)
    except Exception:
        print(f"      [AVISO] {prefixo}_CHAVE_AUTENTICADOR está definida mas não é uma chave")
        print("      válida (esperado: letras A-Z e números 2-7, a 'chave' mostrada na")
        print("      tela de configuração do App Autenticador).")
        return None
    return chave


def login_automatico(page, usuario, senha, timeout=30000):
    """Preenche e envia o formulário de login a partir de um campo de
    senha visível na página (`input[type='password']`) - não presume
    nomes/ids de campo, já que cada tela de login deste projeto pode
    diferir. Acha o campo de usuário como o primeiro <input> de texto
    (text/email, sem type) no MESMO <form> do campo de senha.

    Nunca imprime a senha em nenhuma circunstância. Devolve True só se
    conseguiu preencher e efetivamente submeter o formulário (clique no
    botão ou Enter) - False deixa quem chama decidir se pausa pra login
    manual."""
    try:
        campo_senha = page.locator("input[type='password']:visible").first
        campo_senha.wait_for(timeout=timeout)
    except Exception:
        return False

    form = campo_senha.locator("xpath=ancestor::form[1]")
    campo_usuario = form.locator(
        "input[type='text']:visible, input[type='email']:visible, input:not([type]):visible"
    ).first
    try:
        campo_usuario.fill(usuario, timeout=timeout)
        campo_senha.fill(senha, timeout=timeout)
    except Exception:
        return False

    botao = form.locator("button[type='submit'], input[type='submit']").first
    try:
        if botao.count() > 0:
            botao.first.click(timeout=timeout)
        else:
            campo_senha.press("Enter")
    except Exception:
        return False
    return True


# Caracteres que a busca da tela (List.js 2.x) "escapa" antes de procurar:
# ela transforma "." em "\." e "-" em "\-" e depois procura esse texto
# LITERALMENTE (indexOf), então um nº como "3.190.620" ou "10000021-1"
# nunca é achado. Confirmado lendo o código da List.js 2.3.1 e reproduzido
# localmente com a biblioteca real (24/09/2026).
_CARACTERES_QUEBRAM_BUSCA = re.compile(r"[-\[\]{}()*+?.,\\^$|#]")


def termo_busca_listjs(num_proposta):
    """Texto seguro pra digitar na busca da tela: troca os caracteres
    que a List.js estraga por espaço. Ela trata espaço como "todas estas
    partes", então "3 190 620" ainda filtra bem - e quem decide a linha
    certa continua sendo a comparação EXATA do nº depois do filtro."""
    termo = _CARACTERES_QUEBRAM_BUSCA.sub(" ", num_proposta)
    return re.sub(r"\s+", " ", termo).strip() or num_proposta


def limpar_busca_listjs(page, seletor="input.search[type='search']"):
    """Esvazia o campo de busca e devolve todas as linhas pra página. A
    List.js REMOVE do HTML as linhas que não batem com a busca - sem
    limpar, qualquer checagem feita depois só enxerga o que sobrou do
    filtro."""
    campo = page.locator(seletor)
    if campo.count() == 0:
        return
    if (campo.first.input_value() or "") == "":
        return
    campo.first.fill("")
    campo.first.dispatch_event("keyup")
    page.wait_for_timeout(800)


def diagnosticar_linha_nao_achada(page, num_proposta):
    """Quando localizar_linha devolve None, diz POR QUE: a proposta não
    está na lista, ou está mas escondida por um filtro da tela.

    A tela tem filtros (Financiamento/Consórcio e Agendamento) que
    escondem linhas com display:none, e localizar_linha só enxerga as
    visíveis. Depois de salvar, a página recarrega no filtro padrão -
    então uma proposta de outro tipo simplesmente sumia, e o log só
    dizia 'não achei', sem explicar nada.

    Limpa a busca antes de olhar: a List.js tira do HTML as linhas que
    não batem com o que foi digitado, e sem isso uma busca que falhou
    (ex.: nº com ponto/hífen, 24/09/2026) fazia uma proposta que ESTÁ
    em Em Aberto ser dada como "não está mais em Em Aberto"."""
    limpar_busca_listjs(page)
    return page.evaluate(
        """(numProposta) => {
            const tabela = document.querySelector('#table');
            if (!tabela) return 'a tabela nem estava na página';
            const limpa = (el) => (el ? el.textContent : '').replace(/\\s+/g, '').replace(/⭐️|⭐/g, '').trim();
            const achados = [];
            for (const tr of tabela.querySelectorAll(':scope > tbody.list > tr')) {
                const celula = tr.querySelector('td.num_proposta');
                if (celula && limpa(celula) === numProposta) achados.push(tr);
            }
            if (achados.length === 0) {
                return {codigo: 'fora_da_lista',
                        texto: 'não está mais em Em Aberto (já finalizada ou movida de fila)'};
            }
            if (achados.length > 1) {
                return {codigo: 'duplicada',
                        texto: 'aparece ' + achados.length + 'x na lista com o mesmo nº - '
                               + 'não dá pra saber qual editar, precisa de conferência manual'};
            }
            const tr = achados[0];
            if (tr.offsetParent === null) {
                return {codigo: 'escondida',
                        texto: 'está na lista, mas escondida pelo filtro da tela (tipo='
                               + (tr.getAttribute('data-tipo') || '?') + ', agendamento='
                               + (tr.getAttribute('data-agendamento') || '?') + ')'};
            }
            return {codigo: 'visivel_divergente',
                    texto: 'está visível na lista mas a comparação do nº divergiu'};
        }""",
        num_proposta,
    )


def pedir_confirmacao(*linhas_aviso, palavra="CONFIRMAR"):
    """Mostra um aviso e só devolve True se o usuário digitar a palavra
    exata - trava de segurança pra não enviar nada de verdade pro
    sistema de gestão sem um OK explícito (útil enquanto um fluxo novo
    ainda está em teste).

    Com MODO_EXECUCAO=auto (ver modo_automatico), pula a pergunta e
    confirma sozinho - é o que permite o mesmo script rodar sem
    ninguém na frente da tela, uma vez que o fluxo já esteja validado."""
    if modo_automatico():
        print("\n      [MODO_EXECUCAO=auto] Confirmando automaticamente (sem perguntar):")
        for linha in linhas_aviso:
            print(f"      {linha}")
        return True

    print("\n" + "!" * 60)
    print("!  CONFIRMAÇÃO NECESSÁRIA ANTES DE ENVIAR")
    print("!" * 60)
    for linha in linhas_aviso:
        print(linha)
    resposta = input(f"Digite {palavra} para enviar (qualquer outra coisa cancela): ").strip()
    return resposta == palavra
