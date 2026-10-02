# robo-cadastro-banco-b

> **Sobre este repositório.** Versão de portfólio de um projeto real, desenvolvido em
> ambiente corporativo em 2026. Os nomes da empresa, dos sistemas e dos bancos, as URLs
> e os dados de exemplo foram trocados por nomes genéricos ("Central de Gestão",
> "Plataforma", "Banco A", "Banco B"), então o código não roda contra os endereços de
> exemplo. O histórico de commits é novo e resumido; o caminho real do projeto está em
> [HISTORICO.md](HISTORICO.md).

Robô que entra na Plataforma (cliente **Banco B**), filtra o Relatório
Analítico (Tipo de Inspeção = todos menos AVM, Status = Laudo Aceito,
Laudo Aceito com Ressalvas, Laudo Indeferido, Laudo Devolvido ou
Cancelada, a partir de uma data escolhida
até hoje), clica em
**Exportar para Excel** e ainda confere Cancelada por Cancelada se
teve vistoria de verdade antes de considerar o arquivo pronto — salvo
em `data/exports/`. Na sequência, cadastra as propostas em "Em Aberto"
no sistema de gestão (Central de Gestão) e atribui o engenheiro financeiro
de cada uma.

Adaptado do robô irmão que já faz esse fluxo completo pro **Banco A**
(`robo-cadastro-banco-a`) — mesma lógica de login/sessão/navegação/
filtros/deduplicação, só trocando o cliente selecionado na Plataforma e
as URLs do sistema de gestão que forem específicas do Banco B. O que
já foi confirmado contra o sistema real no fluxo do Banco A foi mantido
igual aqui (mesmos filtros, mesma checagem de vistoria, mesma trava de
segurança na atribuição de engenheiro por CPF). O que é específico da
tela do Banco B e ainda não foi visto de verdade está marcado na
seção **"Pontos que podem precisar de ajuste fino"**, abaixo — este
fluxo, ao contrário do fluxo do Banco A, **ainda não rodou nenhuma vez
contra o sistema real do Banco B**.

> **Guia rápido de execução** (comandos, parâmetros e o que fazer quando
> algo sai do normal): [`COMO_EXECUTAR.md`](COMO_EXECUTAR.md). Este README
> explica o porquê de cada comportamento.

## Instalação (uma vez só)

```bash
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
playwright install chromium
```

## Atualizando o projeto

Para atualizar o projeto:

```powershell
git pull --autostash origin main
```

O `--autostash` guarda e devolve sozinho qualquer alteração local sua
(ex.: se você ainda editar algum arquivo na mão), evitando o erro "seus
arquivos locais seriam sobrescritos" — não precisa mais fazer
`git stash` / `git pull` / `git stash pop` manualmente.

## Uso

```bash
python exportar_status_banco_b.py
```

Vai pedir só a **data inicial** (formato `dd/mm/aaaa`) — a data final é
sempre hoje. No Relatório Analítico da Plataforma, o script aplica esses
filtros antes de exibir/exportar:

- **Tipo de Inspeção**: **Laudão**, **Urbano Alto Padrão** e **Urbano
  Padrão** — ou seja, todos menos **AVM**. (Confirmado na tela real do
  Banco B em 09/09/2026: a lista de tipos aqui é diferente da do Banco A,
  que tinha "Laudo Eletrônico"/"Laudo Físico".)
- **Status de Inspeção**: **Laudo Aceito**, **Laudo Aceito com
  Ressalvas**, **Laudo Indeferido**, **Laudo Devolvido** e
  **Cancelada**. (O status "com Ressalvas" não existe no Banco A e entra
  aqui porque o laudo foi aceito do mesmo jeito — a vistoria aconteceu.)
- **Indeferido e Devolvido (29/09/2026)**: entraram a pedido do usuário.
  Nesses status a vistoria foi feita e o engenheiro tem de ser pago, e
  antes eles ficavam fora do export e nunca eram cadastrados.
  - **Rótulos:** copiados do print real do filtro "STATUS DE INSPEÇÃO"
    do Banco B (29/09/2026). A etiqueta da lista mostra "Laudo
    Devolvido (Prest)", mas no filtro a opção se chama só **"Laudo
    Devolvido"** (não existe opção com "(Prest)"). O filtro marca pelo
    texto exato, então é esse o rótulo usado.
  - **Sem conferência de vistoria:** entram direto, como as aceitas
    (decisão do usuário). Só a Cancelada continua sendo conferida uma a
    uma.
  - **Teste:** `definir_filtro_checkboxes` rodou numa página com os 33
    status do print, na mesma ordem. Marcou exatamente os 5 pedidos (sem
    pegar "Laudo Negado (Prest)" nem outros parecidos) e não mexeu na
    seção seguinte.

Os dois filtros ficam em constantes no topo de
`exportar_status_banco_b.py` (`TIPOS_INSPECAO` e `STATUS_INSPECAO`),
então dá pra incluir/remover um status sem caçar no meio do código.

Depois de exportar, ele ainda abre **cada inspeção Cancelada**
individualmente (aba Laudos, dentro do detalhe) pra conferir se ela
teve vistoria de verdade (pelo menos uma versão de laudo publicada).
**Só as Canceladas são abertas** — um laudo aceito (com ou sem
ressalvas) nunca seria aceito sem vistoria, então abrir essas linhas
seria checar o óbvio e custa caro (cada uma abre um modal, espera
carregar e fecha). Se o período não tiver nenhuma Cancelada, essa etapa
inteira é pulada.

**Quem precisa ser conferido sai do próprio Excel, não da tela** (coluna
`Status` = Cancelada, coluna `Identificador` = o código `TAUxxxx`). O
export vem completo; a tela nem sempre (ver rolagem virtual, mais
abaixo). O terminal mostra quais status vieram no Excel a cada execução —
assim, se a Plataforma algum dia escrever "Cancelada" de outro jeito, dá
pra perceber na hora em vez de a conferência ser pulada em silêncio.
Cancelada **sem** nenhum laudo é removida do Excel final — caso normal
(proposta cancelada antes de qualquer vistoria acontecer), não conta
como erro. Laudo aceito (com ou sem ressalvas) nunca é removido — não
seria aceito sem laudo.

**Cancelada que a tela não mostrou (22/09/2026, porte do robô irmão do
Banco A)**: a lista da Plataforma só desenha parte das linhas (ver rolagem
virtual, mais abaixo) — se uma Cancelada do Excel nunca aparecer na
tela mesmo depois de rolar tudo, a regra continua sendo "não deu pra
conferir → remove por segurança", com uma exceção objetiva: **se a
Plataforma já tem a Data Vistoria preenchida no Excel para essa
Cancelada, ela fica** — a vistoria claramente aconteceu, só a tela que
não mostrou a linha. Sem essa exceção, uma Cancelada com vistoria de
verdade acabava descartada só por azar de rolagem — foi exatamente o
que aconteceu no Banco A (8 propostas cadastradas, com engenheiro
atribuído, presas em Em Aberto sem nunca poder completar, porque a
checagem as removeu do Excel de saída sem nunca terem sido conferidas
de verdade).

> Essa remoção assume que a 1ª coluna do Excel exportado é o mesmo ID
> mostrado na lista da Plataforma (ex.: `INS1004` no fluxo do Banco A — o
> prefixo do Banco B pode ser diferente). Na primeira execução real,
> confira se a contagem de linhas removidas bateu com o esperado (o
> script avisa se não bateu) — se a coluna for outra, me avisa pra eu
> ajustar.

O terminal mostra a **contagem de laudos** em dois momentos: logo após
o download (total baixado) e depois da checagem de Canceladas (total
final, já sem as sem vistoria) — e repete o total final no resumo de
"Concluído!" ao terminar.

O arquivo final fica salvo em
`data/exports/laudos_banco_b_<inicio>_a_<fim>_<timestamp>.<extensão original>`
(a extensão vem do próprio download — `.xlsx`, `.xls` ou o que o sistema
gerar).

A Plataforma inclui de brinde uma aba **"Informações"** só com metadados
do relatório (usuário, período, filtros usados etc.) — não é dado de
vistoria e atrapalha o cadastro depois. O script remove essa aba
automaticamente antes de terminar, mantendo as abas de dados intactas.

**Bug real corrigido (22/09/2026): corrida de tempo ao clicar em
Exportar.** Encontrado primeiro no robô irmão do Banco A (mesma função,
copiada de lá ao portar essa etapa) — o sistema às vezes não baixa
direto ao clicar em Exportar, abre antes um submenu com o formato
(Excel/PDF/CSV). O código checava `opcao_excel.count() == 0` **na
hora**, sem esperar o submenu terminar de renderizar: o clique tinha
funcionado certo, só a contagem rodava cedo demais, e isso virava um
erro falso — "não achei a opção Excel" mesmo ela existindo, só que
ainda não tinha entrado no DOM. Reproduzido localmente com uma página
de teste onde o texto "Excel" só aparece 120ms depois do clique:
`count() == 0` dava falso positivo, `wait_for(state="visible")` acha
certo. Se mesmo assim a opção não aparecer (falha de verdade), agora
salva print + HTML da tela em `data/debug_laudos/` (fora do git) antes
de desistir — evidência real pra próxima vez, em vez de chutar.

Um log de cada execução fica salvo em `logs/`.

### Janela de dias e piso de data (`JANELA_DIAS` / `DATA_MINIMA`)

O filtro da Plataforma é por **data de solicitação** da inspeção, mas a
proposta só aparece no export depois que o laudo é **aceito** — o que
demora dias. Ou seja: uma proposta solicitada no dia D só é capturada se
ainda estiver dentro da janela no dia em que o laudo fica pronto. O que
sobrar fora da janela **nunca mais é exportado** — some de vez, não é
"pega na próxima execução".

O robô irmão do Banco A mediu essa distribuição num export real (358
propostas, diferença entre "Data Criação" e "Data Status") e viu que
uma janela de **4 dias perdia quase metade** (51,1% capturado até 3
dias, só 100% até 12 dias) — por isso passou a usar **20 dias** por
padrão, e o Banco B herdou esse número.

**No Banco B o atraso é maior, e o padrão passou a 30 dias
(30/09/2026).** Medido no export real de 01/09 a 29/09 (497 laudos sem
Cancelada, "Data Criação" x "Data Status"):

| Janela | Capturado |
|---|---|
| 7 dias | 53,5% |
| 12 dias | 86,7% |
| 20 dias | 96,6% (17 laudos perdidos) |
| 25 dias | 99,4% |
| 30 dias | 100% (maior atraso visto: 27 dias) |

Confirmado na prática: em 29/09/2026 uma execução com `JANELA_DIAS=30`
achou **11 propostas que nunca tinham sido cadastradas**, com laudo
aceito depois que tinham saído da janela de 20 dias. Ressalva: o export
medido só alcança 29 dias pra trás, então um atraso maior que isso não
apareceria nele. Vale medir de novo daqui a algumas semanas. Ajustável
sem editar o código: `$env:JANELA_DIAS="45"`.

Existe também um **piso** (`DATA_MINIMA`, formato `dd/mm/aaaa`): nunca
puxar nada anterior a essa data, mesmo que a janela alcance mais pra
trás. É regra de negócio, não técnica — **confirmado com o usuário
(17/09/2026)**: o mesmo piso do Banco A, `01/09/2026`. Ajustável sem
editar o código: `$env:DATA_MINIMA="01/09/2026"` (ou `""` pra desligar
o piso).

## Login

Login automático: se as variáveis de ambiente `PLATAFORMA_USUARIO` e
`PLATAFORMA_SENHA` estiverem configuradas, o robô tenta logar sozinho
quando a sessão salva expira (sem precisar de `HEADLESS=false` nem de
alguém acompanhando a tela). Configurar uma vez:
```powershell
[Environment]::SetEnvironmentVariable("PLATAFORMA_USUARIO", "voce@exemplo.com", "User")
[Environment]::SetEnvironmentVariable("PLATAFORMA_SENHA", "suasenha", "User")
```
Sem essas variáveis configuradas, ou se o login automático falhar, o
fluxo cai para o login manual descrito abaixo.

**Duas falhas reais em produção (22/09/2026), a mesma causa raiz**:
numa execução `--auto` headless, o log mostrou que o robô **conseguiu
logar e selecionar o cliente Banco B** (isso só acontece depois de
autenticar com sucesso), mas a confirmação seguinte — esperar o texto
"GRID DE INSPEÇÃO" aparecer — não bateu dentro do timeout (primeiro 20s,
depois 45s), e o robô tratou uma sessão **genuinamente válida** como
morta. Isso aconteceu duas vezes, uma dentro do login automático e
outra na checagem inicial de sessão no `main()` — ambas esperavam o
painel aparecer com um único `wait_for`, sem alternativa se ele
demorasse mais que o timeout.

**Corrigido na raiz** (mesmo aprendizado do robô irmão do Banco A,
17-18/09/2026): `aguardar_painel_ou_erro()` substitui a espera única
por um polling que checa, a cada 500ms, se o painel apareceu OU se a
tela de erro 401 apareceu — o que vier primeiro — em vez de assumir
"não apareceu = sessão morta" na primeira checagem. Usado em três
pontos: na checagem inicial de sessão (`main()`), dentro do login
automático (com mais uma chance de 30s se o formulário de login nem
foi achado — pode ser só o painel lento, não sessão morta de verdade),
e depois de preencher o login. Reproduzido e validado localmente com
página sintética (painel aparecendo atrasado, nada aparecendo, e URL
de erro 401) antes de subir. Continua salvando um print da tela em
`logs/falha_login_plataforma_<motivo>_<timestamp>.png` quando a falha é
de verdade — evidência real em vez de precisar adivinhar.

**A causa real, revelada pela evidência (22/09/2026 12:19)**: com o
print e a URL no log, a terceira execução mostrou onde o robô parava
de fato: `https://plataforma-laudos.example.com/sistema/index.html#/verificacao-mfa`.
Ou seja, a senha **foi aceita** (o cliente Banco B chegou a ser
selecionado), e a Plataforma pediu **verificação em duas etapas
(MFA)**. Isso não é sessão lenta, nem credencial errada, nem
formulário sumido — e senha nenhuma passa por ali, porque impedir
exatamente isso é o objetivo do MFA. Nenhum dos dois robôs (Banco A e
Banco B) tratava esse caso: a tela era confundida com "não achei o
formulário" e a mensagem final ainda sugeria configurar credenciais,
que não resolveriam nada.

Como o robô lida com isso agora:
- `pagina_pede_mfa()` reconhece a URL `#/verificacao-mfa` e
  `aguardar_painel_ou_erro()` devolve um estado próprio `"mfa"`
  (além de `painel` / `erro` / `nada`). Validado localmente com página
  sintética nos quatro estados.
- Sem tela (`HEADLESS=true` ou `--auto`): a mensagem diz claramente que
  é MFA, que a credencial está certa e que o que falta é o código, e
  salva o print em `logs/falha_login_plataforma_mfa_<timestamp>.png`.
  Não sugere mais configurar `PLATAFORMA_USUARIO`/`PLATAFORMA_SENHA`
  nesse caso.
- Com tela: explica que só falta digitar o código no Chrome.

**Solução definitiva: App Autenticador (23/09/2026)**. O print da tela
real mostrou as opções de verificação da Plataforma: **E-mail** (a única
ativa na conta), **SMS** (nenhum telefone cadastrado) e **App
Autenticador** ("não configurado — acesse a edição de perfil"). **Não
existe "lembrar este dispositivo"**: toda sessão nova passa pelo MFA,
então guardar sessão/perfil do navegador não resolve — só adia até a
próxima expiração.

O App Autenticador usa TOTP (RFC 6238): o código de 6 dígitos é um
cálculo feito a partir de uma chave secreta e do relógio — o mesmo que o
Google/Microsoft Authenticator fazem no celular. Com essa chave na
variável `PLATAFORMA_CHAVE_AUTENTICADOR`, o robô faz o mesmo cálculo e
responde a verificação sozinho: escolhe "App Autenticador", clica
"Avançar", digita o código (um campo só ou uma caixinha por dígito) e
confirma. Se o código for recusado (relógio do Windows um pouco fora),
espera a próxima janela de 30s e tenta mais uma vez. O cálculo usa só a
biblioteca padrão do Python (`comum.gerar_codigo_totp`, sem dependência
nova) e foi validado contra os vetores de teste oficiais da RFC 6238; o
fluxo da tela foi testado com página sintética reproduzindo o passo 1
real (os dois formatos de passo 2, código recusado e depois aceito, e
sem chave configurada). O passo 2 (onde se digita o código) **ainda não
foi visto de verdade** — por isso o robô escreve no log o que essa tela
mostra toda vez que passa por ela, e salva print se algo não bater.

Configurar uma vez (a mesma conta serve pro robô do Banco A). O que o robô
precisa é a **chave fixa** que o QR code carrega — não o código de 6
dígitos, que muda a cada 30s e é calculado a partir dela:
1. Na Plataforma (navegador normal), **edição de perfil** → "Aplicativo
   de Autenticação". Se já estiver ligado, desligue e ligue de novo pra
   a tela mostrar o QR outra vez (e escaneie de novo no celular - a
   entrada antiga do app deixa de valer).
2. Escaneie o QR com o app autenticador do celular — assim uma pessoa
   continua conseguindo logar na mão.
3. Tire um print **só do QR** (Win+Shift+S; no Windows 11 o print fica
   salvo em Imagens\Capturas de Tela) e rode:
   ```powershell
   python -m pip install opencv-python-headless
   python configurar_autenticador.py "C:\caminho\do\print.png"
   ```
   Se a Plataforma mostrar a chave em texto, dá pra colar em vez do print:
   `python configurar_autenticador.py` (a chave colada não aparece na
   tela). O script mostra o código que o robô calcularia agora, pede pra
   você comparar com o celular e **só salva se bater** — na variável
   `PLATAFORMA_CHAVE_AUTENTICADOR` do seu usuário do Windows. A chave
   nunca é impressa nem vai pro código/git. Apague o print depois.
4. Feche e reabra o VS Code inteiro e confira com
   `python testar_autenticador.py` (mostra o código atual, nunca a chave).

Leitura do QR testada localmente com QR de autenticador gerado no mesmo
formato (`otpauth://totp/...`), inclusive dentro de um print de página
inteira com QR pequeno (180px): o leitor padrão do OpenCV **falhou** num
QR desse tamanho de texto, por isso o script tenta o leitor Aruco e
algumas escalas antes de desistir. A gravação no registro do Windows
(`HKEY_CURRENT_USER\Environment`) não pôde ser testada fora do Windows
— o passo 4 confirma.

Segurança: com a chave no computador, senha e segundo fator ficam na
mesma máquina — quem tiver acesso a esse usuário do Windows tem os dois.
É a troca aceita pra rodar sem ninguém na frente; guarde a variável só
no escopo do usuário (`"User"`) e nunca a compartilhe. O MFA foi
"habilitado por padrão pelo administrador da conta" (texto da própria
tela), então vale avisar quem administra a conta da Plataforma.

Sem a chave configurada, o comportamento é o de antes: sem tela, o log
diz `[MFA]` e aponta esta seção; com tela, uma pessoa digita o código.

**Ajustes de 23/09/2026 (caso real às 10:01)**: na pausa manual, o
usuário apertou ENTER no terminal com o Chrome ainda na tela de MFA
(tinha escolhido receber o código, mas ainda não o havia digitado) e o
robô encerrou na primeira checagem. Agora, enquanto a tela for de MFA
(ou o painel simplesmente não tiver aparecido), o robô pausa de novo e
explica o que falta, até 5 vezes, em vez de jogar a execução fora. E,
sempre que detecta MFA, escreve no log **o que a tela mostra** (texto
visível, campos, botões, checkboxes e se estão marcados) — pra decidir
a solução definitiva em cima do que a Plataforma realmente pede (código
por e-mail? app autenticador? SMS? existe "lembrar este dispositivo"?)
em vez de adivinhar. Só lê a tela; não digita nada.

A Plataforma também pode devolver uma tela própria de erro (401 "ACESSO
NÃO AUTORIZADO", em vez de redirecionar pro login) quando a sessão
salva perde validade — comportamento confirmado no robô irmão do Banco A
(mesma plataforma). O robô reconhece essa tela e descarta a sessão
salva. Ao recomeçar, abre um **contexto novo do navegador** (sem
`storage_state`) em vez de só limpar cookies — a Plataforma é uma SPA e
guarda o token de sessão no `localStorage`, não só em cookie; um
`clear_cookies()` sozinho deixava o token velho lá e o reload caía no
mesmo 401 de novo.

Na primeira vez, a sessão salva (`sessao_plataforma_banco_b.json`)
ainda não existe, então:

1. Rode o script pedindo pra ver a janela do navegador, com a variável de
   ambiente `HEADLESS` (não edite o arquivo — ver caixa abaixo):
   ```powershell
   $env:HEADLESS="false"; python exportar_status_banco_b.py
   ```
2. Uma janela do Chrome vai abrir e o robô vai pausar pedindo pra você
   fazer login manualmente (e-mail e senha, e completar qualquer
   verificação de segurança pedida).
3. Confirme que o cliente **Banco B** foi selecionado (o robô tenta
   selecionar sozinho o segundo card da tela de escolha de cliente).
4. Aperte ENTER no terminal quando a tela "GRID DE INSPEÇÃO" aparecer.
5. Pronto — a sessão fica salva. Nas próximas vezes, rode normal (sem a
   variável) e ele não pede login de novo (até expirar).

O arquivo de sessão nunca deve ser commitado (já está no `.gitignore`).

> **Por que variável de ambiente e não editar o arquivo:** todos os
> scripts deste projeto respeitam a variável `HEADLESS` (`comum.py`).
> Assim, o arquivo nunca muda por causa disso e o `git pull` nunca trava
> por conflito local.

## Pontos que podem precisar de ajuste fino

Este fluxo foi montado adaptando o robô já validado do Banco A, mas
**ainda não rodou nenhuma vez contra o sistema real do Banco B**.
Algumas partes vieram confirmadas de outro robô (não são chute) e
outras são suposições que precisam de um primeiro teste real:

- ✅ **Seleção do cliente Banco B, login, caminho de menu até o
  Relatório Analítico e preenchimento do período** — **confirmados na
  tela real em 09/09/2026**: o robô logou, selecionou o Banco B
  (`INDICE_CLIENTE_BANCO_B = 1`), chegou no Relatório Analítico e
  preencheu as datas sozinho.
- ✅ **Filtros de Tipo e Status de Inspeção** — as listas do Banco B
  são **diferentes** das do Banco A e já foram corrigidas a partir da tela
  real (ver seção "Uso" acima). Ficam nas constantes `TIPOS_INSPECAO` e
  `STATUS_INSPECAO`, no topo de `exportar_status_banco_b.py`.
- ✅ **Exportar para Excel** — **confirmado em teste real (09/09/2026)**:
  o download aconteceu, a aba "Informações" foi removida e o arquivo
  saiu em `data/exports/`.
- ✅ **Conferência das Canceladas** — **confirmada em teste real
  (09/09/2026)**: a inspeção TAU5142 (Cancelada, sem nenhum laudo) foi
  detectada e removida do Excel (7 → 6 laudos).
- ⚠️ **Rolagem virtual da lista**: a lista da Plataforma só desenha as
  linhas que cabem na altura visível — no 1º teste o Excel veio com 7
  laudos mas a tela só entregou 4 linhas. Duas proteções foram
  adicionadas: o robô **rola a lista** até parar de aparecer linha
  nova, e a lista de quem precisa ser conferido vem **do Excel**, não
  da tela (ver abaixo). Se mesmo assim uma Cancelada do Excel não
  aparecer na tela, ela é **removida por segurança** e o terminal diz
  qual foi.
- **URL do sistema de gestão** (`GESTAO_BANCO_B_URL` em
  `cadastrar_central_gestao.py` e `atribuir_engenheiro_central_gestao.py`):
  ⚠️ **suposição ainda não confirmada** — assume
  `https://central-gestao.example.com/gestao-banco-b`, seguindo o mesmo padrão da
  URL do Banco A (`/gestao-banco-a`). **Não precisa acertar de primeira**: se
  essa URL não responder (404, erro, ou cair no login), o robô procura
  sozinho no menu do sistema um link com "banco_b" no endereço e usa
  o que funcionar, mostrando no terminal qual foi. Se não achar nenhum,
  ele **para sem enviar nada** e pede a URL certa — nunca age numa tela
  que não confirmou ser a certa.
- ✅ **Seletores das telas Cadastrar e Consultar** — **conferidos contra
  o HTML real do Banco B (09/09/2026)** via `mapear_central_gestao.py`.
  Vários são **diferentes** dos do Banco A, apesar de ser o mesmo sistema:

  | | Banco A | Banco B |
  |---|---|---|
  | campo de arquivo | `#excelFilePlataforma` | `#excelFile` |
  | campo oculto (JSON) | `#dadosExcelPlataforma` | `#dadosExcel` |
  | formulário | `#excelFormPlataforma` | action `/storeExcel` |
  | caixa de busca | `input.search[name='search']` | `input.search[type='search']` |
  | aba Cadastrar | `#tab2-tab` | `#tab2-tab` (igual) |

  Usar os nomes do Banco A fazia o robô **esperar 5 minutos por um elemento
  inexistente e desistir em silêncio**, sem enviar nada — foi
  exatamente o que travou a Etapa 2 no primeiro teste.

  > O formulário é escolhido pela **action** (`/storeExcel`), não pelo
  > id: a página tem **dois** formulários com `id="excelForm"`
  > (Plataforma e Cetip). Escolher pelo id poderia mandar o Excel da
  > Plataforma para o formulário do Cetip.

- ⚠️ **A página importa só a PRIMEIRA aba do Excel.** O JS da tela faz
  `workbook.SheetNames[0]` — as demais abas somem sem aviso. Por isso a
  Etapa 1 remove a aba "Informações" (se ela ficasse em primeiro, subiria
  metadado em vez de dado) e a Etapa 2 **avisa antes de enviar** se o
  arquivo tiver mais de uma aba.
- **Nomes das colunas do Excel** ("Inspetor", "CPF Inspetor", "Nº
  Proposta"): o relatório vem do mesmo sistema (Plataforma), então os
  nomes devem se repetir — mas confirme na primeira execução real que
  as colunas saem com esses mesmos títulos.

Rode primeiro com uma data inicial recente (poucos dias, poucos
resultados) pra validar cada etapa antes de rodar com uma data inicial
bem pra trás ou ligar o modo automático.

## Etapa 2: cadastrar no sistema de gestão (Central de Gestão)

`cadastrar_central_gestao.py` loga em `https://central-gestao.example.com/login`
(sessão salva em `sessao_central_gestao.json`, mesmo padrão de login manual
uma vez com `$env:HEADLESS="false"` na primeira execução), vai até
`https://central-gestao.example.com/gestao-banco-b` (ver aviso acima), abre a aba
**Cadastrar** e envia o Excel mais recente de `data/exports/` na seção
**PLATAFORMA**.

```bash
python cadastrar_central_gestao.py
```

(opcionalmente aceita o caminho de um Excel específico como argumento, se
não quiser usar o mais recente)

**Confirmação antes de enviar**: como o fluxo inteiro ainda está em
teste pro Banco B, o script mostra o arquivo e a quantidade de linhas
de dados e só segue se você digitar `CONFIRMAR` — qualquer outra
resposta cancela sem enviar nada. Isso acontece **antes** de abrir o
navegador.

**Atenção com `/gestao-banco-b` sem login**: esse sistema não
redireciona pra tela de login quando você acessa uma página protegida
sem sessão válida — ele quebra com um erro de servidor (`Attempt to
read property "group" on null`, porque `Auth::user()` fica nulo).
Por isso os scripts sempre passam pela tela de login (`/login`) primeiro,
e só então navegam pra `/gestao-banco-b` - nunca vão direto pra lá sem
confirmar sessão.

**A tela de "Em Aberto" é lenta** (confirmado no fluxo do Banco A) — tanto
pra abrir quanto pra trocar de aba dentro dela. Duas proteções, não uma
só:
1. Timeouts bem folgados (minutos, não segundos) esperando os elementos
   aparecerem, em vez de um valor curto que travaria à toa num dia mais
   lento.
2. **Além disso**, uma pausa fixa de **10s**
   (`ESPERA_CARREGAMENTO_GESTAO`) depois da página "aparecer" e antes de
   clicar em qualquer coisa — confirmado no fluxo do Banco A que a tela
   aparenta carregada bem antes do JS/jQuery por trás terminar de
   inicializar; interagir cedo demais trava, e nenhum timeout de espera
   por elemento resolve isso (o elemento já apareceu, só não está
   funcional ainda). Isso roda toda vez que a página é carregada -
   inclusive uma vez por proposta nas Etapas 3 e 4, já que cada Salvar
   recarrega a página inteira (é um formulário de verdade, não AJAX) -
   por isso esse número pesa bastante no tempo total num Excel com
   muitas linhas.

   **Calibragem**: 45s era um chute inicial ("30s a 1 minuto"), nunca
   medido. 25s rodou **sem nenhuma falha em produção** (18/09/2026) nas
   Etapas 4 e 5, inclusive na Etapa 4 com 6 guias em paralelo (o caso
   mais pesado). Padrão desceu pra **20s** em 21/09/2026 pra economizar
   relógio, e pra **10s** em 24/09/2026, depois de uma execução completa
   em produção com 10s (Etapas 2 a 5, 6 guias em paralelo) rodar sem
   nenhuma falha. Se aparecer trava em série, o próximo passo de volta é
   15000.

   Dá pra ajustar sem editar o código, pela variável de ambiente
   `ESPERA_CARREGAMENTO_GESTAO_MS` (em milissegundos):
   ```powershell
   $env:ESPERA_CARREGAMENTO_GESTAO_MS="45000"; python atribuir_engenheiro_central_gestao.py
   ```
   Se um dia a tela voltar a travar (erros de "não consegui localizar a
   proposta" ou timeout no Salvar aparecendo em série), sobe o número
   sem mexer no código.

Outra otimização, essa automática e sem trade-off: rodando com
`HEADLESS=true` (padrão), os três scripts que automatizam de verdade
(`exportar_status_banco_b.py`, `cadastrar_central_gestao.py`,
`atribuir_engenheiro_central_gestao.py`) bloqueiam o carregamento de
imagens/fontes/vídeo das páginas — o robô nunca olha pra tela, só
interage com o HTML, então baixar ícones/logos/fontes só consome tempo
à toa. Não bloqueia CSS nem JS (evita quebrar layout/comportamento).
Só vale com `HEADLESS=true` — rodando com `HEADLESS=false` (janela
visível, pra acompanhar) as páginas continuam com a aparência normal.

**Duplicidade — duas camadas**:

1. Confirmado no fluxo do Banco A — reenviar um Excel com propostas já
   cadastradas **não duplica** os registros em "Em Aberto". O próprio
   sistema identifica pelo número da proposta e ignora as que já
   existem. Essa é a garantia de verdade; o Excel inteiro pode ser
   reenviado sem medo, mesmo com sobreposição de período entre
   execuções (assumindo que a gestão do Banco B se comporta igual à
   do Banco A - mesmo sistema).
2. Além disso, o script mantém um log local
   (`propostas_ja_enviadas_em_aberto.txt`, fora do git — nunca
   versionado) com toda proposta que **ele mesmo** já enviou com
   sucesso. Antes de cada envio, gera uma **cópia filtrada** do Excel
   (em `data/envios/`, nunca sobrescrevendo o export original) sem as
   propostas que já estão nesse log, e mostra quantas são realmente
   novas. A coluna de Nº de Proposta é achada pelo cabeçalho (procura
   por "proposta" no texto da 1ª linha), não por um índice fixo.

   > **Por que uma cópia em pasta separada, e não sobrescrever o
   > export**: o robô irmão do Banco A fazia isso até 17/09/2026 e o
   > arquivo original (lido depois pelas Etapas 3, 4 e 5) ficava só com
   > as propostas da última rodada — um bug silencioso, sem nenhum erro
   > na tela. Corrigido aqui antes de acontecer: `data/envios/` fica
   > fora de `data/exports/` de propósito, pra nunca ser confundida com
   > "o Excel mais recente" pelas etapas seguintes.

   Pra pular essa deduplicação de propósito (ex.: reenviar uma proposta
   que o log já conhece, num teste controlado), use
   `$env:IGNORAR_LOG_LOCAL="1"` — ver `preparar_teste_reenvio.py`
   abaixo. Fora desse teste, não use.

**Login automático**: com `CENTRAL_USUARIO`/`CENTRAL_SENHA` configurados (mesmo
mecanismo do login da Plataforma, ver seção "Login" acima), o robô tenta
logar sozinho na Central de Gestão quando a sessão salva expira.

### Testando reenvio de proposta já finalizada (`preparar_teste_reenvio.py`)

A deduplicação do lado do sistema ("reenviar não duplica") só tinha sido
confirmada com propostas que ainda estavam em Em Aberto — reenviar uma
proposta **já finalizada** era um caso diferente e não verificado (duas
possibilidades bem diferentes: o sistema ignora, e reenviar em lote é
seguro, ou recria a proposta em Em Aberto, e reenviar um lote grande
ressuscitaria centenas de propostas já finalizadas). Este script monta
um Excel com uma proposta só pra testar isso com segurança:

```powershell
python preparar_teste_reenvio.py <numero_da_proposta>
$env:HEADLESS="false"; $env:IGNORAR_LOG_LOCAL="1"; python cadastrar_central_gestao.py data\envios\teste_<numero>.xlsx
```

**Confirmado (17-18/09/2026)**: o teste equivalente rodou no fluxo
irmão do Banco A (mesmo sistema Central de Gestão por trás dos dois clientes) e
o resultado foi positivo — o sistema **não recriou** a proposta já
finalizada em Em Aberto. Reenviar o Excel inteiro (mesmo com propostas
já finalizadas dentro) é seguro pros dois clientes. O script continua
aqui pra quem quiser reconferir especificamente contra o Banco B.

## Etapa 3: atribuir engenheiro (`atribuir_engenheiro_central_gestao.py`)

Depois que as propostas já estão em "Em Aberto" (Etapa 2), esse script
atribui o engenheiro financeiro de cada uma, usando duas colunas do
Excel da Plataforma: **"Inspetor"** (nome) pra **buscar** o engenheiro, e
**"CPF Inspetor"** (CPF) pra **confirmar** com certeza qual é o certo —
nome nunca decide sozinho, só serve pra trazer candidatos.

```bash
python atribuir_engenheiro_central_gestao.py
```

(opcionalmente aceita o caminho de um Excel específico como argumento, se
não quiser usar o mais recente)

Mecânica (confirmada via captura real da tela **do Banco A**, ainda não
confirmada na tela do Banco B):

1. Na tela **Consultar** de "Em Aberto", cada proposta tem um id interno
   usado nos ids do modal/campo daquela linha. O robô usa a caixa
   "Procurar..." (mesma busca que você usaria na tela) pra achar a
   linha certa pelo Nº de Proposta.
2. Abre o modal **"INCLUSÃO DO ENGENHEIRO FINANCEIRO"** dessa linha
   (o mesmo que abre pelo ícone de atribuir engenheiro).
3. Consulta o mesmo endpoint que a busca da tela usa
   (`/buscar-pagengenheiros?q=<nome>`) pra trazer candidatos pelo nome
   do Inspetor.
4. Preenche o ID do engenheiro e clica em Salvar — só depois de você
   confirmar (ver abaixo).

**Trava de segurança (isso mexe em pagamento de engenheiro)**:
- A busca por nome só **traz candidatos** — só atribui quando
  **exatamente 1** desses candidatos tiver CPF **idêntico** (comparando
  só os dígitos, sem pontuação) ao da coluna "CPF Inspetor" do Excel.
  Zero ou mais de um com CPF batendo = pula essa proposta e marca pra
  você revisar na mão — nunca chuta qual é o engenheiro certo.
- Sempre roda uma **prévia** primeiro (proposta → engenheiro encontrado,
  ou motivo de ter pulado) sem clicar em nada, e só segue pro envio de
  verdade se você digitar `CONFIRMAR`.
- Depois de cada Salvar, busca a proposta de novo na tela e confere se o
  nome do engenheiro aparece — nunca confia só no clique. Se essa
  reconferência der um erro técnico, o resultado fica como
  **VERIFICAR** (não **ERRO**) - o salvamento em si já tinha acontecido
  antes dessa checagem, então "não confirmei" não é o mesmo que
  "falhou". Confira esses casos na tela quando sobrar algum.

**`FORA` não é erro**: antes de atribuir, o robô lê a lista inteira de
"Em Aberto" na tela e marca como `FORA` quem já não está mais lá — é o
estado normal de proposta já finalizada numa rodada anterior (a janela
do export cobre vários dias, então isso é esperado, não exceção). Quem
não é localizado por outro motivo (duplicada, escondida por filtro,
divergência) usa a mesma classificação já usada na Etapa 4 (ver
`comum.diagnosticar_linha_nao_achada`) pra separar isso de um erro de
verdade — assim o resumo final não esconde os poucos erros reais no
meio de muitas propostas já resolvidas.

**Login automático**: com `CENTRAL_USUARIO`/`CENTRAL_SENHA` configurados
(mesmo mecanismo da seção "Login", acima), o robô tenta logar sozinho
quando a sessão salva expira.

Gera um Excel de resultado em
`data/exports/resultado_engenheiros_<timestamp>.xlsx` (proposta, CPF do
inspetor no Excel, engenheiro achado na Central de Gestão, resultado —
OK/FORA/VERIFICAR/ERRO/PULADO —, detalhe), e termina com um resumo no
terminal com a linha "Nada pendente de ação: X de Y" (soma de OK + FORA).

✅ **Testado em produção**: rodou contra o sistema real do Banco B e
atribuiu o engenheiro corretamente na proposta certa, sem duplicata. A
classificação `FORA`/login automático (portados do robô irmão do Banco A)
ainda não passaram por um teste real do Banco B com volume — validar
com um Excel pequeno antes de confiar neles em lote.

### Bug real (24/09/2026): proposta com ponto/hífen no nº dada como "não está mais em Em Aberto"

Nas duas execuções de 24/09, as propostas `3.190.620`, `3.191.625`,
`3.195.055` e `10000021-1` apareceram na prévia da Etapa 3 como "vão
receber engenheiro", mas na hora de salvar o robô disse "não está mais em
Em Aberto" — e a própria Etapa 5, logo depois, listou as quatro **em Em
Aberto, sem engenheiro**. A Etapa 4 errou do mesmo jeito. Ou seja: ficavam
presas pra sempre e ainda eram contadas como "nada pendente".

Causa, confirmada lendo o código da List.js 2.3.1 (a biblioteca da busca
da tela) e reproduzida localmente com ela: a busca "escapa" os caracteres
especiais do texto digitado (`.` vira `\.`, `-` vira `\-`) e depois procura
esse texto **literalmente** — então "3.190.620" nunca é achado. E a
List.js **tira do HTML** as linhas que não batem com a busca: o
diagnóstico, rodando depois, olhava uma tabela vazia e concluía "fora da
lista".

Correção (`comum.termo_busca_listjs` e `comum.limpar_busca_listjs`):
- o robô digita `3 190 620` em vez de `3.190.620` (a List.js trata espaço
  como "todas estas partes"); quem escolhe a linha continua sendo a
  comparação do nº **exato** logo depois — nada muda nessa trava;
- antes de concluir "não está mais em Em Aberto", o diagnóstico limpa a
  busca e olha a lista inteira.

Validado com a List.js real: antes, as 4 davam `fora_da_lista`
(exatamente como no log); depois, as 4 são achadas com o id certo, um nº
parecido de propósito (`3.190.6200`) não é confundido, um nº inexistente
continua `fora_da_lista`, e o diagnóstico logo após uma busca crua que
falhou agora enxerga a linha.

### Aviso falso (29/09 e 02/10/2026): engenheiro com espaço duplo no nome dava "salva mas não confirmada"

Duas propostas voltaram como `VERIFICAR` ("salvo, mas nome não apareceu
na tela depois"): 10000022 (29/09) e 10000023 (02/10). Nas duas, a
atribuição tinha dado certo: a Etapa 5 da mesma execução finalizou a
proposta com aquele engenheiro.

Causa: o nome está cadastrado na Central com **espaço duplo**
("Eduardo De Jesus Novais - Cícero Dantas  BA", "Augusto Albuquerque
Santos Neto -  CarIús CE"). A busca de engenheiros devolve o nome assim,
mas a tela (HTML) mostra um espaço só, e a conferência
(`verificar_atribuicao`) comparava o texto literal. A mesma função
decide "já está atribuído" antes de salvar, então o mesmo erro fazia o
robô salvar de novo um engenheiro que já estava lá.

**Corrigido**: a comparação ignora espaços repetidos dos dois lados.
Validado com a função real numa página com os dois nomes dos casos
reais:

| Caso | Código antigo | Código novo |
|---|---|---|
| 10000023 com espaço duplo | não confirmado | confirmado |
| 10000022 com espaço duplo | não confirmado | confirmado |
| Engenheiro errado na linha | recusado | recusado |
| Parte do nome | confirmado | confirmado |

## Etapa 4: agendar vistoria (`agendar_vistoria_central_gestao.py`)

Preenche a **Data da Vistoria**, o **Horário da Vistoria** e a **Data e
Hora do Agendamento** de cada proposta já em "Em Aberto", usando 3
colunas do Excel da Plataforma: "Data Agendamento", "Horário Agendamento"
e "Data Vistoria".

```bash
python agendar_vistoria_central_gestao.py
```

Adaptado do robô irmão do Banco A, que passou por várias rodadas de ajuste
em execução real: horário arredondado pro slot de 30 em 30 min mais
próximo (o `<select>` não aceita digitação livre), "Data e Hora do
Agendamento" descoberta como sendo só um carimbo do sistema (não um
dado controlável — fica fora da reconferência), e uma checagem prévia
de campos obrigatórios (`conferir_campos_obrigatorios`) que evita um
travamento de minutos por proposta quando o formulário tem campo vazio
que o navegador bloquearia em silêncio.

**Seletores conferidos contra o HTML real do Banco B**
(`mapear_central_gestao.py`, 17/09/2026): os campos do modal
(`#data_input`, `#hora_input`, `#data_ag_input`) são **idênticos** aos
do Banco A. Só o id do modal muda de padrão —
`#editGestaoBancoBModal<id>` no lugar de `#editGestaoBancoAModal<id>`
— e os campos obrigatórios confirmados são Identificador, CEP,
Logradouro e Cidade.

Mesma trava de segurança do resto do fluxo: prévia obrigatória, só
mexe em proposta cujo valor atual ainda não bate com o alvo, e
reconfere lendo a lista depois de salvar (não confia só no clique).
Roda em paralelo (`PARALELISMO_AGENDAMENTO`, padrão 6 guias) — no
Banco A, 12 guias saturaram a tela real, então o robô avisa se você
configurar mais de 6.

**`FORA`, `DESCARTADA` e login automático** (mesmo aprendizado do
robô irmão do Banco A, ver Etapa 3): quem já saiu de Em Aberto vira
`FORA` (não erro); quem só falta o engenheiro atribuído vira
`DESCARTADA` (processo interno, fora do controle do robô); login
automático via `CENTRAL_USUARIO`/`CENTRAL_SENHA`.

**Tipo de Imóvel — diferente do Banco A**: no Banco A esse campo é um
`<select>` com opções fixas, por isso existe lá um dicionário de
tradução do texto do Excel pra cada opção exata. **Confirmado na
captura real do Banco B (17/09/2026)**: aqui é um campo de **texto
livre** (`#edit_tipo`) — não há dicionário nenhum, o robô só preenche
direto com o texto da coluna "Tipo Imovel" do Excel quando o campo
está vazio no sistema (nunca sobrescreve valor já preenchido; se o
Excel também estiver vazio, pula a proposta em vez de chutar).

**Desempate de proposta com mais de uma inspeção (Identificador)**:
o robô do Banco A usa um campo oculto (`docid`) no HTML de cada linha
pra distinguir duas inspeções da mesma proposta. **Conferido na
captura real do Banco B e esse campo não existe** — não dá pra
saber se é porque o Banco B não tem esse caso ou porque a amostra
(6 linhas) não pegou um. Por segurança, o robô aqui **não tenta
desempatar**: se aparecer mais de uma linha com o mesmo nº de
proposta, cai no fallback já seguro (classificado como `duplicada`,
vira `PULADO` pra revisão manual — nunca ação na linha errada). O
dedup por Identificador (evitar descartar a 2ª inspeção da mesma
proposta no Excel) foi portado normalmente, já que isso não depende
do HTML.

⚠️ **Ainda não rodou contra o sistema real do Banco B** — os
seletores foram conferidos contra uma captura real da tela, mas o
fluxo completo de preencher/salvar precisa do primeiro teste real
antes de confiar nele com volume.

### Bug real (28/09/2026): modal Editar que não fechava travava a busca da proposta seguinte

**Execução de 28/09/2026 às 11:48** (Etapa 4, Guia 1): depois de
"Já está com esses dados - pulando" na 10000019, a proposta seguinte
(10000020) deu `[ERRO] Não consegui nem localizar a proposta na tela:
Locator.click: Timeout 300000ms exceeded`. O log do Playwright mostrou
548 tentativas de clicar na caixa de busca, todas com
`<div class="col-md-3">…</div> from <div class="table-responsive scrollbar">…</div> subtree intercepts pointer events`.
Não se perdeu nada: essa proposta já tinha saído de Em Aberto às 10:09.

Causa: na captura real da tela (`exploracao/01_janela_acoes`), os
modais ficam **dentro da tabela** (cada linha carrega os seus), e os
campos deles são `div.col-md-3`. O que cobria a busca era o **modal
Editar da proposta anterior, que continuou aberto**. Reproduzido com o
Bootstrap 5 real numa página com a mesma estrutura:
- um fechar que chega antes da animação de abertura terminar
  (`shown.bs.modal`) é **ignorado**, e o modal fica aberto de vez;
- fechando 50, 150 ou 250 ms depois de abrir, o modal ficou aberto nas 3
  vezes, e a busca travou com a mesma mensagem de "intercepts pointer
  events" vinda do `table-responsive`;
- na sequência normal do robô, o clique cai uns 100 ms depois da
  animação, porque o Playwright espera o botão parar de se mexer. Por
  isso, das 53 vezes em que o robô já tinha fechado o modal sem salvar,
  só esta deu errado.

**Corrigido** em `_fechar_modal_editar`, usado nos 4 casos em que a
Etapa 4 fecha sem salvar (já estava certo, DESCARTADA, PULADO por campo
obrigatório e PULADO por Tipo de Imóvel):
- depois de clicar em fechar, **confere que o modal sumiu**;
- se não sumiu em 5s, clica de novo (a animação já acabou), até 3
  vezes;
- se ainda assim não fechar, recarrega a tela (`[AVISO] O modal Editar
  não fechou...`), pra nunca deixar a próxima proposta presa atrás dele.

Validado na mesma página, com o código antigo e o novo, fazendo o
Bootstrap ignorar o fechar como faz durante a animação:

| Caso | Código antigo | Código novo |
|---|---|---|
| Sequência normal | ok | ok |
| 1º fechar ignorado (a corrida) | busca trava (mesma mensagem) | ok, fecha na 2ª tentativa |
| Fechar sempre ignorado | busca trava | ok, recarrega a tela |

## Etapa 5: finalizar vistorias completas (`finalizar_vistorias_central_gestao.py`)

Marca o checkbox de cada proposta que já está **completa** (engenheiro
+ Data de Agendamento + Data de Vistoria preenchidas) e clica no botão
verde **Finalizar Selecionados**, tirando-as de "Em Aberto".

```bash
python finalizar_vistorias_central_gestao.py
```

Diferente das outras etapas, **não é guiada pelo Excel** — o critério é
lido direto da tela. Assim ela pega qualquer proposta pronta, mesmo de
lotes anteriores, e nunca finaliza uma incompleta.

**Quem fica de fora vem explicado (22/09/2026, porte do robô irmão do
Banco A)**: em vez de só contar "Nx sem engenheiro, Ny sem datas", o robô
cruza cada proposta incompleta com o Excel mais recente e diz por quê:
`sem engenheiro` é processo interno (nada a fazer); `no Excel COM data
de vistoria` é uma bandeira de atenção real — a Etapa 4 deveria ter
agendado essa e não agendou, vale investigar; `no Excel SEM data` é
normal, a Plataforma ainda não tem a data, completa sozinha numa rodada
futura; e quem não está no export nem entra na conta (fora do período
da última exportação). A explicação também vai pro Excel de resultado,
não só pro terminal.

**Seletores conferidos contra o HTML real do Banco B**: a tabela é
`#table` (Banco A usa `#tabelaBancoA`), o resto é **idêntico** ao Banco A —
`input.checkbox-proposta`, `#formFinalizarSelecionados`, botão com
`title="Finalizar Selecionados"`. Uma diferença real encontrada na
captura: a *action* do formulário é
`https://central-gestao.example.com/gestao-banco_b/atualizar-status` (repare o
**hífen** em "gestao-banco_b" — diferente do resto das URLs do
Banco B, que não têm hífen). Isso não precisou ir pro código: o robô
só clica no botão da própria tela, nunca monta essa URL na mão.

Sem paralelismo, de propósito — marcar checkbox não recarrega a
página, então não há ganho em várias guias aqui.

**Login automático**: com `CENTRAL_USUARIO`/`CENTRAL_SENHA` configurados
(mesmo mecanismo das outras etapas), tenta logar sozinho quando a
sessão salva expira.

### Bug real (18/09/2026): "Finalizar Selecionados" nunca enviava nada

**Confirmado em execução real**: 184 propostas completas na prévia,
`CONFIRMAR` digitado, checkboxes marcados certinho (`184 checkbox(es)
marcados na página`) — e depois do clique em Finalizar, **5 minutos de
timeout** esperando uma navegação que nunca chegava. Resultado: **0
finalizadas**, as 184 continuaram intactas em Em Aberto (falha segura —
nada foi perdido nem duplicado, só não avançou).

Causa raiz, confirmada no JS real capturado da tela: o botão chama
`confirmarFinalizacao()`, que — quando tem proposta marcada — dispara
um **diálogo nativo do navegador**, `confirm('Tem certeza que deseja
finalizar as propostas selecionadas?')`, e só monta o formulário e
envia se for aceito. O Playwright **descarta automaticamente** qualquer
diálogo desses quando não há um handler `dialog` registrado — equivale
a clicar em "Cancelar" sozinho. O clique no botão não dava erro nenhum
(por isso passou despercebido até agora), mas o `confirm()` sempre
voltava `false` e o formulário nunca era enviado de verdade.

**Corrigido**: a página desta etapa agora registra
`page.on("dialog", ...)` que aceita automaticamente qualquer diálogo
nativo (`_aceitar_dialogo`) — mesmo padrão que teria pego o `alert()`
de "nenhuma proposta selecionada" também, se algum dia acontecesse.
Único ponto do fluxo com esse tipo de diálogo: as outras etapas
(Salvar da Etapa 3/4, Enviar da Etapa 2) não usam `confirm()`/`alert()`
nativo, confirmado tanto na captura da tela quanto pelas centenas de
salvamentos reais que já funcionaram sem esse problema.

✅ **Confirmado em produção logo em seguida (18/09/2026)**: o log
mostrou `[diálogo do navegador] '...' - aceitando automaticamente.`
depois do clique, a página navegou normal, e o resultado final foi
**184 finalizada(s) e confirmada(s), 0 com erro** — as mesmas 184 que
tinham travado no teste anterior.

## Fluxo completo num comando só (`executar_fluxo_banco_b.py`)

Roda as cinco etapas em sequência (exporta → cadastra → atribui
engenheiro → agenda vistoria → finaliza completas) com um único
comando, sem lógica nova — só chama, em ordem, o `main()` de cada
script acima. Cada script continua funcionando sozinho normalmente
também, se preferir rodar um de cada vez.

```bash
python executar_fluxo_banco_b.py
```

Se a Etapa 1 não trouxer nenhum resultado novo (ex.: já rodou há pouco
e não passou tempo suficiente pra ter inspeção nova), as Etapas 2 a 5
são puladas nessa execução — sem sentido cadastrar/atribuir em cima de
Excel velho.

### O login é resolvido uma vez só

Mesmo aprendizado do robô irmão do Banco A: as Etapas 2 a 5 compartilham o
**mesmo arquivo de sessão** (`sessao_central_gestao.json`), mas antes
cada uma descobria a sessão vencida por conta própria — em modo
visível, quatro pausas pedindo login manual na mesma execução; em
headless, quatro erros seguidos, terminando em "Concluído" sem ter
feito nada. Agora `preparar_sessao_central()` resolve o login **uma vez,
logo depois da Etapa 1**, e só então as Etapas 2 a 5 rodam (encontrando
a sessão já válida). Se não conseguir logar, o fluxo **para ali** com a
mensagem certa, em vez de arrastar quatro etapas até o mesmo beco.

Cada etapa também roda isolada (`rodar_etapa`): um erro inesperado numa
delas não derruba as seguintes — importa principalmente pra Etapa 5,
que não depende do Excel desta rodada.

Trava contra execução simultânea (`fluxo_em_execucao.lock`, criada e
apagada pelo próprio fluxo, fora do git): se o Agendador disparar de
novo enquanto a execução anterior ainda roda, a nova desiste em vez de
atropelar.

### Bug real (21/09/2026): com `--auto`, as Etapas 2, 3 e 4 não faziam nada

**Visto numa execução real do robô irmão do Banco A** (mesmo código de
orquestrador): a Etapa 1 exportou 579 laudos normalmente e, logo em
seguida, as três etapas seguintes morreram com
`[ERRO] Não achei nenhum Excel pra usar` — com o export recém-criado
parado ali na pasta. A Etapa 5 (única que não depende do Excel) rodou
normal, então o fluxo terminava imprimindo "Concluído" e parecendo ter
funcionado.

Causa: as Etapas 2, 3 e 4 rodam **no mesmo processo** do orquestrador e
leem `sys.argv[1]` pra saber se você passou um Excel específico na mão.
Com `--auto` ainda em `sys.argv`, as três liam a string `"--auto"` como
se fosse o **caminho do arquivo**, não achavam, e desistiam. Só
acontecia no modo `--auto` — é justamente o modo da tarefa agendada,
que roda sem ninguém olhando.

**Corrigido**: o orquestrador agora consome o `--auto`
(`sys.argv.remove("--auto")`) depois de ligar o modo automático, então
as etapas voltam a cair no comportamento certo (usar o Excel mais
recente de `data/exports/`). Correção feita na raiz — quem introduz a
flag é quem a remove — em vez de espalhar tratamento de flag pelas três
etapas.

> O mesmo bug existe no robô do Banco A (foi de lá que o mecanismo
> `--auto` veio). Vale avisar aquele fluxo pra corrigir também.

### Bug real (25/09/2026): um timeout da Plataforma perdia a execução inteira, e o Agendador registrava sucesso

**Execução agendada de 25/09/2026 às 17:15** (a primeira sem ninguém
olhando): a Etapa 1 parou logo no passo `[1/5] Verificando sessão
salva...` com `Page.goto: Timeout 30000ms exceeded` ao abrir
`plataforma-laudos.example.com/sistema/index.html#/home`. Sem Excel, as Etapas 2 a 5
foram puladas (certo), mas a execução inteira foi perdida por uma
lentidão passageira: a mesma abertura funcionou normalmente na execução
seguinte (28/09, 10:09, fluxo completo sem erro). E o processo terminou
com **código de saída 0**, então o Agendador de Tarefas registrou
"êxito" e um aviso de falha baseado nisso nunca dispararia.

**Correção 1: nova tentativa ao abrir a página** (`comum.abrir_pagina`).
A abertura da tela de entrada de cada etapa (Plataforma na Etapa 1; login
da Central no orquestrador e nas Etapas 2 a 5) tenta de novo quando dá
estouro de tempo ou erro de rede (`net::ERR_...`): até 3 tentativas, com
30s de pausa entre elas. Qualquer outro erro sobe na hora, como antes. Se
as 3 falharem, o erro original sobe sem mudança e cada etapa o trata do
mesmo jeito de sempre. Ajuste sem editar o código:
`TENTATIVAS_ABRIR_PAGINA` (padrão 3) e `PAUSA_NOVA_TENTATIVA_S` (padrão
30).

Reproduzido antes de corrigir (28/09/2026) com um servidor local que
demora mais que o timeout pra responder: o `page.goto` puro estoura com
o mesmo `Page.goto: Timeout ... exceeded` do log real. No mesmo teste,
`abrir_pagina` abre na 2ª tentativa; com o site lento o tempo todo, faz
3 pedidos e sobe o mesmo `TimeoutError`; com a porta fechada
(`net::ERR_CONNECTION_REFUSED`), tenta 3 vezes e sobe o erro; com uma URL
inválida, sobe na hora, sem nova tentativa.

**Correção 2: código de saída que diz o que aconteceu.** A última linha
do `fluxo_completo_*.txt` passa a ser `[FLUXO] Resultado: ...`, e o
processo sai com:

| Código | Significa | Precisa de ação? |
|---|---|---|
| `0` | Rodou tudo, nenhum `[ERRO]` no log | Não |
| `1` | Execução perdida: Etapa 1 sem Excel, sem login na Central, ou uma etapa quebrou | **Sim**, abrir o log |
| `2` | Rodou até o fim, mas alguma etapa/proposta deu `[ERRO]` | **Sim**, conferir as linhas `[ERRO]` |
| `3` | Outra execução ainda estava rodando (trava) - não fez nada | Só se repetir |

O `2` usa o marcador que as etapas já imprimem em toda falha de verdade
(`[ERRO]` / `[ERRO INESPERADO]`, inclusive nas guias paralelas).
`PULADO`, `FORA`, `DESCARTADA` e "salva mas não confirmada" não usam esse
marcador e continuam não contando como erro.

Validado de dois jeitos (28/09/2026):
- O `main()` real, com as etapas trocadas por versões falsas (sem tocar
  Plataforma/Central), saiu com o código esperado em 7 cenários: tudo ok (0),
  Etapa 1 com o timeout de 25/09 (1), Central sem login (1), login da Central
  quebrando (1), `[ERRO]` numa guia da Etapa 3 (2), Etapa 4 quebrando (1)
  e trava ocupada (3). Nos casos 0 a 2, a trava foi liberada.
- A mesma regra aplicada aos 17 `fluxo_completo_*.txt` reais de setembro:
  - `1` nas 6 execuções de 22/09, que pararam no MFA recém-ligado, e na
    de 25/09 às 17:19;
  - `2` na de 18/09, com 74 `[ERRO]` do bug do Finalizar;
  - `0` em todas as execuções limpas de 21, 23, 24, 25 e 28/09.

  A de 21/09 às 16:46 foi interrompida no meio (o log termina de repente).
  Processo interrompido já sai com código diferente de 0 pelo próprio
  Windows.

### Dois modos (variável de ambiente `MODO_EXECUCAO`)

- **Padrão** (sem definir a variável): pede a data inicial da Etapa 1 e
  exige digitar `CONFIRMAR` antes de qualquer envio real nas Etapas 2 a
  5 — exatamente como rodar cada script separado hoje. **Use esse modo
  até confiar no fluxo inteiro**, rodando várias vezes e conferindo os
  resultados.
- **`auto`** (ou `--auto` na linha de comando): calcula a data inicial
  sozinho (últimos 30 dias por padrão — ver **Janela de dias e piso de
  data** na seção Uso — sobrepõe com `JANELA_DIAS`) e confirma tudo
  sozinho, sem pausar esperando ninguém. É o modo pra rodar sem
  supervisão, de tempos em tempos — **só ligar depois de validar bem o
  fluxo no modo padrão**.
  ```powershell
  python executar_fluxo_banco_b.py --auto
  ```

  **"Sem pausar" vale inclusive com a janela visível** (21/09/2026). As
  Etapas 3, 4 e 5 já tratavam isso, mas a Etapa 1, a Etapa 2 e o login
  único do orquestrador só olhavam o `HEADLESS` — então rodar
  `--auto` com `$env:HEADLESS="false"` ainda parava pedindo ENTER no
  fim da Etapa 1 e da Etapa 2, e pediria login manual se a sessão
  tivesse vencido. Agora as seis guardas usam o mesmo critério
  (`HEADLESS or modo_automatico()`): em modo automático o robô **nunca**
  espera alguém, com ou sem janela. Se não conseguir logar sozinho, ele
  encerra a etapa com a mensagem no log em vez de ficar parado.

### Agendando de 6 em 6 horas no Windows (Agendador de Tarefas)

Só faça isso **depois de validar bem o fluxo no modo padrão**.

**Antes de agendar, os três pré-requisitos** (sem eles a tarefa roda e
não faz nada, ou para num login):

1. Credenciais fixadas no Windows (`PLATAFORMA_USUARIO`/`PLATAFORMA_SENHA`,
   `CENTRAL_USUARIO`/`CENTRAL_SENHA` — ver **Login**).
2. Paralelismo: o padrão já é 6 guias nas Etapas 3 e 4 (validado em
   produção) - não precisa configurar nada. Só defina
   `PARALELISMO_ATRIBUICAO` / `PARALELISMO_AGENDAMENTO` se quiser outro
   número (ver `COMO_EXECUTAR.md`).
3. **Um ensaio do modo automático na mão**, no mesmo terminal, com as
   sessões apagadas pra forçar o login automático (a única parte que só
   se prova com sessão vencida de verdade):
   ```powershell
   Remove-Item sessao_plataforma_banco_b.json, sessao_central_gestao.json -ErrorAction SilentlyContinue
   python executar_fluxo_banco_b.py --auto
   ```
   Tem que ir do começo ao fim sem parar, com `Login automático OK` nos
   dois sistemas. Se parar, o log diz onde — não agende antes de resolver.

**Aí sim, agendar.** PowerShell **como Administrador** (ajuste o caminho
da pasta e do `.venv` se for diferente):

```powershell
$pasta = "$env:USERPROFILE\caminho\para\robo-cadastro-banco-b"
$acao = New-ScheduledTaskAction -Execute "$pasta\.venv\Scripts\python.exe" `
    -Argument "executar_fluxo_banco_b.py --auto" `
    -WorkingDirectory $pasta
$gatilho = New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Hours 6) -RepetitionDuration ([TimeSpan]::MaxValue)
$configuracoes = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Hours 5)
Register-ScheduledTask -TaskName "FluxoBancoBCentralGestao" -Action $acao -Trigger $gatilho -Settings $configuracoes
```

O que cada parte garante:

- `--auto` liga o modo automático **só nessa execução** — as execuções
  manuais continuam pedindo `CONFIRMAR`. Nunca fixe `MODO_EXECUCAO`
  como variável do Windows.
- `-MultipleInstances IgnoreNew` + a trava `fluxo_em_execucao.lock`: se
  uma execução ainda estiver rodando quando a próxima disparar, a nova
  desiste em vez de atropelar.
- `-ExecutionTimeLimit 5h`: uma execução travada é morta antes da
  seguinte.
- Sem `HEADLESS` definido, roda sem janela (mais rápido). Em modo
  automático o robô **nunca** para esperando ENTER ou login manual —
  se não conseguir logar sozinho, encerra a etapa com a mensagem no log.

Pra conferir: `Get-ScheduledTaskInfo -TaskName "FluxoBancoBCentralGestao"`.
Pra pausar/remover: `Disable-ScheduledTask` / `Unregister-ScheduledTask -TaskName "FluxoBancoBCentralGestao"`.
Cada execução gera `logs/fluxo_completo_<timestamp>.txt` (resumo das 5
etapas) além dos logs de cada etapa. O que olhar num log de execução
agendada: a última linha, `[FLUXO] Resultado: ...`, e a linha
`Nada pendente de ação: X de Y` de cada etapa. O "Resultado da última
execução" do Agendador mostra o mesmo código de saída (0 = ok; 1, 2 ou 3
= ver a tabela em **Bug real (25/09/2026)** acima). É nele que o aviso
de falha deve se basear.

## Ferramenta de mapeamento (`mapear_central_gestao.py`)

Usada pra descobrir os seletores certos sem chutar, caso alguma das
suposições marcadas acima (URL de gestão, seletores da tela Cadastrar
ou de Em Aberto) não bater com o Banco B:

```bash
python mapear_central_gestao.py
```

Ele abre o Chrome visível. Você navega manualmente até cada tela que quer
mapear, volta no terminal e aperta **ENTER** — isso já captura a tela
atual (com um nome automático tipo `etapa_1`; se quiser um nome melhor,
digite ele antes do ENTER). Pra encerrar, digite `sair` e ENTER. Cada
captura salva em `exploracao/<numero>_<nome>/`:

- `screenshot.png` — foto da tela
- `frame_0.html`, `frame_1.html`, ... — HTML de cada frame/iframe da página
- `resumo.json` — lista de botões, links e campos já filtrados

No fim, **compacte a pasta inteira `exploracao` em um `.zip`** e manda
aqui no chat — uso isso pra escrever/ajustar os seletores certos, de
primeira, sem ida e volta.

A pasta `exploracao/` nunca deve ser commitada (já está no `.gitignore`) —
tem capturas de telas internas do sistema.

## Status do fluxo

Estado em 18/09/2026. As cinco etapas já rodaram contra o sistema real
do Banco B, na sequência completa (Etapa 1 → 2 → 3 → 4 → 5), depois
de portar pro Banco B o que o robô irmão do Banco A aprendeu em produção
(login automático, tratamento do 401 da Plataforma, janela de 20 dias,
piso de data, `data/envios`, `FORA`/`DESCARTADA`, login único no
orquestrador):

- ✅ **Etapa 1** (exportar da Plataforma) — **testada em produção**:
  janela de 20 dias e piso `01/09/2026` aplicados, filtros corretos, 224
  laudos baixados, 3 Canceladas sem vistoria detectadas e removidas
  (213 no Excel final). Login: senha automática funciona; o MFA da
  Plataforma (sem "lembrar dispositivo") passa a ser respondido sozinho
  via App Autenticador + `PLATAFORMA_CHAVE_AUTENTICADOR` (ver **Login**) -
  testado com página sintética, **ainda não contra a tela real**. Tela
  de erro 401 ainda não foi vista de verdade.
- ✅ **Etapa 2** (cadastrar em Em Aberto) — **testada em produção**: 208
  propostas novas cadastradas com sucesso, cópia filtrada indo pra
  `data/envios/` (não sobrescreveu o export original). Login automático
  não foi exercitado nessa execução (sessão já estava válida). O caso
  de reenviar uma proposta **já finalizada** foi testado no robô irmão
  do Banco A (mesmo sistema Central de Gestão) e confirmado seguro — ver
  **Testando reenvio de proposta já finalizada**, acima.
- ✅ **Etapa 3** (atribuir engenheiro) — **testada em produção**: busca
  por nome, decisão por CPF, salvar e reconferir confirmados contra o
  sistema real, 196 atribuídas com sucesso, `14 FORA` corretamente
  identificadas (proposta já finalizada), `Nada pendente de ação: 208
  de 213`. Paralelismo (6 guias) também testado.
- ✅ **Etapa 4** (agendar vistoria) — rodou contra o sistema real;
  confirmado indiretamente pela Etapa 5 (184 propostas chegaram com Dt
  Ag./Dt Vist. preenchidas corretamente). A 1ª tentativa foi
  interrompida no meio por fechamento acidental do VSCode - reexecutar
  no mesmo Excel é seguro (`_valores_batem` pula quem já foi salvo, só
  processa o resto). `MAPA_TIPO_IMOVEL`/select não se aplicou (Tipo de
  Imóvel é texto livre no Banco B, diferente do Banco A - ver **Etapa
  4**, acima); ainda não vimos um caso real de preenchimento automático
  desse campo (nenhuma proposta chegou com ele vazio nos testes até
  aqui).
- ✅ **Etapa 5** (finalizar completas) — **testada em produção
  (18/09/2026)**: 184 finalizadas e confirmadas, 0 erro, 10 puladas
  (incompletas, corretamente identificadas). Corrigido um bug real
  antes de fechar essa validação: o botão "Finalizar Selecionados"
  dispara um `confirm()` **nativo do navegador**, que o Playwright
  descarta sozinho sem um handler registrado - o clique não dava erro,
  mas nada era finalizado de verdade (ver **Bug real (18/09/2026)**,
  acima).
- O orquestrador `executar_fluxo_banco_b.py` tem a mesma mecânica
  (login único, isolamento por etapa, trava, `--auto`) das etapas já
  validadas individualmente acima, mas **ainda não rodou de ponta a
  ponta numa única execução sem interrupção manual** - a única tentativa
  foi cortada pelo fechamento do VSCode no meio da Etapa 4.
- Só ligar o modo `auto` (agendamento sem supervisão) depois de rodar o
  modo padrão várias vezes e confirmar que as cinco etapas funcionam de
  ponta a ponta — inclusive o ensaio com sessão apagada descrito em
  **Agendando de 6 em 6 horas**, acima.
