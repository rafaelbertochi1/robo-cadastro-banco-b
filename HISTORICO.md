# Histórico do projeto

Este repositório é a versão de portfólio de um projeto real, desenvolvido em
ambiente corporativo entre 09/09/2026 e 02/10/2026. O repositório original é
privado e teve 43 commits de trabalho nesse período.

O histórico de commits que você vê aqui é **novo**: o código foi publicado em
poucos commits, um por etapa do fluxo, depois de anonimizado. Este arquivo
conta como o projeto de fato evoluiu.

## O que foi anonimizado

- Nome da empresa e do seu sistema interno: aparece como "Central de Gestão".
- Sistema de laudos de terceiros: aparece como "Plataforma".
- Bancos clientes: "Banco A" e "Banco B".
- URLs e endpoints: trocados por endereços `example.com`.
- Números de proposta e outros dados que apareciam em comentários e exemplos:
  trocados por valores fictícios.

Por causa disso o código não roda contra os endereços de exemplo. A lógica, a
estrutura e a documentação (incluindo os registros de incidentes e decisões no
README) são as do projeto real.

## Como o projeto evoluiu

### 1. Fluxo inicial (09/09 a 10/09)

O robô nasceu como irmão do
[robo-cadastro-banco-a](https://github.com/rafaelbertochi1/robo-cadastro-banco-a),
já com as três primeiras etapas: exportar o relatório da Plataforma, cadastrar
na Central de Gestão e atribuir o engenheiro. Nessa fase:

- filtros de tipo e status de inspeção ajustados à lista real do Banco B;
- correção de um falso positivo na checagem de vistoria das canceladas;
- seletores das Etapas 2 e 3 corrigidos com o HTML real da tela do Banco B;
- conferência de proposta já atribuída antes de salvar de novo.

### 2. Paridade com o robô do Banco A (17/09 a 18/09)

- correções e paralelismo trazidos do robô irmão;
- Etapa 4 (agendar vistoria) e Etapa 5 (finalizar vistorias completas);
- login automático, tratamento da tela 401 e janela de datas maior;
- orquestrador com login único, isolamento por etapa e trava contra execução
  simultânea;
- correção da Etapa 5: o botão de finalizar não enviava nada, por causa de um
  diálogo nativo do navegador que ninguém respondia.

### 3. Modo automático (21/09 a 22/09)

- correção do `--auto`, em que as Etapas 2, 3 e 4 liam a própria flag como se
  fosse o caminho do Excel;
- fim das pausas que esperavam o usuário;
- correção de uma corrida de tempo na Etapa 1;
- sessão válida deixou de ser tratada como morta quando o painel estava lento.

### 4. Verificação em duas etapas (22/09 a 24/09)

Este robô foi o primeiro a resolver a verificação em duas etapas da Plataforma,
e a solução foi depois levada aos outros projetos:

- reconhecimento da tela de MFA;
- resposta automática via App Autenticador (TOTP);
- `configurar_autenticador.py`, que lê o QR code e guarda a chave;
- `COMO_EXECUTAR.md`, um guia rápido de comandos.

### 5. Casos reais de produção (24/09 a 02/10)

- número de proposta com ponto ou hífen deixou de ser dado como ausente;
- nova tentativa quando a página não abre, e código de erro quando o fluxo
  falha;
- exportação de mais status de laudo e janela padrão de 30 dias;
- conferência do engenheiro ignorando espaços repetidos no nome.

## Projetos relacionados

- [robo-cadastro-banco-a](https://github.com/rafaelbertochi1/robo-cadastro-banco-a):
  o robô irmão, para o Banco A, onde o fluxo foi desenvolvido primeiro.
- [pipeline-laudos-banco-a](https://github.com/rafaelbertochi1/pipeline-laudos-banco-a) e
  [pipeline-laudos-banco-b](https://github.com/rafaelbertochi1/pipeline-laudos-banco-b):
  pipelines de download e extração dos laudos em PDF.
