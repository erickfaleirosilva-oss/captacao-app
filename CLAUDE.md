# QualifiQ (captacao-app) - Guia do Claude

Leia antes de qualquer tarefa neste repositório.
Dono: Erick Faleiro (CEO Your Vacation). Fuso: America/Sao_Paulo.
Erick não é programador: explicar em linguagem simples e resumida, e executar sozinho tudo que não depender dele.

ATENÇÃO: este repositório é PÚBLICO. Nunca registrar aqui senhas, tokens, dados de clientes ou descrições de falhas de segurança.

## Estrutura
- index.html: o app inteiro (HTML + CSS + JS em um arquivo, ~6.800 linhas).
- apps-script-backend.js e _COLAR_NO_APPS_SCRIPT.txt: backend Google Apps Script (mesmo conteúdo).
- sw.js, manifest.json, ícones: PWA.
- Hospedagem: GitHub Pages, branch main. Push na main publica em 1 a 2 min.

## Seções do JavaScript (index.html)
- CONFIG: captadores, salas, regras, webhookUrl.
- API: post(lead) e getLeads().
- calcResult(): lógica Q / NQ / PARCIAL. Ribeirão Preto tem regra própria em calcRibeiraoPreto().
- goTo(tab): navegação entre home, hist, cdp, equipe, ddp, timerange.
- Telas: showLogin, showEquipe, showCdp, showHist, showDdp, showTimeRange.

## Cargos
captador, sublider, lider, supervisor_adm (leitura total, sem ações destrutivas), ceo.

## Operações
São Pedro, Caldas Novas, Atibaia, Ribeirão Preto.

## Fluxo de edição
1. Ler o trecho relevante antes de editar.
2. Editar com strings únicas.
3. Validar o JS: extrair o conteúdo do <script> e rodar node --check.
4. Commit com mensagem clara e push na main.
5. Conferir no ar depois de 2 min (aba anônima).

## Cuidados
- Captador desativado = linha comentada com // em CONFIG.captadores. Reativar = remover o //.
- Atibaia com operacao null em registros antigos é proposital.
- Backend retorna { status: 'ok' }: checar res.status === 'ok', não res.ok.
- Template literals dentro de innerHTML: erro de sintaxe deixa o app em branco.
- Reimplantar Apps Script: Implantar > Gerenciar implantações > editar > Nova versão. Assim a URL do webhook não muda. "Nova implantação" gera URL nova e exige trocar CONFIG.webhookUrl.
- Usuários podem ver versão antiga por cache: recarregar forçado (Cmd+Shift+R).

## Documentação completa
Fica fora do repositório, na pasta local do Erick (projects/captacao-app): PROMPT-IA-HANDOFF.md, ESTADO-ATUAL.md, RESUMO-HANDOFF-OUTRA-IA.md, PLAYBOOK-QUALIFIQ.md.
