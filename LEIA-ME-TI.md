# Qualificador de Captação — Guia para T.I

## Visão geral

App web de página única (HTML puro, sem framework, sem dependências externas) usado pelos captadores da Your Vacation para qualificar clientes em campo.

---

## Arquivos do projeto

| Arquivo | O que é | Quando editar |
|---|---|---|
| `index.html` | **Todo o app** — HTML + CSS + JavaScript em um único arquivo | Qualquer mudança no app |
| `apps-script-backend.js` | Código do backend Google Sheets (Google Apps Script) | Mudar estrutura da planilha ou lógica de sync |
| `deploy.sh` | Script de deploy automático via git | Raramente |
| `_headers` | Headers HTTP para Netlify (não usado atualmente) | Ignorar |

**O único arquivo que você vai editar na maioria das vezes é o `index.html`.**

---

## Como está hospedado

- **Hospedagem:** GitHub Pages (gratuito, deploy automático)
- **Repositório:** https://github.com/erickfaleirosilva-oss/captacao-app
- **URL de produção:** https://erickfaleirosilva-oss.github.io/captacao-app/
- **Branch:** `main` — qualquer push para `main` atualiza o app em ~1-2 minutos

### Deploy (como publicar uma atualização)

```bash
cd /Users/erickfaleiro/Desktop/IA/YourVacation/projects/captacao-app
git add -A
git commit -m "descrição da mudança"
git push origin main
```

Ou simplesmente execute o script:

```bash
bash deploy.sh
```

---

## Backend — Google Sheets

- **Planilha:** https://docs.google.com/spreadsheets/d/16wO36S_m8eGDkScnTbS_BtEvJldFjht9DgsKRczJGK0
- **Webhook (Apps Script):** `https://script.google.com/macros/s/AKfycbwdGNL10ZtmZ3TgsZDJVyBlGF9NB7WR1PSzeWb1_ifgtB7bSkl5QjqUsLpAoDrJ3v_dKQ/exec`
- **Função:** recebe cada lead via POST quando o captador salva uma pesquisa; também fornece GET para sync do CDP

Se precisar re-implantar o Apps Script (ex: mudou o código):
1. Abra a planilha → Extensões → Apps Script
2. Cole o conteúdo de `apps-script-backend.js`
3. Implantar → Gerenciar implantações → Editar → Nova versão → Implantar
4. Copie a nova URL e atualize `CONFIG.webhookUrl` no `index.html`

---

## Estrutura do `index.html`

O arquivo tem ~1.900 linhas divididas em 3 blocos:

```
<html>
  <head>
    <style> ... </style>       ← Todo o CSS (linhas ~1–400)
  </head>
  <body>
    <!-- Telas HTML -->        ← Estrutura das telas (linhas ~400–750)
    <script>
      // Configuração         ← CONFIG, captadores, cargos (linhas ~760–800)
      // Estado global        ← variáveis: leads, captador, step, ans... (linhas ~790–795)
      // Lógica de perguntas  ← array QS com as 8 perguntas (linhas ~966–1035)
      // Qualificação         ← calcResult(), buildFail_low/high() (linhas ~1040–1190)
      // Navegação            ← hideAll(), show(), goTo() (linhas ~1190–1210)
      // Quiz                 ← renderStep(), nextStep(), prevStep() (linhas ~1280–1380)
      // Resultado            ← showResult(), saveAndNew(), skipNew() (linhas ~1380–1460)
      // Histórico            ← showHist(), renderHistList() (linhas ~1455–1560)
      // CDP                  ← showCdp(), renderCdp() (linhas ~1615–1720)
      // Equipe               ← showEquipe(), addMembro() (linhas ~1720–1800)
      // API/Sheets           ← objeto API com getLeads() e postLead() (linhas ~1800–1850)
    </script>
  </body>
</html>
```

---

## Configurações principais (dentro do `index.html`)

### CONFIG (início do bloco `<script>`)

```javascript
const CONFIG = {
  webhookUrl: 'https://script.google.com/macros/s/...', // URL do Sheets
  parcial: { mensagem: '...' },
  // ...
};
```

### Lista de captadores

Procure por `CONFIG.captadores` — é um array com nome, cargo e sala de cada membro fixo da equipe. Membros adicionais ficam em `localStorage` como `yv_equipe_extras`.

### Perguntas do quiz

Procure por `const QS = [` — array com 8 objetos, um por pergunta. Cada objeto tem:
- `id` — identificador (ex: `'renda'`)
- `meta` — título exibido (ex: `'Renda · 7/8'`)
- `type` — `'text'`, `'tel'` ou `'chips'`
- `text` — pergunta exibida
- `opts` — opções (só para `chips`): `{ icon, label, sub, value }`

### Regras de qualificação

Procure por `function calcResult()` — contém toda a lógica de Q / NQ / PARCIAL.

---

## Fluxo do app

```
Login → Seleção de sala → Quiz (8 perguntas) → Resultado (Q/NQ/PARCIAL)
                                                      ↓
                                              Salvar → Sheets + localStorage
```

### Perguntas (ordem atual)
1. Nome do cliente
2. Perfil (casal casado / união estável / noivos / solteiro / multiprop.)
3. Idade
4. Casa própria
5. Veículo (ano)
6. Cartão de crédito
7. Renda mensal
8. Telefone com DDD

---

## Cargos e permissões

| Cargo | O que vê/faz |
|---|---|
| `captador` | Faz quiz, vê histórico próprio, CDP próprio |
| `lider` | Tudo do captador + aba Equipe + histórico/CDP de todos |
| `ceo` | Igual ao líder |

---

## Dados — como são armazenados

- **Por dispositivo:** `localStorage` e `sessionStorage` com chave `yv_leads`
- **Central:** Google Sheets (sincronizado via webhook a cada pesquisa salva)
- **CDP e Histórico:** ao abrir, o app busca do Sheets e faz merge com os locais

---

## Dicas para editar

1. **Sempre valide o JS antes de fazer deploy:**
   ```bash
   node --check index.html  # não funciona direto para .html
   # Use este comando:
   python3 -c "
   import re
   with open('index.html','r') as f: c=f.read()
   m=re.search(r'<script>(.*?)</script>',c,re.DOTALL)
   open('/tmp/check.js','w').write(m.group(1))
   " && node --check /tmp/check.js
   ```

2. **Cuidado com template literals** — dentro do `innerHTML` do JS, use `\`` e `\${}` corretamente. Um erro de sintaxe deixa o app em branco.

3. **O app não tem build step** — o que você editar no `index.html` é exatamente o que vai para produção.

4. **Para testar localmente:** abra o `index.html` direto no navegador (`File > Open`). Funciona sem servidor.

5. **Cache do GitHub Pages** — após o deploy, aguarde 1-2 minutos e teste em aba anônima.
