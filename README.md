# ClusterCRM

> Segmentação K-Means · Perfis com IA · Mensagens WhatsApp

**[→ Ver demo ao vivo](https://banco-clientes-gamma.vercel.app)**

---

## O problema que resolve

Times de CRM em bancos e fintechs gastam horas segmentando clientes manualmente no Excel — e ainda assim as mensagens de relacionamento são genéricas. O ClusterCRM automatiza esse processo em minutos: carrega a base, agrupa por comportamento, gera perfis e cria mensagens personalizadas prontas para disparar no WhatsApp.

---

## Como funciona

1. **Importe** — CSV ou Excel com dados dos clientes
2. **Segmente** — K-Means agrupa automaticamente por comportamento financeiro (renda, saldo, engajamento, uso do app…)
3. **Analise** — Claude API lê as médias de cada cluster e gera nome, descrição e perfil do grupo
4. **Dispare** — Mensagens WhatsApp personalizadas por segmento, prontas para copiar ou abrir direto no app

---

## Stack

```
React · Vite · K-Means (browser) · Claude API (Anthropic) · XLSX
```

---

## Funcionalidades

- Segmentação K-Means rodando direto no browser (sem backend)
- 2 a 6 clusters configuráveis
- Perfis sugeridos editáveis — o usuário nomeia e descreve cada cluster antes de rodar a IA
- Mensagens WhatsApp geradas pela Claude API com tom adaptado ao perfil
- Export em CSV, JSON e XLSX em todas as abas
- Aba "AI WhatsApp" pisca quando as mensagens estão prontas
- Dark mode responsivo — desktop e mobile
- Demo pré-carregada com 40 clientes fictícios

---

## Rodando localmente

```bash
git clone https://github.com/renatasqu/banco-clientes
cd banco-clientes
npm install
npm run dev
```

Acesse `http://localhost:5173` e use o botão **"USAR MEUS DADOS"** para importar sua base.

---

## Formato do CSV

| Coluna | Tipo | Obrigatória |
|---|---|---|
| nome | texto | não |
| idade | número | sim |
| renda_mensal | número | sim |
| saldo_medio | número | sim |
| valor_investido | número | não |
| score_engajamento | número | sim |
| uso_app | número | não |

Qualquer coluna numérica é usada no K-Means. Quanto mais colunas, melhor a segmentação.

---

## Sobre

Projeto de portfólio — [Renata Queiroz](https://linkedin.com/in/renatasampaioqueiroz)

Chef executiva → AI Product Builder · MIT Professional Education in AI · Fundadora do O Combinado
