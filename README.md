# Painel Navigator — Board Academy

Painel de closers do **funil Navigator**, no formato da tabela de métricas (Meta / Realizado / Gap / Previsto).

## Estrutura

```
.
├── api/index.py       # backend Flask (serverless function da Vercel)
├── index.html         # front estático
├── requirements.txt
└── vercel.json
```

## Deploy na Vercel

1. Sobe esses 4 arquivos num repo do GitHub.
2. Na Vercel: **Add New → Project → importa o repo**. Não precisa mexer em build settings — a Vercel detecta `api/*.py` como Python Function e serve o `index.html` como estático.
3. Em **Settings → Environment Variables**, adiciona:

| Variável | Obrigatória | Para que serve |
|---|---|---|
| `PIPE_API_KEY` | **sim** | token da API do Pipedrive |
| `FUNIL_NOME` | não | nome do funil. Default: `navigator` |
| `SUBAREAS_META` | não | subáreas do COLAB que compõem a meta. Default: `ascensao` |
| `PAINEL_SENHA` | não | se preenchida, a API só responde com `?k=<senha>` na URL |
| `META_FIXA` | não | sobrescreve a meta somada da planilha (ex: `1180000`) |
| `URL_COLAB` / `URL_METAS` / `URL_FERIADOS` | não | já vêm com as planilhas atuais no default |

4. **Redeploy** depois de salvar as variáveis (a Vercel não injeta env em deploy já feito).

## Conferindo se está tudo certo

Abre `https://SEU-PROJETO.vercel.app/api/health`. Ele responde:

```json
{
  "token_configurado": true,
  "funil_procurado": "navigator",
  "funil_encontrado": "Navigator",
  "pipeline_id": 12,
  "funis_disponiveis": ["Sniper", "Elite", "Navigator", "..."],
  "hoje_br": "2026-09-08"
}
```

Se `funil_encontrado` vier `null`, a lista `funis_disponiveis` mostra os nomes reais —
copia o certo para a variável `FUNIL_NOME`.

## Rodando local

```bash
pip install -r requirements.txt
export PIPE_API_KEY=xxxxx
python api/index.py          # http://localhost:5000/api/navigator
```

Para ver o HTML local, sobe um estático em paralelo (`python -m http.server 8000`) e
troca o `fetch('/api/navigator...')` por `http://localhost:5000/api/navigator...`.

## Fórmulas

| Linha | Cálculo |
|---|---|
| Meta Mês | Σ meta financeira dos **closers** (meta de reunião = 0) com Subarea = Ascensão |
| Meta Dia | Meta Mês ÷ DU total |
| Realizado Bruto | Σ `value` dos deals ganhos no funil, `won_time` em BRT dentro do mês |
| Realizado Multiplicador | Σ do campo multiplicador desses mesmos deals |
| Deveria (100%) — MTD | Meta Dia × DU passados |
| Atingimento | Realizado Multiplicador ÷ Meta Mês |
| Gap 100% | Meta Mês − Realizado Multiplicador |
| Meta/Dia 100% | Gap 100% ÷ DU restantes |
| Meta/Dia 100% (Bruto) | (Meta Mês − Realizado Bruto) ÷ DU restantes |
| Previsto (hoje/ontem) | p20×0,20 + p50×0,50 + p70×0,70 dos deals **abertos** com `expected_close_date` naquele dia |
| Em Aberto Hoje | soma de **todos** os abertos de hoje, inclusive fora dos buckets |
| Entrou (hoje/ontem) | realizado bruto e multiplicado daquele dia |

**DU restantes = DU total − DU passados** (o dia de hoje conta como passado).
"Ontem" é o **último dia útil anterior**, pulando fim de semana e feriado.

## Notas

- Todas as datas usam UTC−3 calculado no código, então não depende do fuso do servidor da Vercel.
- Cache em memória de 2 min (Pipedrive) e 5 min (planilhas) por instância, para não estourar
  o rate limit do token quando várias abas ficarem abertas.
- Deals abertos com probabilidade fora de 20/50/70 ou sem probabilidade **não entram no Previsto**
  (entram só no "Em Aberto"). Quando isso acontece, aparece um aviso amarelo no rodapé do painel.
