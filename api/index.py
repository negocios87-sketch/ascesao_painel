"""
Board Academy — Painel NAVIGATOR (Closers)
Funil: Navigator  |  Deploy: Vercel (serverless)

Regras de cálculo (validadas com o Rodrigo):
  Meta Mês          = soma das metas financeiras dos CLOSERS com Subarea = Ascensão
  Meta Dia          = Meta Mês / DU total
  Realizado Bruto   = Σ value dos deals GANHOS no funil Navigator (won_time BRT no mês)
  Realizado Multi   = Σ campo multiplicador desses mesmos deals
  Deveria (100%)    = Meta Dia × DU passados
  Atingimento       = Realizado Multi / Meta Mês
  Gap 100%          = Meta Mês − Realizado Multi
  Meta/Dia 100%     = Gap 100% / DU restantes        (DU restantes = DU total − DU passados)
  Meta/Dia (Bruto)  = (Meta Mês − Realizado Bruto) / DU restantes
  Previsto          = p20×0,20 + p50×0,50 + p70×0,70 dos deals ABERTOS naquele dia
  Em Aberto         = soma de TODOS os deals abertos naquele dia (inclusive fora dos buckets)
  Entrou            = realizado (bruto e multi) daquele dia

Variáveis de ambiente:
  PIPE_API_KEY   (obrigatória)  token da API do Pipedrive
  FUNIL_NOME     (opcional)     default "navigator"
  SUBAREAS_META  (opcional)     default "ascensao"  — separadas por vírgula
  PAINEL_SENHA   (opcional)     se setada, exige ?k=<senha> nas chamadas da API
  META_FIXA      (opcional)     sobrescreve a meta somada da planilha
"""

from flask import Flask, jsonify, request, Response
import requests as req
import os
import io
import csv
import math
import time
import calendar
import unicodedata
from datetime import date, datetime, timedelta

app = Flask(__name__)

# ── CONFIG ────────────────────────────────────────────────────
API_KEY = os.environ.get("PIPE_API_KEY", "")
BASE_V1 = "https://boardacademy.pipedrive.com/api/v1"
BASE_V2 = "https://boardacademy.pipedrive.com/api/v2"

CF_MULTIPLICADOR = "7e0e43c2734751f77be292a72527f638a850ad50"

FUNIL_NOME   = os.environ.get("FUNIL_NOME", "navigator")
PAINEL_SENHA = os.environ.get("PAINEL_SENHA", "")
META_FIXA    = os.environ.get("META_FIXA", "")

URL_COLAB = os.environ.get("URL_COLAB", "https://docs.google.com/spreadsheets/d/e/2PACX-1vSvwO3Ag2f2cbkVgR1pJZp6fANQcbualGKlAG50fmOljuEGKZ1gJBbSAjRdO3SomXUEVQOWnTvlfHRd/pub?gid=1782440078&single=true&output=csv")
URL_METAS = os.environ.get("URL_METAS", "https://docs.google.com/spreadsheets/d/e/2PACX-1vSvwO3Ag2f2cbkVgR1pJZp6fANQcbualGKlAG50fmOljuEGKZ1gJBbSAjRdO3SomXUEVQOWnTvlfHRd/pub?gid=0&single=true&output=csv")
URL_FERIADOS = os.environ.get("URL_FERIADOS", "https://docs.google.com/spreadsheets/d/e/2PACX-1vSvwO3Ag2f2cbkVgR1pJZp6fANQcbualGKlAG50fmOljuEGKZ1gJBbSAjRdO3SomXUEVQOWnTvlfHRd/pub?gid=1010928978&single=true&output=csv")

EXCLUIR_PESSOAS = {"priscila ribeiro"}

# Rótulo exibido no painel (só cosmético — o funil buscado continua sendo FUNIL_NOME)
ROTULO_FUNIL = os.environ.get("ROTULO_FUNIL", "ASCENSÃO/MGM")

# Quem aparece em cada tabela. Nomes separados por vírgula, como estão no Pipedrive.
CLOSERS_LISTA = os.environ.get("CLOSERS", "Denise Mussolin,Mylena Oliveira")
SDRS_LISTA    = os.environ.get("SDRS",    "Raphaela Moutinho")
# Só as pendências deste dono aparecem na seção "Pipe a arrumar" (vazio = todos)
PENDENCIAS_DONO = os.environ.get("PENDENCIAS_DONO", "Denise Mussolin")
# Responsáveis ignorados na contagem E no detalhamento de reuniões
EXCLUIR_REU = os.environ.get("EXCLUIR_REU", "Denise Mussolin")
# Quem FAZ a reunião — o negócio precisa ser dela para a reunião validar.
# (o responsável pela atividade é quem AGENDOU, normalmente a SDR)
DONO_REUNIAO = os.environ.get("DONO_REUNIAO", "Denise Mussolin")

# Usados só no cálculo de SDR (mesmos filtros e campo do painel gerente_comercial)
FILTER_ACTIVITIES = int(os.environ.get("FILTER_ACTIVITIES", "1310451"))
FILTER_DEALS_RV   = int(os.environ.get("FILTER_DEALS_RV",   "1431880"))
CF_QUALIFICADOR   = "a6f13cc27c8d041f3af4091283ce0d4fe0913875"


def _lista_norm(raw):
    return [norm(x) for x in raw.split(",") if x.strip()]


def _subareas():
    raw = os.environ.get("SUBAREAS_META", "ascensao")
    return {norm(s) for s in raw.split(",") if s.strip()}


# ── HELPERS ───────────────────────────────────────────────────
def norm(s):
    if not s:
        return ""
    s = str(s).strip().lower()
    return unicodedata.normalize("NFD", s).encode("ascii", "ignore").decode()


def arred(v):
    try:
        f = float(v)
        return 0.0 if math.isnan(f) or math.isinf(f) else round(f, 2)
    except Exception:
        return 0.0


def safe_div(a, b):
    try:
        return float(a) / float(b) if b else 0.0
    except Exception:
        return 0.0


def hoje_br():
    """Data de hoje no fuso de São Paulo (UTC-3), independente do fuso do servidor."""
    return (datetime.utcnow() - timedelta(hours=3)).date()


def agora_br():
    return datetime.utcnow() - timedelta(hours=3)


def cf(deal, key):
    val = deal.get(key)
    if val is None:
        # API v2 devolve os campos personalizados dentro de custom_fields
        val = (deal.get("custom_fields") or {}).get(key)
    if val is None:
        return None
    if isinstance(val, dict):
        return val.get("value") or val.get("label")
    return val


def won_time_br(deal):
    wt = deal.get("won_time") or ""
    if not wt:
        return ""
    try:
        dt = datetime.fromisoformat(str(wt).replace("Z", "+00:00"))
        return (dt - timedelta(hours=3)).strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return str(wt)


def owner_name(deal, users):
    uid = deal.get("user_id")
    if isinstance(uid, dict):
        return uid.get("name", "") or ""
    oid = deal.get("owner_id")
    if isinstance(oid, dict):
        return oid.get("name", "") or ""
    return users.get(oid or uid, "") or ""


def du_mes_total(ano, mes, feriados=frozenset()):
    return sum(1 for d in range(1, calendar.monthrange(ano, mes)[1] + 1)
               if date(ano, mes, d).weekday() < 5 and date(ano, mes, d) not in feriados)


def du_passados(ano, mes, feriados=frozenset()):
    h = hoje_br()
    ultimo = calendar.monthrange(ano, mes)[1]
    return max(sum(1 for d in range(1, min(h.day, ultimo) + 1)
                   if date(ano, mes, d).weekday() < 5 and date(ano, mes, d) not in feriados), 1)


def due_br(act):
    """
    (data, hora) da atividade no fuso de São Paulo.

    A API do Pipedrive devolve due_date/due_time em UTC: uma reunião marcada para
    07/09 às 21:00 BRT chega como due_date 2026-09-08 e due_time 00:00. Sem esta
    conversão, tudo depois das 21:00 cai no dia seguinte e as horas saem 3h adiantadas.

    Atividade de dia inteiro (sem due_time) não tem hora para converter — a data fica
    como está.
    """
    d = str(act.get("due_date", "") or "")[:10]
    t = str(act.get("due_time", "") or "")[:5]
    if not d:
        return "", None
    if not t:
        return d, None
    try:
        dt = datetime.strptime(f"{d} {t}", "%Y-%m-%d %H:%M") - timedelta(hours=3)
        return dt.strftime("%Y-%m-%d"), dt.strftime("%H:%M")
    except Exception:
        return d, t


def ultimo_du(ref, feriados=frozenset()):
    """Último dia útil ANTES de `ref` (pula fim de semana e feriado)."""
    d = ref - timedelta(days=1)
    for _ in range(30):
        if d.weekday() < 5 and d not in feriados:
            return d
        d -= timedelta(days=1)
    return ref - timedelta(days=1)


def limpar_nans(obj):
    if isinstance(obj, dict):
        return {k: limpar_nans(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [limpar_nans(v) for v in obj]
    if isinstance(obj, float) and (math.isnan(obj) or math.isinf(obj)):
        return None
    return obj


# ── CACHE EM MEMÓRIA (por instância) ──────────────────────────
_CACHE = {}


def cached(key, ttl, fn):
    now = time.time()
    hit = _CACHE.get(key)
    if hit and (now - hit[0]) < ttl:
        return hit[1]
    val = fn()
    _CACHE[key] = (now, val)
    return val


# ── SHEETS (sem pandas) ───────────────────────────────────────
def ler_sheet(url):
    r = req.get(url, timeout=15)
    r.encoding = "utf-8"
    r.raise_for_status()
    return list(csv.DictReader(io.StringIO(r.text)))


def find_col(rows, pred, default=None):
    if not rows:
        return default
    for k in rows[0].keys():
        if k and pred(norm(k)):
            return k
    return default


def to_int(v):
    try:
        return int(float(str(v).strip()))
    except Exception:
        return 0


def to_num_br(v):
    """Converte '1.180.000,00' / 'R$ 56.190' / '1180000' em float."""
    try:
        if v is None:
            return 0.0
        s = str(v).replace("R$", "").replace(" ", "").strip()
        if not s:
            return 0.0
        if "," in s:
            s = s.replace(".", "").replace(",", ".")
        return float(s)
    except Exception:
        return 0.0


def buscar_feriados():
    def _fetch():
        try:
            rows = ler_sheet(URL_FERIADOS)
            fer = set()
            for row in rows:
                val = str(list(row.values())[0] or "").strip()
                for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%m/%d/%Y"):
                    try:
                        fer.add(datetime.strptime(val, fmt).date())
                        break
                    except Exception:
                        continue
            return fer
        except Exception:
            return set()
    return cached("feriados", 3600, _fetch)


def buscar_colaboradores(mes, ano):
    def _fetch():
        rows = ler_sheet(URL_COLAB)
        c_nome = find_col(rows, lambda c: c == "nome", "Nome")
        c_sub = find_col(rows, lambda c: c == "subarea")
        c_status = find_col(rows, lambda c: "status" in c)
        c_mes = find_col(rows, lambda c: "mes" in c and "ref" in c)
        c_ano = find_col(rows, lambda c: "ano" in c and "ref" in c)

        base = rows
        if c_mes and c_ano:
            filtrado = [r for r in rows
                        if to_int(r.get(c_mes)) == mes and to_int(r.get(c_ano)) == ano]
            if filtrado:
                base = filtrado

        vistos, out = set(), []
        for r in base:
            nome = str(r.get(c_nome, "") or "").strip()
            nn = norm(nome)
            if not nn or nn in vistos:
                continue
            if c_status and norm(r.get(c_status, "")) != "ativo":
                continue
            vistos.add(nn)
            out.append({
                "nome": nome,
                "nome_norm": nn,
                "subarea": str(r.get(c_sub, "") or "").strip() if c_sub else "",
            })
        return out
    return cached(f"colab:{ano}-{mes}", 300, _fetch)


def buscar_metas(ano, mes):
    def _fetch():
        rows = ler_sheet(URL_METAS)
        c_ano = find_col(rows, lambda c: c == "ano")
        c_mes = find_col(rows, lambda c: c == "mes")
        c_nome = find_col(rows, lambda c: c == "nome")
        c_reu = find_col(rows, lambda c: "reuni" in c and "meta" in c)
        c_fin = find_col(rows, lambda c: "financ" in c)
        c_du = find_col(rows, lambda c: "util" in c)

        out = []
        for r in rows:
            if c_ano and to_int(r.get(c_ano)) != ano:
                continue
            if c_mes and to_int(r.get(c_mes)) != mes:
                continue
            nome = str(r.get(c_nome, "") or "").strip() if c_nome else ""
            out.append({
                "nome": nome,
                "nome_norm": norm(nome),
                "meta_reu": to_num_br(r.get(c_reu)) if c_reu else 0.0,
                "meta_fin": to_num_br(r.get(c_fin)) if c_fin else 0.0,
                "dias_uteis": to_int(r.get(c_du)) if c_du else 0,
            })
        return out
    return cached(f"metas:{ano}-{mes}", 300, _fetch)


# ── PIPEDRIVE ─────────────────────────────────────────────────
def buscar_users():
    def _fetch():
        r = req.get(f"{BASE_V1}/users", params={"api_token": API_KEY}, timeout=20)
        r.raise_for_status()
        return {u["id"]: u.get("name", "") for u in (r.json().get("data") or [])}
    return cached("users", 900, _fetch)


def buscar_pipeline_id():
    """Descobre o id do funil pelo nome (FUNIL_NOME)."""
    def _fetch():
        r = req.get(f"{BASE_V1}/pipelines", params={"api_token": API_KEY}, timeout=20)
        r.raise_for_status()
        pipes = r.json().get("data") or []
        alvo = norm(FUNIL_NOME)
        for p in pipes:
            if norm(p.get("name")) == alvo:
                return {"id": p["id"], "nome": p.get("name"),
                        "todos": [p.get("name") for p in pipes]}
        for p in pipes:                      # fallback: match parcial
            if alvo in norm(p.get("name")):
                return {"id": p["id"], "nome": p.get("name"),
                        "todos": [p.get("name") for p in pipes]}
        return {"id": None, "nome": None, "todos": [p.get("name") for p in pipes]}
    return cached("pipeline", 3600, _fetch)


def buscar_ganhos(mes, ano, pipeline_id):
    """
    Deals GANHOS do funil, com won_time (BRT) dentro do mês.
    Ordena por won_time DESC e para assim que passa do mês alvo — igual ao painel atual.
    O pipeline_id vai na query E é revalidado em Python (defensivo).
    """
    def _fetch():
        mes_str = f"{ano}-{mes:02d}"
        todos, start = [], 0
        while True:
            r = req.get(f"{BASE_V1}/deals", params={
                "pipeline_id": pipeline_id,
                "status": "won",
                "sort": "won_time DESC",
                "limit": 500,
                "start": start,
                "api_token": API_KEY,
            }, timeout=30)
            r.raise_for_status()
            data = r.json()
            lote = data.get("data") or []
            achou_antigo = False
            for d in lote:
                if d.get("pipeline_id") != pipeline_id:
                    continue
                wt = won_time_br(d)[:7]
                if wt == mes_str:
                    todos.append(d)
                elif wt and wt < mes_str:
                    achou_antigo = True
            mais = (data.get("additional_data", {})
                        .get("pagination", {})
                        .get("more_items_in_collection", False))
            if not mais or not lote or achou_antigo:
                break
            start += 500
        return todos
    return cached(f"ganhos:{pipeline_id}:{ano}-{mes}", 120, _fetch)


def buscar_abertos(pipeline_id):
    """Deals ABERTOS do funil (pipeline aberto costuma ser pequeno)."""
    def _fetch():
        todos, cursor = [], None
        while True:
            params = {"pipeline_id": pipeline_id, "status": "open", "limit": 500}
            if cursor:
                params["cursor"] = cursor
            r = req.get(f"{BASE_V2}/deals", params=params,
                        headers={"x-api-token": API_KEY}, timeout=30)
            r.raise_for_status()
            data = r.json()
            lote = data.get("data") or []
            todos += [d for d in lote if d.get("pipeline_id") == pipeline_id]
            cursor = (data.get("additional_data") or {}).get("next_cursor")
            if not cursor or not lote:
                break
        return todos
    return cached(f"abertos:{pipeline_id}", 120, _fetch)



def buscar_pipelines_mapa():
    """Mapa pipeline_id -> nome do funil."""
    def _fetch():
        r = req.get(f"{BASE_V1}/pipelines", params={"api_token": API_KEY}, timeout=20)
        r.raise_for_status()
        return {p["id"]: p.get("name", "") for p in (r.json().get("data") or [])}
    return cached("pipes_mapa", 3600, _fetch)


def buscar_qual_ids():
    """Mapa nome_normalizado -> id da opção no campo Qualificador."""
    def _fetch():
        r = req.get(f"{BASE_V1}/dealFields", params={"api_token": API_KEY}, timeout=25)
        r.raise_for_status()
        for field in (r.json().get("data") or []):
            if field.get("key") == CF_QUALIFICADOR:
                return {norm(o.get("label", "")): str(o.get("id"))
                        for o in (field.get("options") or [])}
        return {}
    return cached("qual_ids", 900, _fetch)


def buscar_activities_mes(mes, ano):
    """Atividades do filtro de reuniões, com due_date dentro do mês."""
    def _fetch():
        # A janela em UTC precisa ser mais larga que o mês em BR: uma atividade de
        # 30/09 21:00 BRT chega como 01/10 00:00 UTC. Filtra largo aqui e converte
        # com due_br() na hora de usar.
        ini = date(ano, mes, 1).isoformat()
        prox = date(ano + (1 if mes == 12 else 0), 1 if mes == 12 else mes + 1, 1)
        fim = prox.isoformat()
        todos, cursor = [], None
        while True:
            params = {"filter_id": FILTER_ACTIVITIES, "limit": 500}
            if cursor:
                params["cursor"] = cursor
            r = req.get(f"{BASE_V2}/activities", params=params,
                        headers={"x-api-token": API_KEY}, timeout=30)
            r.raise_for_status()
            data = r.json()
            lote = data.get("data") or []
            todos += [a for a in lote
                      if ini <= str(a.get("due_date", "") or "")[:10] <= fim]
            cursor = (data.get("additional_data") or {}).get("next_cursor")
            if not cursor or not lote:
                break
        return todos
    return cached(f"acts:{ano}-{mes}", 300, _fetch)


def buscar_deals_rv():
    """Deals do filtro de Reunião Validada: ids válidos + mapa deal -> owner."""
    def _fetch():
        ids, mapa, start = set(), {}, 0
        while True:
            r = req.get(f"{BASE_V1}/deals", params={
                "filter_id": FILTER_DEALS_RV, "status": "all_not_deleted",
                "limit": 500, "start": start, "api_token": API_KEY,
            }, timeout=30)
            r.raise_for_status()
            data = r.json()
            lote = data.get("data") or []
            for d in lote:
                ids.add(d["id"])
                uid = d.get("user_id")
                mapa[d["id"]] = {
                    "owner": uid.get("id") if isinstance(uid, dict) else uid,
                    "pipeline_id": d.get("pipeline_id"),
                    "titulo": d.get("title", ""),
                    "status": d.get("status"),
                }
            mais = (data.get("additional_data", {})
                        .get("pagination", {})
                        .get("more_items_in_collection", False))
            if not mais or not lote:
                break
            start += 500
        return ids, mapa
    return cached("deals_rv", 300, _fetch)


def calcular_sdrs(mes, ano, metas, ganhos, users, du_total, du_pass):
    """
    Métricas de SDR, mesma régua do painel gerente_comercial:
      validadas     = atividade concluída, cujo responsável != dono do deal,
                      e cujo deal está no filtro de Reunião Validada
      deveria_estar = meta_reu / DU total x DU passados
      pct_ganhos    = valor com multiplicador / meta financeira
      pct_final     = pct_reu x PESO_REU + pct_ganhos x PESO_FIN  (70/30 desde mai/2026)

    O financeiro sai dos ganhos DESTE funil — o painel é do funil, não da pessoa.
    """
    alvo = _lista_norm(SDRS_LISTA)
    if not alvo:
        return []

    if (ano > 2026) or (ano == 2026 and mes >= 5):
        PESO_REU, PESO_FIN = 0.70, 0.30
    else:
        PESO_REU, PESO_FIN = 0.50, 0.50

    acts = buscar_activities_mes(mes, ano)
    ids_rv, mapa_deal_owner = buscar_deals_rv()
    qual_ids = buscar_qual_ids()

    # os ganhos do mês também entram no mapa deal -> owner
    for d in ganhos:
        if d["id"] not in mapa_deal_owner:
            uid = d.get("user_id")
            mapa_deal_owner[d["id"]] = {
                "owner": uid.get("id") if isinstance(uid, dict) else uid,
                "pipeline_id": d.get("pipeline_id"),
                "titulo": d.get("title", ""),
            }

    nome_por_uid = {uid: nome for uid, nome in users.items()}
    uid_por_nome = {norm(nome): uid for uid, nome in users.items()}
    metas_por_nome = {m["nome_norm"]: m for m in metas}

    mes_str = f"{ano}-{mes:02d}"
    acts_por_owner = {}
    for a in acts:
        if due_br(a)[0][:7] != mes_str:     # fora do mês depois de converter
            continue
        acts_por_owner.setdefault(str(a.get("owner_id", "")), []).append(a)

    def valida(a):
        if not (a.get("done") is True or a.get("status") == "done"):
            return False
        deal_id = a.get("deal_id")
        dono_act = str(a.get("owner_id", ""))
        dono_deal = str((mapa_deal_owner.get(deal_id) or {}).get("owner", "")) if deal_id else ""
        if dono_act and dono_deal and dono_act == dono_deal:
            return False
        if deal_id and deal_id not in ids_rv:
            return False
        return True

    out = []
    for nn in alvo:
        uid = uid_por_nome.get(nn)
        nome = nome_por_uid.get(uid, nn.title())
        m = metas_por_nome.get(nn, {})
        meta_reu = m.get("meta_reu", 0.0)
        meta_fin = m.get("meta_fin", 0.0)

        validadas = len([a for a in acts_por_owner.get(str(uid), []) if valida(a)])
        deveria = arred(safe_div(meta_reu, du_total) * du_pass)
        pct_reu = arred(safe_div(validadas, meta_reu) * 100) if meta_reu else None

        qid = qual_ids.get(nn)
        deals_sdr = [d for d in ganhos
                     if qid and str(cf(d, CF_QUALIFICADOR)) == str(qid)]
        valor_bruto = sum(float(d.get("value") or 0) for d in deals_sdr)
        valor_multi = sum(float(cf(d, CF_MULTIPLICADOR) or 0) for d in deals_sdr)
        pct_ganhos = arred(safe_div(valor_multi, meta_fin) * 100) if meta_fin else None

        pct_final = (arred((pct_reu or 0) * PESO_REU + (pct_ganhos or 0) * PESO_FIN)
                     if (meta_reu or meta_fin) else None)

        out.append({
            "nome": nome,
            "encontrado_no_pipedrive": uid is not None,
            "meta_reuniao": arred(meta_reu),
            "meta_diaria": arred(safe_div(meta_reu, du_total)),
            "validadas": validadas,
            "deveria_estar": deveria,
            "faltam": arred(deveria - validadas),
            "pct_reu": pct_reu,
            "meta_ganho": arred(meta_fin),
            "qtd_ganhos": len(deals_sdr),
            "valor_ganho": arred(valor_bruto),
            "valor_ganho_multi": arred(valor_multi),
            "pct_ganhos": pct_ganhos,
            "ticket_medio": arred(safe_div(valor_bruto, len(deals_sdr))) if deals_sdr else 0.0,
            "pct_final": pct_final,
            "pesos": f"{int(PESO_REU*100)}/{int(PESO_FIN*100)}",
        })
    return out


# ── CÁLCULO ───────────────────────────────────────────────────
def bucket_dias(deals_abertos, dias):
    """Buckets 20/50/70 dos abertos cuja data prevista cai em `dias` (set ou lista)."""
    alvo = {dias} if isinstance(dias, str) else set(dias)
    p20 = p50 = p70 = outros = sem_prob = 0.0
    qtd = 0
    for d in deals_abertos:
        if str(d.get("expected_close_date") or "")[:10] not in alvo:
            continue
        v = float(d.get("value") or 0)
        qtd += 1
        pr = d.get("probability")
        if pr == 20:
            p20 += v
        elif pr == 50:
            p50 += v
        elif pr == 70:
            p70 += v
        elif pr is None:
            sem_prob += v
        else:
            outros += v
    return {
        "p20": arred(p20), "p50": arred(p50), "p70": arred(p70),
        "fora_bucket": arred(outros), "sem_probabilidade": arred(sem_prob),
        "previsto": arred(p20 * 0.20 + p50 * 0.50 + p70 * 0.70),
        "em_aberto": arred(p20 + p50 + p70 + outros + sem_prob),
        "qtd": qtd,
    }


def bucket_dia(deals_abertos, dia_str):
    """Atalho para um único dia."""
    return bucket_dias(deals_abertos, dia_str)


PROBS_VALIDAS = (20, 50, 70)

MOTIVOS = {
    "sem_data":          "Sem data prevista de fechamento",
    "sem_probabilidade": "Sem probabilidade preenchida",
    "prob_fora":         "Probabilidade fora de 20/50/70",
}


def montar_pendencias(abertos, users, so_do_dono=""):
    """
    Deals ABERTOS do funil com informacao faltando - e a lista que a gestora usa
    para arrumar o pipe no Pipedrive. Tres problemas, em ordem de gravidade:

      1. sem_data          -> o deal nao aparece em NENHUM dia do painel
      2. sem_probabilidade -> entra em "Em Aberto", fica fora do "Previsto"
      3. prob_fora         -> idem; a regua do painel so pondera 20/50/70
    """
    dono_alvo = norm(so_do_dono)
    itens = []
    for d in abertos:
        if dono_alvo and norm(owner_name(d, users)) != dono_alvo:
            continue
        dt = str(d.get("expected_close_date") or "")[:10]
        pr = d.get("probability")

        if not dt:
            motivo = "sem_data"
        elif pr is None:
            motivo = "sem_probabilidade"
        elif pr not in PROBS_VALIDAS:
            motivo = "prob_fora"
        else:
            continue

        itens.append({
            "id": d.get("id"),
            "titulo": d.get("title") or "(sem titulo)",
            "dono": owner_name(d, users) or "- sem dono -",
            "valor": arred(float(d.get("value") or 0)),
            "probabilidade": pr,
            "prev_fechamento": dt or None,
            "motivo": motivo,
            "motivo_label": MOTIVOS[motivo],
            "url": f"https://boardacademy.pipedrive.com/deal/{d.get('id')}",
        })

    ordem = {"sem_data": 0, "sem_probabilidade": 1, "prob_fora": 2}
    itens.sort(key=lambda x: (ordem[x["motivo"]], -x["valor"]))

    considerados = [d for d in abertos
                    if not dono_alvo or norm(owner_name(d, users)) == dono_alvo]

    resumo = {}
    for k, label in MOTIVOS.items():
        sel = [i for i in itens if i["motivo"] == k]
        resumo[k] = {"label": label, "qtd": len(sel),
                     "valor": arred(sum(i["valor"] for i in sel))}

    return {
        "itens": itens,
        "resumo": resumo,
        "qtd_total": len(itens),
        "valor_total": arred(sum(i["valor"] for i in itens)),
        "qtd_abertos": len(considerados),
        "valor_abertos": arred(sum(float(d.get("value") or 0) for d in considerados)),
        "filtrado_por": so_do_dono or None,
    }



# ── AS TRÊS FRENTES DA EQUIPE ASCENSÃO/MGM ───────────────────
# Navigator  = alunos já existentes (funil Navigator)
# MGM        = venda nova              (funil MGM, id 92)
# Renovação  = também no funil Navigator, separado por uma TAG (campo personalizado)
PIPELINE_MGM = int(os.environ.get("PIPELINE_MGM", "92"))

# Enquanto a tag não for configurada, Renovação fica zerada e o Navigator vem inteiro.
# TAG_RENOVACAO_CAMPO = a chave (hash) do campo personalizado no Pipedrive
# TAG_RENOVACAO_VALOR = o valor/label que identifica renovação
TAG_RENOVACAO_CAMPO = os.environ.get(
    "TAG_RENOVACAO_CAMPO", "54fc9258843cdf7ea126b6c5aca9d4dc93a3a718")
TAG_RENOVACAO_VALOR = os.environ.get("TAG_RENOVACAO_VALOR", "Renovacao_IC")

# Só negócios nesta etapa entram na previsão (20/50/70) e no pipe do dia.
# Casa por trecho do nome, então "Negocia" pega "Negociação" em qualquer funil.
ETAPA_PREVISAO = os.environ.get("ETAPA_PREVISAO", "Negocia")

# Metas por frente e por mês.
# Set/26 revisado para 500k: os 40k a mais entraram no Navigator (300k -> 340k),
# batendo com a meta da Denise na planilha METAS.
METAS_FRENTES = {
    (2026,  9): {"navigator": 340000, "mgm": 120000, "renovacao": 40000},
    (2026, 10): {"navigator": 313600, "mgm": 127300, "renovacao": 47300},
    (2026, 11): {"navigator": 327300, "mgm": 134500, "renovacao": 54500},
    (2026, 12): {"navigator": 340900, "mgm": 141800, "renovacao": 61800},
}

NOMES_FRENTES = {
    "navigator": "Ascensão (Navigator)",
    "mgm":       "Novos Negócios (MGM)",
    "renovacao": "Renovação",
}


def valores_tag_renovacao():
    """
    Aceita tanto o label ("Renovacao_IC") quanto o id da opção — campos de seleção
    do Pipedrive devolvem o id, não o texto. Resolve os dois via /dealFields.
    """
    def _fetch():
        aceitos = {norm(TAG_RENOVACAO_VALOR), str(TAG_RENOVACAO_VALOR).strip()}
        try:
            r = req.get(f"{BASE_V1}/dealFields", params={"api_token": API_KEY}, timeout=25)
            r.raise_for_status()
            for field in (r.json().get("data") or []):
                if field.get("key") != TAG_RENOVACAO_CAMPO:
                    continue
                for o in (field.get("options") or []):
                    if norm(o.get("label", "")) == norm(TAG_RENOVACAO_VALOR):
                        aceitos.add(str(o.get("id")))
        except Exception:
            pass
        return {a for a in aceitos if a}
    return cached("tag_renov", 900, _fetch)


def eh_renovacao(deal):
    """True se o negócio carrega a tag de Renovação."""
    if not TAG_RENOVACAO_CAMPO or not TAG_RENOVACAO_VALOR:
        return False
    val = cf(deal, TAG_RENOVACAO_CAMPO)
    if val is None:
        return False
    aceitos = valores_tag_renovacao()

    def bate(v):
        if isinstance(v, dict):
            return (str(v.get("id")) in aceitos
                    or norm(str(v.get("label", ""))) in aceitos)
        return str(v).strip() in aceitos or norm(str(v)) in aceitos

    if isinstance(val, list):
        return any(bate(v) for v in val)
    return bate(val)


def buscar_stages_mapa():
    """Mapa stage_id -> nome da etapa."""
    def _fetch():
        r = req.get(f"{BASE_V1}/stages", params={"api_token": API_KEY, "limit": 500},
                    timeout=20)
        r.raise_for_status()
        return {s["id"]: s.get("name", "") for s in (r.json().get("data") or [])}
    return cached("stages", 3600, _fetch)


def stages_previsao():
    """Ids das etapas que contam para a previsão. None = filtro desligado."""
    alvo = norm(ETAPA_PREVISAO)
    if not alvo:
        return None
    ids = {sid for sid, nome in buscar_stages_mapa().items() if alvo in norm(nome)}
    return ids or None          # nenhuma etapa casou: não filtra nada


def so_em_negociacao(deals):
    """Filtra para as etapas de previsão. Devolve (lista, filtrou?)."""
    ids = stages_previsao()
    if ids is None:
        return list(deals), False
    return [d for d in deals if d.get("stage_id") in ids], True


def contar_reunioes(mes, ano, users, dias, deals_extra=None, pipelines_escopo=None):
    """
    Reuniões por dia.

    Quem define a reunião é o DONO DO NEGÓCIO (DONO_REUNIAO — a Denise, que conduz),
    não quem agendou. Reunião agendada por qualquer pessoa conta, desde que o negócio
    seja dela e esteja em funil do escopo.

      agendadas  = atividades com negócio da Denise, em funil do escopo, com due_date
                   naquele dia (BRT), deduplicadas por negócio
      realizadas = as concluídas que também estão no filtro de Reunião Validada

    Descartes ficam registrados: "fora_escopo" (outro funil) e "fora_dono" (negócio de
    outra pessoa), para o painel poder mostrar quanto ficou de fora e por quê.

    Com DONO_REUNIAO vazio, volta à régua antiga: filtra pelo responsável (CLOSERS+SDRS
    menos EXCLUIR_REU) e exige responsável != dono do negócio.
    """
    dono_alvo = norm(DONO_REUNIAO)
    por_dono = bool(dono_alvo)

    equipe = set(_lista_norm(CLOSERS_LISTA)) | set(_lista_norm(SDRS_LISTA))
    excluidos = set(_lista_norm(EXCLUIR_REU))
    uids_equipe = {str(uid) for uid, nome in users.items()
                   if norm(nome) in equipe and norm(nome) not in excluidos}
    nome_por_uid = {str(uid): nome for uid, nome in users.items()}

    fora_escopo = fora_dono = 0
    contagem = {d: {"agendadas": 0, "realizadas": 0} for d in dias}
    detalhe = {d: [] for d in dias}
    if not por_dono and not uids_equipe:
        return {"contagem": contagem, "detalhe": detalhe, "fora_escopo": 0, "fora_dono": 0}

    acts = buscar_activities_mes(mes, ano)
    ids_rv, mapa_deal = buscar_deals_rv()
    pipes = buscar_pipelines_mapa()

    for d in (deals_extra or []):
        uid = d.get("user_id")
        dono = (uid.get("id") if isinstance(uid, dict)
                else (d.get("owner_id") or {}).get("id")
                if isinstance(d.get("owner_id"), dict) else d.get("owner_id"))
        mapa_deal.setdefault(d.get("id"), {
            "owner": dono, "pipeline_id": d.get("pipeline_id"), "titulo": d.get("title", ""),
            "status": d.get("status"),
        })

    for a in acts:
        resp_uid = str(a.get("owner_id", ""))
        if not por_dono and resp_uid not in uids_equipe:
            continue

        dia, hora = due_br(a)                      # UTC -> BRT
        if dia not in contagem:
            continue

        deal_id = a.get("deal_id")
        if not deal_id:                            # sem negócio vinculado, não conta
            continue
        info = mapa_deal.get(deal_id) or {}

        pid_deal = info.get("pipeline_id")
        if pipelines_escopo and pid_deal is not None and pid_deal not in pipelines_escopo:
            fora_escopo += 1
            continue

        dono_uid = str(info.get("owner", ""))
        nome_dono = nome_por_uid.get(dono_uid, "")
        if por_dono and norm(nome_dono) != dono_alvo:
            fora_dono += 1
            continue

        concluida = a.get("done") is True or a.get("status") == "done"
        motivo = None
        if not concluida:
            status = "pendente"
        elif not por_dono and resp_uid and dono_uid and resp_uid == dono_uid:
            status, motivo = "nao_validada", "quem agendou é o dono do negócio"
        elif deal_id not in ids_rv:
            status, motivo = "nao_validada", "fora do filtro de Reunião Validada"
        else:
            status = "realizada"

        detalhe[dia].append({
            "deal_id": deal_id,
            "responsavel": nome_por_uid.get(resp_uid, "—"),
            "hora": hora or None,
            "funil": pipes.get(pid_deal) or None,
            "assunto": a.get("subject") or "(sem assunto)",
            "deal": info.get("titulo") or None,
            "deal_url": f"https://boardacademy.pipedrive.com/deal/{deal_id}",
            "status": status,
            "motivo": motivo,
            "ganho": info.get("status") == "won",
        })

    # ── dedup: mesmo dia + mesmo negócio = uma reunião só ─────
    ordem = {"realizada": 0, "nao_validada": 1, "pendente": 2}
    duplicadas = {}
    for dia, itens in detalhe.items():
        melhor = {}
        for it in itens:
            chave = it["deal_id"]
            atual = melhor.get(chave)
            if atual is None or ((ordem[it["status"]], it["hora"] or "99:99")
                                 < (ordem[atual["status"]], atual["hora"] or "99:99")):
                melhor[chave] = it
        duplicadas[dia] = len(itens) - len(melhor)
        lista = sorted(melhor.values(),
                       key=lambda x: (x["hora"] or "99:99", ordem[x["status"]]))
        detalhe[dia] = lista
        contagem[dia] = {
            "agendadas": len(lista),
            "realizadas": sum(1 for i in lista if i["status"] == "realizada"),
            "duplicadas_ocultas": duplicadas[dia],
        }

    # ── taxa de conversão ────────────────────────────────────
    # negócios DISTINTOS com reunião validada no período, e quantos deles estão Ganho.
    # Um mesmo negócio com duas reuniões no mês conta uma vez só.
    deals_com_reu, deals_ganhos = set(), set()
    for itens in detalhe.values():
        for it in itens:
            if it["status"] != "realizada":
                continue
            deals_com_reu.add(it["deal_id"])
            if it.get("ganho"):
                deals_ganhos.add(it["deal_id"])
    conversao = {
        "deals_com_reuniao": len(deals_com_reu),
        "deals_ganhos": len(deals_ganhos),
        "taxa": (arred(safe_div(len(deals_ganhos), len(deals_com_reu)) * 100)
                 if deals_com_reu else None),
    }

    return {"contagem": contagem, "detalhe": detalhe, "conversao": conversao,
            "fora_escopo": fora_escopo, "fora_dono": fora_dono}


def calcular_navigator(mes=None, ano=None):
    hoje = hoje_br()
    mes = mes or hoje.month
    ano = ano or hoje.year

    feriados = buscar_feriados()
    metas = buscar_metas(ano, mes)
    colab = buscar_colaboradores(mes, ano)
    subareas = _subareas()

    # ── Meta Mês = soma dos CLOSERS da Ascensão ────────────────
    nome_to_sub = {c["nome_norm"]: norm(c["subarea"]) for c in colab}
    meta_mes, closers_meta = 0.0, []
    for m in metas:
        nn = m["nome_norm"]
        if not nn or nn in EXCLUIR_PESSOAS:
            continue
        if not (m["meta_reu"] == 0 and m["meta_fin"] > 0):   # closer
            continue
        if nome_to_sub.get(nn) not in subareas:
            continue
        meta_mes += m["meta_fin"]
        closers_meta.append({"nome": m["nome"], "meta": arred(m["meta_fin"])})

    if META_FIXA:
        meta_mes = to_num_br(META_FIXA)

    # ── Dias úteis ────────────────────────────────────────────
    du_calc = du_mes_total(ano, mes, feriados)
    du_sheet = next((m["dias_uteis"] for m in metas if m["dias_uteis"] > 0), 0)
    du_total = du_sheet if du_sheet > 0 else du_calc

    mes_fechado = (ano < hoje.year) or (ano == hoje.year and mes < hoje.month)
    mes_futuro = (ano > hoje.year) or (ano == hoje.year and mes > hoje.month)
    if mes_fechado:
        du_pass = du_total
    elif mes_futuro:
        du_pass = 0
    else:
        du_pass = min(du_passados(ano, mes, feriados), du_total)
    du_rest = max(du_total - du_pass, 0)

    # ── Pipedrive ─────────────────────────────────────────────
    pipe = buscar_pipeline_id()
    pid = pipe["id"]
    if not pid:
        return {
            "erro": f"Funil '{FUNIL_NOME}' não encontrado no Pipedrive.",
            "funis_disponiveis": pipe["todos"],
        }

    users = buscar_users()

    # O Painel do Mês é CONSOLIDADO: Navigator + MGM juntos.
    # A separação por frente vive só na aba "Resumo — 3 Frentes".
    ganhos_nav = buscar_ganhos(mes, ano, pid)
    abertos_nav = buscar_abertos(pid)

    erro_mgm = None
    ganhos_mgm, abertos_mgm = [], []
    if PIPELINE_MGM:
        try:
            ganhos_mgm = buscar_ganhos(mes, ano, PIPELINE_MGM)
            abertos_mgm = buscar_abertos(PIPELINE_MGM)
        except Exception as e:
            erro_mgm = f"{type(e).__name__}: {e}"

    ganhos = ganhos_nav + ganhos_mgm
    abertos_todos = abertos_nav + abertos_mgm
    # A previsão (20/50/70) e o pipe do dia só consideram a etapa de Negociação
    abertos, filtrou_etapa = so_em_negociacao(abertos_todos)

    composicao = {
        "etapa_previsao": ETAPA_PREVISAO if filtrou_etapa else None,
        "abertos_negociacao": len(abertos),
        "abertos_total": len(abertos_todos),
        "navigator": {"qtd": len(ganhos_nav),
                      "valor": arred(sum(float(d.get("value") or 0) for d in ganhos_nav)),
                      "abertos": len(abertos_nav)},
        "mgm": {"qtd": len(ganhos_mgm),
                "valor": arred(sum(float(d.get("value") or 0) for d in ganhos_mgm)),
                "abertos": len(abertos_mgm)},
    }

    # ── Realizado ─────────────────────────────────────────────
    real_bruto = sum(float(d.get("value") or 0) for d in ganhos)
    real_multi = sum(float(cf(d, CF_MULTIPLICADOR) or 0) for d in ganhos)

    por_dia = {}
    for d in ganhos:
        dia = won_time_br(d)[:10]
        if not dia:
            continue
        acc = por_dia.setdefault(dia, {"bruto": 0.0, "multi": 0.0, "qtd": 0})
        acc["bruto"] += float(d.get("value") or 0)
        acc["multi"] += float(cf(d, CF_MULTIPLICADOR) or 0)
        acc["qtd"] += 1

    hoje_str = hoje.strftime("%Y-%m-%d")
    dia_ontem = ultimo_du(hoje, feriados)
    ontem_str = dia_ontem.strftime("%Y-%m-%d")

    # Semana corrente: domingo -> sábado (a tabela mostra a semana; os cards, o dia)
    dias_desde_dom = (hoje.weekday() + 1) % 7
    dom = hoje - timedelta(days=dias_desde_dom)
    sab = dom + timedelta(days=6)
    dias_semana = [(dom + timedelta(days=i)).strftime("%Y-%m-%d") for i in range(7)]

    b_hoje = bucket_dia(abertos, hoje_str)
    b_ontem = bucket_dia(abertos, ontem_str)
    b_semana = bucket_dias(abertos, dias_semana)
    g_semana = {"bruto": 0.0, "multi": 0.0, "qtd": 0}
    for d in dias_semana:
        acc = por_dia.get(d)
        if acc:
            g_semana["bruto"] += acc["bruto"]
            g_semana["multi"] += acc["multi"]
            g_semana["qtd"] += acc["qtd"]

    g_hoje = por_dia.get(hoje_str, {"bruto": 0.0, "multi": 0.0, "qtd": 0})
    g_ontem = por_dia.get(ontem_str, {"bruto": 0.0, "multi": 0.0, "qtd": 0})

    # ── Métricas do painel ────────────────────────────────────
    meta_dia = safe_div(meta_mes, du_total)
    deveria_mtd = meta_dia * du_pass
    gap_100 = meta_mes - real_multi
    gap_100_bruto = meta_mes - real_bruto

    metricas = {
        "meta_mes":           arred(meta_mes),
        "meta_dia":           arred(meta_dia),
        "real_bruto":         arred(real_bruto),
        "real_multi":         arred(real_multi),
        "deveria_mtd":        arred(deveria_mtd),
        "pct_mes_decorrido":  arred(safe_div(du_pass, du_total) * 100),
        "atingimento":        arred(safe_div(real_multi, meta_mes) * 100),
        "atingimento_bruto":  arred(safe_div(real_bruto, meta_mes) * 100),
        "gap_100":            arred(gap_100),
        "meta_dia_100":       arred(safe_div(gap_100, du_rest)) if du_rest else 0.0,
        "meta_dia_100_bruto": arred(safe_div(gap_100_bruto, du_rest)) if du_rest else 0.0,
        "previsto_ontem":     b_ontem["previsto"],
        "entrou_ontem_multi": arred(g_ontem["multi"]),
        "entrou_ontem_bruto": arred(g_ontem["bruto"]),
        "previsto_hoje":      b_hoje["previsto"],
        "em_aberto_hoje":     b_hoje["em_aberto"],
        "entrou_hoje_multi":  arred(g_hoje["multi"]),
        "entrou_hoje_bruto":  arred(g_hoje["bruto"]),
        "qtd_ganhos_mes":     len(ganhos),
        "ticket_medio":       arred(safe_div(real_bruto, len(ganhos))) if ganhos else 0.0,
    }

    # ── Quebra por closer (dono do deal) ──────────────────────
    por_closer = {}
    for d in ganhos:
        nome = owner_name(d, users) or "— sem dono —"
        acc = por_closer.setdefault(nome, {
            "nome": nome, "bruto": 0.0, "multi": 0.0, "qtd": 0,
            "aberto_hoje": 0.0, "previsto_hoje": 0.0})
        acc["bruto"] += float(d.get("value") or 0)
        acc["multi"] += float(cf(d, CF_MULTIPLICADOR) or 0)
        acc["qtd"] += 1

    for d in abertos:
        if str(d.get("expected_close_date") or "")[:10] != hoje_str:
            continue
        nome = owner_name(d, users) or "— sem dono —"
        acc = por_closer.setdefault(nome, {
            "nome": nome, "bruto": 0.0, "multi": 0.0, "qtd": 0,
            "aberto_hoje": 0.0, "previsto_hoje": 0.0})
        v = float(d.get("value") or 0)
        pr = d.get("probability")
        acc["aberto_hoje"] += v
        if pr in (20, 50, 70):
            acc["previsto_hoje"] += v * (pr / 100.0)

    metas_por_nome = {norm(c["nome"]): c["meta"] for c in closers_meta}
    closers = []
    for nome, v in por_closer.items():
        meta_ind = metas_por_nome.get(norm(nome), 0.0)
        closers.append({
            "nome": nome,
            "meta": arred(meta_ind),
            "real_bruto": arred(v["bruto"]),
            "real_multi": arred(v["multi"]),
            "pct": arred(safe_div(v["multi"], meta_ind) * 100) if meta_ind else None,
            "qtd": v["qtd"],
            "ticket": arred(safe_div(v["bruto"], v["qtd"])) if v["qtd"] else 0.0,
            "aberto_hoje": arred(v["aberto_hoje"]),
            "previsto_hoje": arred(v["previsto_hoje"]),
        })
    lista_closers = _lista_norm(CLOSERS_LISTA)
    if lista_closers:
        closers = [c for c in closers if norm(c["nome"]) in lista_closers]
    closers.sort(key=lambda x: -x["real_multi"])

    # ── SDRs (régua do painel gerente_comercial) ──────────────
    sdrs, erro_sdr = [], None
    try:
        sdrs = calcular_sdrs(mes, ano, metas, ganhos, users, du_total, du_pass)
    except Exception as e:
        erro_sdr = f"{type(e).__name__}: {e}"

    # ── Reuniões de hoje e do último DU ───────────────────────
    reunioes, erro_reu = {"contagem": {}, "detalhe": {}}, None
    dias_mes = []
    try:
        import calendar as _cal
        ultimo_dia = _cal.monthrange(ano, mes)[1]
        dias_mes = [date(ano, mes, i).strftime("%Y-%m-%d") for i in range(1, ultimo_dia + 1)]
        # o mês inteiro: os totais do mês, da semana e dos cards saem de uma chamada só
        dias_reu = sorted(set(dias_mes) | {ontem_str, hoje_str} | set(dias_semana))
        escopo = {pid} | ({PIPELINE_MGM} if PIPELINE_MGM else set())
        reunioes = contar_reunioes(mes, ano, users, dias_reu,
                                   deals_extra=ganhos + abertos_todos,
                                   pipelines_escopo=escopo)
    except Exception as e:
        erro_reu = f"{type(e).__name__}: {e}"

    cont_reu = reunioes.get("contagem") or {}
    reu_mes = {
        "agendadas":  sum((cont_reu.get(d) or {}).get("agendadas", 0)  for d in dias_mes),
        "realizadas": sum((cont_reu.get(d) or {}).get("realizadas", 0) for d in dias_mes),
    }

    # ── Pendências do pipe + alertas ──────────────────────────
    pend = montar_pendencias(abertos, users, so_do_dono=PENDENCIAS_DONO)

    alertas = []
    if not closers_meta and not META_FIXA:
        alertas.append(
            "Nenhum closer com Subarea = " + "/".join(sorted(subareas))
            + " e meta financeira no METAS — a Meta Mês veio zerada.")
    if not filtrou_etapa and ETAPA_PREVISAO:
        alertas.append(f"Nenhuma etapa com \"{ETAPA_PREVISAO}\" no nome foi encontrada — "
                       "a previsão está considerando todas as etapas.")
    if erro_mgm:
        alertas.append("Não consegui ler o funil MGM agora, os números estão só com o "
                       "Navigator — " + erro_mgm)
    if erro_sdr:
        alertas.append("Não consegui calcular as métricas de SDR agora — " + erro_sdr)
    if erro_reu:
        alertas.append("Não consegui contar as reuniões agora — " + erro_reu)
    for chave in ("sem_data", "sem_probabilidade", "prob_fora"):
        r = pend["resumo"][chave]
        if r["qtd"]:
            valor_fmt = f"{r['valor']:,.0f}".replace(",", ".")
            alertas.append(f"{r['qtd']} deal(s) · R$ {valor_fmt} — {r['label']}.")

    return {
        "funil": pipe["nome"],
        "rotulo": ROTULO_FUNIL,
        "pipeline_id": pid,
        "consolidado": True,
        "composicao": composicao,
        "periodo": {
            "mes": mes, "ano": ano,
            "du_total": du_total, "du_passados": du_pass, "du_restantes": du_rest,
            "hoje": hoje_str, "ontem": ontem_str,
            "atualizado_em": agora_br().strftime("%d/%m/%Y %H:%M"),
        },
        "metricas": metricas,
        "closers": closers,
        "sdrs": sdrs,
        "reunioes": {
            "hoje":  (reunioes.get("contagem") or {}).get(hoje_str,  {"agendadas": 0, "realizadas": 0, "duplicadas_ocultas": 0}),
            "ontem": (reunioes.get("contagem") or {}).get(ontem_str, {"agendadas": 0, "realizadas": 0, "duplicadas_ocultas": 0}),
            "mes": reu_mes,
            "conversao": reunioes.get("conversao") or {
                "deals_com_reuniao": 0, "deals_ganhos": 0, "taxa": 0.0},
            "lista_hoje":  (reunioes.get("detalhe") or {}).get(hoje_str, []),
            "lista_ontem": (reunioes.get("detalhe") or {}).get(ontem_str, []),
            "excluidos": EXCLUIR_REU,
            "dono": DONO_REUNIAO,
            "fora_escopo": (reunioes.get("fora_escopo") or 0),
            "fora_dono": (reunioes.get("fora_dono") or 0),
        },
        "semana": {
            "ini": dom.strftime("%Y-%m-%d"),
            "fim": sab.strftime("%Y-%m-%d"),
            "p20": b_semana["p20"], "p50": b_semana["p50"], "p70": b_semana["p70"],
            "previsto": b_semana["previsto"], "em_aberto": b_semana["em_aberto"],
            "qtd_abertos": b_semana["qtd"],
            "entrou_bruto": arred(g_semana["bruto"]),
            "entrou_multi": arred(g_semana["multi"]),
            "qtd_vendas": g_semana["qtd"],
            "reunioes_agendadas": sum(
                ((reunioes.get("contagem") or {}).get(d) or {}).get("agendadas", 0)
                for d in dias_semana),
            "reunioes_validadas": sum(
                ((reunioes.get("contagem") or {}).get(d) or {}).get("realizadas", 0)
                for d in dias_semana),
        },
        "detalhe_hoje": b_hoje,
        "detalhe_ontem": b_ontem,
        "meta_composicao": sorted(closers_meta, key=lambda x: -x["meta"]),
        "pendencias": pend,
        "alertas": alertas,
    }



# ── RESUMO DAS TRÊS FRENTES ───────────────────────────────────
def calcular_resumo(mes=None, ano=None):
    """
    Consolida as três frentes da equipe Ascensão/MGM:
      navigator  → funil Navigator, sem a tag de Renovação
      mgm        → funil MGM (PIPELINE_MGM)
      renovacao  → funil Navigator, com a tag de Renovação
    """
    hoje = hoje_br()
    mes = mes or hoje.month
    ano = ano or hoje.year
    hoje_str = hoje.strftime("%Y-%m-%d")

    feriados = buscar_feriados()
    metas_sheet = buscar_metas(ano, mes)

    du_calc = du_mes_total(ano, mes, feriados)
    du_sheet = next((m["dias_uteis"] for m in metas_sheet if m["dias_uteis"] > 0), 0)
    du_total = du_sheet if du_sheet > 0 else du_calc

    mes_fechado = (ano < hoje.year) or (ano == hoje.year and mes < hoje.month)
    mes_futuro = (ano > hoje.year) or (ano == hoje.year and mes > hoje.month)
    if mes_fechado:
        du_pass = du_total
    elif mes_futuro:
        du_pass = 0
    else:
        du_pass = min(du_passados(ano, mes, feriados), du_total)
    du_rest = max(du_total - du_pass, 0)

    pipe = buscar_pipeline_id()
    pid_nav = pipe["id"]
    if not pid_nav:
        return {"erro": f"Funil '{FUNIL_NOME}' não encontrado.",
                "funis_disponiveis": pipe["todos"]}

    users = buscar_users()
    ganhos_nav = buscar_ganhos(mes, ano, pid_nav)
    abertos_nav = buscar_abertos(pid_nav)

    erro_mgm = None
    ganhos_mgm, abertos_mgm = [], []
    try:
        ganhos_mgm = buscar_ganhos(mes, ano, PIPELINE_MGM)
        abertos_mgm = buscar_abertos(PIPELINE_MGM)
    except Exception as e:
        erro_mgm = f"{type(e).__name__}: {e}"

    tag_ativa = bool(TAG_RENOVACAO_CAMPO and TAG_RENOVACAO_VALOR)
    # o pipe considerado é só o da etapa de Negociação
    abertos_nav_neg, filtrou_etapa = so_em_negociacao(abertos_nav)
    abertos_mgm_neg, _ = so_em_negociacao(abertos_mgm)

    grupos = {
        "navigator": {
            "ganhos":  [d for d in ganhos_nav if not eh_renovacao(d)],
            "abertos": [d for d in abertos_nav_neg if not eh_renovacao(d)],
        },
        "mgm": {"ganhos": ganhos_mgm, "abertos": abertos_mgm_neg},
        "renovacao": {
            "ganhos":  [d for d in ganhos_nav if eh_renovacao(d)],
            "abertos": [d for d in abertos_nav_neg if eh_renovacao(d)],
        },
    }

    metas_mes = METAS_FRENTES.get((ano, mes), {})
    frentes = []
    for chave in ("navigator", "mgm", "renovacao"):
        g = grupos[chave]["ganhos"]
        a = grupos[chave]["abertos"]
        meta = float(metas_mes.get(chave, 0))
        bruto = sum(float(d.get("value") or 0) for d in g)
        multi = sum(float(cf(d, CF_MULTIPLICADOR) or 0) for d in g)
        b = bucket_dia(a, hoje_str)
        em_aberto_total = sum(float(d.get("value") or 0) for d in a)
        gap = meta - multi
        frentes.append({
            "chave": chave,
            "nome": NOMES_FRENTES[chave],
            "meta": arred(meta),
            "meta_dia": arred(safe_div(meta, du_total)),
            "deveria_mtd": arred(safe_div(meta, du_total) * du_pass),
            "real_bruto": arred(bruto),
            "real_multi": arred(multi),
            "pct": arred(safe_div(multi, meta) * 100) if meta else None,
            "gap": arred(gap),
            "meta_dia_rest": arred(safe_div(gap, du_rest)) if du_rest and meta else 0.0,
            "qtd": len(g),
            "ticket": arred(safe_div(bruto, len(g))) if g else 0.0,
            "pipe_aberto": arred(em_aberto_total),
            "previsto_hoje": b["previsto"],
            "em_aberto_hoje": b["em_aberto"],
            "qtd_abertos": len(a),
        })

    def soma(campo):
        return arred(sum(f[campo] or 0 for f in frentes))

    total_meta = soma("meta")
    total_multi = soma("real_multi")
    total = {
        "nome": "TOTAL",
        "meta": total_meta,
        "meta_dia": arred(safe_div(total_meta, du_total)),
        "deveria_mtd": arred(safe_div(total_meta, du_total) * du_pass),
        "real_bruto": soma("real_bruto"),
        "real_multi": total_multi,
        "pct": arred(safe_div(total_multi, total_meta) * 100) if total_meta else None,
        "gap": arred(total_meta - total_multi),
        "meta_dia_rest": arred(safe_div(total_meta - total_multi, du_rest)) if du_rest and total_meta else 0.0,
        "qtd": sum(f["qtd"] for f in frentes),
        "ticket": arred(safe_div(soma("real_bruto"), sum(f["qtd"] for f in frentes))) if sum(f["qtd"] for f in frentes) else 0.0,
        "pipe_aberto": soma("pipe_aberto"),
        "previsto_hoje": soma("previsto_hoje"),
        "em_aberto_hoje": soma("em_aberto_hoje"),
        "qtd_abertos": sum(f["qtd_abertos"] for f in frentes),
    }

    # projeção Set–Dez, para contexto
    projecao = []
    for (a_, m_), vals in sorted(METAS_FRENTES.items()):
        projecao.append({
            "ano": a_, "mes": m_,
            "rotulo": ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"][m_-1] + f"/{str(a_)[2:]}",
            "navigator": vals["navigator"], "mgm": vals["mgm"], "renovacao": vals["renovacao"],
            "total": vals["navigator"] + vals["mgm"] + vals["renovacao"],
            "atual": (a_ == ano and m_ == mes),
        })

    meta_planilha = sum(
        m["meta_fin"] for m in metas_sheet
        if m["meta_reu"] == 0 and m["meta_fin"] > 0 and m["nome_norm"] not in EXCLUIR_PESSOAS
        and norm(next((c["subarea"] for c in buscar_colaboradores(mes, ano)
                       if c["nome_norm"] == m["nome_norm"]), "")) in _subareas())

    alertas = []
    if not filtrou_etapa and ETAPA_PREVISAO:
        alertas.append(f"Nenhuma etapa com \"{ETAPA_PREVISAO}\" no nome foi encontrada — "
                       "o pipe está considerando todas as etapas.")
    if not tag_ativa:
        alertas.append("A tag de Renovação ainda não foi configurada — a frente aparece zerada "
                       "e o Navigator está vindo inteiro. Preencha TAG_RENOVACAO_CAMPO e "
                       "TAG_RENOVACAO_VALOR nas variáveis de ambiente.")
    if not metas_mes:
        alertas.append(f"Não tenho metas por frente cadastradas para {mes:02d}/{ano} — "
                       "as metas do resumo vieram zeradas.")
    if erro_mgm:
        alertas.append("Não consegui ler o funil MGM agora — " + erro_mgm)
    if metas_mes and meta_planilha and abs(meta_planilha - total_meta) > 1:
        alertas.append(f"As metas por frente somam R$ {total_meta:,.0f}".replace(",", ".")
                       + f", mas a planilha METAS traz R$ {meta_planilha:,.0f}".replace(",", ".")
                       + " para os closers da Ascensão. Vale conferir qual das duas manda.")

    return {
        "periodo": {
            "mes": mes, "ano": ano,
            "du_total": du_total, "du_passados": du_pass, "du_restantes": du_rest,
            "hoje": hoje_str,
            "atualizado_em": agora_br().strftime("%d/%m/%Y %H:%M"),
        },
        "frentes": frentes,
        "total": total,
        "projecao": projecao,
        "meta_planilha": arred(meta_planilha),
        "tag_renovacao_ativa": tag_ativa,
        "etapa_previsao": ETAPA_PREVISAO if filtrou_etapa else None,
        "alertas": alertas,
    }



# ── ABA GRÁFICOS ──────────────────────────────────────────────
def calcular_graficos(mes=None, ano=None):
    """
    Séries diárias para a aba Gráficos:
      · reuniões por dia, separadas por frente
      · vendas brutas por dia, separadas por frente
      · jacaré: meta acumulada (linear por dia útil) x realizado acumulado
      · atingimento por frente

    A frente de uma REUNIÃO vem do negócio vinculado: funil MGM -> mgm;
    funil Navigator -> renovacao se tiver a tag, senão navigator.
    """
    import calendar as cal_mod

    hoje = hoje_br()
    mes = mes or hoje.month
    ano = ano or hoje.year
    hoje_str = hoje.strftime("%Y-%m-%d")

    feriados = buscar_feriados()
    metas_sheet = buscar_metas(ano, mes)
    colab = buscar_colaboradores(mes, ano)
    subareas = _subareas()

    # meta consolidada do mês (mesma régua do Painel do Mês)
    nome_to_sub = {c["nome_norm"]: norm(c["subarea"]) for c in colab}
    meta_mes = 0.0
    for m in metas_sheet:
        nn = m["nome_norm"]
        if not nn or nn in EXCLUIR_PESSOAS:
            continue
        if not (m["meta_reu"] == 0 and m["meta_fin"] > 0):
            continue
        if nome_to_sub.get(nn) not in subareas:
            continue
        meta_mes += m["meta_fin"]
    if META_FIXA:
        meta_mes = to_num_br(META_FIXA)

    du_calc = du_mes_total(ano, mes, feriados)
    du_sheet = next((m["dias_uteis"] for m in metas_sheet if m["dias_uteis"] > 0), 0)
    du_total = du_sheet if du_sheet > 0 else du_calc

    pipe = buscar_pipeline_id()
    pid = pipe["id"]
    if not pid:
        return {"erro": f"Funil '{FUNIL_NOME}' não encontrado.",
                "funis_disponiveis": pipe["todos"]}

    users = buscar_users()
    ganhos_nav = buscar_ganhos(mes, ano, pid)
    abertos_nav = buscar_abertos(pid)
    ganhos_mgm, abertos_mgm = [], []
    erro_mgm = None
    if PIPELINE_MGM:
        try:
            ganhos_mgm = buscar_ganhos(mes, ano, PIPELINE_MGM)
            abertos_mgm = buscar_abertos(PIPELINE_MGM)
        except Exception as e:
            erro_mgm = f"{type(e).__name__}: {e}"

    # negócio -> frente
    frente_por_deal = {}
    for d in ganhos_nav + abertos_nav:
        frente_por_deal[d.get("id")] = "renovacao" if eh_renovacao(d) else "navigator"
    for d in ganhos_mgm + abertos_mgm:
        frente_por_deal[d.get("id")] = "mgm"

    ultimo = cal_mod.monthrange(ano, mes)[1]
    dias = [date(ano, mes, d).strftime("%Y-%m-%d") for d in range(1, ultimo + 1)]
    chaves = ("navigator", "mgm", "renovacao")

    # ── vendas brutas por dia, por frente ─────────────────────
    vendas = {k: {d: 0.0 for d in dias} for k in chaves}
    vendas_multi_dia = {d: 0.0 for d in dias}
    for d in ganhos_nav + ganhos_mgm:
        dia = won_time_br(d)[:10]
        if dia not in vendas["navigator"]:
            continue
        f = frente_por_deal.get(d.get("id"), "navigator")
        vendas[f][dia] += float(d.get("value") or 0)
        vendas_multi_dia[dia] += float(cf(d, CF_MULTIPLICADOR) or 0)

    # ── reuniões por dia, por frente ──────────────────────────
    # agendadas = tudo que estava na agenda; validadas = o subconjunto que passou na régua
    agendadas = {k: {d: 0 for d in dias} for k in chaves}
    reunioes  = {k: {d: 0 for d in dias} for k in chaves}
    titulos   = {d: [] for d in dias}          # para o tooltip do gráfico
    com_reu   = {k: set() for k in chaves}     # negócios distintos com reunião validada
    ganhou    = {k: set() for k in chaves}     # ... e que estão Ganho
    fora_escopo = fora_dono = 0
    erro_reu = None
    try:
        escopo = {pid} | ({PIPELINE_MGM} if PIPELINE_MGM else set())
        r = contar_reunioes(mes, ano, users, dias,
                            deals_extra=ganhos_nav + ganhos_mgm + abertos_nav + abertos_mgm,
                            pipelines_escopo=escopo)
        for dia, itens in (r.get("detalhe") or {}).items():
            if dia not in agendadas["navigator"]:
                continue
            for it in itens:
                f = frente_por_deal.get(it.get("deal_id"))
                if not f:
                    f = "mgm" if (it.get("funil") or "").strip().upper() == "MGM" else "navigator"
                agendadas[f][dia] += 1
                if it["status"] == "realizada":
                    reunioes[f][dia] += 1
                    com_reu[f].add(it.get("deal_id"))
                    if it.get("ganho"):
                        ganhou[f].add(it.get("deal_id"))
                nome_resp = (it.get("responsavel") or "").split(" ")[0]
                titulos[dia].append({
                    "t": (it.get("deal") or it.get("assunto") or "(sem título)")[:60],
                    "r": nome_resp,
                    "f": f,
                    "s": it["status"],
                    "m": it.get("motivo"),
                    "h": it.get("hora"),
                })
        fora_escopo = r.get("fora_escopo", 0)
        fora_dono = r.get("fora_dono", 0)
        for d in titulos:
            titulos[d].sort(key=lambda x: (x["h"] or "99:99"))
    except Exception as e:
        erro_reu = f"{type(e).__name__}: {e}"

    # ── jacaré: meta acumulada x realizado acumulado ──────────
    meta_dia = safe_div(meta_mes, du_total)
    jacare, acum_du, acum_real, acum_bruto = [], 0, 0.0, 0.0
    for dia in dias:
        dt = date(ano, mes, int(dia[8:10]))
        if dt.weekday() < 5 and dt not in feriados:
            acum_du += 1
        acum_real += vendas_multi_dia[dia]
        acum_bruto += sum(vendas[k][dia] for k in chaves)
        futuro = dia > hoje_str
        jacare.append({
            "dia": dia,
            "meta_mtd": arred(min(meta_dia * acum_du, meta_mes)),
            "real_mtd": None if futuro else arred(acum_real),
            "bruto_mtd": None if futuro else arred(acum_bruto),
            "e_du": dt.weekday() < 5 and dt not in feriados,
        })

    # ── atingimento por frente ────────────────────────────────
    metas_frentes = METAS_FRENTES.get((ano, mes), {})
    atingimento = []
    for k in chaves:
        meta_f = float(metas_frentes.get(k, 0))
        bruto = sum(vendas[k].values())
        deals_f = [d for d in ganhos_nav + ganhos_mgm
                   if frente_por_deal.get(d.get("id")) == k]
        multi = sum(float(cf(d, CF_MULTIPLICADOR) or 0) for d in deals_f)
        atingimento.append({
            "chave": k, "nome": NOMES_FRENTES[k],
            "meta": arred(meta_f),
            "bruto": arred(bruto), "multi": arred(multi),
            "pct": arred(safe_div(multi, meta_f) * 100) if meta_f else None,
            "vendas": len(deals_f),
            "reunioes": sum(reunioes[k].values()),
            "deals_com_reuniao": len(com_reu[k]),
            "deals_ganhos": len(ganhou[k]),
            "conversao": (arred(safe_div(len(ganhou[k]), len(com_reu[k])) * 100)
                          if com_reu[k] else None),
        })

    alertas = []
    if erro_mgm:
        alertas.append("Não consegui ler o funil MGM agora — " + erro_mgm)
    if erro_reu:
        alertas.append("Não consegui montar a série de reuniões — " + erro_reu)

    return {
        "periodo": {"mes": mes, "ano": ano, "hoje": hoje_str,
                    "du_total": du_total, "meta_mes": arred(meta_mes),
                    "atualizado_em": agora_br().strftime("%d/%m/%Y %H:%M")},
        "dias": dias,
        "frentes": [{"chave": k, "nome": NOMES_FRENTES[k]} for k in chaves],
        "reunioes": {k: [reunioes[k][d] for d in dias] for k in chaves},
        "conversao_mes": {
            "deals_com_reuniao": sum(len(com_reu[k]) for k in chaves),
            "deals_ganhos": sum(len(ganhou[k]) for k in chaves),
            "taxa": (arred(safe_div(sum(len(ganhou[k]) for k in chaves),
                                    sum(len(com_reu[k]) for k in chaves)) * 100)
                     if any(com_reu[k] for k in chaves) else None),
        },
        "titulos": {d: titulos[d] for d in dias if titulos[d]},
        "regra_reuniao": {"dono": DONO_REUNIAO, "fora_escopo": fora_escopo,
                          "fora_dono": fora_dono},
        "vendas":   {k: [arred(vendas[k][d]) for d in dias] for k in chaves},
        "jacare": jacare,
        "atingimento": atingimento,
        "alertas": alertas,
    }


# ── ABA FORECAST ──────────────────────────────────────────────
def calcular_forecast(mes=None, ano=None):
    """
    Forecast dia a dia do mês: para cada dia, os negócios ABERTOS cuja data prevista
    de fechamento cai naquele dia, ponderados pela probabilidade (20/50/70), com a
    lista dos negócios por trás de cada dia e link direto para o Pipedrive.

    Mesma régua do resto do painel: só a etapa de Negociação, Navigator + MGM.
    Dias passados também mostram o que de fato entrou (ganho) naquele dia, para
    dar para comparar o que estava previsto com o que aconteceu.
    """
    import calendar as cal_mod

    hoje = hoje_br()
    mes = mes or hoje.month
    ano = ano or hoje.year
    hoje_str = hoje.strftime("%Y-%m-%d")

    feriados = buscar_feriados()
    metas_sheet = buscar_metas(ano, mes)
    colab = buscar_colaboradores(mes, ano)
    subareas = _subareas()

    nome_to_sub = {c["nome_norm"]: norm(c["subarea"]) for c in colab}
    meta_mes = 0.0
    for m in metas_sheet:
        nn = m["nome_norm"]
        if not nn or nn in EXCLUIR_PESSOAS:
            continue
        if not (m["meta_reu"] == 0 and m["meta_fin"] > 0):
            continue
        if nome_to_sub.get(nn) not in subareas:
            continue
        meta_mes += m["meta_fin"]
    if META_FIXA:
        meta_mes = to_num_br(META_FIXA)

    du_calc = du_mes_total(ano, mes, feriados)
    du_sheet = next((m["dias_uteis"] for m in metas_sheet if m["dias_uteis"] > 0), 0)
    du_total = du_sheet if du_sheet > 0 else du_calc

    pipe = buscar_pipeline_id()
    pid = pipe["id"]
    if not pid:
        return {"erro": f"Funil '{FUNIL_NOME}' não encontrado.",
                "funis_disponiveis": pipe["todos"]}

    users = buscar_users()
    ganhos_nav = buscar_ganhos(mes, ano, pid)
    abertos_nav = buscar_abertos(pid)
    ganhos_mgm, abertos_mgm, erro_mgm = [], [], None
    if PIPELINE_MGM:
        try:
            ganhos_mgm = buscar_ganhos(mes, ano, PIPELINE_MGM)
            abertos_mgm = buscar_abertos(PIPELINE_MGM)
        except Exception as e:
            erro_mgm = f"{type(e).__name__}: {e}"

    ganhos = ganhos_nav + ganhos_mgm
    abertos_todos = abertos_nav + abertos_mgm
    abertos, filtrou_etapa = so_em_negociacao(abertos_todos)

    frente_por_deal = {}
    for d in ganhos_nav + abertos_nav:
        frente_por_deal[d.get("id")] = "renovacao" if eh_renovacao(d) else "navigator"
    for d in ganhos_mgm + abertos_mgm:
        frente_por_deal[d.get("id")] = "mgm"

    ultimo = cal_mod.monthrange(ano, mes)[1]
    dias = [date(ano, mes, i).strftime("%Y-%m-%d") for i in range(1, ultimo + 1)]

    # ── o que já entrou, por dia ──────────────────────────────
    ganho_dia = {d: {"bruto": 0.0, "multi": 0.0, "qtd": 0} for d in dias}
    for d in ganhos:
        dia = won_time_br(d)[:10]
        if dia not in ganho_dia:
            continue
        ganho_dia[dia]["bruto"] += float(d.get("value") or 0)
        ganho_dia[dia]["multi"] += float(cf(d, CF_MULTIPLICADOR) or 0)
        ganho_dia[dia]["qtd"] += 1

    # ── o pipe aberto, por dia previsto ───────────────────────
    PESO = {20: 0.20, 50: 0.50, 70: 0.70}
    por_dia = {d: {"p20": 0.0, "p50": 0.0, "p70": 0.0, "outros": 0.0,
                   "sem_prob": 0.0, "itens": []} for d in dias}
    fora_do_mes = {"qtd": 0, "valor": 0.0, "previsto": 0.0, "itens": []}
    sem_data = {"qtd": 0, "valor": 0.0, "itens": []}

    def _item(d):
        pr = d.get("probability")
        v = float(d.get("value") or 0)
        did = d.get("id")
        return {
            "id": did,
            "titulo": d.get("title") or "(sem título)",
            "dono": owner_name(d, users) or "— sem dono —",
            "valor": arred(v),
            "probabilidade": pr,
            "ponderado": arred(v * PESO.get(pr, 0.0)),
            "frente": frente_por_deal.get(did) or "navigator",
            "prev": str(d.get("expected_close_date") or "")[:10] or None,
            "url": f"https://boardacademy.pipedrive.com/deal/{did}",
        }

    for d in abertos:
        dt = str(d.get("expected_close_date") or "")[:10]
        it = _item(d)
        v = float(d.get("value") or 0)
        if not dt:
            sem_data["qtd"] += 1
            sem_data["valor"] += v
            sem_data["itens"].append(it)
            continue
        if dt not in por_dia:
            fora_do_mes["qtd"] += 1
            fora_do_mes["valor"] += v
            fora_do_mes["previsto"] += it["ponderado"]
            fora_do_mes["itens"].append(it)
            continue
        pr = d.get("probability")
        alvo = por_dia[dt]
        if pr in PESO:
            alvo["p%d" % pr] += v
        elif pr is None:
            alvo["sem_prob"] += v
        else:
            alvo["outros"] += v
        alvo["itens"].append(it)

    linhas = []
    for dia in dias:
        a = por_dia[dia]
        dt = date(ano, mes, int(dia[8:10]))
        previsto = a["p20"] * 0.20 + a["p50"] * 0.50 + a["p70"] * 0.70
        em_aberto = a["p20"] + a["p50"] + a["p70"] + a["outros"] + a["sem_prob"]
        g = ganho_dia[dia]
        a["itens"].sort(key=lambda x: -x["ponderado"])
        linhas.append({
            "dia": dia,
            "e_du": dt.weekday() < 5 and dt not in feriados,
            "passado": dia < hoje_str,
            "hoje": dia == hoje_str,
            "p20": arred(a["p20"]), "p50": arred(a["p50"]), "p70": arred(a["p70"]),
            "sem_probabilidade": arred(a["sem_prob"]),
            "fora_bucket": arred(a["outros"]),
            "previsto": arred(previsto),
            "em_aberto": arred(em_aberto),
            "qtd": len(a["itens"]),
            "ganho_bruto": arred(g["bruto"]),
            "ganho_multi": arred(g["multi"]),
            "ganho_qtd": g["qtd"],
            "itens": a["itens"],
        })

    real_multi = sum(float(cf(d, CF_MULTIPLICADOR) or 0) for d in ganhos)
    real_bruto = sum(float(d.get("value") or 0) for d in ganhos)
    fut = [l for l in linhas if not l["passado"]]
    prev_restante = sum(l["previsto"] for l in fut)
    aberto_restante = sum(l["em_aberto"] for l in fut)

    sem_data["itens"].sort(key=lambda x: -x["valor"])
    fora_do_mes["itens"].sort(key=lambda x: -x["ponderado"])

    total = {
        "meta_mes": arred(meta_mes),
        "real_bruto": arred(real_bruto),
        "real_multi": arred(real_multi),
        "previsto_restante": arred(prev_restante),
        "aberto_restante": arred(aberto_restante),
        "projecao": arred(real_multi + prev_restante),
        "pct_projecao": arred(safe_div(real_multi + prev_restante, meta_mes) * 100) if meta_mes else None,
        "gap_projecao": arred(meta_mes - (real_multi + prev_restante)),
        "previsto_mes_todo": arred(sum(l["previsto"] for l in linhas)),
        "qtd_abertos_mes": sum(l["qtd"] for l in linhas),
    }

    alertas = []
    if erro_mgm:
        alertas.append("Não consegui ler o funil MGM agora — " + erro_mgm)
    if not filtrou_etapa and ETAPA_PREVISAO:
        alertas.append(f"Nenhuma etapa com \"{ETAPA_PREVISAO}\" no nome foi encontrada — "
                       "o forecast está considerando todas as etapas.")
    if sem_data["qtd"]:
        alertas.append(f"{sem_data['qtd']} negócio(s) · R$ {sem_data['valor']:,.0f}".replace(",", ".")
                       + " sem data prevista — ficam fora de todos os dias do forecast.")

    return {
        "periodo": {"mes": mes, "ano": ano, "hoje": hoje_str,
                    "du_total": du_total, "meta_mes": arred(meta_mes),
                    "etapa": ETAPA_PREVISAO if filtrou_etapa else None,
                    "atualizado_em": agora_br().strftime("%d/%m/%Y %H:%M")},
        "linhas": linhas,
        "total": total,
        "frentes": [{"chave": k, "nome": NOMES_FRENTES[k]} for k in
                    ("navigator", "mgm", "renovacao")],
        "sem_data": {"qtd": sem_data["qtd"], "valor": arred(sem_data["valor"]),
                     "itens": sem_data["itens"]},
        "fora_do_mes": {"qtd": fora_do_mes["qtd"], "valor": arred(fora_do_mes["valor"]),
                        "previsto": arred(fora_do_mes["previsto"]),
                        "itens": fora_do_mes["itens"]},
        "alertas": alertas,
    }


# ── AUTH OPCIONAL ─────────────────────────────────────────────
def autorizado():
    if not PAINEL_SENHA:
        return True
    k = request.args.get("k", "") or request.headers.get("X-Painel-Key", "")
    return k == PAINEL_SENHA


# ── ROTAS ─────────────────────────────────────────────────────
def _handler(fn=None):
    if not API_KEY:
        return jsonify({"erro": "PIPE_API_KEY não configurada nas variáveis de ambiente."}), 500
    if not autorizado():
        return jsonify({"erro": "Não autorizado."}), 401
    try:
        mes = request.args.get("mes", type=int)
        ano = request.args.get("ano", type=int)
        data = (fn or calcular_navigator)(mes=mes, ano=ano)
        status = 400 if data.get("erro") else 200
        return jsonify(limpar_nans(data)), status
    except req.HTTPError as e:
        return jsonify({"erro": f"Pipedrive respondeu {e.response.status_code}",
                        "detalhe": e.response.text[:400]}), 502
    except Exception as e:
        import traceback
        return jsonify({"erro": str(e), "trace": traceback.format_exc()}), 500


@app.route("/", defaults={"p": ""}, methods=["GET"])
@app.route("/<path:p>", methods=["GET"])
def roteador(p):
    """
    A Vercel entrega TUDO para esta função (o path que o Flask vê pode ser /, /api/index
    ou o path original — varia). Por isso a decisão é pelo parâmetro ?fmt=, que nenhuma
    reescrita altera, e o HTML mora aqui dentro em vez de depender de arquivo estático.

        /                                → o painel (HTML)
        /api/index?fmt=json&mes=&ano=    → dados do Painel do Mês (JSON)
        /api/index?fmt=resumo            → as 3 frentes (JSON)
        /api/index?fmt=graficos          → séries dos gráficos (JSON)
        /api/index?fmt=forecast          → forecast dia a dia (JSON)
        /api/index?fmt=health            → diagnóstico (JSON)
    """
    modo = (request.args.get("fmt") or "").strip().lower()
    alvo = (p or "").strip("/").lower()

    if modo == "health" or alvo.endswith("health"):
        return health_payload()
    if modo == "resumo" or alvo.endswith("resumo"):
        return _handler(calcular_resumo)
    if modo == "graficos" or alvo.endswith("graficos"):
        return _handler(calcular_graficos)
    if modo == "forecast" or alvo.endswith("forecast"):
        return _handler(calcular_forecast)
    if modo == "json" or alvo.endswith("navigator"):
        return _handler()
    return Response(PAGINA_HTML, mimetype="text/html; charset=utf-8")


def health_payload():
    """Diagnóstico. Nunca estoura: se o Pipedrive falhar, devolve o motivo em JSON."""
    pipe, erro_pipe = {"id": None, "nome": None, "todos": []}, None
    if API_KEY:
        try:
            pipe = buscar_pipeline_id()
        except req.HTTPError as e:
            erro_pipe = (f"Pipedrive respondeu {e.response.status_code} — "
                         "token inválido ou sem permissão.")
        except Exception as e:
            erro_pipe = f"{type(e).__name__}: {e}"

    return jsonify({
        "ok": bool(API_KEY) and bool(pipe.get("id")),
        "token_configurado": bool(API_KEY),
        "erro_pipedrive": erro_pipe,
        "funil_procurado": FUNIL_NOME,
        "funil_encontrado": pipe.get("nome"),
        "pipeline_id": pipe.get("id"),
        "funis_disponiveis": pipe.get("todos"),
        "senha_ativa": bool(PAINEL_SENHA),
        "hoje_br": hoje_br().isoformat(),
        "path_recebido": request.path,
    })


# ── PÁGINA (HTML embutido: a Vercel não serve estático neste projeto) ──────
# Para mexer no visual do painel, edite daqui para baixo.
PAGINA_HTML = r"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Ascensão/MGM — Board Academy</title>
<style>
  *,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
  :root{
    --gold:#B8860B; --gold-bg:#FFF9E6;
    --navy:#1A1A2E; --bg:#F4F5F7; --white:#fff;
    --text:#1E1E2E; --muted:#6B7280; --border:#E2E4E9;
    --green:#0D7A3E; --green-bg:#F2F9F5; --green-hd:#EAF4EE;
    --red:#B91C1C; --red-bg:#FEF2F2;
    --amber:#92400E; --amber-bg:#FFFBEB;
    --blue-bg:#F4F7FE;
    --shadow:0 1px 3px rgba(0,0,0,.08);
  }
  body{background:var(--bg);color:var(--text);font-size:13px;
       font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif}

  /* HEADER */
  .header{background:var(--navy);height:56px;padding:0 24px;display:flex;
          align-items:center;justify-content:space-between;position:sticky;top:0;z-index:10;
          box-shadow:0 4px 6px rgba(0,0,0,.07);flex-wrap:wrap}
  .header-left{display:flex;align-items:center;gap:22px}
  .brand{display:flex;align-items:center;gap:10px}
  .brand-bar{width:3px;height:22px;background:var(--gold);border-radius:2px}
  .brand-text{color:#fff;font-size:14px;font-weight:700;letter-spacing:.5px}
  .brand-sub{color:var(--gold);font-size:10px;letter-spacing:1.5px;font-weight:600;text-transform:uppercase}
  .periodo-badge{background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.12);
                 border-radius:6px;padding:4px 12px;color:rgba(255,255,255,.75);font-size:12px}
  .periodo-badge strong{color:var(--gold)}
  .header-right{display:flex;align-items:center;gap:10px}
  .header-right select{background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.15);
                       border-radius:5px;color:#fff;font-size:12px;padding:5px 10px;cursor:pointer;outline:none}
  .header-right select option{background:#1A1A2E;color:#fff}
  .update-info{color:rgba(255,255,255,.45);font-size:11px}
  .btn-reload{background:transparent;border:1px solid rgba(255,255,255,.2);border-radius:5px;
              color:rgba(255,255,255,.6);font-size:11px;padding:5px 12px;cursor:pointer;transition:all .2s}
  .btn-reload:hover{border-color:var(--gold);color:var(--gold)}

  .main{padding:22px 24px;max-width:1180px;margin:0 auto}

  .block-title{font-size:11px;font-weight:700;color:var(--gold);letter-spacing:1px;
               text-transform:uppercase;margin-bottom:12px;display:flex;align-items:center;gap:8px}
  .block-title::before{content:'';width:3px;height:12px;background:var(--gold);border-radius:2px}
  .block-title .rule{flex:1;height:1px;background:var(--border)}

  .card{background:var(--white);border:1px solid var(--border);border-radius:8px;
        box-shadow:var(--shadow);overflow:hidden;margin-bottom:26px}
  .table-scroll{overflow-x:auto}
  table{width:100%;border-collapse:collapse;font-size:13px}

  /* TABELA PRINCIPAL (métrica x valor) */
  .kpi thead th{background:#F7F8FA;color:var(--muted);font-size:10px;font-weight:700;
                letter-spacing:1px;text-transform:uppercase;padding:11px 16px;text-align:left;
                border-bottom:1px solid var(--border);white-space:nowrap}
  .kpi thead th.val{text-align:right;background:var(--green-hd);color:var(--green);
                    font-size:12px;letter-spacing:1.5px;min-width:200px}
  .kpi td{padding:10px 16px;border-bottom:1px solid #EFF1F4}
  .kpi td.val{text-align:right;background:var(--green-bg);font-variant-numeric:tabular-nums;
              font-weight:600;white-space:nowrap}
  .kpi tr:last-child td{border-bottom:none}
  .kpi tr.sep td{border-top:2px solid var(--border)}
  .kpi tr.g-ontem td:first-child{background:#FAFAFB}
  .kpi tr.g-hoje  td:first-child{background:var(--blue-bg)}
  .kpi tr.g-hoje  td.val{background:#EDF3FC}
  .kpi tr.destaque td{font-weight:700}
  .kpi tr.destaque td:first-child{color:var(--navy)}
  .kpi .hint{color:var(--muted);font-size:11px;font-weight:400;margin-left:6px}

  .pct{display:inline-block;padding:2px 8px;border-radius:4px;font-size:12px;font-weight:700;min-width:56px;text-align:center}
  .pct-green{background:#ECFDF5;color:var(--green)}
  .pct-amber{background:var(--amber-bg);color:var(--amber)}
  .pct-red{background:var(--red-bg);color:var(--red)}
  .neg{color:var(--red)} .pos{color:var(--green)}
  .zero{color:var(--muted);font-weight:400}

  /* TABELA DE CLOSERS */
  .closers thead tr{background:var(--navy)}
  .closers thead th{color:rgba(255,255,255,.7);font-size:10px;font-weight:600;letter-spacing:.5px;
                    text-transform:uppercase;padding:9px 14px;text-align:right;white-space:nowrap}
  .closers thead th:first-child{text-align:left}
  .closers td{padding:10px 14px;text-align:right;border-bottom:1px solid #EFF1F4;
              font-variant-numeric:tabular-nums;white-space:nowrap}
  .closers td:first-child{text-align:left;font-weight:600}
  .closers tbody tr:hover{background:#F8F9FF}
  .closers tr.total td{background:var(--gold-bg);border-top:2px solid var(--gold);font-weight:700;color:var(--navy)}
  .closers tr.total td:first-child{color:var(--gold)}

  /* AVISOS */
  .alerta{background:var(--amber-bg);border:1px solid #F5DFA8;border-left:3px solid var(--gold);
          border-radius:6px;padding:10px 14px;font-size:12px;color:var(--amber);margin-bottom:10px}
  details.meta-comp{background:var(--white);border:1px solid var(--border);border-radius:8px;
                    padding:12px 16px;font-size:12px;box-shadow:var(--shadow)}
  details.meta-comp summary{cursor:pointer;font-weight:600;color:var(--navy);font-size:12px;outline:none}
  details.meta-comp ul{margin:10px 0 0 18px;color:var(--muted);line-height:1.8}
  .rodape{text-align:right;color:var(--muted);font-size:11px;font-style:italic;margin-top:14px}


  .graf{background:var(--white);border:1px solid var(--border);border-radius:8px;
        box-shadow:var(--shadow);padding:16px 18px 10px;margin-bottom:20px}
  .graf-tit{font-size:13px;font-weight:700;color:var(--navy);margin-bottom:2px}
  .graf-sub{font-size:11px;color:var(--muted);margin-bottom:12px}
  .graf-box{position:relative;height:320px}
  .graf-box.alto{height:360px}

  /* janelinha com o total do mês, dentro do gráfico */
  .graf-total{position:absolute;top:2px;left:4px;z-index:3;pointer-events:none;
              background:rgba(255,255,255,.93);border:1px solid var(--border);
              border-radius:6px;padding:6px 10px;box-shadow:0 1px 3px rgba(0,0,0,.06)}
  .graf-total .gt-rot{font-size:9px;font-weight:700;color:var(--muted);
                      letter-spacing:.6px;text-transform:uppercase;line-height:1.2}
  .graf-total .gt-val{font-size:17px;font-weight:800;color:var(--navy);line-height:1.15;margin-top:1px}
  .graf-total .gt-pct{font-size:11px;font-weight:700;color:var(--muted)}
  .graf-total .gt-sub{font-size:10px;color:var(--muted);line-height:1.3;margin-top:3px}
  .graf-total .gt-pos{color:#0D7A3E;font-weight:700}
  .graf-total .gt-neg{color:#B42318;font-weight:700}
  /* variante do jacaré: MTD + atingimento lado a lado */
  .graf-total .gt-cols{display:flex;gap:18px}
  .graf-total .gt-cols .gt-val{font-size:15px}
  .graf-total .gt-cols .gt-col+.gt-col{border-left:1px solid var(--border);padding-left:18px;margin-left:0}

  /* ABAS */
  .tabs{background:var(--white);border-bottom:1px solid var(--border);padding:0 24px;display:flex}
  .tab{padding:12px 18px;font-size:12px;font-weight:700;cursor:pointer;color:var(--muted);
       border-bottom:2px solid transparent;margin-bottom:-1px;transition:all .15s;letter-spacing:.3px}
  .tab:hover{color:var(--navy)}
  .tab.on{color:var(--gold);border-bottom-color:var(--gold)}
  .view{display:none}
  .view.on{display:block}

  /* RESUMO DAS FRENTES */
  .frentes thead tr{background:var(--navy)}
  .frentes thead th{color:rgba(255,255,255,.7);font-size:10px;font-weight:600;letter-spacing:.5px;
                    text-transform:uppercase;padding:10px 14px;text-align:right;white-space:nowrap}
  .frentes thead th:first-child{text-align:left}
  .frentes td{padding:11px 14px;text-align:right;border-bottom:1px solid #EFF1F4;
              font-variant-numeric:tabular-nums;white-space:nowrap}
  .frentes td:first-child{text-align:left;font-weight:700;color:var(--navy)}
  .frentes tbody tr:hover{background:#F8F9FF}
  .frentes tr.total td{background:var(--gold-bg);border-top:2px solid var(--gold);font-weight:700}
  .frentes tr.total td:first-child{color:var(--gold)}
  .frentes tr.zerada td:first-child{color:var(--muted)}
  .frentes tr.zerada td{color:var(--muted)}
  .cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:12px;margin-bottom:22px}
  .kcard{background:var(--white);border:1px solid var(--border);border-left:3px solid var(--gold);
         border-radius:8px;padding:14px 16px;box-shadow:var(--shadow)}
  .kcard .rot{font-size:10px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.8px}
  .kcard .val{font-size:22px;font-weight:700;color:var(--navy);margin-top:6px;font-variant-numeric:tabular-nums}
  .kcard .sub{font-size:11px;color:var(--muted);margin-top:3px}
  .kcard.hoje{border-left-color:#1E3A8A;background:var(--blue-bg)}
  .kcard.ontem{border-left-color:var(--muted)}

  .btn-det{float:right;background:transparent;border:1px solid var(--border);border-radius:4px;
           color:var(--muted);font-size:9px;font-weight:700;padding:2px 7px;cursor:pointer;
           text-transform:uppercase;letter-spacing:.4px;transition:all .15s}
  .btn-det:hover{border-color:var(--gold);color:var(--gold)}
  .det-box{background:var(--white);border:1px solid var(--border);border-radius:8px;
           box-shadow:var(--shadow);margin-bottom:22px;overflow:hidden}
  .det-head{background:var(--navy);padding:9px 16px;display:flex;align-items:center;
            justify-content:space-between;gap:10px;flex-wrap:wrap}
  .det-head .t{color:var(--gold);font-size:10px;font-weight:700;letter-spacing:.8px;text-transform:uppercase}
  .det-head .x{background:transparent;border:1px solid rgba(255,255,255,.25);border-radius:4px;
               color:rgba(255,255,255,.65);font-size:10px;padding:3px 9px;cursor:pointer}
  .det-head .x:hover{border-color:var(--gold);color:var(--gold)}
  .det table{width:100%;border-collapse:collapse;font-size:12px}
  .det th{background:#F7F8FA;color:var(--muted);font-size:10px;font-weight:700;letter-spacing:.5px;
          text-transform:uppercase;padding:8px 14px;text-align:left;border-bottom:1px solid var(--border)}
  .det td{padding:9px 14px;border-bottom:1px solid #EFF1F4;vertical-align:middle}
  .det tr:hover{background:#F8F9FF}
  .det .hora{font-variant-numeric:tabular-nums;font-weight:700;color:var(--navy);white-space:nowrap}
  .det a{color:var(--navy);text-decoration:none;font-weight:600}
  .det a:hover{text-decoration:underline}
  .st{display:inline-block;padding:2px 8px;border-radius:4px;font-size:10px;font-weight:700;
      border:1px solid;white-space:nowrap}
  .st-realizada{background:#ECFDF5;color:var(--green);border-color:#A7D8BE}
  .st-nao_validada{background:var(--amber-bg);color:var(--amber);border-color:#F0D9A8}
  .st-pendente{background:var(--blue-bg);color:#1E3A8A;border-color:#C9D4E8}
  .st-motivo{font-size:10px;color:var(--muted);margin-top:3px;line-height:1.3}
  .det-vazio{padding:22px;text-align:center;color:var(--muted);font-size:12px}
  .det-nota{padding:10px 16px;font-size:11px;color:var(--muted);border-top:1px solid var(--border);line-height:1.6}

  .nota-comp{font-size:11px;color:var(--muted);padding:9px 4px 0;line-height:1.6}
  .nota-comp b{color:var(--navy)}

  .peso-nota{font-size:10px;font-weight:600;color:var(--muted);text-transform:none;letter-spacing:0}

  /* PIPE A ARRUMAR */
  .chips{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:12px}
  .chip{border-radius:6px;padding:7px 12px;font-size:12px;border:1px solid;background:var(--white)}
  .chip b{font-size:14px;margin-right:4px}
  .chip-sem_data{border-color:#E8B4B4;color:var(--red);background:var(--red-bg)}
  .chip-sem_probabilidade{border-color:#F0D9A8;color:var(--amber);background:var(--amber-bg)}
  .chip-prob_fora{border-color:#C9D4E8;color:#1E3A8A;background:var(--blue-bg)}
  .pend thead tr{background:var(--navy)}
  .pend thead th{color:rgba(255,255,255,.7);font-size:10px;font-weight:600;letter-spacing:.5px;
                 text-transform:uppercase;padding:9px 14px;text-align:left;white-space:nowrap}
  .pend thead th.num{text-align:right}
  .pend td{padding:9px 14px;border-bottom:1px solid #EFF1F4;font-size:12px;vertical-align:middle}
  .pend td.num{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}
  .pend tbody tr:hover{background:#F8F9FF}
  .pend .tit a{color:var(--navy);text-decoration:none;font-weight:600}
  .pend .tit a:hover{text-decoration:underline}
  .pend .falta{color:var(--red);font-weight:700}
  .tag{display:inline-block;padding:2px 8px;border-radius:4px;font-size:10px;font-weight:700;
       letter-spacing:.3px;white-space:nowrap;border:1px solid}
  .tag-sem_data{background:var(--red-bg);color:var(--red);border-color:#E8B4B4}
  .tag-sem_probabilidade{background:var(--amber-bg);color:var(--amber);border-color:#F0D9A8}
  .tag-prob_fora{background:var(--blue-bg);color:#1E3A8A;border-color:#C9D4E8}
  .pend tr.total td{background:var(--gold-bg);border-top:2px solid var(--gold);font-weight:700}
  .pend-vazio{padding:22px;text-align:center;color:var(--green);font-size:12px;font-weight:600}

  /* FORECAST */
  .fc{width:100%;border-collapse:collapse}
  .fc thead tr{background:var(--navy)}
  .fc thead th{color:rgba(255,255,255,.7);font-size:10px;font-weight:600;letter-spacing:.5px;
               text-transform:uppercase;padding:9px 12px;text-align:left;white-space:nowrap}
  .fc thead th.num{text-align:right}
  .fc td{padding:7px 12px;border-bottom:1px solid #EFF1F4;font-size:12px;white-space:nowrap}
  .fc td.num{text-align:right;font-variant-numeric:tabular-nums}
  .fc td.dia{font-weight:600;color:var(--navy)}
  .fc td.prev{background:var(--gold-bg)}
  .fc td.real{background:var(--green-bg);color:var(--green);font-weight:600}
  .fc tr.fds td{background:#FAFAFC;color:var(--muted)}
  .fc tr.vazia td{color:#B9BDC7}
  .fc tr.passado td.dia{color:var(--muted);font-weight:500}
  .fc tr.e-hoje td{background:var(--blue-bg);border-top:1px solid #C9D4E8;border-bottom:1px solid #C9D4E8}
  .fc tr.e-hoje td.dia{color:var(--navy);font-weight:800}
  .fc .hoje-tag{background:var(--navy);color:#fff;font-size:9px;padding:1px 5px;
                border-radius:3px;letter-spacing:.5px;font-weight:700;margin-left:2px}
  .fc tr.total td{background:var(--gold-bg);border-top:2px solid var(--gold);font-weight:700}
  .fc-bt{background:transparent;border:1px solid var(--border);border-radius:4px;width:20px;height:20px;
         line-height:1;font-size:9px;color:var(--muted);cursor:pointer;margin-right:7px;padding:0}
  .fc-bt:hover:not(:disabled){border-color:var(--gold);color:var(--gold)}
  .fc-bt:disabled{opacity:.3;cursor:default}
  .fc-bt.aberto{background:var(--navy);border-color:var(--navy);color:#fff}
  .fc-det-linha td{padding:0;background:#F8F9FC;border-bottom:1px solid var(--border)}
  .fc-det{width:100%;border-collapse:collapse;margin:0}
  .fc-det thead th{background:#EEF0F5;color:var(--muted);font-size:9px;font-weight:700;
                   letter-spacing:.5px;text-transform:uppercase;padding:7px 14px;text-align:left}
  .fc-det thead th.num{text-align:right}
  .fc-det td{padding:7px 14px;border-bottom:1px solid #E8EAF0;font-size:12px}
  .fc-det td.num{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}
  .fc-det .tit a{color:var(--navy);text-decoration:none;font-weight:600}
  .fc-det .tit a:hover{text-decoration:underline}
  .fc-det .falta{color:var(--red);font-weight:700}
  .fr{display:inline-block;padding:1px 7px;border-radius:3px;font-size:10px;font-weight:700;
      border:1px solid;white-space:nowrap}
  .fr-navigator{background:#F2F3F7;color:var(--navy);border-color:#D3D6E0}
  .fr-mgm{background:var(--gold-bg);color:var(--amber);border-color:#F0D9A8}
  .fr-renovacao{background:var(--blue-bg);color:#43586F;border-color:#C9D4E8}
  .fc-nota{padding:12px 16px;font-size:11px;color:var(--muted);line-height:1.7;
           border-top:1px solid var(--border);background:#FCFCFD}
  .fc-aviso{padding:12px 16px;font-size:12px;color:var(--text);background:var(--amber-bg);
            border-bottom:1px solid #F0D9A8}
  .fc-vazio{padding:18px;text-align:center;color:var(--muted);font-size:12px}

  .loading{display:flex;align-items:center;justify-content:center;gap:12px;padding:60px;color:var(--muted)}
  .spinner{width:18px;height:18px;border:2px solid var(--border);border-top-color:var(--gold);
           border-radius:50%;animation:spin .7s linear infinite}
  @keyframes spin{to{transform:rotate(360deg)}}
  .erro{padding:40px;text-align:center;color:var(--red);font-size:13px}
  .erro code{display:block;margin-top:10px;font-size:11px;color:var(--muted);white-space:pre-wrap;text-align:left}
</style>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/chartjs-plugin-datalabels/2.2.0/chartjs-plugin-datalabels.min.js"></script>
</head>
<body>

<header class="header">
  <div class="header-left">
    <div class="brand">
      <div class="brand-bar"></div>
      <div>
        <div class="brand-text">BOARD ACADEMY</div>
        <div class="brand-sub" id="brand-sub">Ascensão/MGM · Consolidado</div>
      </div>
    </div>
    <div class="periodo-badge" id="periodo-badge">Carregando…</div>
  </div>
  <div class="header-right">
    <select id="sel-mes" onchange="trocouPeriodo()"></select>
    <select id="sel-ano" onchange="trocouPeriodo()"></select>
    <button class="btn-reload" id="btn-reload" onclick="atualizarTudo()">Atualizar</button>
    <span class="update-info" id="update-info"></span>
  </div>
</header>

<div class="tabs">
  <div class="tab on" id="tab-mes"    onclick="setView('mes')">Painel do Mês</div>
  <div class="tab"    id="tab-resumo" onclick="setView('resumo')">Resumo — 3 Frentes</div>
  <div class="tab"    id="tab-graficos" onclick="setView('graficos')">Gráficos</div>
  <div class="tab"    id="tab-forecast" onclick="setView('forecast')">Forecast</div>
</div>

<div class="main">
  <div class="view on" id="view-mes">
    <div id="conteudo">
      <div class="card"><div class="loading"><div class="spinner"></div>Buscando dados…</div></div>
    </div>
  </div>
  <div class="view" id="view-resumo">
    <div id="conteudo-resumo">
      <div class="card"><div class="loading"><div class="spinner"></div>Carregando resumo…</div></div>
    </div>
  </div>
  <div class="view" id="view-graficos">
    <div id="conteudo-graficos">
      <div class="card"><div class="loading"><div class="spinner"></div>Carregando gráficos…</div></div>
    </div>
  </div>
  <div class="view" id="view-forecast">
    <div id="conteudo-forecast">
      <div class="card"><div class="loading"><div class="spinner"></div>Carregando forecast…</div></div>
    </div>
  </div>
</div>

<script>
const MESES = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho',
               'Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];

// senha opcional: propagada da URL (?k=...) para a chamada da API
const KEY = new URLSearchParams(location.search).get('k') || '';

const R = v => 'R$ ' + Number(v||0).toLocaleString('pt-BR',{maximumFractionDigits:0});
const N = v => Number(v||0).toLocaleString('pt-BR');
const fmtDia = s => s ? s.slice(8,10) + '/' + s.slice(5,7) : '';
const P = v => v == null ? '—' : Number(v).toFixed(1).replace('.',',') + '%';
const esc = s => String(s == null ? '' : s)
  .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');

function pctTag(v){
  if (v == null) return '<span class="zero">—</span>';
  const c = v >= 100 ? 'pct-green' : v >= 60 ? 'pct-amber' : 'pct-red';
  return `<span class="pct ${c}">${P(v)}</span>`;
}
function money(v, opts={}){
  const n = Number(v||0);
  if (n === 0 && !opts.forcar) return '<span class="zero">R$ 0</span>';
  return R(n);
}

// ── seletores de período ──────────────────────────────────────
(function initSelects(){
  const hoje = new Date();
  const brNow = new Date(hoje.getTime() - (hoje.getTimezoneOffset()+180)*60000);
  const sm = document.getElementById('sel-mes');
  const sa = document.getElementById('sel-ano');
  MESES.forEach((m,i) => sm.add(new Option(m, i+1)));
  for (let a = brNow.getFullYear()-1; a <= brNow.getFullYear()+1; a++) sa.add(new Option(a, a));
  sm.value = brNow.getMonth()+1;
  sa.value = brNow.getFullYear();
})();

// ── render ────────────────────────────────────────────────────
function linha(rotulo, valorHtml, cls='', hint=''){
  return `<tr class="${cls}">
    <td>${rotulo}${hint ? `<span class="hint">${hint}</span>` : ''}</td>
    <td class="val">${valorHtml}</td>
  </tr>`;
}

function render(d){
  dadosPainel = d;
  const p = d.periodo, m = d.metricas;

  document.getElementById('periodo-badge').innerHTML =
    `${MESES[p.mes-1]} ${p.ano} &nbsp;·&nbsp; <strong>${p.du_passados}</strong>/${p.du_total} dias úteis
     &nbsp;·&nbsp; <strong>${p.du_restantes}</strong> restantes`;
  document.getElementById('update-info').textContent = 'Atualizado: ' + p.atualizado_em;


  // ── respostas rápidas: ontem e hoje ──
  const reu = d.reunioes || {hoje:{}, ontem:{}};
  const dh = d.detalhe_hoje || {};
  const cFunis = d.composicao || {};
  const sem = d.semana || {};
  const rmes = reu.mes || {agendadas:0, realizadas:0};
  const conv = reu.conversao || {deals_com_reuniao:0, deals_ganhos:0, taxa:0};
  const respostas = `
    <div class="cards">
      <div class="kcard ontem">
        <div class="rot">Reuniões — Último DU
          <button class="btn-det" onclick="verReunioes('ontem')">Detalhamento</button></div>
        <div class="val">${N(reu.ontem.realizadas || 0)}</div>
        <div class="sub">validadas de ${N(reu.ontem.agendadas || 0)} agendadas · ${p.ontem}</div>
      </div>
      <div class="kcard ontem">
        <div class="rot">Vendas — Último DU</div>
        <div class="val">${R(m.entrou_ontem_multi)}</div>
        <div class="sub">c/ multiplicador · bruto ${R(m.entrou_ontem_bruto)}</div>
      </div>
      <div class="kcard hoje">
        <div class="rot">Reuniões — Hoje
          <button class="btn-det" onclick="verReunioes('hoje')">Detalhamento</button></div>
        <div class="val">${N(reu.hoje.agendadas || 0)}</div>
        <div class="sub">agendadas · ${N(reu.hoje.realizadas || 0)} já validadas</div>
      </div>
      <div class="kcard hoje">
        <div class="rot">Vendas — Hoje</div>
        <div class="val">${R(m.entrou_hoje_multi)}</div>
        <div class="sub">c/ multiplicador · bruto ${R(m.entrou_hoje_bruto)}</div>
      </div>
    </div>
    <div id="det-reunioes"></div>
    <div class="cards">
      <div class="kcard hoje"><div class="rot">Pipe 70% — Hoje</div>
        <div class="val">${R(dh.p70 || 0)}</div><div class="sub">pondera ${R((dh.p70||0)*0.7)}</div></div>
      <div class="kcard hoje"><div class="rot">Pipe 50% — Hoje</div>
        <div class="val">${R(dh.p50 || 0)}</div><div class="sub">pondera ${R((dh.p50||0)*0.5)}</div></div>
      <div class="kcard hoje"><div class="rot">Pipe 20% — Hoje</div>
        <div class="val">${R(dh.p20 || 0)}</div><div class="sub">pondera ${R((dh.p20||0)*0.2)}</div></div>
      <div class="kcard"><div class="rot">Previsto Hoje</div>
        <div class="val">${R(m.previsto_hoje)}</div>
        <div class="sub">de ${R(m.em_aberto_hoje)} em aberto (${dh.qtd || 0} negócios)${
          cFunis.etapa_previsao ? ` · só negócios em Negociação` : ''}</div></div>
    </div>`;

  const gapCls   = m.gap_100 > 0 ? 'neg' : 'pos';
  const devHint  = `(${P(m.pct_mes_decorrido)} do mês)`;

  const notaComp = d.consolidado ? `
    <div class="nota-comp">Consolidado: <b>Navigator</b> ${N(cFunis.navigator?.qtd || 0)} venda(s) ·
      ${R(cFunis.navigator?.valor || 0)} &nbsp;+&nbsp; <b>MGM</b> ${N(cFunis.mgm?.qtd || 0)} venda(s) ·
      ${R(cFunis.mgm?.valor || 0)}. A separação por frente fica na aba Resumo.${
        cFunis.etapa_previsao ? ` A previsão do dia considera só os ${N(cFunis.abertos_negociacao || 0)} negócios
        na etapa <b>${esc(cFunis.etapa_previsao)}</b>, de ${N(cFunis.abertos_total || 0)} abertos no total.` : ''}</div>` : '';

  const kpi = `
  <div class="card">
    <div class="table-scroll">
      <table class="kpi">
        <thead>
          <tr><th>Métrica</th><th class="val">${(d.rotulo || d.funil || '').toUpperCase()}</th></tr>
        </thead>
        <tbody>
          ${linha('Meta Mês',                 money(m.meta_mes,{forcar:1}), 'destaque')}
          ${linha('Meta Dia',                 money(m.meta_dia,{forcar:1}))}
          ${linha('Realizado Bruto',          money(m.real_bruto))}
          ${linha('Realizado Multiplicador',  money(m.real_multi))}
          ${linha('Deveria (100%) — MTD',     `${money(m.deveria_mtd,{forcar:1})} <span class="hint">${devHint}</span>`, 'sep')}
          ${linha('Atingimento',              pctTag(m.atingimento), 'destaque', 'c/ multiplicador')}
          ${linha('Gap 100%',                 `<span class="${gapCls}">${R(m.gap_100)}</span>`, '', 'c/ multiplicador')}
          ${linha('Meta/Dia 100% (c/ Multiplicador)', money(m.meta_dia_100,{forcar:1}), '', `÷ ${p.du_restantes} DU`)}
          ${linha('Meta/Dia 100% (Bruto)',            money(m.meta_dia_100_bruto,{forcar:1}), '', `÷ ${p.du_restantes} DU`)}

          ${linha(`<b>Semana</b> <span class="hint">${sem.ini ? fmtDia(sem.ini) + ' a ' + fmtDia(sem.fim) : ''}</span>`, '', 'sep g-hoje destaque')}
          ${linha('Entrou na Semana (Multiplicador)', money(sem.entrou_multi), 'g-hoje')}
          ${linha('Entrou na Semana (Bruto)',         money(sem.entrou_bruto), 'g-hoje')}
          ${linha('Vendas na Semana',                 N(sem.qtd_vendas || 0), 'g-hoje')}
          ${linha('Previsto da Semana',               money(sem.previsto), 'g-hoje', 'soma dos ponderados abaixo')}
          ${linha('Em Aberto na Semana',              `${money(sem.em_aberto)} <span class="hint">${N(sem.qtd_abertos || 0)} negócio(s)</span>`, 'g-hoje')}
          ${linha('&nbsp;&nbsp;· Pipe 70%', `${money(sem.p70)} <span class="hint">pondera ${R((sem.p70||0)*0.7)}</span>`, 'g-hoje')}
          ${linha('&nbsp;&nbsp;· Pipe 50%', `${money(sem.p50)} <span class="hint">pondera ${R((sem.p50||0)*0.5)}</span>`, 'g-hoje')}
          ${linha('&nbsp;&nbsp;· Pipe 20%', `${money(sem.p20)} <span class="hint">pondera ${R((sem.p20||0)*0.2)}</span>`, 'g-hoje')}
          ${linha('Reuniões na Semana',               `${N(sem.reunioes_validadas || 0)} <span class="hint">validadas de ${N(sem.reunioes_agendadas || 0)} agendadas</span>`, 'g-hoje')}
          ${linha('Entrou Hoje (Multiplicador)',  money(m.entrou_hoje_multi), 'g-hoje')}
          ${linha('Entrou Hoje (Bruto)',          money(m.entrou_hoje_bruto), 'g-hoje')}

          ${linha('Reuniões Validadas no Mês', `${N(rmes.realizadas || 0)} <span class="hint">de ${N(rmes.agendadas || 0)} agendadas</span>`, 'sep')}
          ${linha('Taxa de Conversão das Reuniões', pctTag(conv.taxa), 'destaque',
                  `${N(conv.deals_ganhos || 0)} ganho(s) de ${N(conv.deals_com_reuniao || 0)} negócio(s) com reunião validada`)}
          ${linha('Vendas no mês',  `${N(m.qtd_ganhos_mes)}`)}
          ${linha('Ticket médio',   money(m.ticket_medio))}
        </tbody>
      </table>
    </div>
  </div>${notaComp}`;

  // ── closers ──
  let closersHtml = '';
  if ((d.closers || []).length) {
    const tb = d.closers.map(c => `
      <tr>
        <td>${c.nome}</td>
        <td>${c.meta > 0 ? R(c.meta) : '<span class="zero">—</span>'}</td>
        <td>${money(c.real_bruto)}</td>
        <td>${money(c.real_multi)}</td>
        <td>${pctTag(c.pct)}</td>
        <td>${N(c.qtd)}</td>
        <td>${money(c.ticket)}</td>
        <td>${money(c.previsto_hoje)}</td>
        <td>${money(c.aberto_hoje)}</td>
      </tr>`).join('');

    const t = d.closers.reduce((a,c) => ({
      meta:a.meta+(c.meta||0), bruto:a.bruto+c.real_bruto, multi:a.multi+c.real_multi,
      qtd:a.qtd+c.qtd, prev:a.prev+c.previsto_hoje, ab:a.ab+c.aberto_hoje
    }), {meta:0,bruto:0,multi:0,qtd:0,prev:0,ab:0});

    closersHtml = `
    <div class="block-title">Por Closer<div class="rule"></div></div>
    <div class="card">
      <div class="table-scroll">
        <table class="closers">
          <thead><tr>
            <th>Closer</th><th>Meta</th><th>Bruto</th><th>Multiplicador</th>
            <th>% Ating.</th><th>Vol.</th><th>Ticket</th>
            <th>Previsto Hoje</th><th>Em Aberto Hoje</th>
          </tr></thead>
          <tbody>
            ${tb}
            <tr class="total">
              <td>TOTAL</td>
              <td>${t.meta > 0 ? R(t.meta) : '—'}</td>
              <td>${R(t.bruto)}</td>
              <td>${R(t.multi)}</td>
              <td>${pctTag(t.meta > 0 ? t.multi/t.meta*100 : null)}</td>
              <td>${N(t.qtd)}</td>
              <td>${R(t.qtd ? t.bruto/t.qtd : 0)}</td>
              <td>${R(t.prev)}</td>
              <td>${R(t.ab)}</td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>`;
  }


  // ── SDRs ──
  let sdrHtml = '';
  if ((d.sdrs || []).length) {
    const linhas = d.sdrs.map(s => `
      <tr>
        <td>${esc(s.nome)}${s.encontrado_no_pipedrive ? '' : ' <span class="tag tag-sem_data">não achei no Pipedrive</span>'}</td>
        <td>${s.meta_reuniao > 0 ? N(s.meta_reuniao) : '<span class="zero">—</span>'}</td>
        <td>${s.meta_diaria > 0 ? s.meta_diaria.toFixed(1).replace('.',',') : '<span class="zero">—</span>'}</td>
        <td><b>${N(s.validadas)}</b></td>
        <td>${s.deveria_estar > 0 ? Math.round(s.deveria_estar) : '<span class="zero">—</span>'}</td>
        <td>${s.meta_reuniao > 0 ? (s.faltam <= 0
              ? `<span class="pos">+${N(Math.abs(Math.round(s.faltam)))}</span>`
              : `<span class="neg">${N(Math.round(s.faltam))}</span>`) : '<span class="zero">—</span>'}</td>
        <td>${pctTag(s.pct_reu)}</td>
        <td>${s.meta_ganho > 0 ? R(s.meta_ganho) : '<span class="zero">—</span>'}</td>
        <td>${N(s.qtd_ganhos)}</td>
        <td>${money(s.valor_ganho)}</td>
        <td>${money(s.valor_ganho_multi)}</td>
        <td>${pctTag(s.pct_ganhos)}</td>
        <td>${pctTag(s.pct_final)}</td>
      </tr>`).join('');
    const pesos = d.sdrs[0].pesos || '70/30';
    sdrHtml = `
      <div class="block-title">Por SDR<div class="rule"></div>
        <span class="peso-nota">% Final = % Reunião x ${pesos.split('/')[0]}% + % Ganhos x ${pesos.split('/')[1]}%</span>
      </div>
      <div class="card">
        <div class="table-scroll">
          <table class="closers">
            <thead><tr>
              <th>SDR</th><th>Meta Reu.</th><th>Meta/Dia</th><th>Validadas</th>
              <th>Deveria Estar</th><th>Faltam</th><th>% Reu.</th>
              <th>Meta Ganho</th><th>Qtd</th><th>R$ Ganho</th><th>R$ c/ Multi</th>
              <th>% Ganhos</th><th>% Final</th>
            </tr></thead>
            <tbody>${linhas}</tbody>
          </table>
        </div>
      </div>`;
  }

  // ── pipe a arrumar ──
  let pendHtml = '';
  const pend = d.pendencias;
  if (pend) {
    const ordem = ['sem_data','sem_probabilidade','prob_fora'];
    const chips = ordem.map(k => {
      const r = pend.resumo[k];
      if (!r || !r.qtd) return '';
      return `<div class="chip chip-${k}"><b>${r.qtd}</b> ${r.label} · ${R(r.valor)}</div>`;
    }).join('');

    const corpo = pend.qtd_total
      ? pend.itens.map(i => `
          <tr>
            <td class="tit"><a href="${i.url}" target="_blank" rel="noopener">${esc(i.titulo)} ↗</a></td>
            <td>${esc(i.dono)}</td>
            <td class="num">${R(i.valor)}</td>
            <td class="num">${i.prev_fechamento || '<span class="falta">faltando</span>'}</td>
            <td class="num">${i.probabilidade == null ? '<span class="falta">faltando</span>' : i.probabilidade + '%'}</td>
            <td><span class="tag tag-${i.motivo}">${i.motivo_label}</span></td>
          </tr>`).join('') +
        `<tr class="total">
            <td>TOTAL A CORRIGIR</td><td>${pend.qtd_total} deals</td>
            <td class="num">${R(pend.valor_total)}</td>
            <td class="num" colspan="3">${P(pend.valor_abertos ? pend.valor_total/pend.valor_abertos*100 : 0)} do pipe aberto (${R(pend.valor_abertos)})</td>
         </tr>`
      : '';

    pendHtml = `
      <div class="block-title">Pipe a arrumar${pend.filtrado_por ? ` — ${esc(pend.filtrado_por)}` : ''}<div class="rule"></div></div>
      ${chips ? `<div class="chips">${chips}</div>` : ''}
      <div class="card">
        ${pend.qtd_total ? `
        <div class="table-scroll">
          <table class="pend">
            <thead><tr>
              <th>Negócio</th><th>Dono</th><th class="num">Valor</th>
              <th class="num">Prev. Fechamento</th><th class="num">Prob.</th><th>O que falta</th>
            </tr></thead>
            <tbody>${corpo}</tbody>
          </table>
        </div>` : '<div class="pend-vazio">✓ Pipe limpo — todos os negócios abertos têm data e probabilidade preenchidas.</div>'}
      </div>`;
  }

  // ── alertas + composição da meta ──
  const alertas = (d.alertas || []).map(a => `<div class="alerta">⚠ ${a}</div>`).join('');
  const comp = (d.meta_composicao || []).length
    ? `<details class="meta-comp">
         <summary>Como a Meta Mês de ${R(m.meta_mes)} foi montada (${d.meta_composicao.length} closers)</summary>
         <ul>${d.meta_composicao.map(c => `<li>${c.nome} — ${R(c.meta)}</li>`).join('')}</ul>
       </details>`
    : '';

  document.getElementById('conteudo').innerHTML =
    `${respostas}<div class="block-title">Painel do Mês<div class="rule"></div></div>${kpi}${closersHtml}${sdrHtml}${pendHtml}${alertas}${comp}
     <div class="rodape">Funil ${d.funil} (pipeline ${d.pipeline_id}) · atualizado em ${p.atualizado_em}</div>`;
}

// ── fetch ─────────────────────────────────────────────────────
async function carregar(){
  document.getElementById('conteudo').innerHTML =
    '<div class="card"><div class="loading"><div class="spinner"></div>Buscando dados do Navigator…</div></div>';
  try {
    const mes = document.getElementById('sel-mes').value;
    const ano = document.getElementById('sel-ano').value;
    // /api/index é o path canônico da function na Vercel — sempre chega no Python.
    // fmt=json é o que diz ao backend para devolver os dados do painel.
    const q = new URLSearchParams({fmt: 'json', mes, ano});
    if (KEY) q.set('k', KEY);
    const res  = await fetch('/api/index?' + q.toString());
    const data = await res.json();
    if (data.erro) {
      const extra = data.funis_disponiveis
        ? '<code>Funis encontrados no Pipedrive:\n· ' + data.funis_disponiveis.join('\n· ') + '</code>'
        : (data.trace ? `<code>${data.trace}</code>` : '');
      document.getElementById('conteudo').innerHTML =
        `<div class="card"><div class="erro">${data.erro}${extra}</div></div>`;
      return;
    }
    render(data);
  } catch (e) {
    document.getElementById('conteudo').innerHTML =
      `<div class="card"><div class="erro">Falha ao carregar: ${e.message}</div></div>`;
  }
}

// ── DETALHAMENTO DE REUNIÕES ──────────────────────────────────
let dadosPainel = null;
const ST_LABEL = {realizada:'Validada', nao_validada:'Não validada', pendente:'Pendente'};

let DONO_REU = 'Denise Mussolin';
function verReunioes(qual){
  const box = document.getElementById('det-reunioes');
  if (!box || !dadosPainel) return;
  if (box.dataset.aberto === qual) { box.innerHTML = ''; box.dataset.aberto = ''; return; }
  box.dataset.aberto = qual;

  const r = dadosPainel.reunioes || {};
  if (r.dono) DONO_REU = r.dono;
  const itens = qual === 'hoje' ? (r.lista_hoje || []) : (r.lista_ontem || []);
  const dia = qual === 'hoje' ? dadosPainel.periodo.hoje : dadosPainel.periodo.ontem;
  const dup = (qual === 'hoje' ? r.hoje : r.ontem)?.duplicadas_ocultas || 0;
  const rot = qual === 'hoje' ? 'Hoje' : 'Último dia útil';

  const linhas = itens.map(i => `
    <tr>
      <td class="hora">${i.hora || '—'}</td>
      <td>${esc(i.responsavel)}</td>
      <td>${i.funil ? esc(i.funil) : '<span class="zero">—</span>'}</td>
      <td>${i.deal_url ? `<a href="${i.deal_url}" target="_blank" rel="noopener">${esc(i.deal || i.assunto)} ↗</a>` : esc(i.assunto)}</td>
      <td><span class="st st-${i.status}">${ST_LABEL[i.status]}</span></td>
    </tr>`).join('');

  box.innerHTML = `
    <div class="det-box det">
      <div class="det-head">
        <span class="t">Reuniões — ${rot} · ${dia} · ${itens.length} registro(s)</span>
        <button class="x" onclick="verReunioes('${qual}')">fechar</button>
      </div>
      ${itens.length ? `
        <table>
          <thead><tr><th>Hora</th><th>Responsável</th><th>Funil</th><th>Negócio</th><th>Status</th></tr></thead>
          <tbody>${linhas}</tbody>
        </table>` : '<div class="det-vazio">Nenhuma reunião registrada nesse dia.</div>'}
      <div class="det-nota">
        <b>Validada</b> = concluída, com responsável diferente do dono do negócio e dentro do filtro de Reunião Validada — é o número que o card conta.
        <b>Não validada</b> = concluída, mas reprovada em uma dessas duas regras.
        <b>Pendente</b> = ainda não marcada como concluída.
        Conta toda reunião em negócio de <b>${esc(DONO_REU)}</b>, que é quem conduz — não importa quem agendou.
        ${r.excluidos ? `Reuniões com <b>${esc(r.excluidos)}</b> como responsável ficam de fora.` : ''}
        ${dup > 0 ? `<br><b>${dup}</b> atividade(s) duplicada(s) no mesmo negócio foram ocultadas — no Pipedrive existe mais de uma para o mesmo negócio neste dia.` : ''}
        ${r.fora_escopo > 0 ? `<br><b>${r.fora_escopo}</b> reunião(ões) do mês ficaram de fora por estarem em funil que não é Navigator nem MGM.` : ''}
        ${r.fora_dono > 0 ? `<br><b>${r.fora_dono}</b> reunião(ões) do mês ficaram de fora por serem em negócio de outra pessoa.` : ''}
      </div>
    </div>`;
}

// ── GRÁFICOS ──────────────────────────────────────────────────
const COR_FRENTE = { navigator: '#1A1A2E', mgm: '#B8860B', renovacao: '#7A8CA8' };
let CHARTS = {};

function destroiCharts(){
  Object.values(CHARTS).forEach(c => { try { c.destroy(); } catch(e){} });
  CHARTS = {};
}
const diaCurto = s => s.slice(8,10) + '/' + s.slice(5,7);
const kBRL  = v => v >= 1000 ? 'R$ ' + (v/1000).toFixed(0) + 'k' : 'R$ ' + Math.round(v);
const ST_CURTO = { realizada: 'validada', nao_validada: 'não validada', pendente: 'pendente' };

// lista de reuniões do dia para o rodapé do tooltip
function detalheDoDia(dados, idx, apenasValidadas){
  const dia = dados.dias[idx];
  let itens = (dados.titulos || {})[dia] || [];
  if (apenasValidadas) itens = itens.filter(i => i.s === 'realizada');
  if (!itens.length) return '';
  const LIM = 8;
  const linhas = itens.slice(0, LIM).map(i =>
    `${i.h || '--:--'}  ${i.t}${i.r ? '  · ' + i.r : ''}` +
    (apenasValidadas ? '' : `  [${ST_CURTO[i.s] || i.s}${i.m ? ': ' + i.m : ''}]`));
  if (itens.length > LIM) linhas.push(`+ ${itens.length - LIM} outra(s)`);
  return linhas;
}

const kCurto = v => v >= 1000 ? (v/1000).toFixed(0) + 'k' : String(Math.round(v));  // cabe dentro da barra

// ── janelinha de total do mês, desenhada por cima do gráfico ───
const CURTO_FRENTE = { navigator: 'Ascensão', mgm: 'MGM', renovacao: 'Renovação' };

function somaSerie(mapa, fr){
  const partes = fr.map(f => ({
    chave: f.chave,
    nome: CURTO_FRENTE[f.chave] || f.nome,
    v: ((mapa || {})[f.chave] || []).reduce((a, b) => a + (Number(b) || 0), 0)
  }));
  return { tot: partes.reduce((a, p) => a + p.v, 0), partes };
}

function caixaTotal(rotulo, mapa, fr, fmt){
  const s = somaSerie(mapa, fr);
  const quebra = s.partes.filter(p => p.v > 0).map(p =>
    `<span style="color:${COR_FRENTE[p.chave] || '#6B7280'}">●</span> ${esc(p.nome)} <b>${fmt(p.v)}</b>`
  ).join(' &nbsp;·&nbsp; ') || 'nada registrado no mês';
  return `<div class="graf-total">
      <div class="gt-rot">${esc(rotulo)}</div>
      <div class="gt-val">${fmt(s.tot)}</div>
      <div class="gt-sub">${quebra}</div>
    </div>`;
}

function caixaJacare(d){
  const j = d.jacare || [];
  let iUlt = -1;
  for (let i = 0; i < j.length; i++) if (j[i].real_mtd != null) iUlt = i;
  const real     = iUlt >= 0 ? Number(j[iUlt].real_mtd || 0) : 0;
  const bruto    = iUlt >= 0 ? Number(j[iUlt].bruto_mtd || 0) : 0;
  const metaHoje = iUlt >= 0 ? Number(j[iUlt].meta_mtd || 0) : 0;
  const metaMes  = Number(d.periodo.meta_mes || 0);
  const gap      = real - metaHoje;
  const pctMes   = metaMes  ? (real / metaMes)  * 100 : null;   // atingimento da meta cheia
  const pctMtd   = metaHoje ? (real / metaHoje) * 100 : null;   // atingimento do ritmo (MTD)
  const cls      = gap >= 0 ? 'gt-pos' : 'gt-neg';
  const sinal    = gap >= 0 ? '+' : '−';
  return `<div class="graf-total">
      <div class="gt-cols">
        <div class="gt-col">
          <div class="gt-rot">Realizado MTD <span class="gt-pct">c/ multi</span></div>
          <div class="gt-val">${R(real)}</div>
        </div>
        <div class="gt-col">
          <div class="gt-rot">Bruto MTD</div>
          <div class="gt-val">${R(bruto)}</div>
        </div>
        <div class="gt-col">
          <div class="gt-rot">Meta MTD</div>
          <div class="gt-val">${R(metaHoje)}</div>
        </div>
        <div class="gt-col">
          <div class="gt-rot">Ating. MTD</div>
          <div class="gt-val ${cls}">${P(pctMtd)}</div>
        </div>
      </div>
      <div class="gt-sub"><span class="${cls}">${sinal}${R(Math.abs(gap))}</span> vs. ritmo de hoje
        &nbsp;·&nbsp; meta do mês ${R(metaMes)} (<b>${P(pctMes)}</b> atingido)</div>
    </div>`;
}

function baseEmpilhado(rotulos, series, opts){
  return {
    type: 'bar',
    data: { labels: rotulos.map(diaCurto), datasets: series },
    options: {
      responsive: true, maintainAspectRatio: false,
      layout: { padding: { top: 42 } },   // espaço para a janelinha de total
      interaction: { mode: 'index', intersect: false },
      scales: {
        x: { stacked: true, grid: { display: false },
             ticks: { color: '#6B7280', font: { size: 9 }, maxRotation: 0, autoSkip: false } },
        y: { stacked: true, beginAtZero: true,
             grid: { color: 'rgba(0,0,0,.05)' },
             ticks: { color: '#6B7280', font: { size: 10 }, callback: opts.tickY } }
      },
      plugins: {
        legend: { position: 'top', align: 'end',
          labels: { boxWidth: 12, boxHeight: 12, font: { size: 11, weight: '600' },
                    color: '#374151', usePointStyle: true, pointStyle: 'rect' } },
        tooltip: { backgroundColor: '#1A1A2E', padding: 11, cornerRadius: 6,
          titleFont: { size: 12 }, bodyFont: { size: 11.5 }, footerFont: { size: 11, weight: '400' },
          bodySpacing: 3, footerMarginTop: 8, footerColor: '#C9CBD6', boxPadding: 3,
          filter: item => item.parsed.y > 0,
          callbacks: {
            label: c => ' ' + c.dataset.label + ': ' + opts.fmt(c.parsed.y),
            footer: items => opts.detalhe ? opts.detalhe(items[0].dataIndex) : ''
          } },
        datalabels: {
          color: '#fff', font: { size: 9, weight: '700' },
          formatter: v => v > 0 ? opts.rotulo(v) : '',
          display: ctx => ctx.dataset.data[ctx.dataIndex] > 0
        }
      }
    },
    plugins: [ChartDataLabels]
  };
}

function renderGraficos(d){
  destroiCharts();
  const fr = d.frentes;
  const alertas = (d.alertas || []).map(a => `<div class="alerta">⚠ ${esc(a)}</div>`).join('');

  const cardsAting = d.atingimento.map(a => `
    <div class="kcard">
      <div class="rot">${esc(a.nome)}</div>
      <div class="val">${a.pct == null ? '—' : P(a.pct)}</div>
      <div class="sub">${R(a.multi)} de ${R(a.meta)}<br>
        ${N(a.vendas)} venda(s) · ${N(a.reunioes)} reunião(ões) validada(s)<br>
        conversão ${P(a.conversao)} <span class="hint">${N(a.deals_ganhos)} de ${N(a.deals_com_reuniao)} negócio(s)</span></div>
    </div>`).join('');

  const cm = d.conversao_mes || {};

  document.getElementById('conteudo-graficos').innerHTML = `
    <div class="block-title">Atingimento por frente<div class="rule"></div></div>
    <div class="cards">${cardsAting}</div>
    <div class="nota-comp">Conversão consolidada: <b>${P(cm.taxa)}</b> —
      ${N(cm.deals_ganhos || 0)} negócio(s) Ganho de ${N(cm.deals_com_reuniao || 0)} com reunião validada no mês.</div>

    <div class="graf">
      <div class="graf-tit">Reuniões validadas por dia</div>
      <div class="graf-sub">O subconjunto que passou na régua de validação ·
        passe o mouse na barra para ver os negócios</div>
      <div class="graf-box"><canvas id="g-reunioes"></canvas>
        ${caixaTotal('Validadas no mês', d.reunioes, fr, v => N(v))}</div>
    </div>

    <div class="graf">
      <div class="graf-tit">Vendas por dia</div>
      <div class="graf-sub">Valor bruto das vendas, empilhado por frente</div>
      <div class="graf-box"><canvas id="g-vendas"></canvas>
        ${caixaTotal('Vendas brutas no mês', d.vendas, fr, v => R(v))}</div>
    </div>

    <div class="graf">
      <div class="graf-tit">Meta x Realizado acumulado</div>
      <div class="graf-sub">Meta de ${R(d.periodo.meta_mes)} distribuída pelos ${N(d.periodo.du_total)} dias úteis,
        contra o realizado acumulado — linha cheia é o <b>bruto</b>, pontilhada é o valor
        <b>com multiplicador</b> (é essa que bate contra a meta)</div>
      <div class="graf-box alto"><canvas id="g-jacare"></canvas>
        ${caixaJacare(d)}</div>
    </div>
    ${alertas}
    <div class="rodape">Atualizado em ${d.periodo.atualizado_em}</div>`;

  if (typeof Chart === 'undefined') {
    document.getElementById('conteudo-graficos').insertAdjacentHTML('afterbegin',
      '<div class="alerta">⚠ Não consegui carregar a biblioteca de gráficos (Chart.js). Verifique a conexão.</div>');
    return;
  }

  // 1 — reuniões validadas
  CHARTS.reu = new Chart(document.getElementById('g-reunioes'), baseEmpilhado(
    d.dias,
    fr.map(f => ({ label: f.nome, data: d.reunioes[f.chave],
                   backgroundColor: COR_FRENTE[f.chave], borderWidth: 0,
                   borderRadius: 2, maxBarThickness: 30 })),
    { fmt: v => N(v), rotulo: v => N(v), tickY: v => N(v),
      detalhe: i => detalheDoDia(d, i, true) }
  ));

  // 2 — vendas brutas
  CHARTS.ven = new Chart(document.getElementById('g-vendas'), baseEmpilhado(
    d.dias,
    fr.map(f => ({ label: f.nome, data: d.vendas[f.chave],
                   backgroundColor: COR_FRENTE[f.chave], borderWidth: 0,
                   borderRadius: 2, maxBarThickness: 30 })),
    { fmt: v => R(v), rotulo: v => kCurto(v), tickY: v => kBRL(v) }
  ));

  // 3 — jacaré
  CHARTS.jac = new Chart(document.getElementById('g-jacare'), {
    type: 'line',
    data: {
      labels: d.jacare.map(j => diaCurto(j.dia)),
      datasets: [
        { label: 'Meta acumulada', data: d.jacare.map(j => j.meta_mtd),
          borderColor: '#B8860B', borderWidth: 2, borderDash: [5,4],
          pointRadius: 0, pointHoverRadius: 5, tension: 0, fill: false, order: 3 },
        { label: 'Realizado bruto', data: d.jacare.map(j => j.bruto_mtd),
          borderColor: '#8A8FA3', borderWidth: 1.8, pointRadius: 0, pointHoverRadius: 5,
          tension: 0.15, spanGaps: false, fill: false, order: 2 },
        { label: 'Realizado c/ multiplicador', data: d.jacare.map(j => j.real_mtd),
          borderColor: '#1A1A2E', borderWidth: 2.5, borderDash: [6,3],
          pointRadius: 0, pointHoverRadius: 6,
          tension: 0.15, spanGaps: false, order: 1,
          fill: { target: 0, above: 'rgba(13,122,62,.10)', below: 'rgba(160,160,170,.14)' } }
      ]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      layout: { padding: { top: 42 } },   // espaço para a janelinha de total
      interaction: { mode: 'index', intersect: false },
      scales: {
        x: { grid: { display: false },
             ticks: { color: '#6B7280', font: { size: 9 }, maxRotation: 0, autoSkip: true, maxTicksLimit: 16 } },
        y: { beginAtZero: true, grid: { color: 'rgba(0,0,0,.05)' },
             ticks: { color: '#6B7280', font: { size: 10 }, callback: v => kBRL(v) } }
      },
      plugins: {
        datalabels: { display: false },
        legend: { position: 'top', align: 'end',
          labels: { boxWidth: 20, boxHeight: 2, font: { size: 11, weight: '600' },
                    color: '#374151', usePointStyle: true, pointStyle: 'line' } },
        tooltip: { backgroundColor: '#1A1A2E', padding: 12, cornerRadius: 6,
          callbacks: {
            label: c => ' ' + c.dataset.label + ': ' + (c.parsed.y == null ? '—' : R(c.parsed.y)),
            footer: items => {
              const meta = items.find(i => i.datasetIndex === 0);
              const real = items.find(i => i.datasetIndex === 2);   // c/ multiplicador
              if (!meta || !real || real.parsed.y == null || !meta.parsed.y) return '';
              const g = real.parsed.y - meta.parsed.y;
              return (g >= 0 ? 'Acima da meta em ' : 'Abaixo da meta em ') + R(Math.abs(g));
            }
          } }
      }
    },
    plugins: [ChartDataLabels]
  });
}

async function carregarGraficos(){
  document.getElementById('conteudo-graficos').innerHTML =
    '<div class="card"><div class="loading"><div class="spinner"></div>Carregando gráficos…</div></div>';
  try {
    const q = new URLSearchParams({fmt:'graficos',
      mes: document.getElementById('sel-mes').value,
      ano: document.getElementById('sel-ano').value});
    if (KEY) q.set('k', KEY);
    const data = await (await fetch('/api/index?' + q)).json();
    if (data.erro) throw new Error(data.erro);
    renderGraficos(data);
    graficosCarregados = true;
  } catch(e) {
    document.getElementById('conteudo-graficos').innerHTML =
      `<div class="card"><div class="erro">Erro: ${esc(e.message)}</div></div>`;
  }
}

// ── ABAS ──────────────────────────────────────────────────────
const ABAS = ['mes','resumo','graficos','forecast'];
let resumoCarregado = false, graficosCarregados = false, forecastCarregado = false;

// Atualiza as QUATRO abas de uma vez. É o que o botão Atualizar faz, e é o que
// roda ao trocar o período. Não existe atualização automática por tempo.
let atualizando = false, iniciado = false;
async function atualizarTudo(){
  if (atualizando) return;
  atualizando = true;
  const btn = document.getElementById('btn-reload');
  const rot = btn ? btn.textContent : '';
  if (btn) { btn.textContent = 'Atualizando…'; btn.disabled = true; }
  try {
    await Promise.all([carregar(), carregarResumo(), carregarGraficos(), carregarForecast()]);
  } finally {
    if (btn) { btn.textContent = rot || 'Atualizar'; btn.disabled = false; }
    atualizando = false;
  }
}

function trocouPeriodo(){ atualizarTudo(); }

function setView(v){
  ABAS.forEach(k => {
    document.getElementById('view-' + k).classList.toggle('on', k === v);
    document.getElementById('tab-' + k).classList.toggle('on', k === v);
  });
  location.hash = v === 'mes' ? '' : '#' + v;
  // rede de segurança: se a carga dessa aba falhou, tenta de novo ao abrir.
  // Durante o boot/refresh geral não faz nada — atualizarTudo() já está buscando.
  if (!iniciado || atualizando) return;
  if (v === 'resumo'   && !resumoCarregado)    carregarResumo();
  if (v === 'graficos' && !graficosCarregados) carregarGraficos();
  if (v === 'forecast' && !forecastCarregado)  carregarForecast();
}

// ── RESUMO DAS 3 FRENTES ──────────────────────────────────────
function renderResumo(d){
  const p = d.periodo;
  const linhaFrente = (f, cls='') => `
    <tr class="${cls}${(f.meta === 0 && f.real_multi === 0) ? ' zerada' : ''}">
      <td>${esc(f.nome)}</td>
      <td>${f.meta > 0 ? R(f.meta) : '<span class="zero">—</span>'}</td>
      <td>${f.deveria_mtd > 0 ? R(f.deveria_mtd) : '<span class="zero">—</span>'}</td>
      <td>${money(f.real_bruto)}</td>
      <td>${money(f.real_multi)}</td>
      <td>${pctTag(f.pct)}</td>
      <td>${f.meta > 0 ? `<span class="${f.gap > 0 ? 'neg' : 'pos'}">${R(f.gap)}</span>` : '<span class="zero">—</span>'}</td>
      <td>${f.meta_dia_rest > 0 ? R(f.meta_dia_rest) : '<span class="zero">—</span>'}</td>
      <td>${N(f.qtd)}</td>
      <td>${money(f.pipe_aberto)}</td>
      <td>${money(f.previsto_hoje)}</td>
    </tr>`;


  const alertas = (d.alertas || []).map(a => `<div class="alerta">⚠ ${esc(a)}</div>`).join('');

  document.getElementById('conteudo-resumo').innerHTML = `
    <div class="cards">
      <div class="kcard"><div class="rot">Meta do Mês</div><div class="val">${R(d.total.meta)}</div>
        <div class="sub">três frentes somadas</div></div>
      <div class="kcard"><div class="rot">Realizado</div><div class="val">${R(d.total.real_multi)}</div>
        <div class="sub">c/ multiplicador · bruto ${R(d.total.real_bruto)}</div></div>
      <div class="kcard"><div class="rot">Atingimento</div><div class="val">${P(d.total.pct)}</div>
        <div class="sub">deveria estar em ${R(d.total.deveria_mtd)}</div></div>
      <div class="kcard"><div class="rot">Falta por dia</div><div class="val">${R(d.total.meta_dia_rest)}</div>
        <div class="sub">nos ${p.du_restantes} DU restantes</div></div>
    </div>

    <div class="block-title">Por frente — ${MESES[p.mes-1]} ${p.ano}<div class="rule"></div>
      ${d.etapa_previsao ? `<span class="peso-nota">pipe e previsto: só etapa ${esc(d.etapa_previsao)}</span>` : ''}
    </div>
    <div class="card"><div class="table-scroll">
      <table class="frentes">
        <thead><tr>
          <th>Frente</th><th>Meta</th><th>Deveria (MTD)</th><th>Bruto</th><th>Multiplicador</th>
          <th>% Ating.</th><th>Gap</th><th>Falta/Dia</th><th>Vol.</th>
          <th>Pipe em Negociação</th><th>Previsto Hoje</th>
        </tr></thead>
        <tbody>
          ${d.frentes.map(f => linhaFrente(f)).join('')}
          ${linhaFrente(d.total, 'total')}
        </tbody>
      </table>
    </div></div>

    ${alertas}
    <div class="rodape">Atualizado em ${p.atualizado_em} · ${p.du_passados}/${p.du_total} dias úteis</div>`;
}

async function carregarResumo(){
  document.getElementById('conteudo-resumo').innerHTML =
    '<div class="card"><div class="loading"><div class="spinner"></div>Carregando resumo…</div></div>';
  try {
    const q = new URLSearchParams({fmt:'resumo',
      mes: document.getElementById('sel-mes').value,
      ano: document.getElementById('sel-ano').value});
    if (KEY) q.set('k', KEY);
    const data = await (await fetch('/api/index?' + q)).json();
    if (data.erro) throw new Error(data.erro);
    renderResumo(data);
    resumoCarregado = true;
  } catch(e) {
    document.getElementById('conteudo-resumo').innerHTML =
      `<div class="card"><div class="erro">Erro: ${esc(e.message)}</div></div>`;
  }
}

// ── FORECAST DIA A DIA ────────────────────────────────────────
let dadosForecast = null;

function listaNegocios(itens, vazio){
  if (!itens || !itens.length) return `<div class="fc-vazio">${vazio}</div>`;
  const linhas = itens.map(i => `
    <tr>
      <td class="tit"><a href="${i.url}" target="_blank" rel="noopener">${esc(i.titulo)} ↗</a></td>
      <td>${esc(i.dono)}</td>
      <td><span class="fr fr-${i.frente}">${esc(CURTO_FRENTE[i.frente] || i.frente)}</span></td>
      <td class="num">${i.probabilidade == null
          ? '<span class="falta">sem prob.</span>' : i.probabilidade + '%'}</td>
      <td class="num">${R(i.valor)}</td>
      <td class="num"><b>${i.ponderado ? R(i.ponderado) : '<span class="zero">—</span>'}</b></td>
    </tr>`).join('');
  return `
    <table class="fc-det">
      <thead><tr>
        <th>Negócio</th><th>Dono</th><th>Frente</th>
        <th class="num">Prob.</th><th class="num">Valor</th><th class="num">Ponderado</th>
      </tr></thead>
      <tbody>${linhas}</tbody>
    </table>`;
}

function toggleDia(i){
  const tr  = document.getElementById('fc-det-' + i);
  const bt  = document.getElementById('fc-bt-' + i);
  if (!tr) return;
  const abrir = tr.hidden;
  tr.hidden = !abrir;
  if (bt) { bt.textContent = abrir ? '▼' : '▶'; bt.classList.toggle('aberto', abrir); }
  if (abrir && !tr.dataset.pronto) {
    tr.querySelector('td').innerHTML =
      listaNegocios((dadosForecast.linhas[i] || {}).itens, 'Nenhum negócio previsto para este dia.');
    tr.dataset.pronto = '1';
  }
}

function renderForecast(d){
  dadosForecast = d;
  const p = d.periodo, t = d.total;
  const alertas = (d.alertas || []).map(a => `<div class="alerta">⚠ ${esc(a)}</div>`).join('');
  const clsProj = (t.pct_projecao != null && t.pct_projecao >= 100) ? 'pos' : 'neg';

  const cards = `
    <div class="cards">
      <div class="kcard"><div class="rot">Realizado no mês</div>
        <div class="val">${R(t.real_multi)}</div>
        <div class="sub">c/ multiplicador · bruto ${R(t.real_bruto)}</div></div>
      <div class="kcard hoje"><div class="rot">Previsto que falta</div>
        <div class="val">${R(t.previsto_restante)}</div>
        <div class="sub">ponderado de ${R(t.aberto_restante)} em aberto de hoje em diante</div></div>
      <div class="kcard"><div class="rot">Projeção do mês</div>
        <div class="val">${R(t.projecao)}</div>
        <div class="sub">realizado + previsto ponderado</div></div>
      <div class="kcard"><div class="rot">Projeção x Meta</div>
        <div class="val">${P(t.pct_projecao)}</div>
        <div class="sub">meta ${R(t.meta_mes)} · <span class="${clsProj}">${
          t.gap_projecao > 0 ? 'faltam ' + R(t.gap_projecao) : 'sobra ' + R(Math.abs(t.gap_projecao))}</span></div></div>
    </div>`;

  const linhas = d.linhas.map((l, i) => {
    const vazio = l.qtd === 0 && l.ganho_qtd === 0;
    const cls = [l.hoje ? 'e-hoje' : '', l.passado ? 'passado' : '',
                 !l.e_du ? 'fds' : '', vazio ? 'vazia' : ''].filter(Boolean).join(' ');
    return `
      <tr class="${cls}">
        <td class="dia">
          <button class="fc-bt" id="fc-bt-${i}" onclick="toggleDia(${i})"
                  ${l.qtd ? '' : 'disabled'}>▶</button>
          ${fmtDia(l.dia)}${l.hoje ? ' <span class="hoje-tag">hoje</span>' : ''}
          ${l.e_du ? '' : ' <span class="hint">n/ útil</span>'}
        </td>
        <td class="num">${l.qtd ? N(l.qtd) : '<span class="zero">—</span>'}</td>
        <td class="num">${l.p70 ? R(l.p70) : '<span class="zero">—</span>'}</td>
        <td class="num">${l.p50 ? R(l.p50) : '<span class="zero">—</span>'}</td>
        <td class="num">${l.p20 ? R(l.p20) : '<span class="zero">—</span>'}</td>
        <td class="num">${l.em_aberto ? R(l.em_aberto) : '<span class="zero">—</span>'}</td>
        <td class="num prev"><b>${l.previsto ? R(l.previsto) : '<span class="zero">—</span>'}</b></td>
        <td class="num real">${l.ganho_multi
            ? `${R(l.ganho_multi)} <span class="hint">${N(l.ganho_qtd)} venda(s)</span>`
            : '<span class="zero">—</span>'}</td>
      </tr>
      <tr class="fc-det-linha" id="fc-det-${i}" hidden><td colspan="8"></td></tr>`;
  }).join('');

  const totReal = d.linhas.reduce((a,l) => a + l.ganho_multi, 0);
  const totPrev = d.linhas.reduce((a,l) => a + l.previsto, 0);
  const totAb   = d.linhas.reduce((a,l) => a + l.em_aberto, 0);
  const totQtd  = d.linhas.reduce((a,l) => a + l.qtd, 0);

  const semData = d.sem_data.qtd ? `
    <div class="block-title">Sem data prevista — ficam fora do forecast<div class="rule"></div></div>
    <div class="card">
      <div class="fc-aviso">${N(d.sem_data.qtd)} negócio(s) · ${R(d.sem_data.valor)} em aberto
        sem data de fechamento. Enquanto a data não entrar no Pipedrive, esse valor não
        aparece em nenhum dia da tabela acima.</div>
      ${listaNegocios(d.sem_data.itens, '')}
    </div>` : '';

  const foraMes = d.fora_do_mes.qtd ? `
    <div class="block-title">Previstos para fora deste mês<div class="rule"></div></div>
    <div class="card">
      <div class="fc-aviso">${N(d.fora_do_mes.qtd)} negócio(s) · ${R(d.fora_do_mes.valor)}
        (ponderado ${R(d.fora_do_mes.previsto)}) com data prevista fora de
        ${MESES[p.mes-1]}/${p.ano}.</div>
      ${listaNegocios(d.fora_do_mes.itens, '')}
    </div>` : '';

  document.getElementById('conteudo-forecast').innerHTML = `
    ${cards}
    <div class="block-title">Forecast dia a dia — ${MESES[p.mes-1]} ${p.ano}<div class="rule"></div></div>
    <div class="card">
      <div class="table-scroll">
        <table class="fc">
          <thead><tr>
            <th>Dia</th><th class="num">Negócios</th>
            <th class="num">70%</th><th class="num">50%</th><th class="num">20%</th>
            <th class="num">Em aberto</th><th class="num">Previsto</th><th class="num">Entrou</th>
          </tr></thead>
          <tbody>
            ${linhas}
            <tr class="total">
              <td>TOTAL</td><td class="num">${N(totQtd)}</td>
              <td class="num" colspan="3">—</td>
              <td class="num">${R(totAb)}</td>
              <td class="num">${R(totPrev)}</td>
              <td class="num">${R(totReal)}</td>
            </tr>
          </tbody>
        </table>
      </div>
      <div class="fc-nota">
        Clique na setinha de um dia para ver os negócios por trás dele — o título leva
        direto para o negócio no Pipedrive.
        <b>Previsto</b> = 70% × 0,7 + 50% × 0,5 + 20% × 0,2, pela data prevista de fechamento.
        <b>Entrou</b> = o que de fato foi ganho naquele dia, com multiplicador.
        ${p.etapa ? `Só entram negócios na etapa <b>${esc(p.etapa)}</b>.` : ''}
      </div>
    </div>
    ${semData}${foraMes}${alertas}
    <div class="rodape">Atualizado em ${p.atualizado_em}</div>`;
}

async function carregarForecast(){
  document.getElementById('conteudo-forecast').innerHTML =
    '<div class="card"><div class="loading"><div class="spinner"></div>Carregando forecast…</div></div>';
  try {
    const q = new URLSearchParams({fmt:'forecast',
      mes: document.getElementById('sel-mes').value,
      ano: document.getElementById('sel-ano').value});
    if (KEY) q.set('k', KEY);
    const data = await (await fetch('/api/index?' + q)).json();
    if (data.erro) throw new Error(data.erro);
    renderForecast(data);
    forecastCarregado = true;
  } catch(e) {
    document.getElementById('conteudo-forecast').innerHTML =
      `<div class="card"><div class="erro">Erro: ${esc(e.message)}</div></div>`;
  }
}

// ── bootstrap: carrega tudo uma vez ao abrir. Sem refresh automático. ──
const abaInicial = (location.hash || '').replace('#','');
if (ABAS.includes(abaInicial)) setView(abaInicial);
iniciado = true;
atualizarTudo();
</script>
</body>
</html>
"""


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)
