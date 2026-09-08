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

from flask import Flask, jsonify, request
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


# ── CÁLCULO ───────────────────────────────────────────────────
def bucket_dia(deals_abertos, dia_str):
    """Separa os abertos de um dia nos buckets 20/50/70 e calcula previsto/em aberto."""
    p20 = p50 = p70 = outros = sem_prob = 0.0
    qtd = 0
    for d in deals_abertos:
        if str(d.get("expected_close_date") or "")[:10] != dia_str:
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
    ganhos = buscar_ganhos(mes, ano, pid)
    abertos = buscar_abertos(pid)

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

    b_hoje = bucket_dia(abertos, hoje_str)
    b_ontem = bucket_dia(abertos, ontem_str)
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
    closers.sort(key=lambda x: -x["real_multi"])

    # ── Diagnóstico (aparece no rodapé só se houver ruído) ────
    alertas = []
    if not closers_meta and not META_FIXA:
        alertas.append(
            "Nenhum closer com Subarea = "
            + "/".join(sorted(subareas))
            + " e meta financeira no METAS — a Meta Mês veio zerada.")
    if b_hoje["fora_bucket"] > 0 or b_ontem["fora_bucket"] > 0:
        alertas.append(
            f"R$ {b_hoje['fora_bucket'] + b_ontem['fora_bucket']:,.0f} em deals abertos com "
            "probabilidade fora de 20/50/70 — entram em 'Em Aberto' mas não no 'Previsto'."
            .replace(",", "."))
    if b_hoje["sem_probabilidade"] > 0 or b_ontem["sem_probabilidade"] > 0:
        alertas.append(
            f"R$ {b_hoje['sem_probabilidade'] + b_ontem['sem_probabilidade']:,.0f} em deals abertos "
            "SEM probabilidade preenchida — não entram no 'Previsto'."
            .replace(",", "."))

    return {
        "funil": pipe["nome"],
        "pipeline_id": pid,
        "periodo": {
            "mes": mes, "ano": ano,
            "du_total": du_total, "du_passados": du_pass, "du_restantes": du_rest,
            "hoje": hoje_str, "ontem": ontem_str,
            "atualizado_em": agora_br().strftime("%d/%m/%Y %H:%M"),
        },
        "metricas": metricas,
        "closers": closers,
        "detalhe_hoje": b_hoje,
        "detalhe_ontem": b_ontem,
        "meta_composicao": sorted(closers_meta, key=lambda x: -x["meta"]),
        "alertas": alertas,
    }


# ── AUTH OPCIONAL ─────────────────────────────────────────────
def autorizado():
    if not PAINEL_SENHA:
        return True
    k = request.args.get("k", "") or request.headers.get("X-Painel-Key", "")
    return k == PAINEL_SENHA


# ── ROTAS ─────────────────────────────────────────────────────
def _handler():
    if not API_KEY:
        return jsonify({"erro": "PIPE_API_KEY não configurada nas variáveis de ambiente."}), 500
    if not autorizado():
        return jsonify({"erro": "Não autorizado."}), 401
    try:
        mes = request.args.get("mes", type=int)
        ano = request.args.get("ano", type=int)
        data = calcular_navigator(mes=mes, ano=ano)
        status = 400 if data.get("erro") else 200
        return jsonify(limpar_nans(data)), status
    except req.HTTPError as e:
        return jsonify({"erro": f"Pipedrive respondeu {e.response.status_code}",
                        "detalhe": e.response.text[:400]}), 502
    except Exception as e:
        import traceback
        return jsonify({"erro": str(e), "trace": traceback.format_exc()}), 500


@app.route("/api/navigator")
def api_navigator():
    return _handler()


@app.route("/api/index")          # fallback caso o rewrite da Vercel não preserve o path
def api_index():
    return _handler()


@app.route("/api/health")
def health():
    pipe = buscar_pipeline_id() if API_KEY else {"id": None, "nome": None, "todos": []}
    return jsonify({
        "ok": bool(API_KEY),
        "token_configurado": bool(API_KEY),
        "funil_procurado": FUNIL_NOME,
        "funil_encontrado": pipe.get("nome"),
        "pipeline_id": pipe.get("id"),
        "funis_disponiveis": pipe.get("todos"),
        "senha_ativa": bool(PAINEL_SENHA),
        "hoje_br": hoje_br().isoformat(),
    })


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)
