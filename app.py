from flask import Flask, render_template, request, jsonify
import pandas as pd
import plotly.graph_objects as go
import plotly.utils
import json
import os
import math
from datetime import datetime

app = Flask(__name__)

EXCEL_PATH  = "data/dados.xlsx"
SDPO_PATH   = "data/SDPO.xlsx"
PNR_PATH    = "data/PNR.xlsx"
PLANOS_PATH = "data/planos.json"

MESES = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun",
         "Jul", "Ago", "Set", "Out", "Nov", "Dez"]

TRIMESTRES = ["Q1", "Q2", "Q3", "Q4"]

LIMIAR_ALERTA  = 95.0
LIMIAR_CRITICO = 80.0


# ── HELPERS ────────────────────────────────────────────────────────────────────

def safe_num(val):
    if val is None:
        return None
    try:
        v = float(val)
        return None if (math.isnan(v) or math.isinf(v)) else v
    except (TypeError, ValueError):
        return None


def cor_ating(ating):
    if ating is None:
        return "neutral"
    if ating >= LIMIAR_ALERTA:
        return "green"
    elif ating >= LIMIAR_CRITICO:
        return "yellow"
    return "red"


def calcular_cor_borda(ating):
    if ating is None:
        return "#484f58", "neutral"
    if ating >= LIMIAR_ALERTA:
        return "#3fb950", "green"
    elif ating >= LIMIAR_CRITICO:
        return "#d29922", "yellow"
    return "#f85149", "red"


def cor_spo(pilar, pct):
    if pct is None:
        return "neutral"
    if pct < 62:
        return "red"
    p = (pilar or "").strip().lower()
    if "comercial" in p:
        if pct >= 70:
            return "blue"
        if pct >= 65:
            return "yellow"
        return "green"
    if any(x in p for x in [
        "nivel de serviço", "nível de serviço",
        "nivel de servilo", "financeiro", "gente"
    ]):
        if pct >= 73:
            return "blue"
        if pct >= 68:
            return "yellow"
        return "green"
    if any(x in p for x in ["segurança", "seguranca"]):
        if pct >= 80:
            return "blue"
        if pct >= 73:
            return "yellow"
        return "green"
    return "green"


def selo_spo(cor):
    mapa = {
        "red":    "spo_nao_qualificada.png",
        "green":  "spo_qualificada.png",
        "yellow": "spo_certificada.png",
        "blue":   "spo_sustentavel.png",
    }
    return mapa.get(cor, "")


def carregar_planos():
    if not os.path.exists(PLANOS_PATH):
        return []
    try:
        with open(PLANOS_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []


def salvar_planos(planos):
    with open(PLANOS_PATH, "w", encoding="utf-8") as f:
        json.dump(planos, f, ensure_ascii=False, indent=2)


def carregar_dados():
    xl = pd.ExcelFile(EXCEL_PATH)
    setores = {}
    for aba in xl.sheet_names:
        df = xl.parse(aba)
        df.columns = df.columns.str.strip()
        for col in df.columns:
            if col.startswith("Meta_") or col.startswith("Resultado_"):
                df[col] = pd.to_numeric(df[col], errors="coerce")
        setores[aba] = df
    return setores


def ultimo_mes_com_dados(setores):
    ultimo = None
    for mes in MESES:
        col = f"Meta_{mes}"
        for df in setores.values():
            if col not in df.columns:
                continue
            serie = pd.to_numeric(df[col], errors="coerce").dropna()
            if len(serie) > 0:
                ultimo = mes
                break
    return ultimo or MESES[0]


def media_setor(df):
    pcts = []
    for _, row in df.iterrows():
        for m in MESES:
            meta = safe_num(row.get(f"Meta_{m}"))
            res  = safe_num(row.get(f"Resultado_{m}"))
            if meta and meta > 0 and res is not None:
                pcts.append((res / meta) * 100)
    return round(sum(pcts) / len(pcts), 1) if pcts else None


def calcular_ranking(setores):
    ranking = []
    for nome, df in setores.items():
        med = media_setor(df)
        if med is not None:
            _, classe = calcular_cor_borda(med)
            ranking.append({"setor": nome, "media": med, "classe": classe})
    return sorted(ranking, key=lambda x: x["media"], reverse=True)


def detectar_alertas(setores):
    mes_ref = ultimo_mes_com_dados(setores)
    alertas = []
    for setor, df in setores.items():
        for _, row in df.iterrows():
            indicador = str(row.get("Indicador", "")).strip()
            if not indicador:
                continue
            meta = safe_num(row.get(f"Meta_{mes_ref}"))
            res  = safe_num(row.get(f"Resultado_{mes_ref}"))
            if meta is None or meta <= 0 or res is None:
                continue
            ating  = round((res / meta) * 100, 1)
            desvio = round(res - meta, 2)
            if ating >= LIMIAR_ALERTA:
                continue
            nivel = "CRÍTICO" if ating < LIMIAR_CRITICO else "ALERTA"
            alertas.append({
                "setor":       setor,
                "indicador":   indicador,
                "mes":         mes_ref,
                "meta":        meta,
                "resultado":   res,
                "atingimento": ating,
                "desvio":      desvio,
                "nivel":       nivel,
            })
    alertas.sort(key=lambda a: (0 if a["nivel"] == "CRÍTICO" else 1,
                                a["atingimento"]))
    return alertas, mes_ref


# ── GRÁFICO PRINCIPAL (linha) + YTD (barra lateral) ────────────────────────────

def gerar_grafico(df, indicador):
    row = df[df["Indicador"] == indicador].iloc[0]

    metas      = [safe_num(row.get(f"Meta_{m}"))      for m in MESES]
    resultados = [safe_num(row.get(f"Resultado_{m}")) for m in MESES]

    # ── Estocagem (só Comercial) ──────────────────────────────
    estocagem = []
    for i in range(len(MESES)):
        col = "Estocagem" if i == 0 else f"Estocagem.{i}"
        estocagem.append(safe_num(row.get(col)))
    tem_estocagem = any(v is not None for v in estocagem)

    # ── YTD ───────────────────────────────────────────────────
    meta_ytd_acc = 0.0
    real_ytd_acc = 0.0
    tem_ytd      = False
    for meta, res in zip(metas, resultados):
        if meta is not None and res is not None:
            meta_ytd_acc += meta
            real_ytd_acc += res
            tem_ytd       = True

    pct_ytd = round((real_ytd_acc / meta_ytd_acc) * 100, 1) \
        if tem_ytd and meta_ytd_acc > 0 else None

    cor_ytd_hex = (
        "#3fb950" if pct_ytd is not None and pct_ytd >= LIMIAR_ALERTA else
        "#d29922" if pct_ytd is not None and pct_ytd >= LIMIAR_CRITICO else
        "#f85149"
    )

    # ── Atingimento médio ──────────────────────────────────────
    pcts = []
    for meta, res in zip(metas, resultados):
        if meta and meta > 0 and res is not None:
            pcts.append((res / meta) * 100)
    ating_medio = round(sum(pcts) / len(pcts), 1) if pcts else None
    cor_borda, classe_borda = calcular_cor_borda(ating_medio)

    # ── Figura ────────────────────────────────────────────────
    fig = go.Figure()

    fig.add_trace(go.Scatter(
        x=MESES, y=metas,
        mode="lines+markers", name="Meta",
        line=dict(color="#4f8ef7", width=1.5, dash="dash"),
        marker=dict(size=5, color="#4f8ef7"),
        connectgaps=False
    ))

    fig.add_trace(go.Scatter(
        x=MESES, y=resultados,
        mode="lines+markers", name="Resultado",
        line=dict(color="#3fb950", width=1.5),
        marker=dict(size=5, color="#3fb950"),
        connectgaps=False
    ))

    # Anotações de Estocagem abaixo do eixo X
    annotations = []
    if tem_estocagem:
        # Encontra o menor valor no gráfico para posicionar a estocagem abaixo
        todos_y = [v for v in metas + resultados if v is not None]
        y_min   = min(todos_y) if todos_y else 0
        y_range = max(todos_y) - y_min if len(todos_y) > 1 else y_min
        # Ajusta a posição vertical para ficar abaixo do eixo X
        y_pos   = y_min - (y_range * 0.12) if y_range > 0 else -0.22 # Fallback para yref="paper" se y_range for 0

        for mes, val in zip(MESES, estocagem):
            if val is None:
                continue
            cor_val = "#3fb950" if val > 0 else "#f85149"
            texto   = f"{val/1000:.1f}k" if abs(val) >= 1000 \
                      else f"{int(val):,}".replace(",", ".")
            annotations.append(dict(
                x=mes, y=y_pos,
                xref="x", yref="y", # Usar yref="y" para posicionar em relação ao eixo Y
                text=f"<b>{texto}</b>",
                showarrow=False,
                font=dict(size=9, color=cor_val),
                xanchor="center",
                yanchor="top",
            ))

    fig.update_layout(
        paper_bgcolor="#161b22",
        plot_bgcolor="#161b22",
        font=dict(color="#7d8590", size=10,
                  family="Inter, Segoe UI, sans-serif"),
        legend=dict(
            orientation="h", yanchor="bottom", y=1.02,
            xanchor="right", x=1, font=dict(size=10),
            bgcolor="rgba(0,0,0,0)"
        ),
        margin=dict(l=40, r=10, t=30,
                    b=55 if tem_estocagem else 30), # Aumenta margem inferior se tiver estocagem
        xaxis=dict(
            gridcolor="rgba(255,255,255,.06)",
            linecolor="rgba(255,255,255,.06)",
            tickfont=dict(size=9),
        ),
        yaxis=dict(
            gridcolor="rgba(255,255,255,.06)",
            linecolor="rgba(255,255,255,.06)",
            tickfont=dict(size=9),
        ),
        annotations=annotations,
        showlegend=True,
        hovermode="x unified",
    )

    return (
        json.dumps(fig, cls=plotly.utils.PlotlyJSONEncoder),
        ating_medio,
        cor_borda,
        classe_borda,
        pct_ytd,
        round(meta_ytd_acc, 1) if tem_ytd else None,
        round(real_ytd_acc, 1) if tem_ytd else None,
        cor_ytd_hex,
    )


# ── GRÁFICO RANKING ────────────────────────────────────────────────────────────

def gerar_grafico_ranking(setores, top=5):
    todos = []
    for setor, df in setores.items():
        for _, row in df.iterrows():
            pcts = []
            for m in MESES:
                meta = safe_num(row.get(f"Meta_{m}"))
                res  = safe_num(row.get(f"Resultado_{m}"))
                if meta and meta > 0 and res is not None:
                    pcts.append((res / meta) * 100)
            if pcts:
                todos.append({
                    "label": f"{row.get('Indicador','?')} ({setor})",
                    "media": round(sum(pcts) / len(pcts), 1)
                })

    if not todos:
        vazio = go.Figure()
        vazio.update_layout(paper_bgcolor="#161b22",
                            plot_bgcolor="#161b22",
                            font=dict(color="#7d8590"))
        j = json.dumps(vazio, cls=plotly.utils.PlotlyJSONEncoder)
        return j, j

    todos.sort(key=lambda x: x["media"], reverse=True)
    melhores = todos[:top]
    piores   = list(reversed(todos[-top:]))

    def make_bar(items, cor, titulo):
        labels = [i["label"] for i in items]
        values = [i["media"] for i in items]
        fig = go.Figure(go.Bar(
            x=values, y=labels, orientation="h",
            marker=dict(color=cor, opacity=0.85,
                        line=dict(width=0)),
            text=[f"{v}%" for v in values],
            textposition="outside",
            textfont=dict(size=10, color="#e6edf3")
        ))
        fig.update_layout(
            title=dict(text=titulo,
                       font=dict(size=12, color="#e6edf3"), x=0),
            paper_bgcolor="#161b22", plot_bgcolor="#161b22",
            font=dict(color="#7d8590", size=10),
            margin=dict(l=10, r=70, t=40, b=10),
            xaxis=dict(gridcolor="rgba(255,255,255,.06)",
                       range=[0, max(values) * 1.25],
                       ticksuffix="%", tickfont=dict(size=9)),
            yaxis=dict(gridcolor="rgba(255,255,255,.06)",
                       tickfont=dict(size=9), automargin=True),
            showlegend=False
        )
        return json.dumps(fig, cls=plotly.utils.PlotlyJSONEncoder)

    return (
        make_bar(melhores, "#3fb950", "🏆 Top 5 Melhores Resultados"),
        make_bar(piores,   "#f85149", "⚠️ Top 5 Piores Resultados"),
    )


# ── SDPO ───────────────────────────────────────────────────────────────────────

def carregar_sdpo():
    result = {
        "SPO":    {"qs_ativos": [], "rows": []},
        "DPO":    [],
        "pontos": {"Q1": None, "Q2": None, "Q3": None, "Q4": None},
    }
    if not os.path.exists(SDPO_PATH):
        return result
    try:
        xl = pd.ExcelFile(SDPO_PATH)
    except Exception:
        return result

    if "SPO" in xl.sheet_names:
        df = xl.parse("SPO")
        df.columns = [str(c).strip() for c in df.columns]
        MESES_Q = {
            "Q1": ["Jan", "Fev", "Mar"],
            "Q2": ["Abr", "Mai", "Jun"],
            "Q3": ["Jul", "Ago", "Set"],
            "Q4": ["Out", "Nov", "Dez"],
        }
        qs_ativos = []
        for q, meses in MESES_Q.items():
            col_pts = f"{q}_Pts"
            if col_pts in df.columns:
                vals = pd.to_numeric(df[col_pts], errors="coerce").dropna()
                if len(vals) > 0:
                    if q not in qs_ativos:
                        qs_ativos.append(q)
                    continue
            for mes in meses:
                col_m = f"{mes}_Pct"
                if col_m in df.columns:
                    vals = pd.to_numeric(df[col_m], errors="coerce").dropna()
                    if len(vals) > 0:
                        if q not in qs_ativos:
                            qs_ativos.append(q)
                        break

        rows_spo = []
        for _, row in df.iterrows():
            pilar = str(row.get("Pilar", "")).strip()
            if not pilar or pilar.lower() in ["nan", "pilar"]:
                continue
            pts_possivel = safe_num(row.get("Pts_Possivel"))
            le26         = safe_num(row.get("LE26"))
            ating_f_raw  = safe_num(row.get("Atingimento"))
            ating_f      = round(ating_f_raw, 1) if ating_f_raw is not None else None
            cor_f        = cor_spo(pilar, ating_f)

            qs_data = []
            for q in qs_ativos:
                meses     = MESES_Q[q]
                q_pts     = safe_num(row.get(f"{q}_Pts"))
                q_pct_raw = safe_num(row.get(f"{q}_Pct"))
                q_pct     = round(q_pct_raw, 1) if q_pct_raw is not None else None
                q_fechado = (q_pts is not None or q_pct is not None)
                cor_q     = cor_spo(pilar, q_pct)
                meses_data = []
                for mes in meses:
                    pct_raw = safe_num(row.get(f"{mes}_Pct"))
                    pct     = round(pct_raw, 1) if pct_raw is not None else None
                    meses_data.append({
                        "mes": mes,
                        "pct": pct,
                        "cor": cor_spo(pilar, pct),
                    })
                qs_data.append({
                    "q":       q,
                    "pts":     q_pts,
                    "pct":     q_pct,
                    "cor":     cor_q,
                    "fechado": q_fechado,
                    "meses":   meses_data,
                })

            rows_spo.append({
                "pilar":        pilar,
                "pts_possivel": pts_possivel,
                "qs":           qs_data,
                "le26":         le26,
                "ating":        ating_f,
                "cor":          cor_f,
                "selo":         selo_spo(cor_f),
            })

        result["SPO"] = {"qs_ativos": qs_ativos, "rows": rows_spo}

    if "DPO" in xl.sheet_names:
        df_raw = xl.parse("DPO", header=None)
        header_row = None
        for i, row in df_raw.iterrows():
            vals = [str(v).strip().lower() for v in row.values]
            if "pilar" in vals:
                header_row = i
                break
        if header_row is not None:
            df = xl.parse("DPO", header=header_row)
            df.columns = [str(c).strip() for c in df.columns]
            col_map = {}
            for c in df.columns:
                cl = c.lower()
                if "pilar"  in cl: col_map[c] = "Pilar"
                elif "meta" in cl: col_map[c] = "Meta"
                elif "result" in cl: col_map[c] = "Resultado"
            df.rename(columns=col_map, inplace=True)
            for col in ["Meta", "Resultado"]:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors="coerce")
            rows_dpo = []
            for _, row in df.iterrows():
                pilar = str(row.get("Pilar", "")).strip()
                if not pilar or pilar.lower() in ["nan", "pilar"]:
                    continue
                meta  = safe_num(row.get("Meta"))
                res   = safe_num(row.get("Resultado"))
                ating = round((res / meta) * 100, 1) \
                    if meta and meta > 0 and res is not None else None
                rows_dpo.append({
                    "pilar":     pilar,
                    "meta":      meta,
                    "resultado": res,
                    "ating":     ating,
                    "cor":       cor_ating(ating),
                })
            result["DPO"] = rows_dpo

    if "Pontos" in xl.sheet_names:
        try:
            df_p = xl.parse("Pontos")
            df_p.columns = [str(c).strip() for c in df_p.columns]
            for _, row in df_p.iterrows():
                for t in TRIMESTRES:
                    v = safe_num(row.get(t))
                    if v is not None:
                        result["pontos"][t] = v
                break
        except Exception:
            pass

    return result

# ... (todo o seu código app.py antes de carregar_pnr) ...

def carregar_pnr():
    secoes = []
    all_pnr_items = []  # <--- NOVA LISTA PLANA

    if not os.path.exists(PNR_PATH):
        return secoes, all_pnr_items # Retorna a lista plana vazia também
    try:
        xl  = pd.ExcelFile(PNR_PATH)
        aba = xl.sheet_names[0]
        df  = xl.parse(aba, header=None)
    except Exception:
        return secoes, all_pnr_items # Retorna a lista plana vazia também

    # encontra linha do cabeçalho (onde tem "KPI")
    header_row = None
    for i, row in df.iterrows():
        vals = [str(v).strip().upper() for v in row.values if not pd.isna(v)]
        if "KPI" in vals:
            header_row = i
            break
    if header_row is None:
        return secoes, all_pnr_items

    header = [str(v).strip() if not pd.isna(v) else ""
              for v in df.iloc[header_row].values]

    def col_idx(nome):
        nome = nome.upper()
        for j, h in enumerate(header):
            if nome in h.upper():
                return j
        return None

    idx_num     = col_idx("#")
    idx_kpi     = col_idx("KPI")
    idx_pts_tt  = col_idx("PTS TT")
    idx_pts_rev = col_idx("PONTOS REVENDA")
    idx_ytd     = col_idx("YTD")
    idx_label   = (col_idx("YTD") - 1) if col_idx("YTD") is not None else None
    idx_le      = col_idx("LE")

    def cell(row, idx):
        if idx is None or idx >= len(row):
            return None
        v = row.iloc[idx]
        return None if pd.isna(v) else v

    secao_atual = None
    item_atual  = None

    for i in range(header_row + 1, len(df)):
        row = df.iloc[i]

        raw_num     = cell(row, idx_num)
        raw_kpi     = cell(row, idx_kpi)
        raw_pts_tt  = cell(row, idx_pts_tt)
        raw_pts_rev = cell(row, idx_pts_rev)
        raw_label   = cell(row, idx_label)
        raw_ytd     = cell(row, idx_ytd)
        raw_le      = cell(row, idx_le)

        num   = str(raw_num).strip()   if raw_num   is not None else ""
        kpi   = str(raw_kpi).strip()   if raw_kpi   is not None else ""
        label = str(raw_label).strip() if raw_label is not None else ""

        pts_tt  = safe_num(raw_pts_tt)
        pts_rev = safe_num(raw_pts_rev)
        ytd     = safe_num(raw_ytd)
        le25    = safe_num(raw_le)

        num_limpo = num.replace(".", "").replace(" ", "")
        is_numero = num_limpo.isdigit()
        is_total  = num.upper() == "TOTAL"

        # linha de título de seção
        if num and not is_numero and not is_total and kpi == "":
            secao_atual = {"titulo": num, "pts_tt": pts_tt, "itens": []}
            secoes.append(secao_atual)
            item_atual = None
            continue

        # linha TOTAL
        if is_total:
            secoes.append({
                "titulo":        "TOTAL",
                "pts_tt":        pts_tt,
                "total_revenda": pts_rev,
                "itens":         []
            })
            item_atual = None
            continue

        # início de um item KPI
        if is_numero and kpi:
            item_atual = {
                "num":        num,
                "kpi":        kpi,
                "pts_tt":     pts_tt,
                "pontos_rev": pts_rev,
                "meta_ytd":   None,
                "real_ytd":   None,
                "le25":       le25,
                "ating":      None,
                "cor":        "neutral",
                # Adicionado para o gráfico mensal
                "metas_mensais": {},
                "reais_mensais": {},
                "ytd_label_meta": None, # Para o label do gráfico
                "ytd_label_real": None, # Para o label do gráfico
            }
            if secao_atual:
                secao_atual["itens"].append(item_atual)
            all_pnr_items.append(item_atual) # <--- ADICIONA NA LISTA PLANA
            # se a mesma linha já tiver info de META/REAL
            if label:
                if "META" in label.upper():
                    item_atual["meta_ytd"] = ytd
                    item_atual["ytd_label_meta"] = label
                elif "REAL" in label.upper():
                    item_atual["real_ytd"] = ytd
                    item_atual["ytd_label_real"] = label
            continue

        # linhas seguintes, ainda do mesmo item_atual
        if item_atual and label:
            if "META" in label.upper():
                item_atual["meta_ytd"] = ytd
                item_atual["ytd_label_meta"] = label
            elif "REAL" in label.upper():
                item_atual["real_ytd"] = ytd
                item_atual["ytd_label_real"] = label
                meta = item_atual["meta_ytd"]
                res  = ytd
                if meta and meta != 0 and res is not None:
                    item_atual["ating"] = round((res / meta) * 100, 1)
                    item_atual["cor"]   = cor_ating(item_atual["ating"])

    return secoes, all_pnr_items # <--- RETORNA AS DUAS LISTAS


# ── ROTAS ──────────────────────────────────────────────────────────────────────

@app.route("/")
def home():
    setores  = carregar_dados()
    ranking  = calcular_ranking(setores) # Necessário para a sidebar
    return render_template(
        "home.html",
        setores=list(setores.keys()),
        ranking=ranking # Passa ranking para a sidebar
    )


@app.route("/painel/revenda")
def painel_revenda():
    setores  = carregar_dados()
    alertas, mes_ref = detectar_alertas(setores)
    ranking  = calcular_ranking(setores)
    grafico_melhores, grafico_piores = gerar_grafico_ranking(setores, top=5)
    sdpo     = carregar_sdpo()
    pnr_secoes, all_pnr_items = carregar_pnr() # <--- CHAMA A NOVA FUNÇÃO

    return render_template(
        "metas_revenda.html",
        setores=list(setores.keys()), # Para a sidebar
        alertas=alertas,
        mes_ref=mes_ref,
        ranking=ranking,
        grafico_melhores=grafico_melhores,
        grafico_piores=grafico_piores,
        spo=sdpo.get("SPO", {"qs_ativos": [], "rows": []}),
        dpo_data=sdpo.get("DPO", []),
        pontos=sdpo.get("pontos",
                        {"Q1": None, "Q2": None,
                         "Q3": None, "Q4": None}),
        pnr=pnr_secoes, # Passa as seções como antes
        all_pnr_items_json=json.dumps(all_pnr_items) # <--- PASSA PARA O JS
    )


@app.route("/painel/setores")
def painel_setores():
    setores = carregar_dados()
    ranking = calcular_ranking(setores)
    return render_template(
        "metas_setores.html",
        setores=list(setores.keys()), # Para a sidebar
        ranking=ranking,
    )

@app.route("/painel/area")
def painel_area():
    setores = carregar_dados()
    ranking = calcular_ranking(setores) # Para a sidebar
    return render_template(
        "metas_area.html",
        setores=list(setores.keys()),
        ranking=ranking,
    )

@app.route("/painel/individual")
def painel_individual():
    setores = carregar_dados()
    ranking = calcular_ranking(setores) # Para a sidebar
    return render_template(
        "metas_individual.html",
        setores=list(setores.keys()),
        ranking=ranking,
    )

@app.route("/setor/<nome>")
def setor(nome):
    setores = carregar_dados()
    if nome not in setores:
        return "Setor não encontrado", 404

    df          = setores[nome]
    indicadores = df["Indicador"].tolist()
    graficos    = {}
    metadados   = {}
    med_s       = media_setor(df)
    _, cls_s    = calcular_cor_borda(med_s)
    ranking     = calcular_ranking(setores) # Para a sidebar

    for ind in indicadores:
        g_json, ating, cor, classe, pct_ytd, meta_ytd, real_ytd, cor_ytd = \
            gerar_grafico(df, ind)
        graficos[ind]  = g_json
        metadados[ind] = {
            "atingimento": ating,
            "cor":         cor,
            "classe":      classe,
            "pct_ytd":     pct_ytd,
            "meta_ytd":    meta_ytd,
            "real_ytd":    real_ytd,
            "cor_ytd":     cor_ytd,
        }

    return render_template(
        "setor.html",
        setor=nome,
        indicadores=indicadores,
        graficos=graficos,
        metadados=metadados,
        setores=list(setores.keys()), # Para a sidebar
        media_setor=med_s,
        classe_setor=cls_s,
        ranking=ranking, # Passa ranking para a sidebar
    )


@app.route("/api/planos", methods=["GET"])
def listar_planos():
    return jsonify(carregar_planos())


@app.route("/api/planos", methods=["POST"])
def criar_plano():
    data   = request.get_json()
    planos = carregar_planos()
    novo   = {
        "id":          len(planos) + 1,
        "setor":       data.get("setor", ""),
        "indicador":   data.get("indicador", ""),
        "descricao":   data.get("descricao", ""),
        "responsavel": data.get("responsavel", ""),
        "prazo":       data.get("prazo", ""),
        "status":      "Aberto",
        "criado_em":   datetime.now().strftime("%d/%m/%Y %H:%M")
    }
    planos.append(novo)
    salvar_planos(planos)
    return jsonify({"ok": True, "plano": novo})


@app.route("/api/planos/<int:plano_id>", methods=["DELETE"])
def deletar_plano(plano_id):
    planos = [p for p in carregar_planos() if p["id"] != plano_id]
    salvar_planos(planos)
    return jsonify({"ok": True})


# NOVO ENDPOINT PARA O PNR KPI
@app.route("/api/pnr/kpi/<string:kpi_num>")
def api_pnr_kpi(kpi_num):
    """
    Retorna dados YTD (meta x real) de um KPI do PNR.
    """
    _, all_pnr_items = carregar_pnr() # Carrega a lista plana

    kpi_item = next((item for item in all_pnr_items if item["num"] == kpi_num), None)

    if not kpi_item:
        return jsonify({"error": "KPI não encontrado"}), 404

    # Prepara os dados para o Chart.js
    # Por enquanto, vamos usar apenas os valores YTD como um "mês" sintético
    # Se você tiver dados mensais no Excel, podemos adaptar isso depois.
    ytd_labels = ["YTD"]
    ytd_meta   = [kpi_item["meta_ytd"]]
    ytd_real   = [kpi_item["real_ytd"]]

    return jsonify({
        "kpi_num":    kpi_item["num"],
        "kpi_nome":   kpi_item["kpi"],
        "ytd_labels": ytd_labels,
        "ytd_meta":   ytd_meta,
        "ytd_real":   ytd_real,
        "meta_ytd":   kpi_item["meta_ytd"],
        "real_ytd":   kpi_item["real_ytd"],
        "ating":      kpi_item["ating"],
        "ytd_label_meta": kpi_item.get("ytd_label_meta", "Meta"),
        "ytd_label_real": kpi_item.get("ytd_label_real", "Real"),
    })


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)