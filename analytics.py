# analytics.py — Dash (Plotly)
# Tabela primeiro (com total), filtros Setor/Tipo/Situação, KPIs, gráficos horizontais,
# upload Excel/Parquet, download Excel, modo escuro por template e limpar filtros.
# Performance: Parquet preferido, normalização única, colunas auxiliares *Cmp (category).

import os
import io
import base64
import pandas as pd
from dash import Dash, html, dcc, dash_table, Input, Output, State
import plotly.express as px
import dash_bootstrap_components as dbc

# ----------------- Config -----------------
EXPECTED_COLS = [
    "Descricao", "Col2", "Col3", "Col4", "Interessado",
    "Nr_Processo", "Abertura", "Tipo", "Setor", "Situacao"
]
LOCAL_XLSX = "rptProcAdm.xlsx"
LOCAL_PARQUET = "rptProcAdm.parquet"

# tema base (CSS) estável; modo escuro só via template nos gráficos
BASE_THEME = dbc.themes.FLATLY

# -------------- Leitura/Limpeza --------------
def clean_excel(xls: pd.ExcelFile) -> pd.DataFrame:
    sheet = "rptProcAdm" if "rptProcAdm" in xls.sheet_names else xls.sheet_names[0]
    df = xls.parse(sheet, skiprows=7)
    if df.empty:
        return pd.DataFrame(columns=["Tipo","Setor","Situacao"])
    df.columns = df.iloc[0]
    df = df.iloc[1:].copy()
    df.columns.name = None

    cols = list(df.columns)
    if len(cols) >= 10:
        cols[:10] = EXPECTED_COLS
        df.columns = cols

    if "Tipo" not in df.columns:
        for c in df.columns:
            if "TIPO" in str(c).upper():
                df.rename(columns={c:"Tipo"}, inplace=True); break
    if "Setor" not in df.columns:
        for c in df.columns:
            if "SETOR" in str(c).upper():
                df.rename(columns={c:"Setor"}, inplace=True); break
    if "Situacao" not in df.columns:
        for c in df.columns:
            if "SITUA" in str(c).upper():
                df.rename(columns={c:"Situacao"}, inplace=True); break

    keep = [c for c in ["Tipo","Setor","Situacao"] if c in df.columns]
    df = df[keep].copy() if keep else pd.DataFrame(columns=["Tipo","Setor","Situacao"])

    if not df.empty:
        for c in keep:
            df[c] = df[c].astype(str).str.strip()
        for c in keep:
            df = df[~df[c].str.lower().isin(["", "nan", "none"])]

    return df

def parse_uploaded(contents: str) -> pd.DataFrame:
    ctype, content_string = contents.split(",")
    raw = base64.b64decode(content_string)
    try:
        if "parquet" in ctype or raw[:4] == b"PAR1":
            import pyarrow  # noqa
            return pd.read_parquet(io.BytesIO(raw), engine="pyarrow")
    except Exception:
        pass
    xls = pd.ExcelFile(io.BytesIO(raw))
    return clean_excel(xls)

def load_local_or_sample() -> pd.DataFrame:
    try:
        if os.path.exists(LOCAL_PARQUET):
            return pd.read_parquet(LOCAL_PARQUET)
    except Exception:
        pass
    try:
        if os.path.exists(LOCAL_XLSX):
            xls_local = pd.ExcelFile(LOCAL_XLSX)
            df_local = clean_excel(xls_local)
            if not df_local.empty:
                try:
                    import pyarrow  # noqa
                    df_local.to_parquet(LOCAL_PARQUET, index=False, engine="pyarrow")
                except Exception:
                    pass
                return df_local
    except Exception:
        pass
    return pd.DataFrame([
        {"Setor":"ARQUIVO SRH","Tipo":"CTC","Situacao":"CONCLUSO"},
        {"Setor":"ARQUIVO SRH","Tipo":"CTC","Situacao":"EM ANÁLISE"},
        {"Setor":"ARQUIVO SRH","Tipo":"FICHA FINANCEIRA","Situacao":"EM ANÁLISE"},
        {"Setor":"ARQUIVO SRH","Tipo":"FÉRIAS PRÊMIO","Situacao":"AGUARDANDO ANÁLISE"},
        {"Setor":"JURÍDICO","Tipo":"PROCESSO JUDICIAL","Situacao":"CONCLUSO"},
    ])

def abbreviate(s: str, maxlen: int = 28) -> str:
    s = str(s)
    return (s[:maxlen-1] + "…") if len(s) > maxlen else s

# ----------------- App/Layout -----------------
app = Dash(__name__, external_stylesheets=[BASE_THEME])
server = app.server

DF_BASE = load_local_or_sample()

actions_bar = dbc.Row(
    [
                dbc.Col(dbc.Button(id="btn-clear", children="🧹 Limpar filtros", color="secondary", outline=True, className="w-100"), md=6, xs=12, className="mb-2"),
    ],
    className="g-2"
)

controls = dbc.Card(
    [
        html.Link(id="theme-css", rel="stylesheet", href=BASE_THEME),
        html.H5("Filtros", className="mb-3"),
        dcc.Store(id="store-data"),
        dcc.Store(id="store-dark", data=False),  # False claro, True escuro

        # Importação (botão)
        html.Div(
            [
                dcc.Upload(
                    id="upload-excel",
                    children=dbc.Button("Importar planilha (.xlsx / .parquet)", color="primary", className="w-100"),
                    accept=".xlsx,.parquet",
                    multiple=False,
                    className="w-100",
                    style={"border":"0","display":"block"},
                ),
                html.Small(id="upload-status", className="text-muted"),
            ],
            className="mb-3",
        ),

        dbc.Label("Setor"),
        dcc.Dropdown(
            id="dd-setor",
            options=([{"label": s, "value": s} for s in sorted(DF_BASE["Setor"].unique())] if not DF_BASE.empty else []),
            value=None, placeholder="Selecione um setor", clearable=True,
        ),

        dbc.Label("Tipo (opcional)", className="mt-3"),
        dcc.Dropdown(id="dd-tipo", multi=True, placeholder="Filtrar por um ou mais tipos"),

        dbc.Label("Situação (opcional)", className="mt-3"),
        dcc.Dropdown(id="dd-situacao", multi=True, placeholder="Filtrar por uma ou mais situações"),

        dbc.Label("Mostrar Top N itens", className="mt-3"),
        dcc.Slider(id="slider-topn", min=5, max=50, step=5, value=15, marks=None, tooltip={"placement":"bottom","always_visible":True}),

        html.Hr(className="my-3"),
        actions_bar,

        html.Hr(className="my-3"),
        dcc.Download(id="download-xlsx"),
    ],
    body=True, className="shadow-sm rounded-4",
)

app.layout = dbc.Container(
    [
        dbc.Row([dbc.Col(html.H2("Análise de Processos Administrativos", className="mt-3 mb-1 fw-bold"))]),
        dbc.Row(
            [
                dbc.Col(controls, md=3),
                dbc.Col(
                    [
                        dbc.Card(dbc.CardBody([
                            html.H6("Tabela (quantidades por Setor, Tipo e Situação)", className="mb-2"),
                            html.Div(id="total-processos", className="fw-bold mb-2 text-primary"),
                            dash_table.DataTable(
    id="tabela",
    columns=[{"name": c, "id": c} for c in ["Setor","Tipo","Situacao","Quantidade"]],
    page_size=12,
    sort_action="native",
    filter_action="native",
    style_table={"overflowX": "auto"},
    style_cell={"padding": "8px"},
    style_header={
        "fontWeight": "700",
        "backgroundColor": "#222",  # cabeçalho mais escuro no modo noturno
        "color": "white"
    },
    style_data_conditional=[
        {"if": {"state": "active"}, "fontWeight": "600"},  # célula ativa em negrito
    ],
),

                        ]), className="shadow-sm rounded-4 mb-3"),
                        html.Div(id="kpis", className="mb-3"),
                        dbc.Row(
                            [
                                dbc.Col(dbc.Card(dbc.CardBody([html.H6("Distribuição por Situação (no setor selecionado)", className="mb-2"), dcc.Graph(id="grafico-situacao", style={"height":"360px"})])), md=4),
                                dbc.Col(dbc.Card(dbc.CardBody([html.H6("Distribuição por Tipo (no setor selecionado)", className="mb-2"), dcc.Graph(id="grafico-tipos", style={"height":"360px"})])), md=4),
                                dbc.Col(dbc.Card(dbc.CardBody([html.H6("Comparativo de Setores (total de processos)", className="mb-2"), dcc.Graph(id="grafico-setores", style={"height":"360px"})])), md=4),
                            ],
                            className="gy-3",
                        ),
                    ],
                    md=9,
                ),
            ],
            className="mt-2",
        ),
    ],
    fluid=True,
)

# ----------------- Callbacks -----------------
# Modo escuro (somente template dos gráficos + rótulo do botão)
@app.callback(
    Output("store-dark","data"),
    Output("btn-theme","children"),
    Input("btn-theme","n_clicks"),
    State("store-dark","data"),
    prevent_initial_call=False,
)
def toggle_dark(n_clicks, dark):
    if n_clicks is None:
        return False, "🌙 Modo Noturno"
    new_dark = not bool(dark)
    return new_dark, ("☀️ Modo Claro" if new_dark else "🌙 Modo Noturno")

# Aplica o CSS do tema claro/escuro na página inteira
@app.callback(
    Output("theme-css", "href"),
    Input("store-dark", "data"),
    prevent_initial_call=False,
)
def apply_theme(dark):
    return (dbc.themes.CYBORG if dark else BASE_THEME)


# Inicializa/Upload + normalização
@app.callback(
    Output("store-data","data"),
    Output("dd-setor","options"),
    Output("upload-status","children"),
    Input("upload-excel","contents"),
    State("upload-excel","filename"),
    prevent_initial_call=False,
)
def init_or_upload(contents, filename):
    if contents is not None:
        df = parse_uploaded(contents)
    else:
        df = DF_BASE.copy()

    if not df.empty:
        for c in ["Setor","Tipo","Situacao"]:
            if c in df.columns:
                df[c] = df[c].astype(str).str.strip()
        if "Setor" in df.columns:    df["SetorCmp"]    = df["Setor"].str.lower()
        if "Tipo" in df.columns:     df["TipoCmp"]     = df["Tipo"].str.lower()
        if "Situacao" in df.columns: df["SituacaoCmp"] = df["Situacao"].str.lower()
        for c in [col for col in ["Setor","Tipo","Situacao","SetorCmp","TipoCmp","SituacaoCmp"] if c in df.columns]:
            df[c] = df[c].astype("category")

    setores = sorted(df["Setor"].unique()) if "Setor" in df.columns and not df.empty else []
    status = (f"Arquivo importado: {filename}" if contents else "Usando base local")
    return df.to_json(date_format="iso", orient="split"), [{"label":s,"value":s} for s in setores], status

# Tipos dependem de Setor
@app.callback(
    Output("dd-tipo","options"),
    Output("dd-tipo","value"),
    Input("dd-setor","value"),
    State("store-data","data"),
)
def update_tipos(setor, data_json):
    df = pd.read_json(data_json, orient="split") if data_json else pd.DataFrame(columns=["Tipo","Setor","SetorCmp","TipoCmp"])
    if df.empty: return [], []
    if "SetorCmp" not in df.columns: df["SetorCmp"] = df["Setor"].astype(str).str.strip().str.lower()
    if "Tipo" not in df.columns:     df["Tipo"]     = df["Tipo"].astype(str).str.strip()
    if setor:
        setor_norm = str(setor).strip().lower()
        tipos = sorted(df.loc[df["SetorCmp"]==setor_norm, "Tipo"].dropna().unique().tolist())
    else:
        tipos = sorted(df["Tipo"].dropna().unique().tolist())
    return [{"label":t,"value":t} for t in tipos], []

# Situação depende de Setor
@app.callback(
    Output("dd-situacao","options"),
    Output("dd-situacao","value"),
    Input("dd-setor","value"),
    State("store-data","data"),
)
def update_situacao(setor, data_json):
    df = pd.read_json(data_json, orient="split") if data_json else pd.DataFrame(columns=["Situacao","Setor","SetorCmp","SituacaoCmp"])
    if df.empty: return [], []
    if "SetorCmp" not in df.columns:   df["SetorCmp"]   = df["Setor"].astype(str).str.strip().str.lower()
    if "Situacao" not in df.columns:   df["Situacao"]   = df["Situacao"].astype(str).str.strip()
    if setor:
        setor_norm = str(setor).strip().lower()
        sits = sorted(df.loc[df["SetorCmp"]==setor_norm, "Situacao"].dropna().unique().tolist())
    else:
        sits = sorted(df["Situacao"].dropna().unique().tolist())
    return [{"label":s,"value":s} for s in sits], []

# Limpar filtros
# Limpar filtros  (usa allow_duplicate nos outputs que também são escritos por outros callbacks)
@app.callback(
    Output("dd-setor", "value"),  # só este callback escreve -> não precisa allow_duplicate
    Output("dd-tipo", "value", allow_duplicate=True),       # também é escrito por update_tipos
    Output("dd-situacao", "value", allow_duplicate=True),   # também é escrito por update_situacao
    Output("slider-topn", "value"),  # só este callback escreve
    Input("btn-clear", "n_clicks"),
    prevent_initial_call=True,
)
def clear_filters(n):
    # reset: setor vazio, tipo/situação listas vazias, topN 15
    return None, [], [], 15


# Gráficos/Tabela/KPIs/Total
@app.callback(
    Output("grafico-situacao","figure"),
    Output("grafico-tipos","figure"),
    Output("grafico-setores","figure"),
    Output("tabela","data"),
    Output("kpis","children"),
    Output("total-processos","children"),
    Input("dd-setor","value"),
    Input("dd-tipo","value"),
    Input("dd-situacao","value"),
    Input("slider-topn","value"),
    State("store-data","data"),
    State("store-dark","data"),
)
def update_views(setor, tipos_sel, sits_sel, topn, data_json, dark):
    tipos_sel = tipos_sel or []
    sits_sel  = sits_sel  or []
    df = pd.read_json(data_json, orient="split") if data_json else pd.DataFrame(columns=["Tipo","Setor","Situacao","SetorCmp","TipoCmp","SituacaoCmp"])

    df_filt = df.copy()
    if "SetorCmp" not in df_filt.columns and "Setor" in df_filt.columns:
        df_filt["SetorCmp"] = df_filt["Setor"].astype(str).str.strip().str.lower()
    if "TipoCmp" not in df_filt.columns and "Tipo" in df_filt.columns:
        df_filt["TipoCmp"] = df_filt["Tipo"].astype(str).str.strip().str.lower()
    if "SituacaoCmp" not in df_filt.columns and "Situacao" in df_filt.columns:
        df_filt["SituacaoCmp"] = df_filt["Situacao"].astype(str).str.strip().str.lower()

    if setor:
        setor_norm = str(setor).strip().lower()   # <<< correção: sem .str
        df_filt = df_filt[df_filt["SetorCmp"] == setor_norm]
    if tipos_sel:
        tipos_norm = [str(t).strip().lower() for t in tipos_sel]
        df_filt = df_filt[df_filt["TipoCmp"].isin(tipos_norm)]
    if sits_sel:
        sits_norm = [str(s).strip().lower() for s in sits_sel]
        df_filt = df_filt[df_filt["SituacaoCmp"].isin(sits_norm)]

    gt = (
        df_filt.groupby(["Setor","Tipo","Situacao"]).size().reset_index(name="Quantidade")
        if not df_filt.empty else
        pd.DataFrame({"Setor":[], "Tipo":[], "Situacao":[], "Quantidade":[]})
    )

    total = int(gt["Quantidade"].sum()) if not gt.empty else 0
    total_text = f"Total de processos: {total}" if total else "Nenhum processo encontrado."

    kpis = []
    if not df_filt.empty and "Situacao" in df_filt.columns:
        gk = df_filt.groupby("Situacao").size().reset_index(name="Qtd").sort_values("Qtd", ascending=False).head(3)
        kpis = dbc.Row([
            dbc.Col(dbc.Card(dbc.CardBody([html.Div(row["Situacao"], className="small text-secondary"), html.H4(int(row["Qtd"]))])), md=4)
            for _, row in gk.iterrows()
        ], className="gy-2")

    template = "plotly_dark" if dark else "plotly"

    gS = gt.groupby("Situacao")["Quantidade"].sum().reset_index() if not gt.empty else pd.DataFrame({"Situacao":[], "Quantidade":[]})
    gS = gS.sort_values("Quantidade", ascending=False).head(topn)
    gS["SitShort"] = gS["Situacao"].apply(lambda x: abbreviate(x, 32))
    figS = px.bar(gS, y="SitShort", x="Quantidade", orientation="h", hover_data={"Situacao":True,"SitShort":False}, template=template)
    figS.update_layout(height=360, margin=dict(l=10,r=10,t=10,b=10), yaxis={"categoryorder":"total ascending"})

    gT = gt.groupby("Tipo")["Quantidade"].sum().reset_index() if not gt.empty else pd.DataFrame({"Tipo":[], "Quantidade":[]})
    gT = gT.sort_values("Quantidade", ascending=False).head(topn)
    gT["TipoShort"] = gT["Tipo"].apply(lambda x: abbreviate(x, 32))
    figT = px.bar(gT, y="TipoShort", x="Quantidade", orientation="h", hover_data={"Tipo":True,"TipoShort":False}, template=template)
    figT.update_layout(height=360, margin=dict(l=10,r=10,t=10,b=10), yaxis={"categoryorder":"total ascending"})

    gZ = df.groupby("Setor").size().reset_index(name="Quantidade") if not df.empty else pd.DataFrame({"Setor":[], "Quantidade":[]})
    gZ = gZ.sort_values("Quantidade", ascending=False).head(topn)
    gZ["SetorShort"] = gZ["Setor"].apply(lambda x: abbreviate(x, 32))
    figZ = px.bar(gZ, y="SetorShort", x="Quantidade", orientation="h", hover_data={"Setor":True,"SetorShort":False}, template=template)
    figZ.update_layout(height=360, margin=dict(l=10,r=10,t=10,b=10), yaxis={"categoryorder":"total ascending"})

    return figS, figT, figZ, gt.to_dict("records"), kpis, total_text

# Download Excel
@app.callback(
    Output("download-xlsx","data"),
    Input("btn-download-xlsx","n_clicks"),
    State("tabela","data"),
    State("dd-setor","value"),
    State("dd-tipo","value"),
    State("dd-situacao","value"),
    prevent_initial_call=True,
)
def do_download_xlsx(n_clicks, rows, setor, tipos, situacoes):
    df = pd.DataFrame(rows)
    with pd.ExcelWriter(io.BytesIO(), engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Tabela")
        if not df.empty:
            df.groupby("Tipo")["Quantidade"].sum().reset_index().sort_values("Quantidade", ascending=False)\
              .to_excel(writer, index=False, sheet_name="Por_Tipo")
            df.groupby("Setor")["Quantidade"].sum().reset_index().sort_values("Quantidade", ascending=False)\
              .to_excel(writer, index=False, sheet_name="Por_Setor")
            df.groupby("Situacao")["Quantidade"].sum().reset_index().sort_values("Quantidade", ascending=False)\
              .to_excel(writer, index=False, sheet_name="Por_Situacao")
            df.groupby(["Setor","Situacao"])["Quantidade"].sum().reset_index()\
              .sort_values(["Setor","Quantidade"], ascending=[True,False])\
              .to_excel(writer, index=False, sheet_name="Setor_Situacao")
        meta = pd.DataFrame({
            "Filtro":["Setor","Tipos selecionados","Situações selecionadas"],
            "Valor":[setor if setor else "—", ", ".join(tipos) if tipos else "—", ", ".join(situacoes) if situacoes else "—"]
        })
        meta.to_excel(writer, index=False, sheet_name="Filtros")
        writer.book.active = writer.book["Tabela"]
        bio = writer._handles.handle; bio.seek(0); data = bio.read()
    return dcc.send_bytes(lambda b: b.write(data), filename="dados_processos.xlsx")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8050))
    app.run(host="127.0.0.1", port=port, debug=False)
