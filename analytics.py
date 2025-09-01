# analytics_modernized.py — Versão Modernizada com Design Avançado
# Sistema de Análise de Processos Administrativos - UI/UX Moderna

import os
import io
import base64
import pandas as pd
from io import StringIO
from datetime import datetime
from dash import Dash, html, dcc, dash_table, Input, Output, State, callback_context
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import dash_bootstrap_components as dbc

# ----------------- Configurações Avançadas -----------------
EXPECTED_COLS = [
    "Descricao", "Col2", "Col3", "Col4", "Interessado",
    "Nr_Processo", "Abertura", "Tipo", "Setor", "Situacao"
]
LOCAL_XLSX = "rptProcAdm.xlsx"

# Paleta de cores moderna
COLORS = {
    "primary": "#6366f1",      # Indigo vibrante
    "secondary": "#8b5cf6",    # Violeta
    "success": "#10b981",      # Esmeralda
    "warning": "#f59e0b",      # Âmbar
    "danger": "#ef4444",       # Vermelho
    "info": "#06b6d4",         # Ciano
    "light": "#f8fafc",        # Cinza claro
    "dark": "#1e293b",         # Cinza escuro
    "gradient": "linear-gradient(135deg, #667eea 0%, #764ba2 100%)"
}

# Tema customizado com gradientes e glassmorphism
CUSTOM_CSS = """
/* Importa fontes Google */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

/* Reset e base */
* {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
    -webkit-font-smoothing: antialiased;
    -moz-osx-font-smoothing: grayscale;
}

body {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    min-height: 100vh;
    margin: 0;
}

/* Container principal com glassmorphism */
.main-container {
    backdrop-filter: blur(20px);
    background: rgba(255, 255, 255, 0.1);
    border-radius: 24px;
    border: 1px solid rgba(255, 255, 255, 0.2);
    margin: 20px;
    padding: 30px;
    box-shadow: 0 20px 40px rgba(0, 0, 0, 0.1);
}

/* Cards com efeito glass */
.glass-card {
    backdrop-filter: blur(15px);
    background: rgba(255, 255, 255, 0.95);
    border-radius: 20px;
    border: 1px solid rgba(255, 255, 255, 0.3);
    box-shadow: 0 15px 35px rgba(0, 0, 0, 0.1);
    transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    overflow: hidden;
}

.glass-card:hover {
    transform: translateY(-2px);
    box-shadow: 0 20px 40px rgba(0, 0, 0, 0.15);
}

/* Header com gradiente */
.hero-header {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    color: white;
    padding: 40px 30px;
    border-radius: 24px;
    margin-bottom: 30px;
    position: relative;
    overflow: hidden;
}

.hero-header::before {
    content: '';
    position: absolute;
    top: 0;
    right: 0;
    width: 200px;
    height: 200px;
    background: radial-gradient(circle, rgba(255,255,255,0.1) 0%, transparent 70%);
    border-radius: 50%;
    transform: translate(50px, -50px);
}

/* Sidebar moderna */
.modern-sidebar {
    background: rgba(255, 255, 255, 0.98);
    border-radius: 20px;
    border: 1px solid rgba(255, 255, 255, 0.3);
    box-shadow: 0 10px 30px rgba(0, 0, 0, 0.1);
    padding: 25px;
    backdrop-filter: blur(15px);
    position: sticky;
    top: 20px;
}

/* Botões modernos */
.modern-btn {
    border-radius: 12px !important;
    padding: 12px 24px !important;
    font-weight: 500 !important;
    text-transform: none !important;
    transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important;
    border: none !important;
    position: relative !important;
    overflow: hidden !important;
}

.btn-primary-modern {
    background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 100%) !important;
    color: white !important;
    box-shadow: 0 4px 15px rgba(99, 102, 241, 0.3) !important;
}

.btn-primary-modern:hover {
    transform: translateY(-1px) !important;
    box-shadow: 0 8px 25px rgba(99, 102, 241, 0.4) !important;
}

.btn-secondary-modern {
    background: rgba(255, 255, 255, 0.9) !important;
    color: #374151 !important;
    border: 1px solid rgba(0, 0, 0, 0.1) !important;
    backdrop-filter: blur(10px) !important;
}

/* Dropdowns estilizados */
.Select-control {
    border-radius: 12px !important;
    border: 2px solid rgba(99, 102, 241, 0.2) !important;
    transition: all 0.3s ease !important;
}

.Select-control:hover {
    border-color: rgba(99, 102, 241, 0.4) !important;
}

/* KPI Cards animados */
.kpi-card {
    background: linear-gradient(135deg, rgba(255,255,255,0.9) 0%, rgba(255,255,255,0.7) 100%);
    border-radius: 16px;
    padding: 20px;
    text-align: center;
    border: 1px solid rgba(255, 255, 255, 0.3);
    backdrop-filter: blur(10px);
    transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    position: relative;
    overflow: hidden;
}

.kpi-card::before {
    content: '';
    position: absolute;
    top: 0;
    left: 0;
    right: 0;
    height: 4px;
    background: linear-gradient(90deg, #6366f1, #8b5cf6, #06b6d4);
    border-radius: 16px 16px 0 0;
}

.kpi-card:hover {
    transform: translateY(-3px) scale(1.02);
    box-shadow: 0 15px 30px rgba(0, 0, 0, 0.15);
}

.kpi-value {
    font-size: 2.5rem;
    font-weight: 700;
    color: #1e293b;
    margin: 10px 0;
    background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
}

.kpi-label {
    font-size: 0.9rem;
    color: #64748b;
    font-weight: 500;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}

/* Tabela moderna */
.dash-table-container {
    border-radius: 16px;
    overflow: hidden;
    box-shadow: 0 4px 20px rgba(0, 0, 0, 0.1);
}

/* Upload area com drag & drop visual */
.upload-area {
    border: 2px dashed rgba(99, 102, 241, 0.3);
    border-radius: 16px;
    padding: 30px 20px;
    text-align: center;
    transition: all 0.3s ease;
    background: rgba(99, 102, 241, 0.02);
    cursor: pointer;
}

.upload-area:hover {
    border-color: rgba(99, 102, 241, 0.6);
    background: rgba(99, 102, 241, 0.05);
    transform: scale(1.01);
}

/* Slider customizado */
.rc-slider-track {
    background: linear-gradient(90deg, #6366f1, #8b5cf6) !important;
    height: 6px !important;
}

.rc-slider-handle {
    border: 3px solid #6366f1 !important;
    width: 18px !important;
    height: 18px !important;
    margin-top: -6px !important;
    box-shadow: 0 4px 10px rgba(99, 102, 241, 0.3) !important;
}

/* Animações suaves */
@keyframes fadeInUp {
    from {
        opacity: 0;
        transform: translateY(30px);
    }
    to {
        opacity: 1;
        transform: translateY(0);
    }
}

.animate-in {
    animation: fadeInUp 0.6s cubic-bezier(0.4, 0, 0.2, 1);
}

/* Loading spinner moderno */
.loading-spinner {
    width: 40px;
    height: 40px;
    border: 3px solid rgba(99, 102, 241, 0.3);
    border-radius: 50%;
    border-top-color: #6366f1;
    animation: spin 1s ease-in-out infinite;
}

@keyframes spin {
    to { transform: rotate(360deg); }
}

/* Modo escuro melhorado */
.dark-mode {
    background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
}

.dark-mode .glass-card {
    background: rgba(30, 41, 59, 0.95);
    border-color: rgba(255, 255, 255, 0.1);
}

.dark-mode .modern-sidebar {
    background: rgba(30, 41, 59, 0.98);
}

/* Responsive design */
@media (max-width: 768px) {
    .main-container {
        margin: 10px;
        padding: 20px;
        border-radius: 16px;
    }
    
    .hero-header {
        padding: 20px;
        text-align: center;
    }
    
    .modern-sidebar {
        margin-bottom: 20px;
        position: static;
    }
}
"""

# -------------- Funções de Leitura/Limpeza (mantidas) --------------
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
    xls = pd.ExcelFile(io.BytesIO(raw))
    return clean_excel(xls)

def load_local_or_sample() -> pd.DataFrame:
    try:
        if os.path.exists(LOCAL_XLSX):
            xls_local = pd.ExcelFile(LOCAL_XLSX)
            df_local = clean_excel(xls_local)
            if not df_local.empty:
                return df_local
    except Exception:
        pass

    return pd.DataFrame([
        {"Setor":"ARQUIVO SRH","Tipo":"CTC","Situacao":"CONCLUSO"},
        {"Setor":"ARQUIVO SRH","Tipo":"CTC","Situacao":"EM ANÁLISE"},
        {"Setor":"ARQUIVO SRH","Tipo":"FICHA FINANCEIRA","Situacao":"EM ANÁLISE"},
        {"Setor":"ARQUIVO SRH","Tipo":"FÉRIAS PRÊMIO","Situacao":"AGUARDANDO ANÁLISE"},
        {"Setor":"JURÍDICO","Tipo":"PROCESSO JUDICIAL","Situacao":"CONCLUSO"},
        {"Setor":"RECURSOS HUMANOS","Tipo":"LICENÇA MÉDICA","Situacao":"EM ANÁLISE"},
        {"Setor":"RECURSOS HUMANOS","Tipo":"PROGRESSÃO","Situacao":"DEFERIDO"},
        {"Setor":"FINANCEIRO","Tipo":"REEMBOLSO","Situacao":"AGUARDANDO ANÁLISE"},
        {"Setor":"FINANCEIRO","Tipo":"AUXÍLIO","Situacao":"INDEFERIDO"},
        {"Setor":"PROTOCOLO","Tipo":"CERTIDÃO","Situacao":"CONCLUSO"},
    ])

def abbreviate(s: str, maxlen: int = 28) -> str:
    s = str(s)
    return (s[:maxlen-1] + "…") if len(s) > maxlen else s

# ----------------- App Modernizado -----------------
app = Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])
app.title = "Análise de Processos Administrativos"
server = app.server

# Adicionar CSS customizado como asset interno
app.index_string = f'''
<!DOCTYPE html>
<html>
    <head>
        {{%metas%}}
        <title>{{%title%}}</title>
        {{%favicon%}}
        {{%css%}}
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
        <style>
            {CUSTOM_CSS}
        </style>
    </head>
    <body>
        {{%app_entry%}}
        <footer>
            {{%config%}}
            {{%scripts%}}
            {{%renderer%}}
        </footer>
    </body>
</html>
'''

DF_BASE = load_local_or_sample()

# Componentes modernos
def create_hero_header():
    return html.Div(
        className="hero-header animate-in",
        children=[
            dbc.Row([
                dbc.Col([
                    html.H1("📊 Análise de Processos - SISPREV", 
                           className="display-4 fw-bold mb-3",
                           style={"fontSize": "3rem"}),
                    html.P("Sistema Inteligente de Análise de Processos Administrativos", 
                           className="lead mb-4 opacity-90",
                           style={"fontSize": "1.2rem"}),
                    html.Div([
                                          ], className="mb-3"),
                    html.Small(f"Última atualização: {datetime.now().strftime('%d/%m/%Y às %H:%M')}", 
                              className="opacity-75")
                ], md=8),
                dbc.Col([
                    html.Div([
                        html.I(className="fas fa-chart-bar fa-5x opacity-20"),
                    ], className="text-end")
                ], md=4, className="d-none d-md-block")
            ])
        ]
    )

def create_modern_sidebar():
    return html.Div(
        className="modern-sidebar animate-in",
        children=[
            # Seção de Upload
            html.Div([
                html.H5([html.I(className="fas fa-upload me-2"), "Importar Dados"], 
                       className="mb-3 text-primary fw-bold"),
                dcc.Upload(
                    id="upload-excel",
                    children=html.Div([
                        html.I(className="fas fa-cloud-upload-alt fa-2x mb-2 text-primary"),
                        html.Br(),
                        "Arraste arquivos aqui ou ",
                        html.A("clique para selecionar", className="text-primary fw-bold"),
                        html.Br(),
                        html.Small("Apenas arquivos .xlsx", className="text-muted")
                    ], className="upload-area"),
                    accept=".xlsx",
                    multiple=False,
                ),
                html.Div(id="upload-status", className="mt-2"),
            ], className="mb-4"),
            
            html.Hr(),
            
            # Filtros
            html.Div([
                html.H5([html.I(className="fas fa-filter me-2"), "Filtros"], 
                       className="mb-3 text-primary fw-bold"),
                
                html.Div([
                    dbc.Label("🏢 Setor", className="fw-semibold mb-2"),
                    dcc.Dropdown(
                        id="dd-setor",
                        options=([{"label": s, "value": s} for s in sorted(DF_BASE["Setor"].unique())] if not DF_BASE.empty else []),
                        value=None,
                        placeholder="Selecione um setor",
                        clearable=True,
                        style={"marginBottom": "1rem"}
                    ),
                ]),
                
                html.Div([
                    dbc.Label("📋 Tipo (múltipla seleção)", className="fw-semibold mb-2"),
                    dcc.Dropdown(
                        id="dd-tipo", 
                        multi=True, 
                        placeholder="Filtrar por tipos",
                        style={"marginBottom": "1rem"}
                    ),
                ]),
                
                html.Div([
                    dbc.Label("📊 Situação (múltipla seleção)", className="fw-semibold mb-2"),
                    dcc.Dropdown(
                        id="dd-situacao", 
                        multi=True, 
                        placeholder="Filtrar por situações",
                        style={"marginBottom": "1rem"}
                    ),
                ]),
                
                html.Div([
                    dbc.Label("📈 Top N itens", className="fw-semibold mb-2"),
                    dcc.Slider(
                        id="slider-topn",
                        min=5, max=50, step=5, value=15,
                        marks={5: '5', 15: '15', 30: '30', 50: '50'},
                        tooltip={"placement":"bottom", "always_visible":True}
                    ),
                ], className="mb-3"),
            ], className="mb-4"),
            
            html.Hr(),
            
            # Ações
            html.Div([
                html.H5([html.I(className="fas fa-tools me-2"), "Ações"], 
                       className="mb-3 text-primary fw-bold"),
                
                dbc.Row([
                    dbc.Col([
                        dbc.Button([
                            html.I(className="fas fa-moon me-2"),
                            html.Span("Modo Noturno", id="theme-text")
                        ], id="btn-theme", color="outline-secondary", 
                           className="modern-btn w-100 mb-2")
                    ], xs=12),
                    dbc.Col([
                        dbc.Button([
                            html.I(className="fas fa-broom me-2"),
                            "Limpar Filtros"
                        ], id="btn-clear", color="outline-warning", 
                           className="modern-btn w-100 mb-2")
                    ], xs=12),
                    
                        
                        
                ]),
            ]),
            
            # Stores
            dcc.Store(id="store-data"),
            dcc.Store(id="store-dark", data=False),
            dcc.Download(id="download-xlsx"),
        ]
    )

def create_stats_cards():
    return html.Div(id="stats-section", className="mb-4")

def create_main_content():
    return html.Div([
        # Seção de estatísticas
        create_stats_cards(),
        
        # Tabela principal
        html.Div([
            dbc.Card([
                dbc.CardHeader([
                    html.H5([
                        html.I(className="fas fa-table me-2"),
                        "Tabela Detalhada"
                    ], className="mb-0 text-primary fw-bold")
                ], className="bg-light"),
                dbc.CardBody([
                    html.Div(id="total-processos", className="mb-3"),
                    dash_table.DataTable(
                        id="tabela",
                        columns=[{"name": c, "id": c} for c in ["Setor","Tipo","Situacao","Quantidade"]],
                        page_size=15,
                        sort_action="native",
                        filter_action="native",
                        style_table={"overflowX": "auto"},
                        style_cell={
                            "padding": "12px",
                            "fontFamily": "Inter, sans-serif",
                            "fontSize": "14px",
                        },
                        style_header={
                            "fontWeight": "600",
                            "backgroundColor": COLORS["primary"],
                            "color": "white",
                            "textAlign": "left"
                        },
                        style_data_conditional=[
                            {"if": {"state": "active"}, "backgroundColor": "rgba(99, 102, 241, 0.1)"},
                            {"if": {"state": "selected"}, "backgroundColor": "rgba(99, 102, 241, 0.2)"},
                        ],
                    ),
                ])
            ], className="glass-card animate-in"),
        ], className="mb-4"),
        
        # Seção de gráficos
        html.Div([
            dbc.Row([
                dbc.Col([
                    dbc.Card([
                        dbc.CardHeader([
                            html.H6([
                                html.I(className="fas fa-chart-pie me-2"),
                                "Situações"
                            ], className="mb-0 text-primary fw-bold")
                        ], className="bg-light"),
                        dbc.CardBody([
                            dcc.Graph(id="grafico-situacao", style={"height": "400px"})
                        ])
                    ], className="glass-card animate-in h-100"),
                ], md=4, className="mb-3"),
                dbc.Col([
                    dbc.Card([
                        dbc.CardHeader([
                            html.H6([
                                html.I(className="fas fa-chart-bar me-2"),
                                "Tipos"
                            ], className="mb-0 text-primary fw-bold")
                        ], className="bg-light"),
                        dbc.CardBody([
                            dcc.Graph(id="grafico-tipos", style={"height": "400px"})
                        ])
                    ], className="glass-card animate-in h-100"),
                ], md=4, className="mb-3"),
                dbc.Col([
                    dbc.Card([
                        dbc.CardHeader([
                            html.H6([
                                html.I(className="fas fa-building me-2"),
                                "Setores"
                            ], className="mb-0 text-primary fw-bold")
                        ], className="bg-light"),
                        dbc.CardBody([
                            dcc.Graph(id="grafico-setores", style={"height": "400px"})
                        ])
                    ], className="glass-card animate-in h-100"),
                ], md=4, className="mb-3"),
            ])
        ])
    ])

# Layout principal
app.layout = html.Div([
    html.Div([
        create_hero_header(),
        
        dbc.Row([
            dbc.Col([
                create_modern_sidebar()
            ], md=3),
            dbc.Col([
                create_main_content()
            ], md=9)
        ])
    ], className="main-container", id="main-container")
])

# ----------------- Callbacks Modernizados -----------------

# Modo escuro melhorado
@app.callback(
    Output("store-dark", "data"),
    Output("theme-text", "children"),
    Output("main-container", "className"),
    Input("btn-theme", "n_clicks"),
    State("store-dark", "data"),
    prevent_initial_call=False,
)
def toggle_dark_mode(n_clicks, dark):
    if n_clicks is None:
        return False, "Modo Noturno", "main-container"
    
    new_dark = not bool(dark)
    theme_text = "Modo Claro" if new_dark else "Modo Noturno"
    container_class = "main-container dark-mode" if new_dark else "main-container"
    
    return new_dark, theme_text, container_class

# Upload com feedback visual melhorado
@app.callback(
    Output("store-data", "data"),
    Output("dd-setor", "options"),
    Output("upload-status", "children"),
    Input("upload-excel", "contents"),
    State("upload-excel", "filename"),
    prevent_initial_call=False,
)
def handle_upload(contents, filename):
    if contents is not None:
        try:
            df = parse_uploaded(contents)
            status = dbc.Alert([
                html.I(className="fas fa-check-circle me-2"),
                f"✅ Sucesso! Arquivo '{filename}' carregado com {len(df)} registros"
            ], color="success", dismissable=True, className="mt-2")
        except Exception as e:
            df = DF_BASE.copy()
            status = dbc.Alert([
                html.I(className="fas fa-exclamation-triangle me-2"),
                f"❌ Erro ao carregar '{filename}': {str(e)}"
            ], color="danger", dismissable=True, className="mt-2")
    else:
        df = DF_BASE.copy()
        status = dbc.Alert([
            html.I(className="fas fa-info-circle me-2"),
            "📁 Usando dados de exemplo"
        ], color="info", className="mt-2")

    # Normalização
    if not df.empty:
        for c in ["Setor", "Tipo", "Situacao"]:
            if c in df.columns:
                df[c] = df[c].astype(str).str.strip()
                df[f"{c}Cmp"] = df[c].str.lower().astype("category")

    setores = sorted(df["Setor"].unique()) if "Setor" in df.columns and not df.empty else []
    return (
        df.to_json(date_format="iso", orient="split"),
        [{"label": s, "value": s} for s in setores],
        status
    )

# Stats cards com animação
@app.callback(
    Output("stats-section", "children"),
    Input("dd-setor", "value"),
    Input("dd-tipo", "value"),
    Input("dd-situacao", "value"),
    State("store-data", "data"),
)
def update_stats_cards(setor, tipos_sel, sits_sel, data_json):
    if not data_json:
        return html.Div()
    
    df = pd.read_json(StringIO(data_json), orient="split")
    
    # Aplicar filtros
    df_filt = df.copy()
    if setor and "SetorCmp" in df_filt.columns:
        df_filt = df_filt[df_filt["SetorCmp"] == str(setor).lower()]
    if tipos_sel and "TipoCmp" in df_filt.columns:
        df_filt = df_filt[df_filt["TipoCmp"].isin([str(t).lower() for t in tipos_sel])]
    if sits_sel and "SituacaoCmp" in df_filt.columns:
        df_filt = df_filt[df_filt["SituacaoCmp"].isin([str(s).lower() for s in sits_sel])]
    
    total = len(df_filt)
    
    # Top situações
    if not df_filt.empty and "Situacao" in df_filt.columns:
        top_sits = df_filt.groupby("Situacao").size().reset_index(name="Qtd").sort_values("Qtd", ascending=False).head(3)
        setores_count = df_filt["Setor"].nunique() if "Setor" in df_filt.columns else 0
        tipos_count = df_filt["Tipo"].nunique() if "Tipo" in df_filt.columns else 0
        
        cards = [
            # Total de Processos
            dbc.Col([
                html.Div([
                    html.Div("📊", style={"fontSize": "2rem", "marginBottom": "10px"}),
                    html.Div(f"{total:,}", className="kpi-value"),
                    html.Div("Total de Processos", className="kpi-label")
                ], className="kpi-card")
            ], md=3, sm=6, xs=12),
            
            # Setores Únicos
            dbc.Col([
                html.Div([
                    html.Div("🏢", style={"fontSize": "2rem", "marginBottom": "10px"}),
                    html.Div(f"{setores_count}", className="kpi-value"),
                    html.Div("Setores Ativos", className="kpi-label")
                ], className="kpi-card")
            ], md=3, sm=6, xs=12),
            
            # Tipos Únicos
            dbc.Col([
                html.Div([
                    html.Div("📋", style={"fontSize": "2rem", "marginBottom": "10px"}),
                    html.Div(f"{tipos_count}", className="kpi-value"),
                    html.Div("Tipos Diferentes", className="kpi-label")
                ], className="kpi-card")
            ], md=3, sm=6, xs=12),
            
            # Situação Mais Comum
            dbc.Col([
                html.Div([
                    html.Div("⭐", style={"fontSize": "2rem", "marginBottom": "10px"}),
                    html.Div(f"{top_sits.iloc[0]['Qtd'] if not top_sits.empty else 0}", className="kpi-value"),
                    html.Div(f"Mais Comum: {abbreviate(top_sits.iloc[0]['Situacao'], 15) if not top_sits.empty else 'N/A'}", className="kpi-label")
                ], className="kpi-card")
            ], md=3, sm=6, xs=12),
        ]
    else:
        cards = [
            dbc.Col([
                html.Div([
                    html.Div("📊", style={"fontSize": "2rem", "marginBottom": "10px"}),
                    html.Div("0", className="kpi-value"),
                    html.Div("Nenhum Dado", className="kpi-label")
                ], className="kpi-card")
            ], md=12)
        ]
    
    return dbc.Row(cards, className="g-3 mb-4 animate-in")

# Filtros dependentes melhorados
@app.callback(
    Output("dd-tipo", "options"),
    Output("dd-tipo", "value"),
    Input("dd-setor", "value"),
    State("store-data", "data"),
)
def update_tipos(setor, data_json):
    if not data_json:
        return [], []
    
    df = pd.read_json(StringIO(data_json), orient="split")
    
    if setor and "SetorCmp" in df.columns:
        tipos = sorted(df.loc[df["SetorCmp"] == str(setor).lower(), "Tipo"].dropna().unique())
    else:
        tipos = sorted(df["Tipo"].dropna().unique()) if "Tipo" in df.columns else []
    
    return [{"label": t, "value": t} for t in tipos], []

@app.callback(
    Output("dd-situacao", "options"),
    Output("dd-situacao", "value"),
    Input("dd-setor", "value"),
    State("store-data", "data"),
)
def update_situacoes(setor, data_json):
    if not data_json:
        return [], []
    
    df = pd.read_json(StringIO(data_json), orient="split")
    
    if setor and "SetorCmp" in df.columns:
        sits = sorted(df.loc[df["SetorCmp"] == str(setor).lower(), "Situacao"].dropna().unique())
    else:
        sits = sorted(df["Situacao"].dropna().unique()) if "Situacao" in df.columns else []
    
    return [{"label": s, "value": s} for s in sits], []

# Limpar filtros
@app.callback(
    Output("dd-setor", "value"),
    Output("dd-tipo", "value", allow_duplicate=True),
    Output("dd-situacao", "value", allow_duplicate=True),
    Output("slider-topn", "value"),
    Input("btn-clear", "n_clicks"),
    prevent_initial_call=True,
)
def clear_filters(n):
    return None, [], [], 15

# Callback principal com gráficos modernizados
@app.callback(
    Output("grafico-situacao", "figure"),
    Output("grafico-tipos", "figure"),
    Output("grafico-setores", "figure"),
    Output("tabela", "data"),
    Output("total-processos", "children"),
    Input("dd-setor", "value"),
    Input("dd-tipo", "value"),
    Input("dd-situacao", "value"),
    Input("slider-topn", "value"),
    State("store-data", "data"),
    State("store-dark", "data"),
)
def update_main_content(setor, tipos_sel, sits_sel, topn, data_json, dark):
    if not data_json:
        empty_fig = go.Figure()
        empty_fig.add_annotation(
            text="Nenhum dado disponível",
            xref="paper", yref="paper",
            x=0.5, y=0.5, showarrow=False,
            font=dict(size=16, color="gray")
        )
        return empty_fig, empty_fig, empty_fig, [], "Nenhum dado disponível"
    
    df = pd.read_json(StringIO(data_json), orient="split")
    tipos_sel = tipos_sel or []
    sits_sel = sits_sel or []
    
    # Filtrar dados
    df_filt = df.copy()
    if setor and "SetorCmp" in df_filt.columns:
        df_filt = df_filt[df_filt["SetorCmp"] == str(setor).lower()]
    if tipos_sel and "TipoCmp" in df_filt.columns:
        df_filt = df_filt[df_filt["TipoCmp"].isin([str(t).lower() for t in tipos_sel])]
    if sits_sel and "SituacaoCmp" in df_filt.columns:
        df_filt = df_filt[df_filt["SituacaoCmp"].isin([str(s).lower() for s in sits_sel])]
    
    # Tabela agrupada
    gt = (
        df_filt.groupby(["Setor", "Tipo", "Situacao"]).size().reset_index(name="Quantidade")
        if not df_filt.empty else
        pd.DataFrame({"Setor": [], "Tipo": [], "Situacao": [], "Quantidade": []})
    )
    
    total = int(gt["Quantidade"].sum()) if not gt.empty else 0
    
    # Status com ícones
    if total > 0:
        total_display = dbc.Alert([
            html.I(className="fas fa-chart-line me-2"),
            html.Strong(f"📈 {total:,} processos encontrados"),
            html.Br(),
            html.Small(f"Filtros aplicados: Setor={setor or 'Todos'}, Tipos={len(tipos_sel)}, Situações={len(sits_sel)}")
        ], color="primary", className="mb-0")
    else:
        total_display = dbc.Alert([
            html.I(className="fas fa-search me-2"),
            "🔍 Nenhum processo encontrado com os filtros aplicados"
        ], color="warning", className="mb-0")
    
    # Configuração de tema para gráficos
    template = "plotly_dark" if dark else "plotly_white"
    color_palette = px.colors.qualitative.Set3
    
    # Gráfico de Situações - Melhorado
    gS = gt.groupby("Situacao")["Quantidade"].sum().reset_index() if not gt.empty else pd.DataFrame({"Situacao": [], "Quantidade": []})
    gS = gS.sort_values("Quantidade", ascending=True).head(topn)
    
    figS = px.bar(
        gS, 
        y="Situacao", 
        x="Quantidade", 
        orientation="h",
        template=template,
        color="Quantidade",
        color_continuous_scale="Viridis",
        title="Distribuição por Situação"
    )
    figS.update_layout(
        height=400,
        margin=dict(l=10, r=10, t=40, b=10),
        yaxis={"categoryorder": "total ascending"},
        showlegend=False,
        font=dict(family="Inter, sans-serif"),
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
    )
    figS.update_traces(
        hovertemplate="<b>%{y}</b><br>Quantidade: %{x}<extra></extra>",
        marker_line_color="white",
        marker_line_width=1
    )
    
    # Gráfico de Tipos - Melhorado
    gT = gt.groupby("Tipo")["Quantidade"].sum().reset_index() if not gt.empty else pd.DataFrame({"Tipo": [], "Quantidade": []})
    gT = gT.sort_values("Quantidade", ascending=True).head(topn)
    
    figT = px.bar(
        gT,
        y="Tipo",
        x="Quantidade",
        orientation="h",
        template=template,
        color="Quantidade",
        color_continuous_scale="Plasma",
        title="Distribuição por Tipo"
    )
    figT.update_layout(
        height=400,
        margin=dict(l=10, r=10, t=40, b=10),
        yaxis={"categoryorder": "total ascending"},
        showlegend=False,
        font=dict(family="Inter, sans-serif"),
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
    )
    figT.update_traces(
        hovertemplate="<b>%{y}</b><br>Quantidade: %{x}<extra></extra>",
        marker_line_color="white",
        marker_line_width=1
    )
    
    # Gráfico de Setores - Melhorado (usando dados completos)
    gZ = df.groupby("Setor").size().reset_index(name="Quantidade") if not df.empty else pd.DataFrame({"Setor": [], "Quantidade": []})
    gZ = gZ.sort_values("Quantidade", ascending=True).head(topn)
    
    figZ = px.bar(
        gZ,
        y="Setor",
        x="Quantidade",
        orientation="h",
        template=template,
        color="Quantidade",
        color_continuous_scale="Turbo",
        title="Comparativo de Setores (Total)"
    )
    figZ.update_layout(
        height=400,
        margin=dict(l=10, r=10, t=40, b=10),
        yaxis={"categoryorder": "total ascending"},
        showlegend=False,
        font=dict(family="Inter, sans-serif"),
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
    )
    figZ.update_traces(
        hovertemplate="<b>%{y}</b><br>Quantidade: %{x}<extra></extra>",
        marker_line_color="white",
        marker_line_width=1
    )
    
    return figS, figT, figZ, gt.to_dict("records"), total_display

# Download melhorado
@app.callback(
    Output("download-xlsx", "data"),
    Input("btn-download-xlsx", "n_clicks"),
    State("tabela", "data"),
    State("dd-setor", "value"),
    State("dd-tipo", "value"),
    State("dd-situacao", "value"),
    prevent_initial_call=True,
)
def download_excel(n_clicks, rows, setor, tipos, situacoes):
    df = pd.DataFrame(rows)
    
    with pd.ExcelWriter(io.BytesIO(), engine="openpyxl") as writer:
        # Aba principal
        df.to_excel(writer, index=False, sheet_name="Dados_Filtrados")
        
        if not df.empty:
            # Análises por categoria
            df.groupby("Tipo")["Quantidade"].sum().reset_index()\
              .sort_values("Quantidade", ascending=False)\
              .to_excel(writer, index=False, sheet_name="Analise_Tipos")
            
            df.groupby("Setor")["Quantidade"].sum().reset_index()\
              .sort_values("Quantidade", ascending=False)\
              .to_excel(writer, index=False, sheet_name="Analise_Setores")
            
            df.groupby("Situacao")["Quantidade"].sum().reset_index()\
              .sort_values("Quantidade", ascending=False)\
              .to_excel(writer, index=False, sheet_name="Analise_Situacoes")
            
            # Análise cruzada
            pivot = df.pivot_table(
                values="Quantidade", 
                index="Setor", 
                columns="Situacao", 
                fill_value=0, 
                aggfunc="sum"
            )
            pivot.to_excel(writer, sheet_name="Matriz_Setor_Situacao")
        
        # Metadados do relatório
        meta = pd.DataFrame({
            "Parâmetro": [
                "Data/Hora da Exportação",
                "Setor Filtrado",
                "Tipos Selecionados", 
                "Situações Selecionadas",
                "Total de Registros"
            ],
            "Valor": [
                datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                setor if setor else "Todos os setores",
                ", ".join(tipos) if tipos else "Todos os tipos",
                ", ".join(situacoes) if situacoes else "Todas as situações",
                len(df)
            ]
        })
        meta.to_excel(writer, index=False, sheet_name="Metadados")
        
        # Define a primeira aba como ativa
        writer.book.active = writer.book["Dados_Filtrados"]
        
        # Obtém os dados do arquivo
        bio = writer._handles.handle
        bio.seek(0)
        data = bio.read()
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"relatorio_processos_{timestamp}.xlsx"
    
    return dcc.send_bytes(lambda b: b.write(data), filename=filename)

# Executar aplicação
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port, debug=False)
