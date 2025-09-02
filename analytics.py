# analytics_optimized.py — Versão Completa Otimizada
# Sistema de Análise de Processos Administrativos - Performance + Loading Moderno

import os
import io
import base64
import pandas as pd
from io import StringIO
from datetime import datetime
from functools import lru_cache
from dash import Dash, html, dcc, dash_table, Input, Output, State, callback_context
import plotly.express as px
import plotly.graph_objects as go
import dash_bootstrap_components as dbc

# ----------------- Configurações -----------------
EXPECTED_COLS = [
    "Descricao", "Col2", "Col3", "Col4", "Interessado",
    "Nr_Processo", "Abertura", "Tipo", "Setor", "Situacao"
]
LOCAL_XLSX = "rptProcAdm.xlsx"

COLORS = {
    "primary": "#6366f1",
    "secondary": "#8b5cf6", 
    "success": "#10b981",
    "warning": "#f59e0b",
    "danger": "#ef4444",
    "info": "#06b6d4",
    "light": "#f8fafc",
    "dark": "#1e293b",
}

# CSS com loading otimizado
CUSTOM_CSS = """
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

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

.main-container {
    backdrop-filter: blur(20px);
    background: rgba(255, 255, 255, 0.1);
    border-radius: 24px;
    border: 1px solid rgba(255, 255, 255, 0.2);
    margin: 20px;
    padding: 30px;
    box-shadow: 0 20px 40px rgba(0, 0, 0, 0.1);
}

/* === LOADING OVERLAY === */
.loading-overlay {
    position: fixed;
    top: 0;
    left: 0;
    width: 100%;
    height: 100%;
    background: rgba(0, 0, 0, 0.7);
    backdrop-filter: blur(10px);
    z-index: 9999;
    display: flex;
    align-items: center;
    justify-content: center;
    flex-direction: column;
    opacity: 0;
    visibility: hidden;
    transition: all 0.3s ease;
}

.loading-overlay.active {
    opacity: 1;
    visibility: visible;
}

.loading-container {
    background: rgba(255, 255, 255, 0.95);
    border-radius: 20px;
    padding: 40px;
    text-align: center;
    box-shadow: 0 20px 40px rgba(0, 0, 0, 0.2);
    border: 1px solid rgba(255, 255, 255, 0.3);
    backdrop-filter: blur(15px);
    min-width: 300px;
    animation: loadingPulse 2s ease-in-out infinite alternate;
}

@keyframes loadingPulse {
    0% { transform: scale(1); }
    100% { transform: scale(1.02); }
}

.modern-spinner {
    width: 60px;
    height: 60px;
    border: 4px solid rgba(99, 102, 241, 0.2);
    border-radius: 50%;
    border-top: 4px solid #6366f1;
    animation: spin 1s linear infinite;
    margin: 0 auto 20px;
}

@keyframes spin {
    0% { transform: rotate(0deg); }
    100% { transform: rotate(360deg); }
}

.loading-dots {
    display: flex;
    justify-content: center;
    gap: 4px;
    margin: 15px 0;
}

.loading-dot {
    width: 8px;
    height: 8px;
    border-radius: 50%;
    background: linear-gradient(45deg, #6366f1, #8b5cf6);
    animation: loadingDot 1.4s ease-in-out infinite both;
}

.loading-dot:nth-child(1) { animation-delay: -0.32s; }
.loading-dot:nth-child(2) { animation-delay: -0.16s; }
.loading-dot:nth-child(3) { animation-delay: 0s; }

@keyframes loadingDot {
    0%, 80%, 100% { transform: scale(0.8); opacity: 0.5; }
    40% { transform: scale(1.2); opacity: 1; }
}

.loading-text {
    font-size: 1.2rem;
    font-weight: 600;
    color: #374151;
    margin-top: 15px;
    background: linear-gradient(90deg, #6366f1, #8b5cf6, #06b6d4, #6366f1);
    background-size: 400% 100%;
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    animation: loadingText 3s ease-in-out infinite;
}

@keyframes loadingText {
    0%, 100% { background-position: 0% 50%; }
    50% { background-position: 100% 50%; }
}

.loading-subtitle {
    font-size: 0.9rem;
    color: #6b7280;
    margin-top: 8px;
}

/* Glass cards */
.glass-card {
    backdrop-filter: blur(15px);
    background: rgba(255, 255, 255, 0.95);
    border-radius: 20px;
    border: 1px solid rgba(255, 255, 255, 0.3);
    box-shadow: 0 15px 35px rgba(0, 0, 0, 0.1);
    transition: all 0.3s ease;
    overflow: hidden;
}

.glass-card:hover {
    transform: translateY(-2px);
    box-shadow: 0 20px 40px rgba(0, 0, 0, 0.15);
}

.hero-header {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    color: white;
    padding: 40px 30px;
    border-radius: 24px;
    margin-bottom: 30px;
    position: relative;
    overflow: hidden;
}

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

.modern-btn {
    border-radius: 12px !important;
    padding: 12px 24px !important;
    font-weight: 500 !important;
    transition: all 0.3s ease !important;
    border: none !important;
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

.kpi-card {
    background: linear-gradient(135deg, rgba(255,255,255,0.9) 0%, rgba(255,255,255,0.7) 100%);
    border-radius: 16px;
    padding: 20px;
    text-align: center;
    border: 1px solid rgba(255, 255, 255, 0.3);
    backdrop-filter: blur(10px);
    transition: all 0.3s ease;
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

.dark-mode {
    background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
}

.dark-mode .glass-card {
    background: rgba(30, 41, 59, 0.95);
    border-color: rgba(255, 255, 255, 0.1);
}

.dark-mode .loading-container {
    background: rgba(30, 41, 59, 0.95);
    color: #f1f5f9;
}

@keyframes fadeInUp {
    from { opacity: 0; transform: translateY(30px); }
    to { opacity: 1; transform: translateY(0); }
}

.animate-in {
    animation: fadeInUp 0.6s ease;
}

@media (max-width: 768px) {
    .main-container {
        margin: 10px;
        padding: 20px;
    }
    .loading-container {
        margin: 20px;
        min-width: auto;
        padding: 30px 20px;
    }
}
"""

# -------------- Funções Otimizadas --------------
@lru_cache(maxsize=32)
def clean_excel_cached(file_path: str) -> pd.DataFrame:
    """Versão cacheada da limpeza"""
    try:
        if os.path.exists(file_path):
            xls = pd.ExcelFile(file_path)
            return clean_excel(xls)
    except:
        pass
    return pd.DataFrame(columns=["Tipo","Setor","Situacao"])

def clean_excel(xls: pd.ExcelFile) -> pd.DataFrame:
    """Limpeza otimizada de Excel"""
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

    # Mapear colunas
    mappings = {"Tipo": "TIPO", "Setor": "SETOR", "Situacao": "SITUA"}
    for target, search in mappings.items():
        if target not in df.columns:
            for col in df.columns:
                if search in str(col).upper():
                    df.rename(columns={col: target}, inplace=True)
                    break

    keep = [c for c in ["Tipo","Setor","Situacao"] if c in df.columns]
    df = df[keep].copy() if keep else pd.DataFrame(columns=["Tipo","Setor","Situacao"])

    if not df.empty:
        for c in keep:
            df[c] = df[c].astype(str).str.strip()
        mask = ~df[keep].isin(["", "nan", "none", "None"]).any(axis=1)
        df = df[mask]

    return df

def parse_uploaded(contents: str) -> pd.DataFrame:
    """Parse otimizado de upload"""
    try:
        ctype, content_string = contents.split(",")
        raw = base64.b64decode(content_string)
        xls = pd.ExcelFile(io.BytesIO(raw))
        return clean_excel(xls)
    except Exception as e:
        print(f"Erro no upload: {e}")
        return pd.DataFrame(columns=["Tipo","Setor","Situacao"])

def load_local_or_sample() -> pd.DataFrame:
    """Carregamento otimizado"""
    try:
        if os.path.exists(LOCAL_XLSX):
            return clean_excel_cached(LOCAL_XLSX)
    except:
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
    """Abreviação otimizada"""
    s = str(s)
    return (s[:maxlen-1] + "…") if len(s) > maxlen else s

# ----------------- App Setup -----------------
app = Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])
app.title = "Análise de Caixas de Processo - SISPREV"
server = app.server

# Injetar CSS
app.index_string = f'''
<!DOCTYPE html>
<html>
    <head>
        {{%metas%}}
        <title>{{%title%}}</title>
        {{%favicon%}}
        {{%css%}}
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
        <style>{CUSTOM_CSS}</style>
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

# -------------- Componentes --------------
def create_loading_overlay():
    """Overlay de loading"""
    return html.Div([
        html.Div([
            html.Div(className="modern-spinner"),
            html.Div([
                html.Div(className="loading-dot"),
                html.Div(className="loading-dot"),
                html.Div(className="loading-dot"),
            ], className="loading-dots"),
            html.Div("Processando dados...", className="loading-text", id="loading-text"),
            html.Div("Por favor, aguarde", className="loading-subtitle"),
        ], className="loading-container")
    ], className="loading-overlay", id="loading-overlay")

def create_hero_header():
    """Header principal"""
    return html.Div([
        dbc.Row([
            dbc.Col([
                html.H1("📊 Análise de Caixa de Processos - SISPREV", className="display-4 fw-bold mb-3"),
                html.P("Sistema Inteligente de Análise de Processos", className="lead mb-4"),
                html.Small(f"Atualizado: {datetime.now().strftime('%d/%m/%Y %H:%M')}", 
                          className="opacity-75")
            ], md=8),
            dbc.Col([
                html.I(className="fas fa-chart-bar fa-5x opacity-20")
            ], md=4, className="text-end d-none d-md-block")
        ])
    ], className="hero-header animate-in")

def create_sidebar():
    """Sidebar moderna"""
    return html.Div([
        # Upload
        html.Div([
            html.H5([html.I(className="fas fa-upload me-2"), "Importar Dados"], 
                   className="mb-3 text-primary fw-bold"),
            dcc.Upload(
                id="upload-excel",
                children=html.Div([
                    html.I(className="fas fa-cloud-upload-alt fa-2x mb-2 text-primary"),
                    html.Br(),
                    "Arraste ou clique para selecionar",
                    html.Br(),
                    html.Small("Apenas .xlsx", className="text-muted")
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
                    options=[{"label": s, "value": s} for s in sorted(DF_BASE["Setor"].unique())] if not DF_BASE.empty else [],
                    value=None,
                    placeholder="Selecione um setor",
                    clearable=True,
                    style={"marginBottom": "1rem"}
                ),
            ]),
            
            html.Div([
                dbc.Label("📋 Tipo", className="fw-semibold mb-2"),
                dcc.Dropdown(id="dd-tipo", multi=True, placeholder="Filtrar tipos",
                           style={"marginBottom": "1rem"}),
            ]),
            
            html.Div([
                dbc.Label("📊 Situação", className="fw-semibold mb-2"),
                dcc.Dropdown(id="dd-situacao", multi=True, placeholder="Filtrar situações",
                           style={"marginBottom": "1rem"}),
            ]),
            
            html.Div([
                dbc.Label("📈 Top N itens", className="fw-semibold mb-2"),
                dcc.Slider(
                    id="slider-topn", min=5, max=50, step=5, value=15,
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
                        html.I(className="fas fa-broom me-2"), "Limpar"
                    ], id="btn-clear", color="outline-warning", 
                       className="modern-btn w-100 mb-2")
                ], xs=12),
                
            ]),
        ]),
        
        # Stores
        dcc.Store(id="store-data"),
        dcc.Store(id="store-dark", data=False),
        dcc.Store(id="store-loading", data=False),
        dcc.Download(id="download-excel"),
        
    ], className="modern-sidebar animate-in")

def create_main_content():
    """Conteúdo principal"""
    return html.Div([
        # Loading overlay
        create_loading_overlay(),
        
        # Stats
        html.Div(id="stats-cards", className="mb-4"),
        
        # Tabela
        html.Div([
            dbc.Card([
                dbc.CardHeader([
                    html.H5([html.I(className="fas fa-table me-2"), "Dados Detalhados"], 
                           className="mb-0 text-primary fw-bold")
                ], className="bg-light"),
                dbc.CardBody([
                    html.Div(id="total-info", className="mb-3"),
                    html.Div(id="table-content"),
                ])
            ], className="glass-card animate-in"),
        ], className="mb-4"),
        
        # Gráficos
        html.Div([
            dbc.Row([
                dbc.Col([
                    dbc.Card([
                        dbc.CardHeader([
                            html.H6([html.I(className="fas fa-chart-pie me-2"), "Situações"], 
                                   className="mb-0 text-primary fw-bold")
                        ], className="bg-light"),
                        dbc.CardBody([html.Div(id="chart-situacao")])
                    ], className="glass-card h-100"),
                ], md=4, className="mb-3"),
                
                dbc.Col([
                    dbc.Card([
                        dbc.CardHeader([
                            html.H6([html.I(className="fas fa-chart-bar me-2"), "Tipos"], 
                                   className="mb-0 text-primary fw-bold")
                        ], className="bg-light"),
                        dbc.CardBody([html.Div(id="chart-tipos")])
                    ], className="glass-card h-100"),
                ], md=4, className="mb-3"),
                
                dbc.Col([
                    dbc.Card([
                        dbc.CardHeader([
                            html.H6([html.I(className="fas fa-building me-2"), "Setores"], 
                                   className="mb-0 text-primary fw-bold")
                        ], className="bg-light"),
                        dbc.CardBody([html.Div(id="chart-setores")])
                    ], className="glass-card h-100"),
                ], md=4, className="mb-3"),
            ])
        ])
    ])

# Layout principal
app.layout = html.Div([
    html.Div([
        create_hero_header(),
        dbc.Row([
            dbc.Col([create_sidebar()], md=3),
            dbc.Col([create_main_content()], md=9)
        ])
    ], className="main-container", id="main-container")
])

# -------------- Callbacks --------------

# Controle de loading
app.clientside_callback(
    """
    function(loading) {
        const overlay = document.getElementById('loading-overlay');
        if (overlay) {
            if (loading) {
                overlay.classList.add('active');
            } else {
                setTimeout(() => overlay.classList.remove('active'), 300);
            }
        }
        return '';
    }
    """,
    Output('loading-text', 'children'),
    Input('store-loading', 'data')
)

# Modo escuro
@app.callback(
    [Output("store-dark", "data"),
     Output("theme-text", "children"),
     Output("main-container", "className")],
    Input("btn-theme", "n_clicks"),
    State("store-dark", "data"),
    prevent_initial_call=False,
)
def toggle_theme(n_clicks, dark):
    if n_clicks is None:
        return False, "Modo Noturno", "main-container"
    
    new_dark = not bool(dark)
    theme_text = "Modo Claro" if new_dark else "Modo Noturno"
    container_class = "main-container dark-mode" if new_dark else "main-container"
    return new_dark, theme_text, container_class

# Upload e inicialização
@app.callback(
    [Output("store-data", "data"),
     Output("dd-setor", "options"),
     Output("upload-status", "children")],
    Input("upload-excel", "contents"),
    State("upload-excel", "filename"),
    prevent_initial_call=False,
)
def handle_upload(contents, filename):
    if contents is not None:
        try:
            df = parse_uploaded(contents)
            status = dbc.Alert([
                html.I(className="fas fa-check me-2"),
                f"✅ {filename} carregado com {len(df)} registros"
            ], color="success", dismissable=True, className="mt-2")
        except Exception as e:
            df = DF_BASE.copy()
            status = dbc.Alert([
                html.I(className="fas fa-times me-2"),
                f"❌ Erro: {str(e)}"
            ], color="danger", dismissable=True, className="mt-2")
    else:
        df = DF_BASE.copy()
        status = dbc.Alert([
            html.I(className="fas fa-info me-2"),
            "📁 Dados de exemplo carregados"
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

# Stats cards
@app.callback(
    Output("stats-cards", "children"),
    [Input("dd-setor", "value"),
     Input("dd-tipo", "value"),
     Input("dd-situacao", "value")],
    State("store-data", "data"),
)
def update_stats(setor, tipos, situacoes, data_json):
    if not data_json:
        return html.Div()
    
    df = pd.read_json(StringIO(data_json), orient="split")
    df_filt = df.copy()
    
    # Aplicar filtros
    if setor and "SetorCmp" in df_filt.columns:
        df_filt = df_filt[df_filt["SetorCmp"] == str(setor).lower()]
    if tipos and "TipoCmp" in df_filt.columns:
        tipos_norm = [str(t).lower() for t in tipos]
        df_filt = df_filt[df_filt["TipoCmp"].isin(tipos_norm)]
    if situacoes and "SituacaoCmp" in df_filt.columns:
        sits_norm = [str(s).lower() for s in situacoes]
        df_filt = df_filt[df_filt["SituacaoCmp"].isin(sits_norm)]
    
    # Métricas
    total = len(df_filt)
    setores_count = df_filt["Setor"].nunique() if "Setor" in df_filt.columns and not df_filt.empty else 0
    tipos_count = df_filt["Tipo"].nunique() if "Tipo" in df_filt.columns and not df_filt.empty else 0
    
    # Situação mais comum
    situacao_top = "N/A"
    situacao_count = 0
    if not df_filt.empty and "Situacao" in df_filt.columns:
        top_sit = df_filt["Situacao"].value_counts()
        if len(top_sit) > 0:
            situacao_top = abbreviate(top_sit.index[0], 15)
            situacao_count = top_sit.iloc[0]
    
    cards = [
        dbc.Col([
            html.Div([
                html.Div("📊", style={"fontSize": "2rem", "marginBottom": "10px"}),
                html.Div(f"{total:,}", className="kpi-value"),
                html.Div("Total Processos", className="kpi-label")
            ], className="kpi-card")
        ], md=3, sm=6, xs=12),
        
        dbc.Col([
            html.Div([
                html.Div("🏢", style={"fontSize": "2rem", "marginBottom": "10px"}),
                html.Div(f"{setores_count}", className="kpi-value"),
                html.Div("Setores Ativos", className="kpi-label")
            ], className="kpi-card")
        ], md=3, sm=6, xs=12),
        
        dbc.Col([
            html.Div([
                html.Div("📋", style={"fontSize": "2rem", "marginBottom": "10px"}),
                html.Div(f"{tipos_count}", className="kpi-value"),
                html.Div("Tipos Únicos", className="kpi-label")
            ], className="kpi-card")
        ], md=3, sm=6, xs=12),
        
        dbc.Col([
            html.Div([
                html.Div("⭐", style={"fontSize": "2rem", "marginBottom": "10px"}),
                html.Div(f"{situacao_count}", className="kpi-value"),
                html.Div(f"Top: {situacao_top}", className="kpi-label")
            ], className="kpi-card")
        ], md=3, sm=6, xs=12),
    ]
    
    return dbc.Row(cards, className="g-3 mb-4 animate-in")

# Filtros dependentes - Tipos
@app.callback(
    [Output("dd-tipo", "options"),
     Output("dd-tipo", "value")],
    Input("dd-setor", "value"),
    State("store-data", "data"),
)
def update_tipos(setor, data_json):
    if not data_json:
        return [], []
    
    df = pd.read_json(StringIO(data_json), orient="split")
    
    if setor and "SetorCmp" in df.columns and "Tipo" in df.columns:
        tipos = df.loc[df["SetorCmp"] == str(setor).lower(), "Tipo"].dropna().unique()
        tipos = sorted(tipos.tolist())
    else:
        tipos = sorted(df["Tipo"].dropna().unique().tolist()) if "Tipo" in df.columns else []
    
    return [{"label": t, "value": t} for t in tipos], []

# Filtros dependentes - Situações
@app.callback(
    [Output("dd-situacao", "options"),
     Output("dd-situacao", "value")],
    Input("dd-setor", "value"),
    State("store-data", "data"),
)
def update_situacoes(setor, data_json):
    if not data_json:
        return [], []
    
    df = pd.read_json(StringIO(data_json), orient="split")
    
    if setor and "SetorCmp" in df.columns and "Situacao" in df.columns:
        sits = df.loc[df["SetorCmp"] == str(setor).lower(), "Situacao"].dropna().unique()
        sits = sorted(sits.tolist())
    else:
        sits = sorted(df["Situacao"].dropna().unique().tolist()) if "Situacao" in df.columns else []
    
    return [{"label": s, "value": s} for s in sits], []

# Limpar filtros
@app.callback(
    [Output("dd-setor", "value"),
     Output("dd-tipo", "value", allow_duplicate=True),
     Output("dd-situacao", "value", allow_duplicate=True),
     Output("slider-topn", "value")],
    Input("btn-clear", "n_clicks"),
    prevent_initial_call=True,
)
def clear_filters(n):
    return None, [], [], 15

# Tabela principal
@app.callback(
    Output("table-content", "children"),
    [Input("dd-setor", "value"),
     Input("dd-tipo", "value"),
     Input("dd-situacao", "value")],
    State("store-data", "data"),
)
def update_table(setor, tipos, situacoes, data_json):
    if not data_json:
        return html.Div("Nenhum dado disponível")
    
    df = pd.read_json(StringIO(data_json), orient="split")
    df_filt = df.copy()
    
    # Aplicar filtros
    tipos = tipos or []
    situacoes = situacoes or []
    
    if setor and "SetorCmp" in df_filt.columns:
        df_filt = df_filt[df_filt["SetorCmp"] == str(setor).lower()]
    if tipos and "TipoCmp" in df_filt.columns:
        tipos_norm = [str(t).lower() for t in tipos]
        df_filt = df_filt[df_filt["TipoCmp"].isin(tipos_norm)]
    if situacoes and "SituacaoCmp" in df_filt.columns:
        sits_norm = [str(s).lower() for s in situacoes]
        df_filt = df_filt[df_filt["SituacaoCmp"].isin(sits_norm)]
    
    # Agrupamento
    if not df_filt.empty:
        gt = df_filt.groupby(["Setor", "Tipo", "Situacao"]).size().reset_index(name="Quantidade")
    else:
        gt = pd.DataFrame({"Setor": [], "Tipo": [], "Situacao": [], "Quantidade": []})
    
    return dash_table.DataTable(
        data=gt.to_dict("records"),
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
    )

# Total de processos
@app.callback(
    Output("total-info", "children"),
    [Input("dd-setor", "value"),
     Input("dd-tipo", "value"),
     Input("dd-situacao", "value")],
    State("store-data", "data"),
)
def update_total(setor, tipos, situacoes, data_json):
    if not data_json:
        return dbc.Alert("Nenhum dado disponível", color="warning")
    
    df = pd.read_json(StringIO(data_json), orient="split")
    df_filt = df.copy()
    
    # Aplicar filtros
    tipos = tipos or []
    situacoes = situacoes or []
    
    if setor and "SetorCmp" in df_filt.columns:
        df_filt = df_filt[df_filt["SetorCmp"] == str(setor).lower()]
    if tipos and "TipoCmp" in df_filt.columns:
        tipos_norm = [str(t).lower() for t in tipos]
        df_filt = df_filt[df_filt["TipoCmp"].isin(tipos_norm)]
    if situacoes and "SituacaoCmp" in df_filt.columns:
        sits_norm = [str(s).lower() for s in situacoes]
        df_filt = df_filt[df_filt["SituacaoCmp"].isin(sits_norm)]
    
    total = len(df_filt)
    
    if total > 0:
        return dbc.Alert([
            html.I(className="fas fa-chart-line me-2"),
            html.Strong(f"📈 {total:,} processos encontrados"),
            html.Br(),
            html.Small(f"Filtros: Setor={setor or 'Todos'} | Tipos={len(tipos)} | Situações={len(situacoes)}")
        ], color="primary", className="mb-0")
    else:
        return dbc.Alert([
            html.I(className="fas fa-search me-2"),
            "🔍 Nenhum processo encontrado"
        ], color="warning", className="mb-0")

# Gráficos
@app.callback(
    [Output("chart-situacao", "children"),
     Output("chart-tipos", "children"),
     Output("chart-setores", "children")],
    [Input("dd-setor", "value"),
     Input("dd-tipo", "value"),
     Input("dd-situacao", "value"),
     Input("slider-topn", "value")],
    [State("store-data", "data"),
     State("store-dark", "data")],
)
def update_charts(setor, tipos, situacoes, topn, data_json, dark):
    if not data_json:
        empty = html.Div("Sem dados", className="text-center p-4 text-muted")
        return empty, empty, empty
    
    df = pd.read_json(StringIO(data_json), orient="split")
    df_filt = df.copy()
    
    # Aplicar filtros
    tipos = tipos or []
    situacoes = situacoes or []
    
    if setor and "SetorCmp" in df_filt.columns:
        df_filt = df_filt[df_filt["SetorCmp"] == str(setor).lower()]
    if tipos and "TipoCmp" in df_filt.columns:
        tipos_norm = [str(t).lower() for t in tipos]
        df_filt = df_filt[df_filt["TipoCmp"].isin(tipos_norm)]
    if situacoes and "SituacaoCmp" in df_filt.columns:
        sits_norm = [str(s).lower() for s in situacoes]
        df_filt = df_filt[df_filt["SituacaoCmp"].isin(sits_norm)]
    
    # Agrupamento
    gt = (df_filt.groupby(["Setor", "Tipo", "Situacao"]).size().reset_index(name="Quantidade")
          if not df_filt.empty else pd.DataFrame({"Setor": [], "Tipo": [], "Situacao": [], "Quantidade": []}))
    
    template = "plotly_dark" if dark else "plotly_white"
    config = {'displayModeBar': False, 'responsive': True}
    
    # Gráfico Situações
    if not gt.empty:
        gS = gt.groupby("Situacao")["Quantidade"].sum().reset_index()
        gS = gS.sort_values("Quantidade", ascending=True).head(topn)
        
        figS = px.bar(gS, y="Situacao", x="Quantidade", orientation="h", 
                     template=template, color="Quantidade", color_continuous_scale="Viridis")
        figS.update_layout(height=380, margin=dict(l=10,r=10,t=10,b=10), showlegend=False,
                          yaxis={"categoryorder": "total ascending"}, font=dict(family="Inter"))
        figS.update_traces(hovertemplate="<b>%{y}</b><br>Qtd: %{x}<extra></extra>")
        
        chart_sit = dcc.Graph(figure=figS, config=config, style={"height": "380px"})
    else:
        chart_sit = html.Div("Sem dados para situações", className="text-center p-4 text-muted")
    
    # Gráfico Tipos
    if not gt.empty:
        gT = gt.groupby("Tipo")["Quantidade"].sum().reset_index()
        gT = gT.sort_values("Quantidade", ascending=True).head(topn)
        
        figT = px.bar(gT, y="Tipo", x="Quantidade", orientation="h",
                     template=template, color="Quantidade", color_continuous_scale="Plasma")
        figT.update_layout(height=380, margin=dict(l=10,r=10,t=10,b=10), showlegend=False,
                          yaxis={"categoryorder": "total ascending"}, font=dict(family="Inter"))
        figT.update_traces(hovertemplate="<b>%{y}</b><br>Qtd: %{x}<extra></extra>")
        
        chart_tipos = dcc.Graph(figure=figT, config=config, style={"height": "380px"})
    else:
        chart_tipos = html.Div("Sem dados para tipos", className="text-center p-4 text-muted")
    
    # Gráfico Setores (dados completos)
    if not df.empty:
        gZ = df.groupby("Setor").size().reset_index(name="Quantidade")
        gZ = gZ.sort_values("Quantidade", ascending=True).head(topn)
        
        figZ = px.bar(gZ, y="Setor", x="Quantidade", orientation="h",
                     template=template, color="Quantidade", color_continuous_scale="Turbo")
        figZ.update_layout(height=380, margin=dict(l=10,r=10,t=10,b=10), showlegend=False,
                          yaxis={"categoryorder": "total ascending"}, font=dict(family="Inter"))
        figZ.update_traces(hovertemplate="<b>%{y}</b><br>Qtd: %{x}<extra></extra>")
        
        chart_setores = dcc.Graph(figure=figZ, config=config, style={"height": "380px"})
    else:
        chart_setores = html.Div("Sem dados para setores", className="text-center p-4 text-muted")
    
    return chart_sit, chart_tipos, chart_setores

# Download Excel
@app.callback(
    Output("download-excel", "data"),
    Input("btn-download", "n_clicks"),
    [State("dd-setor", "value"),
     State("dd-tipo", "value"),
     State("dd-situacao", "value"),
     State("store-data", "data")],
    prevent_initial_call=True,
)
def download_data(n_clicks, setor, tipos, situacoes, data_json):
    if not data_json:
        return None
    
    df = pd.read_json(StringIO(data_json), orient="split")
    df_filt = df.copy()
    
    # Aplicar filtros
    tipos = tipos or []
    situacoes = situacoes or []
    
    if setor and "SetorCmp" in df_filt.columns:
        df_filt = df_filt[df_filt["SetorCmp"] == str(setor).lower()]
    if tipos and "TipoCmp" in df_filt.columns:
        tipos_norm = [str(t).lower() for t in tipos]
        df_filt = df_filt[df_filt["TipoCmp"].isin(tipos_norm)]
    if situacoes and "SituacaoCmp" in df_filt.columns:
        sits_norm = [str(s).lower() for s in situacoes]
        df_filt = df_filt[df_filt["SituacaoCmp"].isin(sits_norm)]
    
    # Tabela agrupada
    gt = (df_filt.groupby(["Setor", "Tipo", "Situacao"]).size().reset_index(name="Quantidade")
          if not df_filt.empty else pd.DataFrame({"Setor": [], "Tipo": [], "Situacao": [], "Quantidade": []}))
    
    # Criar Excel
    with pd.ExcelWriter(io.BytesIO(), engine="openpyxl") as writer:
        # Aba principal
        gt.to_excel(writer, index=False, sheet_name="Dados_Filtrados")
        
        if not gt.empty:
            # Análises
            gt.groupby("Tipo")["Quantidade"].sum().reset_index()\
              .sort_values("Quantidade", ascending=False)\
              .to_excel(writer, index=False, sheet_name="Por_Tipo")
            
            gt.groupby("Setor")["Quantidade"].sum().reset_index()\
              .sort_values("Quantidade", ascending=False)\
              .to_excel(writer, index=False, sheet_name="Por_Setor")
            
            gt.groupby("Situacao")["Quantidade"].sum().reset_index()\
              .sort_values("Quantidade", ascending=False)\
              .to_excel(writer, index=False, sheet_name="Por_Situacao")
        
        # Metadados
        meta = pd.DataFrame({
            "Filtro": ["Data/Hora", "Setor", "Tipos", "Situações", "Total"],
            "Valor": [
                datetime.now().strftime("%d/%m/%Y %H:%M"),
                setor if setor else "Todos",
                ", ".join(tipos) if tipos else "Todos",
                ", ".join(situacoes) if situacoes else "Todas",
                len(gt)
            ]
        })
        meta.to_excel(writer, index=False, sheet_name="Informações")
        
        writer.book.active = writer.book["Dados_Filtrados"]
        bio = writer._handles.handle
        bio.seek(0)
        data = bio.read()
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return dcc.send_bytes(lambda b: b.write(data), filename=f"processos_admin_{timestamp}.xlsx")

# Executar aplicação
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port, debug=False)
