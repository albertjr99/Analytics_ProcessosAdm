# analytics.py — Dash (Plotly)
# Painel com Tabela (primeiro, incluindo total), gráficos compactos e download Excel

import os
import io
import base64
import pandas as pd
from dash import Dash, html, dcc, dash_table, Input, Output, State
import plotly.express as px
import dash_bootstrap_components as dbc

EXPECTED_COLS = ['Descricao','Col2','Col3','Col4','Interessado','Nr_Processo','Abertura','Tipo','Setor','Situacao']

# -----------------------------
# Leitura e limpeza do Excel
# -----------------------------

def clean_excel(xls: pd.ExcelFile) -> pd.DataFrame:
    sheet = 'rptProcAdm' if 'rptProcAdm' in xls.sheet_names else xls.sheet_names[0]
    df = xls.parse(sheet, skiprows=7)
    if df.empty:
        return pd.DataFrame(columns=['Tipo', 'Setor'])
    df.columns = df.iloc[0]
    df = df.iloc[1:].copy()
    df.columns.name = None
    cols = list(df.columns)
    if len(cols) >= 10:
        cols[:10] = EXPECTED_COLS
        df.columns = cols
    if 'Tipo' not in df.columns:
        for c in df.columns:
            if 'TIPO' in str(c).upper():
                df.rename(columns={c: 'Tipo'}, inplace=True)
                break
    if 'Setor' not in df.columns:
        for c in df.columns:
            if 'SETOR' in str(c).upper():
                df.rename(columns={c: 'Setor'}, inplace=True)
                break
    keep = [c for c in ['Tipo', 'Setor'] if c in df.columns]
    df = df[keep].copy() if keep else pd.DataFrame(columns=['Tipo', 'Setor'])
    if not df.empty:
        df['Tipo'] = df['Tipo'].astype(str).str.strip()
        df['Setor'] = df['Setor'].astype(str).str.strip()
        df = df[(df['Tipo'] != '') & (df['Setor'] != '')]
        df = df[(df['Tipo'].str.lower() != 'nan') & (df['Setor'].str.lower() != 'nan')]
    return df


def parse_uploaded(contents: str) -> pd.DataFrame:
    _, content_string = contents.split(',')
    decoded = base64.b64decode(content_string)
    xls = pd.ExcelFile(io.BytesIO(decoded))
    return clean_excel(xls)


def load_local_or_sample() -> pd.DataFrame:
    try:
        xls_local = pd.ExcelFile('rptProcAdm.xlsx')
        df_local = clean_excel(xls_local)
        if not df_local.empty:
            return df_local
    except Exception:
        pass
    # Dados de exemplo (para não quebrar a UI)
    return pd.DataFrame([
        {'Setor': 'CPAD - ARQUIVO GERAL IPAJM', 'Tipo': 'CERTIDÃO DE TEMPO DE CONTRIBUIÇÃO - CTC'},
        {'Setor': 'GERENCIA DE FINANCAS', 'Tipo': 'FICHA FINANCEIRA'},
        {'Setor': 'SUBGERENCIA DE RECURSOS HUMANOS', 'Tipo': 'FÉRIAS PRÊMIO'},
    ])


def abbreviate(s: str, maxlen: int = 28) -> str:
    return (str(s)[: maxlen - 1] + '…') if len(str(s)) > maxlen else str(s)

# -----------------------------
# App e Layout
# -----------------------------
app = Dash(__name__, external_stylesheets=[dbc.themes.FLATLY])
server = app.server

DF_BASE = load_local_or_sample()

controls = dbc.Card(
    [
        html.H5('Filtros', className='mb-3'),
        dcc.Store(id='store-data'),
        dcc.Upload(
            id='upload-excel',
            children=html.Div(['Arraste ou ', html.A('clique para enviar'), ' seu Excel (.xlsx)']),
            accept='.xlsx',
            multiple=False,
            className='mb-3',
        ),
        dbc.Label('Setor'),
        dcc.Dropdown(
            id='dd-setor',
            options=[{'label': s, 'value': s} for s in sorted(DF_BASE['Setor'].unique())],
            value=None,
            placeholder='Selecione um setor',
            clearable=True,
        ),
        dbc.Label('Tipo (opcional)', className='mt-3'),
        dcc.Dropdown(id='dd-tipo', multi=True, placeholder='Filtrar por um ou mais tipos'),
        dbc.Label('Mostrar Top N itens', className='mt-3'),
        dcc.Slider(id='slider-topn', min=5, max=50, step=5, value=20, marks=None, tooltip={"placement": "bottom", "always_visible": True}),
        dcc.Download(id='download-xlsx'),
    ],
    body=True,
    className='shadow-sm rounded-4',
)

app.layout = dbc.Container(
    [
        dbc.Row([dbc.Col(html.H2('Análise da Quantidade de Processos Administrativos - SISPREV', className='mt-3 mb-1 fw-bold'))]),
        dbc.Row([
            dbc.Col(controls, md=3),
            dbc.Col([
                # Tabela primeiro
                dbc.Card(dbc.CardBody([
                    html.H6('Tabela (quantidades por Setor e Tipo)', className='mb-2'),
                    html.Div(id='total-processos', className='fw-bold mb-2 text-primary'),
                    dash_table.DataTable(
                        id='tabela',
                        columns=[{'name': c, 'id': c} for c in ['Setor', 'Tipo', 'Quantidade']],
                        page_size=12,
                        sort_action='native',
                        filter_action='native',
                        style_table={'overflowX': 'auto'},
                        style_cell={'padding': '8px'},
                        style_header={'fontWeight': '700'},
                    ),
                ]), className='shadow-sm rounded-4 mb-3'),
                # Gráficos
                dbc.Row([
                    dbc.Col(dbc.Card(dbc.CardBody([
                        html.H6('Distribuição por Tipo (no setor selecionado)', className='mb-2'),
                        dcc.Graph(id='grafico-tipos', style={'height': '360px'}),
                    ]), className='shadow-sm rounded-4'), md=6),
                    dbc.Col(dbc.Card(dbc.CardBody([
                        html.H6('Comparativo de Setores (total de processos)', className='mb-2'),
                        dcc.Graph(id='grafico-setores', style={'height': '360px'}),
                    ]), className='shadow-sm rounded-4'), md=6),
                ], className='gy-3'),
            ], md=9),
        ], className='mt-2'),
    ],
    fluid=True,
)

# -----------------------------
# Callbacks
# -----------------------------
@app.callback(
    Output('store-data', 'data'),
    Output('dd-setor', 'options'),
    Input('upload-excel', 'contents'),
    prevent_initial_call=False,
)
def init_or_upload(contents):
    if contents is not None:
        df = parse_uploaded(contents)
    else:
        df = DF_BASE.copy()
    setores = sorted(df['Setor'].unique()) if not df.empty else []
    return df.to_json(date_format='iso', orient='split'), [{'label': s, 'value': s} for s in setores]

@app.callback(
    Output('dd-tipo', 'options'),
    Output('dd-tipo', 'value'),
    Input('dd-setor', 'value'),
    State('store-data', 'data'),
)
def update_tipos(setor, data_json):
    df = pd.read_json(data_json, orient='split') if data_json else pd.DataFrame(columns=['Tipo', 'Setor'])
    if 'Setor' in df.columns:
        df['Setor'] = df['Setor'].astype(str).str.strip()
    if 'Tipo' in df.columns:
        df['Tipo'] = df['Tipo'].astype(str).str.strip()
    if setor:
        setor_norm = str(setor).strip().lower()
        tipos = sorted(df.loc[df['Setor'].str.strip().str.lower() == setor_norm, 'Tipo'].dropna().unique().tolist())
    else:
        tipos = sorted(df['Tipo'].dropna().unique().tolist())
    return [{'label': t, 'value': t} for t in tipos], []

@app.callback(
    Output('grafico-tipos', 'figure'),
    Output('grafico-setores', 'figure'),
    Output('tabela', 'data'),
    Output('total-processos', 'children'),
    Input('dd-setor', 'value'),
    Input('dd-tipo', 'value'),
    Input('slider-topn', 'value'),
    State('store-data', 'data'),
)
def update_views(setor, tipos_sel, topn, data_json):
    if tipos_sel is None:
        tipos_sel = []
    df = pd.read_json(data_json, orient='split') if data_json else pd.DataFrame(columns=['Tipo', 'Setor'])

    df_filt = df.copy()
    for col in ['Setor', 'Tipo']:
        if col in df_filt.columns:
            df_filt[col] = df_filt[col].astype(str).str.strip()
    if setor:
        setor_norm = str(setor).strip().lower()
        df_filt = df_filt[df_filt['Setor'].str.strip().str.lower() == setor_norm]
    if tipos_sel:
        tipos_norm = [str(t).strip().lower() for t in tipos_sel]
        df_filt = df_filt[df_filt['Tipo'].str.strip().str.lower().isin(tipos_norm)]

    gt = (
        df_filt.groupby(['Setor', 'Tipo']).size().reset_index(name='Quantidade')
        if not df_filt.empty
        else pd.DataFrame({'Setor': [], 'Tipo': [], 'Quantidade': []})
    )

    total = int(gt['Quantidade'].sum()) if not gt.empty else 0
    total_text = f"Total de processos: {total}" if total else "Nenhum processo encontrado."

    g1 = gt.groupby('Tipo')['Quantidade'].sum().reset_index() if not gt.empty else pd.DataFrame({'Tipo': [], 'Quantidade': []})
    g1 = g1.sort_values('Quantidade', ascending=False).head(topn)
    g1['TipoShort'] = g1['Tipo'].apply(lambda x: abbreviate(x, 32))
    fig1 = px.bar(g1, y='TipoShort', x='Quantidade', orientation='h', hover_data={'Tipo': True, 'TipoShort': False})
    fig1.update_layout(height=360, margin=dict(l=10, r=10, t=10, b=10), yaxis={'categoryorder':'total ascending'})

    g2 = df.groupby('Setor').size().reset_index(name='Quantidade') if not df.empty else pd.DataFrame({'Setor': [], 'Quantidade': []})
    g2 = g2.sort_values('Quantidade', ascending=False).head(topn)
    g2['SetorShort'] = g2['Setor'].apply(lambda x: abbreviate(x, 32))
    fig2 = px.bar(g2, y='SetorShort', x='Quantidade', orientation='h', hover_data={'Setor': True, 'SetorShort': False})
    fig2.update_layout(height=360, margin=dict(l=10, r=10, t=10, b=10), yaxis={'categoryorder':'total ascending'})

    return fig1, fig2, gt.to_dict('records'), total_text

@app.callback(
    Output('download-xlsx', 'data'),
    Input('btn-download-xlsx', 'n_clicks'),
    State('tabela', 'data'),
    State('dd-setor', 'value'),
    State('dd-tipo', 'value'),
    prevent_initial_call=True,
)
def do_download_xlsx(n_clicks, rows, setor, tipos):
    df = pd.DataFrame(rows)
    with pd.ExcelWriter(io.BytesIO(), engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Tabela')
        if not df.empty:
            tipos_df = df.groupby('Tipo')['Quantidade'].sum().reset_index().sort_values('Quantidade', ascending=False)
            tipos_df.to_excel(writer, index=False, sheet_name='Por_Tipo')
            setores_df = df.groupby('Setor')['Quantidade'].sum().reset_index().sort_values('Quantidade', ascending=False)
            setores_df.to_excel(writer, index=False, sheet_name='Por_Setor')
        meta = pd.DataFrame({
            'Filtro': ['Setor', 'Tipos selecionados'],
            'Valor': [setor if setor else '—', ', '.join(tipos) if tipos else '—']
        })
        meta.to_excel(writer, index=False, sheet_name='Filtros')
        writer.book.active = writer.book['Tabela']
        bio = writer._handles.handle
        bio.seek(0)
        data = bio.read()
    return dcc.send_bytes(lambda b: b.write(data), filename='dados_processos.xlsx')

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8050))
    app.run(host='127.0.0.1', port=port, debug=True)
