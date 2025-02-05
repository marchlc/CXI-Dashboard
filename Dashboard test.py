import pandas as pd
import dash
from dash import dcc, html, Input, Output
import plotly.express as px
import sys
import plotly.io as pio
from dash.dependencies import State
from fpdf import FPDF

# Caminho do arquivo atualizado
file_path = "/Users/lucasmarchiori/Desktop/Relatorio CXI/CXI Report.xlsx"

# Carregar as abas do Excel
xls = pd.ExcelFile(file_path)

# Ler as abas Sales e Service
sales_df = pd.read_excel(xls, sheet_name="Sales")
service_df = pd.read_excel(xls, sheet_name="Service")

# Ajustar a conversão da data corretamente
sales_df["Recorded Date"] = pd.to_datetime(sales_df["Recorded Date"], format="%m/%d/%Y %I:%M %p", errors="coerce")
service_df["Recorded Date"] = pd.to_datetime(service_df["Recorded Date"], format="%m/%d/%Y %I:%M %p", errors="coerce")

# Renomear colunas para facilitar no Dash
sales_df.rename(columns={"Recorded Date": "Date", "Sales Consultant": "Consultant", "Question Summary": "Question"}, inplace=True)
service_df.rename(columns={"Recorded Date": "Date", "Service Advisor": "Advisor", "Question Summary": "Question"}, inplace=True)

# Adicionar colunas auxiliares de Year e Month
sales_df["Year"] = sales_df["Date"].dt.year
sales_df["Month"] = sales_df["Date"].dt.month
service_df["Year"] = service_df["Date"].dt.year
service_df["Month"] = service_df["Date"].dt.month

# Converter Score para percentual
sales_df["Score"] = sales_df["Score"] * 100
service_df["Score"] = service_df["Score"] * 100

# Função para definir cor com base no Score
def define_color(score):
    if score < 90:
        return "red"
    elif 90 <= score <= 92.99:
        return "yellow"
    else:
        return "green"

# Criar aplicação Dash
app = dash.Dash(__name__)

# Layout do dashboard
app.layout = html.Div([
    html.H1("CXI Performance Dashboard", style={"text-align": "center"}),

    # Filtros
    html.Div([
        html.Label("Select Year:"),
        dcc.Dropdown(
            id="year_filter",
            options=[{"label": str(int(year)), "value": int(year)} for year in sorted(sales_df["Year"].dropna().unique())],
            multi=False,
            value=2024
        ),

        html.Label("Select Month:"),
        dcc.Dropdown(
            id="month_filter",
            options=[{"label": str(int(month)), "value": int(month)} for month in range(1, 13)],
            multi=False,
            value=None
        ),

        html.Label("Select Sales Consultant:"),
        dcc.Dropdown(
            id="sales_consultant_filter",
            options=[{"label": consultant, "value": consultant} for consultant in sales_df["Consultant"].dropna().unique()],
            multi=True
        ),

        html.Label("Select Service Advisor:"),
        dcc.Dropdown(
            id="service_advisor_filter",
            options=[{"label": advisor, "value": advisor} for advisor in service_df["Advisor"].dropna().unique()],
            multi=True
        ),
    ], style={"width": "30%", "display": "inline-block", "padding": "10px"}),

    # Indicadores de Score
    html.Div([
        html.Div([
            html.H3("Sales Performance Score"),
            html.H2(id="sales_score", style={"font-size": "32px", "color": "blue"}),
        ], style={"width": "45%", "display": "inline-block", "text-align": "center", "border": "1px solid black", "padding": "20px"}),

        html.Div([
            html.H3("Service Performance Score"),
            html.H2(id="service_score", style={"font-size": "32px", "color": "blue"}),
        ], style={"width": "45%", "display": "inline-block", "text-align": "center", "border": "1px solid black", "padding": "20px"}),
    ], style={"display": "flex", "justify-content": "space-around", "margin-top": "20px"}),

    # Botão para download do relatório
    html.Button("Download Report", id="download-btn", n_clicks=0, style={"margin": "20px"}),
    dcc.Download(id="download-pdf"),

    # Gráficos
    dcc.Graph(id="sales_score_chart"),
    dcc.Graph(id="service_score_chart"),
    dcc.Graph(id="sales_question_chart"),
    dcc.Graph(id="service_question_chart"),
])

# Callback para atualizar os gráficos e indicadores
@app.callback(
    [Output("sales_score", "children"),
     Output("service_score", "children"),
     Output("sales_score_chart", "figure"),
     Output("service_score_chart", "figure"),
     Output("sales_question_chart", "figure"),
     Output("service_question_chart", "figure")],
    [Input("year_filter", "value"),
     Input("month_filter", "value"),
     Input("sales_consultant_filter", "value"),
     Input("service_advisor_filter", "value")]
)
def update_charts(selected_year, selected_month, selected_sales, selected_service):
    # Filtrar os dados conforme a seleção do usuário
    filtered_sales = sales_df[sales_df["Year"] == selected_year]
    filtered_service = service_df[service_df["Year"] == selected_year]

    if selected_month:
        filtered_sales = filtered_sales[filtered_sales["Month"] == selected_month]
        filtered_service = filtered_service[filtered_service["Month"] == selected_month]

    if selected_sales:
        filtered_sales = filtered_sales[filtered_sales["Consultant"].isin(selected_sales)]
    
    if selected_service:
        filtered_service = filtered_service[filtered_service["Advisor"].isin(selected_service)]

    # Calcular média de Score
    avg_sales_score = filtered_sales["Score"].mean() if not filtered_sales.empty else 0
    avg_service_score = filtered_service["Score"].mean() if not filtered_service.empty else 0

    sales_score_display = f"{avg_sales_score:.2f}%" if avg_sales_score > 0 else "No Data"
    service_score_display = f"{avg_service_score:.2f}%" if avg_service_score > 0 else "No Data"

    return sales_score_display, service_score_display, px.bar(), px.bar(), px.bar(), px.bar()

# Rodar o app na porta especificada
if __name__ == "__main__":
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8050
    app.run_server(debug=True, host="0.0.0.0", port=port)