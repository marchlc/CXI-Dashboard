import pandas as pd
import dash
from dash import dcc, html, Input, Output
import plotly.express as px

# 🔗 Link do Google Sheets (versão CSV exportável)
google_sheets_link = "https://docs.google.com/spreadsheets/d/1B_vqtjzVjBEyTaG4Ng2onq_SLIWccIVi/gviz/tq?tqx=out:csv&sheet=Service"
google_sheets_sales = "https://docs.google.com/spreadsheets/d/1B_vqtjzVjBEyTaG4Ng2onq_SLIWccIVi/gviz/tq?tqx=out:csv&sheet=Sales"

# 🔄 Carregar os dados diretamente do Google Sheets
sales_df = pd.read_csv(google_sheets_sales, dtype=str)
service_df = pd.read_csv(google_sheets_link, dtype=str)

# ✅ Verifica se a coluna "Service Advisor" existe antes de renomear
if "Service Advisor" in service_df.columns:
    service_df.rename(columns={"Service Advisor": "Advisor"}, inplace=True)

# ✅ Verifica se a coluna "Sales Consultant" existe antes de renomear
if "Sales Consultant" in sales_df.columns:
    sales_df.rename(columns={"Sales Consultant": "Consultant"}, inplace=True)

# 🔄 Conversão da data sem definir formato fixo
sales_df["Recorded Date"] = pd.to_datetime(sales_df["Recorded Date"], errors="coerce")
service_df["Recorded Date"] = pd.to_datetime(service_df["Recorded Date"], errors="coerce")

# ✅ Renomeia a coluna de data
sales_df.rename(columns={"Recorded Date": "Date"}, inplace=True)
service_df.rename(columns={"Recorded Date": "Date"}, inplace=True)

# ✅ Adiciona colunas de ano e mês
sales_df["Year"] = sales_df["Date"].dt.year
sales_df["Month"] = sales_df["Date"].dt.month

service_df["Year"] = service_df["Date"].dt.year
service_df["Month"] = service_df["Date"].dt.month

# ✅ Converter Score para número removendo vírgulas, espaços e % (caso existam)
def clean_score(df, column):
    if column in df.columns:
        df[column] = df[column].str.replace("%", "").str.replace(",", ".").str.strip()
        df[column] = pd.to_numeric(df[column], errors="coerce")  # Converte para float
    return df

sales_df = clean_score(sales_df, "Score")
service_df = clean_score(service_df, "Score")

# ✅ Função para definir cor com base no Score
def define_color(score):
    if score < 90:
        return "red"
    elif 90 <= score <= 92.99:
        return "yellow"
    else:
        return "green"

# 🚀 Criar aplicação Dash
app = dash.Dash(__name__)

# 📊 Layout do dashboard
app.layout = html.Div([
    html.H1("CXI Performance Dashboard", style={"text-align": "center"}),

    # 🔎 Filtros
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

    # 📈 Indicadores de Score
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

    # 📌 Legenda de cores
    html.Div([
        html.H3("Score Color Legend"),
        html.P("🔴 Red: 0% to 89.99%"),
        html.P("🟡 Yellow: 90% to 92.99%"),
        html.P("🟢 Green: 93% to 100%"),
    ], style={"margin-top": "20px", "padding": "10px", "border": "1px solid black"}),

    # 📊 Gráficos
    dcc.Graph(id="sales_score_chart"),
    dcc.Graph(id="service_score_chart"),
    dcc.Graph(id="sales_question_chart"),
    dcc.Graph(id="service_question_chart"),
])

# 🔄 Callback para atualizar os indicadores e gráficos
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
    # 📊 Filtrando os dados com base nas seleções do usuário
    filtered_sales = sales_df[sales_df["Year"] == selected_year]
    filtered_service = service_df[service_df["Year"] == selected_year]

    if selected_month:
        filtered_sales = filtered_sales[filtered_sales["Month"] == selected_month]
        filtered_service = filtered_service[filtered_service["Month"] == selected_month]

    if selected_sales:
        filtered_sales = filtered_sales[filtered_sales["Consultant"].isin(selected_sales)]
    
    if selected_service:
        filtered_service = filtered_service[filtered_service["Advisor"].isin(selected_service)]

    # 📊 Criar gráficos
    sales_fig = px.bar(filtered_sales, x="Consultant", y="Score", title="Sales Score")
    service_fig = px.bar(filtered_service, x="Advisor", y="Score", title="Service Score")

    return f"{filtered_sales['Score'].mean():.2f}%", f"{filtered_service['Score'].mean():.2f}%", sales_fig, service_fig, px.bar(), px.bar()

# 🚀 Rodar o app
if __name__ == "__main__":
   app.run_server(debug=True, host="0.0.0.0", port=8050)