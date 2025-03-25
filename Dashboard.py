import pandas as pd
import dash
from dash import dcc, html, Input, Output
import plotly.express as px
import requests
from io import BytesIO

# 🛠 **PASSO 1: Baixar e carregar os dados do Google Sheets**
sheet_url = "https://docs.google.com/spreadsheets/d/16DjXZM4kC0i6bD05IETJ3gIoBZpy2NaF/export?format=xlsx"

try:
    response = requests.get(sheet_url)
    response.raise_for_status()  # Verifica se a requisição foi bem-sucedida
    xls = pd.ExcelFile(BytesIO(response.content))
    print("Arquivo baixado com sucesso!")
    sales_df = pd.read_excel(xls, sheet_name="Sales")
    service_df = pd.read_excel(xls, sheet_name="Service")
except requests.exceptions.RequestException as e:
    print(f"Erro ao baixar o arquivo do Google Sheets: {e}")
    file_path = "/mnt/data/CXI Report.xlsx"  # Usar arquivo local caso o download falhe
    xls = pd.ExcelFile(file_path)
    print("Arquivo carregado localmente!")
    sales_df = pd.read_excel(xls, sheet_name="Sales")
    service_df = pd.read_excel(xls, sheet_name="Service")

# **PASSO 2: Ajustar a conversão de data (Apenas DD/MM/YYYY)**
sales_df["Date"] = pd.to_datetime(sales_df["Recorded Date"], format="%d/%m/%Y", errors="coerce")
service_df["Date"] = pd.to_datetime(service_df["Recorded Date"], format="%d/%m/%Y", errors="coerce")

# **PASSO 3: Criar colunas de Ano e Mês**
sales_df["Year"] = sales_df["Date"].dt.year
sales_df["Month"] = sales_df["Date"].dt.month
service_df["Year"] = service_df["Date"].dt.year
service_df["Month"] = service_df["Date"].dt.month

# **Verificação dos dados após ajuste**
print("Service DataFrame Após Ajuste de Data:")
print(service_df[["Recorded Date", "Date", "Year", "Month"]].head())

# **PASSO 4: Limpar e ajustar a coluna "Score"**
def clean_score_column(df):
    df["Score"] = df["Score"].astype(str).str.replace("%", "").astype(float)
    
    # Se os valores estiverem entre 0 e 1, convertemos para escala de 100
    if df["Score"].max() <= 1:
        df["Score"] = df["Score"] * 100
    
    return df

sales_df = clean_score_column(sales_df)
service_df = clean_score_column(service_df)

# **PASSO 5: Definir cores para os scores corretamente**
def define_color(score):
    if score < 90:
        return "red"
    elif 90 <= score <= 94:
        return "yellow"
    else:
        return "green"

sales_df["Color"] = sales_df["Score"].apply(define_color)
service_df["Color"] = service_df["Score"].apply(define_color)

# **PASSO 6: Criar a aplicação Dash**
app = dash.Dash(__name__)
server = app.server

# **PASSO 7: Layout do dashboard**
app.layout = html.Div([
    html.H1("CXI Performance Dashboard", style={"text-align": "center", "background-color": "#0F3557", "color": "white", "padding": "20px"}),

    # Filtros
    html.Div([
        html.Label("Select Year:"),
        dcc.Dropdown(id="year_filter", options=[{"label": str(int(year)), "value": int(year)} for year in sorted(sales_df["Year"].dropna().unique())], multi=False, value=2024),
        html.Label("Select Month:"),
        dcc.Dropdown(id="month_filter", options=[{"label": str(int(month)), "value": int(month)} for month in range(1, 13)], multi=False, value=None),
        html.Label("Select Sales Consultant:"),
        dcc.Dropdown(id="sales_consultant_filter", options=[{"label": consultant, "value": consultant} for consultant in sales_df["Sales Consultant"].dropna().unique()], multi=True),
        html.Label("Select Service Advisor:"),
        dcc.Dropdown(id="service_advisor_filter", options=[{"label": advisor, "value": advisor} for advisor in service_df["Service Advisor"].dropna().unique()], multi=True),
    ], style={"width": "30%", "display": "inline-block", "padding": "10px"}),

    # Gráficos
    dcc.Graph(id="sales_score_chart"),
    dcc.Graph(id="service_score_chart"),
    dcc.Graph(id="sales_question_chart", style={"height": "500px"}),
    dcc.Graph(id="service_question_chart"),
])

# **PASSO 8: Callback para atualizar os gráficos**
@app.callback(
    [Output("sales_score_chart", "figure"),
     Output("service_score_chart", "figure"),
     Output("sales_question_chart", "figure"),
     Output("service_question_chart", "figure")],
    [Input("year_filter", "value"),
     Input("month_filter", "value"),
     Input("sales_consultant_filter", "value"),
     Input("service_advisor_filter", "value")]
)
def update_charts(selected_year, selected_month, selected_sales, selected_service):
    filtered_sales = sales_df[sales_df["Year"] == selected_year]
    filtered_service = service_df[service_df["Year"] == selected_year]

    if selected_month:
        filtered_sales = filtered_sales[filtered_sales["Month"] == selected_month]
        filtered_service = filtered_service[filtered_service["Month"] == selected_month]

    if selected_sales:
        filtered_sales = filtered_sales[filtered_sales["Sales Consultant"].isin(selected_sales)]
    
    if selected_service:
        filtered_service = filtered_service[filtered_service["Service Advisor"].isin(selected_service)]

    sales_avg = filtered_sales.groupby("Sales Consultant")["Score"].mean().reset_index()
    service_avg = filtered_service.groupby("Service Advisor")["Score"].mean().reset_index()
    sales_question_avg = filtered_sales.groupby("Question Summary")["Score"].mean().reset_index()
    service_question_avg = filtered_service.groupby("Question Summary")["Score"].mean().reset_index()

    def create_chart(df, x_col, y_col, title):
        if df.empty:
            return px.bar(title=f"No Data Available for {title}")
        df["Color"] = df[y_col].apply(define_color)
        fig = px.bar(
            df, x=x_col, y=y_col, title=title,
            text=df[y_col].apply(lambda x: f"{x:.0f}%"),
            color="Color",
            color_discrete_map={"red": "red", "yellow": "yellow", "green": "green"}
        )
        fig.update_traces(textposition="outside")
        fig.update_layout(yaxis_tickformat=".0%", height=500)
        return fig

    return (
        create_chart(sales_avg, "Sales Consultant", "Score", "Average Score by Sales Consultant"),
        create_chart(service_avg, "Service Advisor", "Score", "Average Score by Service Advisor"),
        create_chart(sales_question_avg, "Question Summary", "Score", "Average Score by Question (Sales)"),
        create_chart(service_question_avg, "Question Summary", "Score", "Average Score by Question (Service)")
    )

if __name__ == "__main__":
    app.run_server(debug=True, host="0.0.0.0", port=8050)