import pandas as pd
import dash
from dash import dcc, html, Input, Output
import plotly.express as px

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
sales_df.rename(columns={
    "Recorded Date": "Date",
    "Sales Consultant": "Consultant",
    "Question Summary": "Question",
}, inplace=True)

service_df.rename(columns={
    "Recorded Date": "Date",
    "Service Advisor": "Advisor",
    "Question Summary": "Question",
}, inplace=True)

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

    # Legenda de cores
    html.Div([
        html.H3("Score Color Legend"),
        html.P("🔴 Red: 0% to 89.99%"),
        html.P("🟡 Yellow: 90% to 92.99%"),
        html.P("🟢 Green: 93% to 100%"),
    ], style={"margin-top": "20px", "padding": "10px", "border": "1px solid black"}),

    # Gráficos
    dcc.Graph(id="sales_score_chart"),
    dcc.Graph(id="service_score_chart"),
    dcc.Graph(id="sales_question_chart"),
    dcc.Graph(id="service_question_chart"),
])

# Callback para atualizar os indicadores e gráficos
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

    # Criar gráficos
    def create_bar_chart(df, x_col, y_col, title):
        if df.empty:
            return px.bar(title=f"No Data Available for {title}")
        df["Color"] = df[y_col].apply(define_color)
        fig = px.bar(
            df,
            x=x_col,
            y=y_col,
            title=title,
            labels={y_col: "Average Score (%)"},
            color="Color",
            text=df[y_col].apply(lambda x: f"{x:.2f}%"),
            color_discrete_map={"red": "red", "yellow": "yellow", "green": "green"}
        )
        fig.update_traces(textposition="outside")
        return fig

    sales_avg = filtered_sales.groupby("Consultant")["Score"].mean().reset_index()
    service_avg = filtered_service.groupby("Advisor")["Score"].mean().reset_index()
    sales_question_avg = filtered_sales.groupby("Question")["Score"].mean().reset_index()
    service_question_avg = filtered_service.groupby("Question")["Score"].mean().reset_index()

    sales_fig = create_bar_chart(sales_avg, "Consultant", "Score", "Average Score by Sales Consultant (%)")
    service_fig = create_bar_chart(service_avg, "Advisor", "Score", "Average Score by Service Advisor (%)")
    sales_question_fig = create_bar_chart(sales_question_avg, "Question", "Score", "Average Score by Question (Sales) (%)")
    service_question_fig = create_bar_chart(service_question_avg, "Question", "Score", "Average Score by Question (Service) (%)")

    return sales_score_display, service_score_display, sales_fig, service_fig, sales_question_fig, service_question_fig


# Rodar o app
if __name__ == "__main__":
   app.run_server(debug=True, host="10.16.93.63", port=8050)