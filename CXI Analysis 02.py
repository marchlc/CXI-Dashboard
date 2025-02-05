import pandas as pd

# Caminho do arquivo
file_path = "/Users/lucasmarchiori/Desktop/Relatorio CXI/CXI Report.xlsx"

# Carregar todas as abas do Excel
xls = pd.ExcelFile(file_path)

# Ler as abas Sales e Service
sales_df = pd.read_excel(xls, sheet_name="Sales")
service_df = pd.read_excel(xls, sheet_name="Service")

# Converter a coluna "Recorded Date" para datetime no formato MM/DD/YYYY
sales_df["Recorded Date"] = pd.to_datetime(sales_df["Recorded Date"], format="%m/%d/%Y", errors="coerce")
service_df["Recorded Date"] = pd.to_datetime(service_df["Recorded Date"], format="%m/%d/%Y", errors="coerce")

# Filtrar os dados para 2024 e 2025
sales_2024 = sales_df[sales_df["Recorded Date"].dt.year == 2024]
service_2024 = service_df[service_df["Recorded Date"].dt.year == 2024]

sales_2025 = sales_df[sales_df["Recorded Date"].dt.year == 2025]
service_2025 = service_df[service_df["Recorded Date"].dt.year == 2025]

# Task #1 - Média do Score em 2024 para cada Sales Consultant e Service Advisor
avg_score_sales_2024 = sales_2024.groupby("Sales Consultant")["Score"].mean().reset_index()
avg_score_service_2024 = service_2024.groupby("Service Advisor")["Score"].mean().reset_index()

# Task #2 - Média do Score mensal de 2024 e 2025
avg_score_sales_month_2024 = sales_2024.groupby([sales_2024["Recorded Date"].dt.month, "Sales Consultant"])["Score"].mean().reset_index()
avg_score_service_month_2024 = service_2024.groupby([service_2024["Recorded Date"].dt.month, "Service Advisor"])["Score"].mean().reset_index()

avg_score_sales_month_2025 = sales_2025.groupby([sales_2025["Recorded Date"].dt.month, "Sales Consultant"])["Score"].mean().reset_index()
avg_score_service_month_2025 = service_2025.groupby([service_2025["Recorded Date"].dt.month, "Service Advisor"])["Score"].mean().reset_index()

# Task #3 - Média do Score por pergunta (Question Summary) em 2024
avg_score_question_sales_2024 = sales_2024.groupby("Question Summary")["Score"].mean().reset_index()
avg_score_question_service_2024 = service_2024.groupby("Question Summary")["Score"].mean().reset_index()

# Task #4 - Média do Score por mês de cada pergunta em 2024 e 2025
avg_score_question_month_sales_2024 = sales_2024.groupby([sales_2024["Recorded Date"].dt.month, "Question Summary"])["Score"].mean().reset_index()
avg_score_question_month_service_2024 = service_2024.groupby([service_2024["Recorded Date"].dt.month, "Question Summary"])["Score"].mean().reset_index()

avg_score_question_month_sales_2025 = sales_2025.groupby([sales_2025["Recorded Date"].dt.month, "Question Summary"])["Score"].mean().reset_index()
avg_score_question_month_service_2025 = service_2025.groupby([service_2025["Recorded Date"].dt.month, "Question Summary"])["Score"].mean().reset_index()

# Task #5 - Média do Score por Question Summary por Sales Consultant / Service Advisor
avg_score_question_sales_consultant_2024 = sales_2024.groupby(["Sales Consultant", "Question Summary"])["Score"].mean().reset_index()
avg_score_question_service_advisor_2024 = service_2024.groupby(["Service Advisor", "Question Summary"])["Score"].mean().reset_index()

avg_score_question_month_sales_consultant_2024 = sales_2024.groupby(["Sales Consultant", sales_2024["Recorded Date"].dt.month, "Question Summary"])["Score"].mean().reset_index()
avg_score_question_month_service_advisor_2024 = service_2024.groupby(["Service Advisor", service_2024["Recorded Date"].dt.month, "Question Summary"])["Score"].mean().reset_index()

avg_score_question_month_sales_consultant_2025 = sales_2025.groupby(["Sales Consultant", sales_2025["Recorded Date"].dt.month, "Question Summary"])["Score"].mean().reset_index()
avg_score_question_month_service_advisor_2025 = service_2025.groupby(["Service Advisor", service_2025["Recorded Date"].dt.month, "Question Summary"])["Score"].mean().reset_index()

# Task #6 - Transformação dos valores de Score em percentuais
def transform_to_percentage(df, score_col="Score"):
    df[score_col] = df[score_col] * 100
    return df

dfs_to_transform = [
    avg_score_sales_2024, avg_score_service_2024,
    avg_score_sales_month_2024, avg_score_service_month_2024,
    avg_score_sales_month_2025, avg_score_service_month_2025,
    avg_score_question_sales_2024, avg_score_question_service_2024,
    avg_score_question_month_sales_2024, avg_score_question_month_service_2024,
    avg_score_question_month_sales_2025, avg_score_question_month_service_2025,
    avg_score_question_sales_consultant_2024, avg_score_question_service_advisor_2024,
    avg_score_question_month_sales_consultant_2024, avg_score_question_month_service_advisor_2024,
    avg_score_question_month_sales_consultant_2025, avg_score_question_month_service_advisor_2025
]

for df in dfs_to_transform:
    transform_to_percentage(df)

# Aplicar formatação condicional
def apply_color_coding(df, score_col="Score"):
    df["Color Code"] = pd.cut(
        df[score_col], bins=[0, 89.99, 92.99, 100],
        labels=["Red", "Yellow", "Green"], include_lowest=True
    )
    return df

dfs_to_format = dfs_to_transform

for df in dfs_to_format:
    apply_color_coding(df)

# Salvar os resultados formatados em um novo arquivo Excel
output_file = "/Users/lucasmarchiori/Desktop/Relatorio CXI/CXI_Report_Processed.xlsx"
with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
    avg_score_sales_2024.to_excel(writer, sheet_name="Sales_Avg_2024", index=False)
    avg_score_service_2024.to_excel(writer, sheet_name="Service_Avg_2024", index=False)
    avg_score_sales_month_2024.to_excel(writer, sheet_name="Sales_Monthly_2024", index=False)
    avg_score_service_month_2024.to_excel(writer, sheet_name="Service_Monthly_2024", index=False)
    avg_score_sales_month_2025.to_excel(writer, sheet_name="Sales_Monthly_2025", index=False)
    avg_score_service_month_2025.to_excel(writer, sheet_name="Service_Monthly_2025", index=False)
    avg_score_question_sales_2024.to_excel(writer, sheet_name="Sales_Question_2024", index=False)
    avg_score_question_service_2024.to_excel(writer, sheet_name="Service_Question_2024", index=False)
    avg_score_question_month_sales_2024.to_excel(writer, sheet_name="Sales_Q_Monthly_2024", index=False)
    avg_score_question_month_service_2024.to_excel(writer, sheet_name="Service_Q_Monthly_2024", index=False)
    avg_score_question_month_sales_2025.to_excel(writer, sheet_name="Sales_Q_Monthly_2025", index=False)
    avg_score_question_month_service_2025.to_excel(writer, sheet_name="Service_Q_Monthly_2025", index=False)

print(f"Arquivo processado e salvo em: {output_file}")