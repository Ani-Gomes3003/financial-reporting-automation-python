import pandas as pd
import os

input_folder = "input_files"
output_file = "output/financial_summary_report.xlsx"

print("Starting financial folder automation...")

files = [f for f in os.listdir(input_folder) if f.endswith(".xlsx")]

if not files:
    print("No Excel files found.")
    exit()

dataframes = []

for file in files:
    file_path = os.path.join(input_folder, file)
    print(f"Processing file: {file}")
    
    df = pd.read_excel(file_path)
    
    # Convert Date column
    df["Date"] = pd.to_datetime(df["Date"])
    
    # Add Month column
    df["Month"] = df["Date"].dt.strftime("%B")
    
    # Financial KPIs
    df["Revenue_per_Hour"] = df["Revenue_USD"] / df["Hours_Worked"]
    df["Revenue_per_Transaction"] = df["Revenue_USD"] / df["Transactions"]
    df["Transactions_per_Hour"] = df["Transactions"] / df["Hours_Worked"]
    df["Error_Rate"] = df["Errors"] / df["Transactions"]
    df["Error_Percentage"] = df["Error_Rate"] * 100
    
    dataframes.append(df)

# Combine all data
combined_df = pd.concat(dataframes, ignore_index=True)

# Department Summary
department_summary = combined_df.groupby("Department").agg({
    "Revenue_USD": "sum",
    "Transactions": "sum",
    "Errors": "sum",
    "Revenue_per_Hour": "mean",
    "Error_Percentage": "mean"
}).reset_index()

# Monthly Summary
monthly_summary = combined_df.groupby("Month").agg({
    "Revenue_USD": "sum",
    "Transactions": "sum",
    "Errors": "sum"
}).reset_index()

# Save Excel report
with pd.ExcelWriter(output_file) as writer:
    
    combined_df.to_excel(writer, sheet_name="Detailed_Data", index=False)
    
    department_summary.to_excel(writer, sheet_name="Department_Summary", index=False)
    
    monthly_summary.to_excel(writer, sheet_name="Monthly_Summary", index=False)

print("Financial automation completed successfully.")
print(f"Report saved at: {output_file}")
