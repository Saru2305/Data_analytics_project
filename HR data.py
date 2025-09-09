import pandas as pd
import matplotlib.pyplot as plt

# ---------------- Load Data ----------------
file_path =  r"C:\Users\sarum\Videos\data\Employee Sample Data.xlsx"
data = pd.read_excel(file_path, sheet_name="Employee Data")

# ---------------- Data Cleaning ----------------
# 1. Remove duplicates
data = data.drop_duplicates()

# 2. Handle missing values
data = data.fillna({col: (data[col].mean() if data[col].dtype in ["int64", "float64"] else data[col].mode()[0]) for col in data.columns})

# 3. Standardize column names
data.columns = data.columns.str.strip().str.lower().str.replace(" ", "_")

# 4. Convert date columns to datetime
if "hire_date" in data.columns:
    data["hire_date"] = pd.to_datetime(data["hire_date"], errors="coerce")

if "annual_salary" in data.columns:
    data["annual_salary"]=data["annual_salary"]* 1000

# 5. Remove salary outliers
if "annual_salary" in data.columns:
    lower, upper = data["annual_salary"].quantile([0.01, 0.99])
    data = data[(data["annual_salary"] >= lower) & (data["annual_salary"] <= upper)]

# 6. Trim whitespace in string columns
str_cols = data.select_dtypes(include=["object"]).columns
for col in str_cols:
    data[col] = data[col].str.strip()

# ---------------- Analysis ----------------
analysis = {}

# Basic info
analysis["summary"] = data.describe(include="all")

# Average salary by department
if "department" in data.columns and "annual_salary" in data.columns:
    analysis["avg_salary_by_department"] = data.groupby("department")["annual_salary"].mean().sort_values(ascending=False)

# Gender distribution
if "gender" in data.columns:
    analysis["gender_distribution"] = data["gender"].value_counts()

# Age distribution
if "age" in data.columns:
    analysis["age_distribution"] = data["age"].describe()

# ---------------- Save Results ----------------
output_path = "Employees(1)_Data_Analysis.xlsx"
with pd.ExcelWriter(output_path) as writer:
    data.to_excel(writer, sheet_name="Cleaned_Data", index=False)
    for key, val in analysis.items():
        if isinstance(val, pd.Series):
            val.to_excel(writer, sheet_name=key)
        elif isinstance(val, pd.DataFrame):
            val.to_excel(writer, sheet_name=key)

print(f"Analysis report saved to: {output_path}")

# ---------------- Visualizations ----------------

# 1. Salary distribution
if "annual_salary" in data.columns:
    plt.figure(figsize=(8,5))
    data["annual_salary"].hist(bins=30)
    plt.title("Salary Distribution")
    plt.xlabel("Annual Salary")
    plt.ylabel("Frequency")
    plt.savefig("histogram salary.png")
    plt.close()

# 2. Gender distribution pie chart
if "gender" in data.columns:
    plt.figure(figsize=(6,6))
    data["gender"].value_counts().plot.pie(autopct='%1.1f%%')
    plt.title("Gender Distribution")
    plt.ylabel("")
    plt.savefig("Gender pie.png")
    plt.close()

# 3. Average salary by department
if "department" in data.columns and "annual_salary" in data.columns:
    plt.figure(figsize=(10,6))
    data.groupby("department")["annual_salary"].mean().sort_values().plot(kind="barh")
    plt.title("Average Salary by Department")
    plt.xlabel("Average Salary")
    plt.ylabel("Department")
    plt.savefig("department barchart.png")
    plt.close()

# 4. Age distribution
if "age" in data.columns:
    plt.figure(figsize=(8,5))
    data["age"].hist(bins=20)
    plt.title("Age Distribution")
    plt.xlabel("Age")
    plt.ylabel("Frequency")
    plt.savefig("histogram Age.png")
    plt.close()

print("Analysis report and charts saved.")

