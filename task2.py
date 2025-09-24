# STEP 1: Import libraries
import pandas as pd

# STEP 2: Load dataset
# Replace with your dataset path, e.g. "Superstore.csv"
df = pd.read_csv(r"C:\Users\jivit\OneDrive\Documents\Matplotib & Seaborn\zInternship\Sample_Superstore.csv", encoding='latin1')

# STEP 3: Data cleaning
# Fix column names (remove spaces)
df.columns = df.columns.str.strip().str.replace(" ", "_")

# Convert dates
df['Order_Date'] = pd.to_datetime(df['Order_Date'], errors='coerce')
df['Ship_Date'] = pd.to_datetime(df['Ship_Date'], errors='coerce')

# Drop duplicates
df = df.drop_duplicates()

# STEP 4: Feature engineering
df['Year'] = df['Order_Date'].dt.year
df['Month'] = df['Order_Date'].dt.month
df['Year_Month'] = df['Order_Date'].dt.to_period('M')

# Profit Margin
df['Profit_Margin'] = df['Profit'] / df['Sales']

# STEP 5: Business KPIs (can be shown in Power BI cards)
total_sales = df['Sales'].sum()
total_profit = df['Profit'].sum()
avg_margin = df['Profit_Margin'].mean()

print("Total Sales: ", total_sales)
print("Total Profit: ", total_profit)
print("Average Profit Margin: ", avg_margin)

# STEP 6: Aggregations for visuals (Power BI charts)
# Sales by Category
sales_by_cat = df.groupby("Category")[["Sales","Profit"]].sum().reset_index()

# Monthly Sales Trend
monthly_sales = df.groupby("Year_Month")[["Sales","Profit"]].sum().reset_index()

# Top 10 Products by Sales
top_products = df.groupby("Product_Name")[["Sales","Profit"]].sum().reset_index().sort_values("Sales", ascending=False).head(10)

# Region Analysis
region_perf = df.groupby("Region")[["Sales","Profit"]].sum().reset_index()

# Segment Contribution
segment_perf = df.groupby("Segment")[["Sales"]].sum().reset_index()

# Top Customers by Profit
top_customers = df.groupby("Customer_Name")[["Sales","Profit"]].sum().reset_index().sort_values("Profit", ascending=False).head(10)

# STEP 7: Export clean data for Power BI
df.to_csv("Cleaned_Superstore.csv", index=False)
sales_by_cat.to_csv("Sales_By_Category.csv", index=False)
monthly_sales.to_csv("Monthly_Sales.csv", index=False)
top_products.to_csv("Top_Products.csv", index=False)
region_perf.to_csv("Region_Performance.csv", index=False)
segment_perf.to_csv("Segment_Performance.csv", index=False)
top_customers.to_csv("Top_Customers.csv", index=False)

print("✅ Cleaned datasets exported! Ready to use in Power BI.")