import pandas as pd

# Load dataset
df = pd.read_excel("output/sku_analysis.xlsx")

# 1. Define priority rules based on ABC_Class
priority_mapping = {
    "A": "CRITICAL",
    "B": "HIGH PRIORITY",
    "C": "MEDIUM PRIORITY"
}
df["Inventory_Priority"] = df["ABC_Class"].map(priority_mapping).fillna("LOW PRIORITY")

# 2. SKUs that require management attention
important_priorities = [
    "CRITICAL",
    "HIGH PRIORITY",
    "MEDIUM PRIORITY"
]

# 3. Filter important SKUs
important_skus = df[df["Inventory_Priority"].isin(important_priorities)].copy()

# 4. Sort by priority rank and value
priority_order = {
    "CRITICAL": 1,
    "HIGH PRIORITY": 2,
    "MEDIUM PRIORITY": 3
}

important_skus["Priority_Order"] = important_skus["Inventory_Priority"].map(priority_order)
important_skus = important_skus.sort_values(
    ["Priority_Order", "Outbound_Value_Proxy"], 
    ascending=[True, False]
).drop(columns=["Priority_Order"])

# Save file
important_skus.to_excel("important_skus_to_care.xlsx", index=False)

print("Important SKU file created successfully!")
print(f"Total important SKUs: {len(important_skus)}")
print("\nBreakdown:")
print(important_skus["Inventory_Priority"].value_counts())