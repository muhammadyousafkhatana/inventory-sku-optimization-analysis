import pandas as pd
import numpy as np

df = pd.read_excel(r'stock-keeping-units/sku_data.xlsx')

print(df.head())
print("Dataset shape:", df.shape)
 
# tochecks the size, columns, data types and basic statistics
print("\nData information:")
print(df.info())

print("\nBasic statistics:")
print(df.describe())

print("\nMissing values:")
print(df.isnull().sum())

#to checks how many zero values exist in important columns.
#zero values are important because they may represent missing or unavailable product information.
print("\nZero values in each column:")
print((df == 0).sum())

# This checks whether Outbound Number and Total Outbound
# are consistent with each other.
print("\nOutbound consistency:")

print("Both outbound values are zero:")
print(
    (
        (df["Outbound number"] == 0) &
        (df["Total outbound"] == 0)
    ).sum()
)

print("Outbound number > 0 but Total outbound = 0:")
print(
    (
        (df["Outbound number"] > 0) &
        (df["Total outbound"] == 0)
    ).sum()
)

print("Outbound number = 0 but Total outbound > 0:")
print(
    (
        (df["Outbound number"] == 0) &
        (df["Total outbound"] > 0)
    ).sum()
)

print("Both values > 0:")
print(
    (
        (df["Outbound number"] > 0) &
        (df["Total outbound"] > 0)
    ).sum()
)

# This shows how demand is distributed across the SKUs.
# It also helps identify highly demanded products.
print("\nOutbound Number statistics:")
print(df["Outbound number"].describe())

print("\nTotal Outbound statistics:")
print(df["Total outbound"].describe())

print("\nImportant outbound percentiles:")
print(
    df[["Outbound number", "Total outbound"]].quantile(
        [0.50, 0.75, 0.90, 0.95, 0.99]
    )
)
 
# ACTIVE VS INACTIVE SKUs
# This identifies which SKUs have actual outbound activity
# and which SKUs have no recorded outbound activity.
df["Activity_Status"] = np.where(
    df["Total outbound"] > 0,
    "Active",
    "Inactive"
)

print("\nActivity status:")
print(df["Activity_Status"].value_counts())
 
# PALLETS PER ORDER
# This measures the average number of pallets handled per
# outbound order for active SKUs.
active = df[df["Outbound number"] > 0].copy()

active["Pallets_per_Order"] = (
    active["Total outbound"] /
    active["Outbound number"]
)

print("\nPallets per order:")
print(active["Pallets_per_Order"].describe())
 
# OUTBOUND VALUE PROXY
# This estimates the value associated with outbound volume.
# It is calculated as Total Outbound multiplied by Unit Price.
df["Outbound_Value_Proxy"] = (
    df["Total outbound"] *
    df["Unitprice"]
)

print("\nOutbound value proxy:")
print(df["Outbound_Value_Proxy"].describe())
 
# TOP SKUs BY OUTBOUND VOLUME
# This identifies the products with the highest physical outbound volume. 
print("\nTop 10 SKUs by outbound volume:")

print(
    df[
        [
            "ID",
            "Outbound number",
            "Total outbound",
            "Unitprice"
        ]
    ]
    .sort_values(
        "Total outbound",
        ascending=False
    )
    .head(10)
)
#top SKUs BY OUTBOUND VALUE
# This identifies the products that have the highest estimated
# outbound value and therefore deserve more attention. 
print("\nTop 10 SKUs by outbound value:")

print(
    df[
        [
            "ID",
            "Unitprice",
            "Total outbound",
            "Outbound_Value_Proxy"
        ]
    ]
    .sort_values(
        "Outbound_Value_Proxy",
        ascending=False
    )
    .head(10)
)
# DEMAND CLASSIFICATION
# This divides active SKUs into Low, Medium and High volume
# based on the 33rd and 67th percentiles of active demand. 
df["Demand_Class"] = "Inactive"

active_mask = df["Total outbound"] > 0

active_volume = df.loc[
    active_mask,
    "Total outbound"
]

low_volume_threshold = active_volume.quantile(0.33)
high_volume_threshold = active_volume.quantile(0.67)

print("\nDemand thresholds:")
print("Low:", low_volume_threshold)
print("High:", high_volume_threshold)
 
df.loc[
    active_mask &
    (df["Total outbound"] <= low_volume_threshold),
    "Demand_Class"
] = "Active - Low Volume"
 
df.loc[
    active_mask &
    (df["Total outbound"] > low_volume_threshold) &
    (df["Total outbound"] <= high_volume_threshold),
    "Demand_Class"
] = "Active - Medium Volume"
 
df.loc[
    active_mask &
    (df["Total outbound"] > high_volume_threshold),
    "Demand_Class"
] = "Active - High Volume"
 
print("\nDemand class:")
print(df["Demand_Class"].value_counts())

# FREQUENCY CLASSIFICATION
# This classifies active SKUs according to how frequently they are ordered. 
df["Frequency_Class"] = "Inactive"

active_orders = df.loc[
    active_mask,
    "Outbound number"
]

low_frequency_threshold = active_orders.quantile(0.33)
high_frequency_threshold = active_orders.quantile(0.67)

print("\nFrequency thresholds:")
print("Low:", low_frequency_threshold)
print("High:", high_frequency_threshold)
 
df.loc[
    active_mask &
    (df["Outbound number"] <= low_frequency_threshold),
    "Frequency_Class"
] = "Low Frequency"
 
df.loc[
    active_mask &
    (df["Outbound number"] > low_frequency_threshold) &
    (df["Outbound number"] <= high_frequency_threshold),
    "Frequency_Class"
] = "Medium Frequency"
 
df.loc[
    active_mask &
    (df["Outbound number"] > high_frequency_threshold),
    "Frequency_Class"
] = "High Frequency"
 
print("\nFrequency class:")
print(df["Frequency_Class"].value_counts())

# SKU SEGMENTATION
# This combines demand volume and order frequency to understand
# which SKUs are low, medium or high demand/frequency.
def create_sku_segment(row):

    if row["Activity_Status"] == "Inactive":
        return "Inactive"

    volume = row["Demand_Class"].replace(
        "Active - ", ""
    )

    frequency = row["Frequency_Class"].replace(
        " Frequency", ""
    )

    return volume + " / " + frequency
 
df["SKU_Segment"] = df.apply(
    create_sku_segment,
    axis=1
)

print("\nSKU Segmentation:")
print(
    df["SKU_Segment"]
    .value_counts()
    .sort_index()
)

# VOLUME × FREQUENCY MATRIX
# This shows how many SKUs fall into each combination of
# demand volume and order frequency.  
print("\nVolume x Frequency Matrix:")

print(
    pd.crosstab(
        df["Demand_Class"],
        df["Frequency_Class"]
    )
) 
# ABC ANALYSIS
# This classifies SKUs according to their contribution to
# outbound value. 

abc_df = df.sort_values(
    "Outbound_Value_Proxy",
    ascending=False
).copy()

abc_df["Value_Share"] = (
    abc_df["Outbound_Value_Proxy"] /
    abc_df["Outbound_Value_Proxy"].sum()
)

abc_df["Cumulative_Value_Share"] = (
    abc_df["Value_Share"].cumsum()
)
  
abc_df["ABC_Class"] = np.select(
    [
        abc_df["Cumulative_Value_Share"] <= 0.80,
        abc_df["Cumulative_Value_Share"] <= 0.95
    ],
    [
        "A",
        "B"
    ],
    default="C"
)
 
# Inactive SKUs are separated because they have no outbound activity
abc_df.loc[
    abc_df["Activity_Status"] == "Inactive",
    "ABC_Class"
] = "Inactive"
 
df["ABC_Class"] = abc_df["ABC_Class"]
 
print("\nABC classification:")
print(df["ABC_Class"].value_counts())
# ABC VALUE SUMMARY
# This shows how much outbound value is associated with each ABC category.
abc_summary = (
    df.groupby("ABC_Class")
    .agg(
        SKU_Count=("ID", "count"),
        Total_Value=("Outbound_Value_Proxy", "sum")
    )
)

print("\nABC value summary:")
print(abc_summary)
#INVENTORY PRIORITY
#This identifies the SKUs that need the most management attention
#by combining ABC value importance with demand volume and frequency.
def assign_priority(row):
    if row["Activity_Status"] == "Inactive":
        return "Inactive / Review"
    if (
        row["ABC_Class"] == "A" and
        row["SKU_Segment"] == "High Volume / High"
    ):
        return "CRITICAL"
    if row["ABC_Class"] == "A":
        return "HIGH PRIORITY"
    if (
        row["ABC_Class"] == "B" and
        row["SKU_Segment"] in [
            "High Volume / High",
            "High Volume / Medium",
            "Medium Volume / High"
        ]
    ):
        return "MEDIUM-HIGH PRIORITY"

    if row["ABC_Class"] == "B":
        return "MEDIUM PRIORITY"

    return "LOW PRIORITY"
 
df["Inventory_Priority"] = df.apply(
    assign_priority,
    axis=1
)
print("\nInventory Priority:")
print(
    df["Inventory_Priority"].value_counts()
)
# CRITICAL SKU ANALYSIS
# This identifies the most important SKUs that are both
# high-value and high-volume/high-frequency.
critical = df[
    df["Inventory_Priority"] == "CRITICAL"
].copy()

print("\nCritical SKUs:")
print(
    critical[
        [
            "ID",
            "Unitprice",
            "Outbound number",
            "Total outbound",
            "Outbound_Value_Proxy",
            "ABC_Class",
            "SKU_Segment",
            "Inventory_Priority"
        ]
    ]
    .sort_values(
        "Outbound_Value_Proxy",
        ascending=False
    )
    .head(20)
)
# CRITICAL SKU IMPACT
# This measures how important the critical SKUs are compared
# with the complete dataset in terms of volume and value.
critical_value = critical[
    "Outbound_Value_Proxy"
].sum()

critical_volume = critical[
    "Total outbound"
].sum()

total_value = df[
    "Outbound_Value_Proxy"
].sum()

total_volume = df[
    "Total outbound"
].sum()

print("\nCritical SKU impact:")

print("Critical SKU count:", len(critical))

print(
    "Critical outbound value share:",
    critical_value / total_value * 100,
    "%"
)

print(
    "Critical outbound volume share:",
    critical_volume / total_volume * 100,
    "%"
) 
# CRITICAL SKU DATA QUALITY
# This checks whether the most important SKUs have missing or
# zero product information that could affect decisions.
print("\nCritical SKU data quality:")

print(
    {
        "Zero Unit Price":
            (critical["Unitprice"] == 0).sum(),

        "Zero Expiry":
            (critical["Expire date"] == 0).sum(),

        "Zero Grossweight":
            (critical["Pal grossweight"] == 0).sum(),

        "Zero Pallet Height":
            (critical["Pal height"] == 0).sum(),

        "Zero Units per Pallet":
            (critical["Units per pal"] == 0).sum()
    }
) 
# PRIORITY IMPACT SUMMARY
# This compares each inventory priority group by SKU count, outbound volume and outbound value.
priority_summary = (
    df.groupby("Inventory_Priority")
    .agg(
        SKU_Count=("ID", "count"),
        Total_Outbound=("Total outbound", "sum"),
        Outbound_Value=("Outbound_Value_Proxy", "sum")
    )
    .reset_index()
)

priority_summary["SKU_Share_%"] = (
    priority_summary["SKU_Count"] /
    len(df) * 100
)

priority_summary["Outbound_Volume_Share_%"] = (
    priority_summary["Total_Outbound"] /
    total_volume * 100
)

priority_summary["Outbound_Value_Share_%"] = (
    priority_summary["Outbound_Value"] /
    total_value * 100
)

print("\nPriority impact summary:")
print(
    priority_summary.sort_values(
        "Outbound_Value",
        ascending=False
    ).to_string(index=False)
)

# PARETO ANALYSIS
# This checks how many active SKUs are responsible for reaching 80% of the total outbound value.
pareto_df = df[
    df["Activity_Status"] == "Active"
].copy()

pareto_df = pareto_df.sort_values(
    "Outbound_Value_Proxy",
    ascending=False
).reset_index(drop=True)

pareto_df["Cumulative_Value"] = (
    pareto_df["Outbound_Value_Proxy"].cumsum()
)

pareto_df["Cumulative_Value_%"] = (
    pareto_df["Cumulative_Value"] /
    pareto_df["Outbound_Value_Proxy"].sum()
) * 100

pareto_df["SKU_%"] = (
    (pareto_df.index + 1) /
    len(pareto_df)
) * 100

pareto_80 = pareto_df[
    pareto_df["Cumulative_Value_%"] >= 80
].iloc[0]
 
sku_count_80 = int(
    pareto_80.name + 1
)

sku_percentage_80 = (
    pareto_80["SKU_%"]
)
 
print("\nPareto analysis:")

print(
    "Active SKUs:",
    len(pareto_df)
)

print(
    "SKUs needed to reach 80% of outbound value:",
    sku_count_80
)

print(
    "Percentage of active SKUs:",
    round(sku_percentage_80, 2),
    "%"
)

print(
    "Cumulative value:",
    round(
        pareto_80["Cumulative_Value_%"],
        2
    ),
    "%"
)

# FINAL IMPORTANT SKU LIST
# This gives the final list of SKUs that management should focus on first.
important_skus = df[
    df["Inventory_Priority"].isin(
        [
            "CRITICAL",
            "HIGH PRIORITY",
            "MEDIUM-HIGH PRIORITY"
        ]
    )
].copy()

print("\nImportant SKUs requiring attention:")

print(
    important_skus[
        [
            "ID",
            "Unitprice",
            "Outbound number",
            "Total outbound",
            "Outbound_Value_Proxy",
            "SKU_Segment",
            "ABC_Class",
            "Inventory_Priority"
        ]
    ]
    .sort_values(
        "Outbound_Value_Proxy",
        ascending=False
    )
    .head(30)
)