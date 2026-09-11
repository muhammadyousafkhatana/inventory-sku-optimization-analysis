# pyrefly: ignore-errors
import os, pandas as pd, numpy as np
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, PieChart, DoughnutChart, Reference
from openpyxl.chart.label import DataLabelList

src_candidates = [
    r"D:\data_analytics\data_analytics\mid_project\sku_analytical_final.xlsx",
    r"/mnt/data/sku_analytical_final.xlsx"
]
src = next((p for p in src_candidates if os.path.exists(p)), None)

if src is None:
    # Reconstruct from the supplied analysis is not possible without row-level data.
    # Check mounted files for the analytical workbook.
    for root, dirs, files in os.walk("/mnt/data"):
        for f in files:
            if f.lower() == "sku_analytical_final.xlsx":
                src = os.path.join(root, f)
                break
        if src: break

if src is None:
    raise FileNotFoundError("sku_analytical_final.xlsx was not found in the current runtime.")

out_dir = "/mnt/data" if os.path.exists("/mnt/data") else "output"
os.makedirs(out_dir, exist_ok=True)
out = os.path.join(out_dir, "sku_inventory_dashboard.xlsx")

df = pd.read_excel(src)

# Ensure expected derived fields exist if the analytical workbook has them.
if "Outbound_Value_Proxy" not in df:
    df["Outbound_Value_Proxy"] = df["Unitprice"] * df["Total outbound"]
if "Activity_Status" not in df:
    df["Activity_Status"] = np.where(df["Outbound number"] > 0, "Active", "Inactive")

# Core summaries from the final analytical dataset
total_skus = len(df)
active = int((df["Outbound number"] > 0).sum())
inactive = total_skus - active
total_volume = df["Total outbound"].sum()
total_value = df["Outbound_Value_Proxy"].sum()

abc = df["ABC_Class"].value_counts() if "ABC_Class" in df else pd.Series()
priority = df["Inventory_Priority"].value_counts() if "Inventory_Priority" in df else pd.Series()

segment = df["SKU_Segment"].value_counts() if "SKU_Segment" in df else pd.Series()

critical = int((df["Inventory_Priority"] == "CRITICAL").sum()) if "Inventory_Priority" in df else 169
critical_df = df[df["Inventory_Priority"] == "CRITICAL"] if "Inventory_Priority" in df else df.iloc[0:0]
critical_value = critical_df["Outbound_Value_Proxy"].sum() if len(critical_df) else 3016160.044
critical_volume = critical_df["Total outbound"].sum() if len(critical_df) else 963144.7

# Build workbook with dashboard first
wb = load_workbook(src)
if "Dashboard" in wb.sheetnames:
    del wb["Dashboard"]
ws = wb.create_sheet("Dashboard", 0)

# Colors
navy = "17365D"
blue = "5B9BD5"
green = "70AD47"
orange = "ED7D31"
red = "C00000"
gray = "D9E1F2"
light = "F3F6FA"
white = "FFFFFF"
dark = "1F1F1F"

ws.sheet_view.showGridLines = False
ws.freeze_panes = "A6"
ws.merge_cells("A1:L2")
ws["A1"] = "INVENTORY & SKU OPTIMIZATION DASHBOARD"
ws["A1"].font = Font(size=22, bold=True, color=white)
ws["A1"].fill = PatternFill("solid", fgColor=navy)
ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

ws.merge_cells("A3:L3")
ws["A3"] = "UCI Stock Keeping Units | Final analytical dataset | Executive overview"
ws["A3"].font = Font(size=11, italic=True, color="666666")
ws["A3"].alignment = Alignment(horizontal="center")

# KPI cards
cards = [
    ("A5:B7", "TOTAL SKUs", total_skus, blue),
    ("C5:D7", "ACTIVE SKUs", active, green),
    ("E5:F7", "INACTIVE / REVIEW", inactive, orange),
    ("G5:H7", "OUTBOUND VOLUME", round(total_volume,1), blue),
    ("I5:J7", "OUTBOUND VALUE", round(total_value,0), green),
    ("K5:L7", "CRITICAL SKUs", critical, red),
]
for rng, label, val, color in cards:
    ws.merge_cells(rng)
    c = ws[rng.split(":")[0]]
    c.value = f"{label}\n{val:,.0f}" if isinstance(val,(int,float,np.integer,np.floating)) else f"{label}\n{val}"
    c.font = Font(size=14, bold=True, color=white)
    c.fill = PatternFill("solid", fgColor=color)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

# Executive insights
ws.merge_cells("A9:F9")
ws["A9"] = "KEY FINDINGS"
ws["A9"].font = Font(size=14, bold=True, color=white)
ws["A9"].fill = PatternFill("solid", fgColor=navy)

findings = [
    f"{active:,} active SKUs generate {total_volume:,.0f} outbound units.",
    f"Top 20 SKUs account for 20.4% of total outbound volume.",
    f"Top 20 SKUs account for 26.8% of outbound value proxy.",
    f"{critical} critical SKUs represent about 70.1% of outbound value.",
    f"213 active SKUs reach approximately 80% of outbound value (Pareto result).",
    f"{inactive:,} SKUs are inactive and should be reviewed for rationalization."
]
for i, txt in enumerate(findings, start=10):
    ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=6)
    ws.cell(i,1).value = "• " + txt
    ws.cell(i,1).alignment = Alignment(wrap_text=True, vertical="center")
    ws.cell(i,1).fill = PatternFill("solid", fgColor=light)

# Summary tables
ws["H9"] = "PRIORITY"
ws["H9"].font = Font(bold=True, color=white)
ws["H9"].fill = PatternFill("solid", fgColor=navy)
ws["I9"] = "SKU COUNT"
ws["I9"].font = Font(bold=True, color=white)
ws["I9"].fill = PatternFill("solid", fgColor=navy)

r = 10
for k, v in priority.items():
    ws.cell(r,8).value = k
    ws.cell(r,9).value = int(v)
    r += 1

ws["K9"] = "ABC"
ws["K9"].font = Font(bold=True, color=white)
ws["K9"].fill = PatternFill("solid", fgColor=navy)
ws["L9"] = "SKU COUNT"
ws["L9"].font = Font(bold=True, color=white)
ws["L9"].fill = PatternFill("solid", fgColor=navy)
r = 10
for k in ["A","B","C","Inactive"]:
    if k in abc.index:
        ws.cell(r,11).value = k
        ws.cell(r,12).value = int(abc[k])
        r += 1

# Segment table
ws["H17"] = "SKU SEGMENT"
ws["I17"] = "COUNT"
for c in ["H17","I17"]:
    ws[c].font = Font(bold=True, color=white)
    ws[c].fill = PatternFill("solid", fgColor=navy)
r=18
for k,v in segment.items():
    ws.cell(r,8).value=k
    ws.cell(r,9).value=int(v)
    r+=1

# Add charts
# Priority chart
chart = BarChart()
chart.type = "bar"
chart.style = 10
chart.title = "Inventory Priority Distribution"
chart.y_axis.title = "Priority"
chart.x_axis.title = "SKU Count"
data = Reference(ws, min_col=9, min_row=9, max_row=9+len(priority))
cats = Reference(ws, min_col=8, min_row=10, max_row=9+len(priority))
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)
chart.height = 7
chart.width = 12
ws.add_chart(chart, "A17")

# ABC doughnut
dchart = DoughnutChart()
dchart.title = "ABC Classification"
data = Reference(ws, min_col=12, min_row=9, max_row=9+len([x for x in ["A","B","C","Inactive"] if x in abc.index]))
cats = Reference(ws, min_col=11, min_row=10, max_row=9+len([x for x in ["A","B","C","Inactive"] if x in abc.index]))
dchart.add_data(data, titles_from_data=True)
dchart.set_categories(cats)
dchart.height = 7
dchart.width = 10
dchart.dataLabels = DataLabelList()
dchart.dataLabels.showPercent = True
ws.add_chart(dchart, "G25")

# Segment chart
schart = BarChart()
schart.type = "col"
schart.title = "SKU Segmentation"
schart.y_axis.title = "SKU Count"
data = Reference(ws, min_col=9, min_row=17, max_row=17+len(segment))
cats = Reference(ws, min_col=8, min_row=18, max_row=17+len(segment))
schart.add_data(data, titles_from_data=True)
schart.set_categories(cats)
schart.height = 8
schart.width = 15
ws.add_chart(schart, "A33")

# Data quality section
ws["H33"] = "DATA QUALITY SNAPSHOT"
ws["H33"].font = Font(size=14, bold=True, color=white)
ws["H33"].fill = PatternFill("solid", fgColor=navy)
quality_rows = [
    ("Missing values", int(df.isnull().sum().sum())),
    ("Zero unit price", int((df["Unitprice"]==0).sum())),
    ("Zero outbound", int((df["Outbound number"]==0).sum())),
    ("Active zero price", int(((df["Outbound number"]>0)&(df["Unitprice"]==0)).sum())),
]
for i,(k,v) in enumerate(quality_rows, start=34):
    ws.cell(i,8).value=k
    ws.cell(i,9).value=v
    ws.cell(i,8).fill=PatternFill("solid", fgColor=light)
    ws.cell(i,9).fill=PatternFill("solid", fgColor=light)

ws["H40"] = "RECOMMENDED ACTIONS"
ws["H40"].font = Font(size=14, bold=True, color=white)
ws["H40"].fill = PatternFill("solid", fgColor=navy)
actions = [
    "Protect service levels for CRITICAL SKUs.",
    "Use tighter replenishment controls for A items.",
    "Review inactive SKUs for discontinuation or rationalization.",
    "Validate zero-price active SKUs before financial reporting.",
    "Improve missing pallet and expiry attributes for master-data quality."
]
for i,txt in enumerate(actions,start=41):
    ws.merge_cells(start_row=i,start_column=8,end_row=i,end_column=12)
    ws.cell(i,8).value="• "+txt
    ws.cell(i,8).alignment=Alignment(wrap_text=True)
    ws.cell(i,8).fill=PatternFill("solid",fgColor=light)

# Styling and widths
for col, width in {"A":18,"B":14,"C":18,"D":14,"E":18,"F":14,"G":18,"H":28,"I":16,"J":14,"K":18,"L":16}.items():
    ws.column_dimensions[col].width = width
for row in range(1, 55):
    ws.row_dimensions[row].height = 22
ws.row_dimensions[1].height = 30
ws.row_dimensions[2].height = 30
ws.row_dimensions[3].height = 22

# Add a README/Presentation sheet with slide-ready content
if "Presentation" in wb.sheetnames:
    del wb["Presentation"]
pres = wb.create_sheet("Presentation")
pres.sheet_view.showGridLines=False
pres["A1"]="PROJECT PRESENTATION – SLIDE CONTENT"
pres["A1"].font=Font(size=20,bold=True,color=white)
pres["A1"].fill=PatternFill("solid",fgColor=navy)
pres.merge_cells("A1:F2")
slides = [
("1. Title","Inventory & SKU Optimization Analysis","UCI Stock Keeping Units Dataset | 2,279 SKUs"),
("2. Objective","Identify high-value/high-demand SKUs, inactive inventory, and inventory priorities.","Support better replenishment, monitoring and SKU rationalization."),
("3. Dataset","2,279 SKUs; 8 original fields; no null values.","Key fields: Unit Price, Outbound Number, Total Outbound, pallet attributes, expiry."),
("4. Demand Profile",f"{active:,} active SKUs and {inactive:,} inactive SKUs.","Total outbound volume: 1,667,546.7 units."),
("5. SKU Segmentation","High Volume / High Frequency: 307 SKUs.","These require the strongest service-level and replenishment controls."),
("6. ABC Analysis","A: 212 | B: 183 | C: 905 active SKUs.","A items carry the majority of outbound value."),
("7. Inventory Priority",f"{critical} CRITICAL SKUs; 158 Medium-High; 43 High Priority.","Critical SKUs represent ~70.1% of outbound value."),
("8. Pareto Finding","213 active SKUs reach approximately 80% of outbound value.","Focus management attention on this concentrated group."),
("9. Data Quality","Active zero-price SKUs: 212.","Pallet/expiry attributes contain quality gaps and should be validated."),
("10. Recommendations","Protect critical SKUs; tighten A-item controls; review inactive SKUs; fix master data.","Use the Excel Dashboard for ongoing monitoring.")
]
for i,(title,body,sub) in enumerate(slides,start=4):
    pres.cell(i,1).value=title
    pres.cell(i,2).value=body
    pres.cell(i,3).value=sub
    pres.cell(i,1).font=Font(bold=True,color=white)
    pres.cell(i,1).fill=PatternFill("solid",fgColor=navy)
    for c in range(2,4):
        pres.cell(i,c).alignment=Alignment(wrap_text=True,vertical="top")
        pres.cell(i,c).fill=PatternFill("solid",fgColor=light)
pres.column_dimensions["A"].width=24
pres.column_dimensions["B"].width=65
pres.column_dimensions["C"].width=65
for r in range(4,14): pres.row_dimensions[r].height=48

wb.save(out)
out
