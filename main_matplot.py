import pandas as pd
import matplotlib.pyplot as plt

# Read the Excel file
xls = pd.ExcelFile('data/2023_dataset.xlsx')

pg_spend = {}
ug_spend = {}
agency_spend = {}

# Read each sheet in the Excel file individually
for sheet in xls.sheet_names:
    faculties_to_take = [
        'BUS',
        'DAB',
        'FASS',
        'FEIT',
        'HEA GSH',
        'LAW',
        'SCI',
        'TD'
    ]
    if sheet not in faculties_to_take:
        continue
    # Read the sheet into a DataFrame
    df = pd.read_excel(xls, sheet)
    # In each dataframe, filter out "PG Autumn 2024" lines
    pg_spend[sheet] = round(df.query("Campaign.str.contains('PG Aut 2024') == False and `Segment` == 'PG'")['Media spend'].sum(), 2)
    ug_spend[sheet] = round(df.query("Campaign.str.contains('UG Aut 2024') == False and `Segment` == 'UG'")['Media spend'].sum(), 2)
    agency_spend[sheet] = {
        agency_key: round(df.query("`Campaign`.str.contains('UG Aut 2024') == False and `Segment` == 'UG' and `Agency/In-house` == @agency_key")['Media spend'].sum(), 2)
        for agency_key in df.query("`Agency/In-house`.notna()")['Agency/In-house'].unique()
    }

fig, ax = plt.subplots()
plt.title('PG Spend by Faculty')
ax.bar(pg_spend.keys(), pg_spend.values(), color=['#0d41d1'])
plt.show()

fig, ax = plt.subplots()
plt.title('UG Spend by Faculty')
ax.bar(ug_spend.keys(), ug_spend.values(), color=['#ff2305'])
plt.show()

fig, ax = plt.subplots()
plt.title('Spend by Agency')
for agency, spend in agency_spend.items():
    ax.bar(spend.keys(), spend.values(), label=agency)
ax.legend()
plt.show()

"""
Insights:
- BUS Faculty has the highest spend for PG, close to them are `HEA GSH`, `LAW`, and `SCI`
- FEIT Faculty has the highest spend for UG
- Most of the spending for agencies are done by FEIT and SCI faculties. They use the most budget for agencies
- BUS is the only faculty that reported some in-house spending
"""