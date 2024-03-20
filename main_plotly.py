import pandas as pd
import plotly.express as px
from matplotlib import pyplot as plt

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

df = pd.DataFrame(list(pg_spend.items()), columns=['Faculty', 'Spend'])

# Create the plotly bar plot
fig = px.bar(df, x='Faculty', y='Spend', color='Spend', color_continuous_scale=['#0d41d1', '#0d41d1'])
fig.update_layout(title_text="PG Spend by Faculty")
fig.show()


df = pd.DataFrame(list(ug_spend.items()), columns=['Faculty', 'Spend'])

# Create the plotly bar plot
fig = px.bar(df, x='Faculty', y='Spend', color='Spend', color_continuous_scale=['#ff2305', '#ff2305'])
fig.update_layout(title_text="UG Spend by Faculty")
fig.show()

# Plot spend per agency
data = []
for faculty, agency_spends in agency_spend.items():
    for agency, spend in agency_spends.items():
        data.append([agency, faculty, spend])
df = pd.DataFrame(data, columns=['Agency', 'Faculty', 'Spend'])

# Create the plotly bar plot
fig = px.bar(df, x='Faculty', y='Spend', color='Agency', title='Spend by Agency')
fig.show()

"""
Insights:
- BUS Faculty has the highest spend for PG, close to them are `HEA GSH`, `LAW`, and `SCI`
- FEIT Faculty has the highest spend for UG
- Most of the spending for agencies are done by FEIT and SCI faculties. They use the most budget for agencies
- BUS is the only faculty that reported some in-house spending
"""
