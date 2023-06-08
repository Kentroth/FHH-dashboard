# Import packages
from dash import Dash, html, dash_table, dcc, callback, Output, Input
import pandas as pd
import numpy as np
import dash_mantine_components as dmc
import seaborn as sns
import openpyxl

future_rev = pd.read_excel('future_rev.xlsx')
future_occ = pd.read_excel('future_occ.xlsx')
past_rev = pd.read_excel('past_rev.xlsx')
past_occ = pd.read_excel('past_occ.xlsx')

future_rev = future_rev.round(0)
future_occ = future_occ.round(2)
past_rev = past_rev.round(0)
past_occ = past_occ.round(2)

# Initialize the app - incorporate a Dash Mantine theme
external_stylesheets = [dmc.theme.DEFAULT_COLORS]
app = Dash(__name__, external_stylesheets=external_stylesheets)

# Define the conditional formatting styles

style = []

datasets = [future_rev, future_occ, past_rev, past_occ]

for dataset in datasets:
    for col in dataset.columns[1:32]:
        non_blank_values = dataset[col].loc[(dataset[col].notna()) & (dataset[col] != '')]
        quantiles = non_blank_values.quantile([0.2, 0.4, 0.6, 0.8])

        style.append(
            {
                'if': {
                    'column_id': str(col),
                    'filter_query': '{{{}}} < {} && {{{}}} != ""'.format(col, quantiles[0.2], col)
                },
                'backgroundColor': 'red',
                'color': 'white'
            }
        )

        style.append(
            {
                'if': {
                    'column_id': str(col),
                    'filter_query': '{{{}}} >= {} && {{{}}} < {} && {{{}}} != ""'.format(col, quantiles[0.2], col, quantiles[0.4], col)
                },
                'backgroundColor': 'pink',
                'color': 'white'
            }
        )

        style.append(
            {
                'if': {
                    'column_id': str(col),
                    'filter_query': '{{{}}} >= {} && {{{}}} < {} && {{{}}} != ""'.format(col, quantiles[0.4], col, quantiles[0.6], col)
                },
                'backgroundColor': 'yellow',
                'color': 'black'
            }
        )

        style.append(
            {
                'if': {
                    'column_id': str(col),
                    'filter_query': '{{{}}} >= {} && {{{}}} < {} && {{{}}} != ""'.format(col, quantiles[0.6], col, quantiles[0.8], col)
                },
                'backgroundColor': 'lightgreen',
                'color': 'white'
            }
        )

        style.append(
            {
                'if': {
                    'column_id': str(col),
                    'filter_query': '{{{}}} >= {} && {{{}}} != ""'.format(col, quantiles[0.8], col)
                },
                'backgroundColor': 'green',
                'color': 'white'
            }
        )

# App layout
app.layout = dmc.Container([
    dmc.Title('Firehouse Hostel Dashboard', color="blue", size="h3"),
    dmc.RadioGroup(
        [dmc.Radio(i, value=i) for i in ['future_rev', 'future_occ', 'past_rev', 'past_occ']],
        id='my-dmc-radio-item',
        value='future_rev',
        size="sm"
    ),
    dmc.Grid([
        dmc.Col([
            dash_table.DataTable(
                id='data-table',
                data=[],
                columns=[{"name": str(col), "id": str(col)} for col in future_rev.columns],
                page_size=12,
                style_table={'overflowX': 'auto'},
                style_header={'fontWeight': 'bold'},
                style_data_conditional=style
            )
        ], span=12),
    ]),

], fluid=True)


# Callback to update the data table
@app.callback(
    Output(component_id='data-table', component_property='data'),
    Input(component_id='my-dmc-radio-item', component_property='value')
)
def update_data_table(value):
    if value == 'future_rev':
        return future_rev.to_dict('records')
    elif value == 'future_occ':
        return future_occ.to_dict('records')
    elif value == 'past_rev':
        return past_rev.to_dict('records')
    elif value == 'past_occ':
        return past_occ.to_dict('records')
    else:
        return []

# Run the App
if __name__ == '__main__':
    app.run_server(debug=True)

