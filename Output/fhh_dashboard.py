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
                'backgroundColor': '#ea5545',
                'color': 'white'
            }
        )

        style.append(
            {
                'if': {
                    'column_id': str(col),
                    'filter_query': '{{{}}} >= {} && {{{}}} < {} && {{{}}} != ""'.format(col, quantiles[0.2], col, quantiles[0.4], col)
                },
                'backgroundColor': '#ef9b20',
                'color': 'white'
            }
        )

        style.append(
            {
                'if': {
                    'column_id': str(col),
                    'filter_query': '{{{}}} >= {} && {{{}}} < {} && {{{}}} != ""'.format(col, quantiles[0.4], col, quantiles[0.6], col)
                },
                'backgroundColor': '#ede15b',
                'color': 'black'
            }
        )

        style.append(
            {
                'if': {
                    'column_id': str(col),
                    'filter_query': '{{{}}} >= {} && {{{}}} < {} && {{{}}} != ""'.format(col, quantiles[0.6], col, quantiles[0.8], col)
                },
                'backgroundColor': '#bdcf32',
                'color': 'white'
            }
        )

        style.append(
            {
                'if': {
                    'column_id': str(col),
                    'filter_query': '{{{}}} >= {} && {{{}}} != ""'.format(col, quantiles[0.8], col)
                },
                'backgroundColor': '#87bc45',
                'color': 'white'
            }
        )

# App layout
app.layout = dmc.Container(
    [
        dmc.Title('Firehouse Hostel Dashboard', color="blue", size="h2", style={'text-align': 'center'}),
        dmc.Grid(
            [
                dmc.Col(
                    [
                        dmc.RadioGroup(
                            [dmc.Radio(i, value=i) for i in ['future_rev', 'future_occ', 'past_rev', 'past_occ']],
                            id='my-dmc-radio-item-top',
                            value='future_rev',
                            size="sm",
                            style={'text-align': 'center'}
                        ),
                        dash_table.DataTable(
                            id='data-table-top',
                            data=[],
                            columns=[{"name": str(col), "id": str(col)} for col in future_rev.columns],
                            page_size=12,
                            style_table={'overflowX': 'auto', 'text-align': 'center'},
                            style_header={'fontWeight': 'bold', 'text-align': 'center'},
                            style_data_conditional=style
                        )
                    ],
                    span=12,
                    className='text-center'
                ),
                dmc.Col(
                    [
                        dmc.RadioGroup(
                            [dmc.Radio(i, value=i) for i in ['future_rev', 'future_occ', 'past_rev', 'past_occ']],
                            id='my-dmc-radio-item-middle',
                            value='future_occ',
                            size="sm",
                            style={'text-align': 'center'}
                        ),
                        dash_table.DataTable(
                            id='data-table-middle',
                            data=[],
                            columns=[{"name": str(col), "id": str(col)} for col in future_rev.columns],
                            page_size=12,
                            style_table={'overflowX': 'auto', 'text-align': 'center'},
                            style_header={'fontWeight': 'bold', 'text-align': 'center'},
                            style_data_conditional=style
                        )
                    ],
                    span=12,
                    className='text-center'
                ),
                dmc.Col(
                    [
                        dmc.RadioGroup(
                            [dmc.Radio(i, value=i) for i in ['future_rev', 'future_occ', 'past_rev', 'past_occ']],
                            id='my-dmc-radio-item-bottom',
                            value='past_rev',
                            size="sm",
                            style={'text-align': 'center'}
                        ),
                        dash_table.DataTable(
                            id='data-table-bottom',
                            data=[],
                            columns=[{"name": str(col), "id": str(col)} for col in future_rev.columns],
                            page_size=12,
                            style_table={'overflowX': 'auto', 'text-align': 'center'},
                            style_header={'fontWeight': 'bold', 'text-align': 'center'},
                            style_data_conditional=style
                        )
                    ],
                    span=12,
                    className='text-center'
                ),
            ],
            style={'background-color': '#f5f5f5'}  # Light gray background color
        ),
    ],
    fluid=True,
)


# Callback to update the data table
@app.callback(
    [Output(component_id='data-table-top', component_property='data'),
     Output(component_id='data-table-middle', component_property='data'),
     Output(component_id='data-table-bottom', component_property='data')],
    [Input(component_id='my-dmc-radio-item-top', component_property='value'),
     Input(component_id='my-dmc-radio-item-middle', component_property='value'),
     Input(component_id='my-dmc-radio-item-bottom', component_property='value')]
)
def update_data_table(value_top, value_middle, value_bottom):
    if value_top == 'future_rev':
        data_top = future_rev.to_dict('records')
    elif value_top == 'future_occ':
        data_top = future_occ.to_dict('records')
    elif value_top == 'past_rev':
        data_top = past_rev.to_dict('records')
    elif value_top == 'past_occ':
        data_top = past_occ.to_dict('records')
    else:
        data_top = []

    if value_middle == 'future_rev':
        data_middle = future_rev.to_dict('records')
    elif value_middle == 'future_occ':
        data_middle = future_occ.to_dict('records')
    elif value_middle == 'past_rev':
        data_middle = past_rev.to_dict('records')
    elif value_middle== 'past_occ':
        data_middle = past_occ.to_dict('records')
    else:
        data_middle = []

    if value_bottom == 'future_rev':
        data_bottom = future_rev.to_dict('records')
    elif value_bottom == 'future_occ':
        data_bottom = future_occ.to_dict('records')
    elif value_bottom == 'past_rev':
        data_bottom = past_rev.to_dict('records')
    elif value_bottom == 'past_occ':
        data_bottom = past_occ.to_dict('records')
    else:
        data_bottom = []

    return data_top, data_middle, data_bottom



# Run the App
if __name__ == '__main__':
    app.run_server(debug=True)

