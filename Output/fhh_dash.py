# Import packages
from dash import Dash, html, dash_table, dcc, callback, Output, Input
import pandas as pd
import dash_mantine_components as dmc

# Read Excel files
future_rev = pd.read_excel('future_rev.xlsx')
future_occ = pd.read_excel('future_occ.xlsx')
past_rev = pd.read_excel('past_rev.xlsx')
past_occ = pd.read_excel('past_occ.xlsx')
past_adr = pd.read_excel('past_adr.xlsx')
future_adr = pd.read_excel('future_adr.xlsx')
past_revpab = pd.read_excel('past_revpab.xlsx')
future_revpab = pd.read_excel('future_revpab.xlsx')

# Initialize the app - incorporate a Dash Mantine theme
external_stylesheets = [dmc.theme.DEFAULT_COLORS]
app = Dash(__name__, external_stylesheets=external_stylesheets)

# Define the conditional formatting styles
style = []

datasets = {
    'future_rev': future_rev,
    'future_occ': future_occ,
    'past_rev': past_rev,
    'past_occ': past_occ,
    'past_adr': past_adr,
    'future_adr': future_adr,
    'past_revpab': past_revpab,
    'future_revpab': future_revpab
}

for dataset_name, dataset in datasets.items():
    for col in dataset.columns[1:32]:
        if dataset_name == 'future_rev' or dataset_name == 'past_rev':
            style.append(
                {
                    'if': {
                        'column_id': str(col),
                        'filter_query': '{{{}}} >= 0.01 && {{{}}} <= 500'.format(col, col)
                    },
                    'backgroundColor': '#ea5545',
                    'color': 'white'
                }
            )

            # Add formatting conditions for different ranges
            style.append(
                {
                    'if': {
                        'column_id': str(col),
                        'filter_query': '{{{}}} >= 501 && {{{}}} <= 1000'.format(col, col)
                    },
                    'backgroundColor': '#ef9b20',
                    'color': 'white'
                }
            )

            # Add formatting conditions for different ranges
            style.append(
                {
                    'if': {
                        'column_id': str(col),
                        'filter_query': '{{{}}} >= 1001 && {{{}}} <= 1500'.format(col, col)
                    },
                    'backgroundColor': '#ede15b',
                    'color': 'black'
                }
            )

            # Add formatting conditions for different ranges
            style.append(
                {
                    'if': {
                        'column_id': str(col),
                        'filter_query': '{{{}}} >= 1501 && {{{}}} <= 2000'.format(col, col)
                    },
                    'backgroundColor': '#bdcf32',
                    'color': 'white'
                }
            )

            # Add formatting conditions for different ranges
            style.append(
                {
                    'if': {
                        'column_id': str(col),
                        'filter_query': '{{{}}} >= 2000'.format(col)
                    },
                    'backgroundColor': '#87bc45',
                    'color': 'white'
                }
            )

        elif dataset_name == 'future_occ' or dataset_name == 'past_occ':
            style.append(
                {
                    'if': {
                        'column_id': str(col),
                        'filter_query': '{{{}}} >= 0.01 && {{{}}} <= 0.2'.format(col, col)
                    },
                    'backgroundColor': '#ea5545',
                    'color': 'white'
                }
            )

            # Add formatting conditions for different ranges
            style.append(
                {
                    'if': {
                        'column_id': str(col),
                        'filter_query': '{{{}}} >= 0.21 && {{{}}} <= 0.4'.format(col, col)
                    },
                    'backgroundColor': '#ef9b20',
                    'color': 'white'
                }
            )

            # Add formatting conditions for different ranges
            style.append(
                {
                    'if': {
                        'column_id': str(col),
                        'filter_query': '{{{}}} >= 0.41 && {{{}}} <= 0.6'.format(col, col)
                    },
                    'backgroundColor': '#ede15b',
                    'color': 'black'
                }
            )

            # Add formatting conditions for different ranges
            style.append(
                {
                    'if': {
                        'column_id': str(col),
                        'filter_query': '{{{}}} >= 0.61 && {{{}}} <= 0.8'.format(col, col)
                    },
                    'backgroundColor': '#bdcf32',
                    'color': 'white'
                }
            )

            # Add formatting conditions for different ranges
            style.append(
                {
                    'if': {
                        'column_id': str(col),
                        'filter_query': '{{{}}} >= 0.81 && {{{}}} <= 1'.format(col, col)
                    },
                    'backgroundColor': '#87bc45',
                    'color': 'white'
                }
            )
        elif dataset_name == 'future_adr' or dataset_name == 'past_adr'or dataset_name == 'past_revpab'or dataset_name == 'past_revpab':
            style.append(
                {
                    'if': {
                        'column_id': str(col),
                        'filter_query': '{{{}}} >= 0.01 && {{{}}} <= 10'.format(col, col)
                    },
                    'backgroundColor': '#ea5545',
                    'color': 'white'
                }
            )

            # Add formatting conditions for different ranges
            style.append(
                {
                    'if': {
                        'column_id': str(col),
                        'filter_query': '{{{}}} >= 10.01 && {{{}}} <= 20'.format(col, col)
                    },
                    'backgroundColor': '#ef9b20',
                    'color': 'white'
                }
            )

            # Add formatting conditions for different ranges
            style.append(
                {
                    'if': {
                        'column_id': str(col),
                        'filter_query': '{{{}}} >= 20.01 && {{{}}} <= 30'.format(col, col)
                    },
                    'backgroundColor': '#ede15b',
                    'color': 'black'
                }
            )

            # Add formatting conditions for different ranges
            style.append(
                {
                    'if': {
                        'column_id': str(col),
                        'filter_query': '{{{}}} >= 30.01 && {{{}}} <= 40'.format(col, col)
                    },
                    'backgroundColor': '#bdcf32',
                    'color': 'white'
                }
            )

            # Add formatting conditions for different ranges
            style.append(
                {
                    'if': {
                        'column_id': str(col),
                        'filter_query': '{{{}}} >= 40.01 && {{{}}} <= 1000'.format(col, col)
                    },
                    'backgroundColor': '#87bc45',
                    'color': 'white'
                }
            )
        # Additional style for 0 values
        style.append(
            {
                'if': {
                    'column_id': str(col),
                    'filter_query': '{{{}}} = 0'.format(col)
                },
                'backgroundColor': 'white',
                'color': 'white'
            }
        )

# Round the datasets to desired decimal places
future_rev = future_rev.round(0)
future_occ = future_occ.round(2)
past_rev = past_rev.round(0)
past_occ = past_occ.round(2)
past_adr = past_adr.round(0)
future_adr = future_adr.round(0)
past_revpab = past_revpab.round(0)
future_revpab = future_revpab.round(0)

# App layout
app.layout = dmc.Container(
    [
        dmc.Title('Firehouse Hostel Dashboard', color="blue", size="h2", style={'text-align': 'center'}),
        dmc.Grid(
            [
                dmc.Col(
                    [
                        dmc.RadioGroup(
                            [dmc.Radio(i, value=i) for i in ['future_rev', 'future_occ', 'past_rev', 'past_occ',
                                                             'past_adr', 'future_adr', 'past_revpab', 'future_revpab']],
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
                            [dmc.Radio(i, value=i) for i in ['future_rev', 'future_occ', 'past_rev', 'past_occ',
                                                             'past_adr', 'future_adr', 'past_revpab', 'future_revpab']],
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
                            [dmc.Radio(i, value=i) for i in ['future_rev', 'future_occ', 'past_rev', 'past_occ',
                                                             'past_adr', 'future_adr', 'past_revpab', 'future_revpab']],
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
    elif value_top == 'past_adr':
        data_top = past_adr.to_dict('records')
    elif value_top == 'future_adr':
        data_top = future_adr.to_dict('records')
    elif value_top == 'past_revpab':
        data_top = past_revpab.to_dict('records')
    elif value_top == 'future_revpab':
        data_top = future_revpab.to_dict('records')
    
    if value_middle == 'future_rev':
        data_middle = future_rev.to_dict('records')
    elif value_middle == 'future_occ':
        data_middle = future_occ.to_dict('records')
    elif value_middle == 'past_rev':
        data_middle = past_rev.to_dict('records')
    elif value_middle == 'past_occ':
        data_middle = past_occ.to_dict('records')
    elif value_middle == 'past_adr':
        data_middle = past_adr.to_dict('records')
    elif value_middle == 'future_adr':
        data_middle = future_adr.to_dict('records')
    elif value_middle == 'past_revpab':
        data_middle = past_revpab.to_dict('records')
    elif value_middle == 'future_revpab':
        data_middle = future_revpab.to_dict('records')

    if value_bottom == 'future_rev':
        data_bottom = future_rev.to_dict('records')
    elif value_bottom == 'future_occ':
        data_bottom = future_occ.to_dict('records')
    elif value_bottom == 'past_rev':
        data_bottom = past_rev.to_dict('records')
    elif value_bottom == 'past_occ':
        data_bottom = past_occ.to_dict('records')
    elif value_bottom == 'past_adr':
        data_bottom = past_adr.to_dict('records')
    elif value_bottom == 'future_adr':
        data_bottom = future_adr.to_dict('records')
    elif value_bottom == 'past_revpab':
        data_bottom = past_revpab.to_dict('records')
    elif value_bottom == 'future_revpab':
        data_bottom = future_revpab.to_dict('records')

    return data_top, data_middle, data_bottom

# Run the App
server = app.server

if __name__ == '__main__':
    app.run_server(debug=True)

