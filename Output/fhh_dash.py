# Import packages
from dash import Dash, html, dash_table, dcc, callback, Output, Input
import pandas as pd
import dash_mantine_components as dmc

# Read Excel files
future_rev = pd.read_excel('future_rev.xlsx')
past_rev = pd.read_excel('past_rev.xlsx')
future_occ = pd.read_excel('future_occ.xlsx')
past_occ = pd.read_excel('past_occ.xlsx')
future_adr = pd.read_excel('future_adr.xlsx') 
past_adr = pd.read_excel('past_adr.xlsx')
future_revpab = pd.read_excel('future_revpab.xlsx')
past_revpab = pd.read_excel('past_revpab.xlsx')
future_book = pd.read_excel('future_book.xlsx')
past_book = pd.read_excel('past_book.xlsx')

# Initialize the app - incorporate a Dash Mantine theme
external_stylesheets = [dmc.theme.DEFAULT_COLORS]
app = Dash(__name__, external_stylesheets=external_stylesheets)

# Define the conditional formatting styles
style_rev = []
style_occ = []
style_adr = []
style_revpab = []
style_book = []

for col in future_rev.columns[2:33]:
    style_rev.append(
        {
            'if': {
                'column_id': str(col),
                'filter_query': '{{{}}} = 0'.format(col)
            },
            'backgroundColor': 'white',
            'color': 'white'
        }
    )
    style_rev.append(
        {
            'if': {
                'column_id': str(col),
                'filter_query': '{{{}}} >= 1 && {{{}}} <= 500'.format(col, col)
            },
            'backgroundColor': '#ea5545',
            'color': 'white'
        }
    )
    style_rev.append(
        {
            'if': {
                'column_id': str(col),
                'filter_query': '{{{}}} >= 501 && {{{}}} <= 1000'.format(col, col)
            },
            'backgroundColor': '#ef9b20',
            'color': 'white'
        }
    )
    style_rev.append(
        {
            'if': {
                'column_id': str(col),
                'filter_query': '{{{}}} >= 1001 && {{{}}} <= 1500'.format(col, col)
                    },
            'backgroundColor': '#e3d219',
            'color': 'white'
        }
    )
    style_rev.append(
        {
            'if': {
                'column_id': str(col),
                'filter_query': '{{{}}} >= 1501 && {{{}}} <= 2000'.format(col, col)
                    },
            'backgroundColor': '#bdcf32',
            'color': 'white'
        }
    )
    style_rev.append(
        {
            'if': {
                'column_id': str(col),
                'filter_query': '{{{}}} >= 2001 && {{{}}} <= 3000'.format(col, col)
                    },
            'backgroundColor': '#87bc45',
            'color': 'white'
        }
    )
    style_rev.append(
        {
            'if': {
                'column_id': str(col),
                'filter_query': '{{{}}} >= 3000'.format(col)
                    },
            'backgroundColor': '#364b1b',
            'color': 'white'
        }
    )

for col in future_rev.columns[2:33]:
    style_occ.append(
        {
            'if': {
                'column_id': str(col),
                'filter_query': '{{{}}} = 0'.format(col)
            },
            'backgroundColor': 'white',
            'color': 'white'
        }
    )
    style_occ.append(
        {
            'if': {
                'column_id': str(col),
                'filter_query': '{{{}}} >= .01 && {{{}}} <= .2'.format(col, col)
            },
            'backgroundColor': '#ea5545',
            'color': 'white'
        }
    )
    style_occ.append(
        {
            'if': {
                'column_id': str(col),
                'filter_query': '{{{}}} >= .21 && {{{}}} <= .4'.format(col, col)
            },
            'backgroundColor': '#ef9b20',
            'color': 'white'
        }
    )
    style_occ.append(
        {
            'if': {
                'column_id': str(col),
                'filter_query': '{{{}}} >= .41 && {{{}}} <= .6'.format(col, col)
                    },
            'backgroundColor': '#e3d219',
            'color': 'white'
        }
    )
    style_occ.append(
        {
            'if': {
                'column_id': str(col),
                'filter_query': '{{{}}} >= .61 && {{{}}} <= .8'.format(col, col)
                    },
            'backgroundColor': '#bdcf32',
            'color': 'white'
        }
    )
    style_occ.append(
        {
            'if': {
                'column_id': str(col),
                'filter_query': '{{{}}} >= .8 && {{{}}} <= .85'.format(col, col)
                    },
            'backgroundColor': '#87bc45',
            'color': 'white'
        }
    )
    style_occ.append(
        {
            'if': {
                'column_id': str(col),
                'filter_query': '{{{}}} >= .86'.format(col)
                    },
            'backgroundColor': '#364b1b',
            'color': 'white'
        }
    )

for col in future_rev.columns[2:33]:
    style_adr.append(
        {
            'if': {
                'column_id': str(col),
                'filter_query': '{{{}}} = 0'.format(col)
            },
            'backgroundColor': 'white',
            'color': 'white'
        }
    )
    style_adr.append(
        {
            'if': {
                'column_id': str(col),
                'filter_query': '{{{}}} >= 1 && {{{}}} <= 40'.format(col, col)
            },
            'backgroundColor': '#ea5545',
            'color': 'white'
        }
    )
    style_adr.append(
        {
            'if': {
                'column_id': str(col),
                'filter_query': '{{{}}} >= 40.01 && {{{}}} <= 45'.format(col, col)
            },
            'backgroundColor': '#ef9b20',
            'color': 'white'
        }
    )
    style_adr.append(
        {
            'if': {
                'column_id': str(col),
                'filter_query': '{{{}}} >= 45.1 && {{{}}} <= 50'.format(col, col)
                    },
            'backgroundColor': '#e3d219',
            'color': 'white'
        }
    )
    style_adr.append(
        {
            'if': {
                'column_id': str(col),
                'filter_query': '{{{}}} >= 50.1 && {{{}}} <= 55'.format(col, col)
                    },
            'backgroundColor': '#bdcf32',
            'color': 'white'
        }
    )
    style_adr.append(
        {
            'if': {
                'column_id': str(col),
                'filter_query': '{{{}}} >= 55.1 && {{{}}} <= 65'.format(col, col)
                    },
            'backgroundColor': '#87bc45',
            'color': 'white'
        }
    )
    style_adr.append(
        {
            'if': {
                'column_id': str(col),
                'filter_query': '{{{}}} >= 65.1'.format(col)
                    },
            'backgroundColor': '#364b1b',
            'color': 'white'
        }
    )

for col in future_rev.columns[2:33]:
    style_revpab.append(
        {
            'if': {
                'column_id': str(col),
                'filter_query': '{{{}}} = 0'.format(col)
            },
            'backgroundColor': 'white',
            'color': 'white'
        }
    )
    style_revpab.append(
        {
            'if': {
                'column_id': str(col),
                'filter_query': '{{{}}} >= 1 && {{{}}} <= 10'.format(col, col)
            },
            'backgroundColor': '#ea5545',
            'color': 'white'
        }
    )
    style_revpab.append(
        {
            'if': {
                'column_id': str(col),
                'filter_query': '{{{}}} >= 10.1 && {{{}}} <= 20'.format(col, col)
            },
            'backgroundColor': '#ef9b20',
            'color': 'white'
        }
    )
    style_revpab.append(
        {
            'if': {
                'column_id': str(col),
                'filter_query': '{{{}}} >= 20.1 && {{{}}} <= 30'.format(col, col)
                    },
            'backgroundColor': '#e3d219',
            'color': 'white'
        }
    )
    style_revpab.append(
        {
            'if': {
                'column_id': str(col),
                'filter_query': '{{{}}} >= 30.1 && {{{}}} <= 40'.format(col, col)
                    },
            'backgroundColor': '#bdcf32',
            'color': 'white'
        }
    )
    style_revpab.append(
        {
            'if': {
                'column_id': str(col),
                'filter_query': '{{{}}} >= 40.1 && {{{}}} <= 50'.format(col, col)
                    },
            'backgroundColor': '#87bc45',
            'color': 'white'
        }
    )
    style_revpab.append(
        {
            'if': { 
                'column_id': str(col),
                'filter_query': '{{{}}} >= 50.1'.format(col)
                    },
            'backgroundColor': '#364b1b',
            'color': 'white'
        }
    )
for col in future_book.columns[3:35]:
    style_book.append(
        {
            'if': {
                'column_id': str(col),
                'filter_query': '{{{}}} = 0'.format(col)
            },
            'backgroundColor': 'white',
            'color': 'white'
        }
    )
    style_book.append(
        {
            'if': {
                'column_id': str(col),
                'filter_query': '{{{}}} >= .01 && {{{}}} <= .2'.format(col, col)
            },
            'backgroundColor': '#ea5545',
            'color': 'white'
        }
    )
    style_book.append(
        {
            'if': {
                'column_id': str(col),
                'filter_query': '{{{}}} >= .21 && {{{}}} <= .4'.format(col, col)
            },
            'backgroundColor': '#ef9b20',
            'color': 'white'
        }
    )
    style_book.append(
        {
            'if': {
                'column_id': str(col),
                'filter_query': '{{{}}} >= .41 && {{{}}} <= .6'.format(col, col)
                    },
            'backgroundColor': '#e3d219',
            'color': 'white'
        }
    )
    style_book.append(
        {
            'if': {
                'column_id': str(col),
                'filter_query': '{{{}}} >= .61 && {{{}}} <= .8'.format(col, col)
                    },
            'backgroundColor': '#bdcf32',
            'color': 'white'
        }
    )
    style_book.append(
        {
            'if': {
                'column_id': str(col),
                'filter_query': '{{{}}} >= .8 && {{{}}} <= .82'.format(col, col)
                    },
            'backgroundColor': '#87bc45',
            'color': 'white'
        }
    )
    style_book.append(
        {
            'if': {
                'column_id': str(col),
                'filter_query': '{{{}}} >= .86'.format(col)
                    },
            'backgroundColor': '#364b1b',
            'color': 'white'
        }
    )


# Round the datasets to desired decimal places
future_rev = future_rev.round(0)
past_rev = past_rev.round(0)
future_occ = future_occ.round(2)
past_occ = past_occ.round(2) 
future_adr = future_adr.round(0)
past_adr = past_adr.round(0)
future_revpab = future_revpab.round(0)
past_revpab = past_revpab.round(0)
future_book = future_book.round(2)
past_book = past_book.round(2)


# App layout
app.layout = dmc.Container(
    [
        dmc.Title('Firehouse Hostel Dashboard', color="blue", size="h2", style={'text-align': 'center'}),
        dmc.Grid(
            [
                dmc.Col(
                    [
                        dmc.RadioGroup(
                            [
                                dmc.Radio('Future Rev', value='future_rev'),
                                dmc.Radio('Past Rev', value='past_rev'),
                                dmc.Radio('Future OCC', value='future_occ'),
                                dmc.Radio('Past OCC', value='past_occ'),
                                dmc.Radio('Future ADR', value='future_adr'),
                                dmc.Radio('Past ADR', value='past_adr'),
                                dmc.Radio('Future RevPAB', value='future_revpab'),
                                dmc.Radio('Past RevPAB', value='past_revpab'),
                                dmc.Radio('Future Organic %', value='future_book'),
                                dmc.Radio('Past Organic %', value='past_book'),

                            ],
                            id='my-dmc-radio-item-top',
                            value='future_rev',
                            size="sm",
                            style={'text-align': 'center'}
                        ),
                        dash_table.DataTable(
                            id='data-table-top',
                            data=[],
                            columns=[{"name": str(col), "id": str(col)} for col in future_rev.columns],
                            page_size=13,
                            style_table={'overflowX': 'auto', 'text-align': 'center'},
                            style_header={'fontWeight': 'bold', 'text-align': 'center'},
                            style_data_conditional=[],  # Initially empty, will be dynamically set in callback
                        )
                    ],
                    span=12,
                    className='text-center'
                ),
            dmc.Col(
                    [
                        dmc.RadioGroup(
                            [
                                dmc.Radio('Future Rev', value='future_rev'),
                                dmc.Radio('Past Rev', value='past_rev'),
                                dmc.Radio('Future OCC', value='future_occ'),
                                dmc.Radio('Past OCC', value='past_occ'),
                                dmc.Radio('Future ADR', value='future_adr'),
                                dmc.Radio('Past ADR', value='past_adr'),
                                dmc.Radio('Future RevPAB', value='future_revpab'),
                                dmc.Radio('Past RevPAB', value='past_revpab'),
                                dmc.Radio('Future Organic %', value='future_book'),
                                dmc.Radio('Past Organic %', value='past_book'),

                            ],
                            id='my-dmc-radio-item-mid',
                            value='past_rev',
                            size="sm",
                            style={'text-align': 'center'}
                        ),
                        dash_table.DataTable(
                            id='data-table-mid',
                            data=[],
                            columns=[{"name": str(col), "id": str(col)} for col in future_rev.columns],
                            page_size=13,
                            style_table={'overflowX': 'auto', 'text-align': 'center'},
                            style_header={'fontWeight': 'bold', 'text-align': 'center'},
                            style_data_conditional=[],  # Initially empty, will be dynamically set in callback
                        )
                    ],
                    span=12,
                    className='text-center'
                )
            ],
            style={'background-color': '#f5f5f5'}  # Light gray background color
        ),
    ],
    fluid=True,
)

# Callback to update the data table
@app.callback(
    [Output(component_id='data-table-top', component_property='data'),
     Output(component_id='data-table-top', component_property='style_data_conditional')],
    [Output(component_id='data-table-mid', component_property='data'),
     Output(component_id='data-table-mid', component_property='style_data_conditional')],
    [Input(component_id='my-dmc-radio-item-top', component_property='value')],
    [Input(component_id='my-dmc-radio-item-mid', component_property='value')]
)
def update_data_table(value_top, value_middle):
    if value_top == 'future_rev':
        data_top = future_rev.to_dict('records')
        style_conditional = style_rev
    elif value_top == 'past_rev':
        data_top = past_rev.to_dict('records')
        style_conditional = style_rev
    elif value_top == 'future_occ':
        data_top = future_occ.to_dict('records')
        style_conditional = style_occ
    elif value_top == 'past_occ':
        data_top = past_occ.to_dict('records')
        style_conditional = style_occ
    elif value_top == 'future_adr':
        data_top = future_adr.to_dict('records')
        style_conditional = style_adr
    elif value_top == 'past_adr':
        data_top = past_adr.to_dict('records')
        style_conditional = style_adr
    elif value_top == 'future_revpab':
        data_top = future_revpab.to_dict('records')
        style_conditional = style_revpab
    elif value_top == 'past_revpab':
        data_top = past_revpab.to_dict('records')
        style_conditional = style_revpab
    elif value_top == 'future_book':
        data_top = future_book.to_dict('records')
        style_conditional = style_book
    elif value_top == 'past_book':
        data_top = past_book.to_dict('records')
        style_conditional = style_book
    
    if value_middle == 'future_rev':
        data_mid = future_rev.to_dict('records')
        style_conditional_1 = style_rev
    elif value_middle == 'past_rev':
        data_mid = past_rev.to_dict('records')
        style_conditional_1 = style_rev
    elif value_middle == 'future_occ':
        data_mid = future_occ.to_dict('records')
        style_conditional_1 = style_occ
    elif value_middle == 'past_occ':
        data_mid = past_occ.to_dict('records')
        style_conditional_1 = style_occ
    elif value_middle == 'future_adr':
        data_mid = future_adr.to_dict('records')
        style_conditional_1 = style_adr
    elif value_middle == 'past_adr':
        data_mid = past_adr.to_dict('records')
        style_conditional_1 = style_adr
    elif value_middle == 'future_revpab':
        data_mid = future_revpab.to_dict('records')
        style_conditional_1 = style_revpab
    elif value_middle == 'past_revpab':
        data_mid = past_revpab.to_dict('records')
        style_conditional_1 = style_revpab
    elif value_middle == 'future_book':
        data_mid = future_book.to_dict('records')
        style_conditional_1 = style_book
    elif value_middle == 'past_book':
        data_mid = past_book.to_dict('records')
        style_conditional_1 = style_book

    return data_top, style_conditional, data_mid, style_conditional_1  

# Run the App
server = app.server

if __name__ == '__main__':
    app.run_server(debug=True)