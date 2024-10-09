import dash
from dash import dcc, html, dash_table
from dash.dependencies import Input, Output, State
import pandas as pd
import base64
from io import BytesIO, StringIO
import webview
import threading
from datetime import datetime, timedelta, date
import calendar
import plotly.express as px
import time
import io
from screeninfo import get_monitors

external_stylesheets = ['https://fonts.googleapis.com/css2?family=Roboto:wght@400;700&display=swap']

def dealer_dashboard():
    app = dash.Dash(__name__, external_stylesheets=external_stylesheets)
    app.title = "Dealer Activity Analysis"
    
    # Define global variables
    current_year = datetime.today().year
    
    # Modify your app.layout to include the dcc.Dropdown component
    app.layout = html.Div([
        html.H1("Dealer Activity Analysis", style={'text-align': 'center', 'font-family': 'Roboto', 'margin-top': '20px'}),
    
        dcc.Upload(
            id='upload-data',
            children=html.Div([
                'Drag and Drop or ',
                html.A('Select Files')
            ]),
            style={
                'width': '100%',
                'height': '60px',
                'lineHeight': '60px',
                'borderWidth': '1px',
                'borderStyle': 'dashed',
                'borderRadius': '5px',
                'textAlign': 'center',
                'margin': '20px auto'
            },
            multiple=False
        ),
        
        dcc.Loading(
            id="loading-1",
            type="default",
            children=[
                html.Div(id="loading-output-1")
            ]
        ),
        
        dcc.Loading(
            id="loading-2",
            type="default",
            children=[
                html.Div(id="loading-output-2")
            ]
        ),
    
        html.Div(id='output-data-upload'),
    
        # Dropdown to select sheets from uploaded Excel file
        dcc.Dropdown(id='sheet-selector', placeholder='Select a sheet', style={'width': '80%', 'margin': '20px auto'}),
        
        
        html.Div(
            dcc.DatePickerRange(
                id='my-date-picker-range',
                min_date_allowed=datetime(1995, 8, 5),
                max_date_allowed=datetime.today(),
                initial_visible_month=datetime.today(),
                start_date=datetime(datetime.today().year, 1, 1),  # Set the start_date to January 1st of the current year
                end_date=datetime.today(),
                display_format='DD/MM/YY'  # Specify the display format here
            ),
            style={'text-align': 'center', 'margin': '20px auto'}
        ),
        
        html.Div(id='main-content'),
        
        dcc.Store(id='stored-data'),
    ])
    
    # Modify your callback to update data based on the selected sheet
    @app.callback(
        [
            Output('output-data-upload', 'children'),
            Output("loading-output-1", "children"),
            Output('sheet-selector', 'options'),
            Output('stored-data', 'data')
        ],
        [Input('upload-data', 'contents')],
        [State('upload-data', 'filename')]
    )
    def update_upload_output(contents, filename):
        if contents is None:
            return 'No file uploaded yet. Please upload a file.', '', [], None
    
        content_string = contents.split(',')[1]  # Assuming content_string is the second part after splitting by ','
        
        decoded = base64.b64decode(content_string)
        
        # Determine file type and read data accordingly
        if 'xls' in filename:
            xls = pd.ExcelFile(BytesIO(decoded))
            sheet_names = xls.sheet_names
            options = [{'label': sheet, 'value': sheet} for sheet in sheet_names]
            return f'File "{filename}" loaded successfully with {len(sheet_names)} sheets.', '', options, contents
    
        elif 'csv' in filename:
            df = pd.read_csv(BytesIO(decoded))
            return f'File "{filename}" loaded successfully with {len(df)} rows.', '', [], contents
        
        return 'Unsupported file format.', '', [], None
    
    # Callback to update data based on selected sheet
    """    [
        Output('main-content', 'children'),
        Output("loading-output-2", "children")
        ],"""
    @app.callback(
        [
            Output('main-content', 'children'),
            Output("loading-output-2", "children")
        ],
        [
            Input('sheet-selector', 'value'),
            Input('my-date-picker-range', 'start_date'),
            Input('my-date-picker-range', 'end_date'),
        ],
        [
            State('stored-data', 'data'),
            State('upload-data', 'filename')
        ]
    )
    def update_main_content(sheet_name, start_date, end_date, stored_data, filename):
        if stored_data is None or sheet_name is None:
            return dash.no_update, ''
    
        content_type, content_string = stored_data.split(',')
        decoded = base64.b64decode(content_string)
        
        if 'xls' in filename:
            xls = pd.ExcelFile(BytesIO(decoded))
            df = pd.read_excel(xls, sheet_name=sheet_name)
        elif 'csv' in filename:
            df = pd.read_csv(BytesIO(decoded))
        else:
            return 'Unsupported file format.', ''
        
        # Ensure 'Date' column is in datetime format if needed
        if 'Date' in df.columns and not pd.api.types.is_datetime64_any_dtype(df['Date']):
            df['Date'] = pd.to_datetime(df['Date'], errors='coerce', format='%d/%m/%y')
        
        # Filter DataFrame based on date range
        if start_date and end_date:
            # Parse start_date and end_date without time
            start_date = datetime.strptime(start_date.split('T')[0], '%Y-%m-%d').strftime('%d/%m/%y')
            end_date = datetime.strptime(end_date.split('T')[0], '%Y-%m-%d').strftime('%d/%m/%y')
        
            start_date = datetime.strptime(start_date, '%d/%m/%y')
            end_date = datetime.strptime(end_date, '%d/%m/%y')
            # Filter for the current year and previous year's same date range
            current_year_data = df[(df['Date'].dt.year == start_date.year) & 
                                   (df['Date'] >= start_date) & (df['Date'] <= end_date)]
            
            prev_year_data = df[(df['Date'].dt.year == (start_date.year - 1)) & 
                                (df['Date'] >= start_date.replace(year=start_date.year - 1)) & 
                                (df['Date'] <= end_date.replace(year=start_date.year - 1))]
            
            return show_main_content(df, current_year_data, prev_year_data, start_date, end_date), ''
        
        # If no date range is selected, use entire dataset
        return show_main_content(df, None, None, start_date, end_date), ''
    
    
    # Function to display main content
    def show_main_content(df, current_year_data, prev_year_data, start_date, end_date):
        current_year = datetime.now().year
    
        if current_year_data is None and prev_year_data is None:
            current_year_data = df[df['Date'].dt.year == current_year]
            prev_year_data = df[df['Date'].dt.year == (current_year - 1)]
    
        today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        today_full = today.strftime("%d %B %Y")
        
        if start_date and end_date:
            # Filter for the current year and previous year's same date range
            current_year_data = df[(df['Date'] >= start_date) & (df['Date'] <= end_date)]
            
            prev_year_data = df[(df['Date'] >= start_date.replace(year=start_date.year - 1)) & 
                                (df['Date'] <= end_date.replace(year=start_date.year - 1))]
        
        if start_date:
            start_of_the_week = start_date - timedelta(days=start_date.weekday())
            end_of_the_week = start_of_the_week + timedelta(days=6)
            
            prev_year_start_of_week = start_of_the_week - timedelta(days=365)
            prev_year_end_of_the_week = end_of_the_week - timedelta(days=365)
        else:
            start_of_the_week = today - timedelta(days=today.weekday())
            end_of_the_week = start_of_the_week + timedelta(days=6)
            
            prev_year_start_of_week = start_of_the_week - timedelta(days=365)
            prev_year_end_of_the_week = end_of_the_week - timedelta(days=365)
    
        current_week_data = current_year_data[(current_year_data['Date'] >= start_of_the_week) & 
                                              (current_year_data['Date'] <= end_of_the_week)]
        
        prev_year_week_data = prev_year_data[(prev_year_data['Date'] >= prev_year_start_of_week) & 
                                             (prev_year_data['Date'] <= prev_year_end_of_the_week)]
    
        # Aggregated and formatted data
        def aggregate_data(data):
            agg_data = data.groupby('Date').agg({'Bids': 'sum', 'Dealer Name': 'nunique'}).reset_index()
            agg_data.rename(columns={'Dealer Name': 'Dealer Count'}, inplace=True)
            agg_data['Date'] = agg_data['Date'].dt.strftime('%d/%m/%Y')
            agg_data['Day'] = agg_data['Date'].apply(lambda x: calendar.day_name[datetime.strptime(x, '%d/%m/%Y').weekday()])
            agg_data['Average Bids per Dealer'] = (agg_data['Bids'] / agg_data['Dealer Count']).round(2)
            return agg_data
    
        current_week_bids = aggregate_data(current_week_data)
        prev_year_week_bids = aggregate_data(prev_year_week_data)
    
        # Create time series plots
        fig_current_year = px.line(current_year_data.groupby('Date').agg({'Bids': 'sum'}).reset_index(), x='Date', y='Bids', color_discrete_sequence=['#1f77b4'])
        fig_previous_year = px.line(prev_year_data.groupby('Date').agg({'Bids': 'sum'}).reset_index(), x='Date', y='Bids', color_discrete_sequence=['#ff7f0e'])
            
        df['Year'] = df['Date'].dt.year
        total_bids_per_year = df.groupby('Year')['Bids'].sum()
        
        # Calculate average number of bids per day for each year
        avg_bids_per_day = total_bids_per_year.copy()  # Make a copy to avoid modifying original
        
        # Calculate the number of days in each year
        for year in avg_bids_per_day.index:
            if year == today.year:
                days_in_year = (today - pd.to_datetime(f'{year}-01-01')).days + 1
            else:
                days_in_year = 365 if pd.Timestamp(f'{year}-12-31').dayofyear == 365 else 366
            
            avg_bids_per_day.loc[year] /= days_in_year
        
        avg_bids_per_day = avg_bids_per_day.astype(int)
        # Create a DataFrame to display the results
        result_df = pd.DataFrame({
            'Year': avg_bids_per_day.index,
            'Average Bids per Day': avg_bids_per_day.values
        })
        
        avg_bids_table = dash_table.DataTable(
            id='average-bids-table',
            columns=[{"name": i, "id": i} for i in result_df.columns],
            data=result_df.to_dict('records'),
            style_table={'overflowX': 'auto'},
            style_header={'backgroundColor': 'white', 'color': 'black', 'fontFamily': 'Roboto'},
            style_cell={'backgroundColor': 'white', 'color': 'black', 'fontFamily': 'Roboto', 'textAlign': 'center'}
        )
        
        current_year_data = df[df['Date'].dt.year == current_year]
        current_year_data['Week'] = current_year_data['Date'].dt.strftime('%W')  # Assign week number
        current_year_data['Day'] = current_year_data['Date'].dt.strftime('%A')  # Assign day name
        
        # Calculate total bids per day for each week
        bids_per_day_per_week = current_year_data.groupby(['Week', 'Day'])['Bids'].sum().reset_index()
        
        # Calculate average bids per day for each week
        avg_bids_per_day_per_week = bids_per_day_per_week.groupby('Week')['Bids'].mean().reset_index()
        
        # Create a DataFrame to display the results
        week_avg_bids_per_day_df = pd.DataFrame({
            'Week': avg_bids_per_day_per_week['Week'],
            'Average Bids per Day': avg_bids_per_day_per_week['Bids'].round(0)  # Round to 0 decimal places
        })
        
        # Function to get start and end date of each week
        def get_week_dates(week_number, year):
            from datetime import datetime, timedelta
            
            # Adjust the week number to 1-indexed for datetime calculation
            start = datetime.strptime(f'{year}-W{int(week_number):02d}-1', "%Y-W%W-%w")
            end = start + timedelta(days=6)
            
            return start.strftime("%d/%m/%Y"), end.strftime("%d/%m/%Y")
        
        # Add start and end dates to the DataFrame
        week_avg_bids_per_day_df['Start Date'], week_avg_bids_per_day_df['End Date'] = zip(*week_avg_bids_per_day_df.apply(
            lambda row: get_week_dates(row['Week'], current_year), axis=1))
        
        # Create a DataTable for Week-On-Week Average Bids per Day Table
        week_avg_bids_per_day_table = dash_table.DataTable(
            id='week-average-bids-per-day-table',
            columns=[
                {"name": "Week", "id": "Week"},
                {"name": "Start Date", "id": "Start Date"},
                {"name": "End Date", "id": "End Date"},
                {"name": "Average Bids per Day", "id": "Average Bids per Day"}
            ],
            data=week_avg_bids_per_day_df.to_dict('records'),
            style_table={'maxHeight': '300px', 'overflowX': 'auto'},
            style_header={'backgroundColor': 'white', 'color': 'black', 'fontFamily': 'Roboto'},
            style_cell={'backgroundColor': 'white', 'color': 'black', 'fontFamily': 'Roboto', 'textAlign': 'center'}
        )
    
        fig_current_year.update_layout(
            paper_bgcolor='#ffffff',
            plot_bgcolor='#ffffff',
            font=dict(color='black'),
            xaxis=dict(title=dict(text='Date', font=dict(size=16)), tickfont=dict(size=14), gridcolor='#4d4d4d'),
            yaxis=dict(title=dict(text='Bids', font=dict(size=16)), tickfont=dict(size=14), gridcolor='#2d2d2d')
        )
    
        fig_previous_year.update_layout(
            paper_bgcolor='#ffffff',
            plot_bgcolor='#ffffff',
            font=dict(color='black'),
            xaxis=dict(title=dict(text='Date', font=dict(size=16)), tickfont=dict(size=14), gridcolor='#4d4d4d'),
            yaxis=dict(title=dict(text='Bids', font=dict(size=16)), tickfont=dict(size=14), gridcolor='#2d2d2d')
        )
    
        # Graphs for current and previous weeks
        fig_current_week = px.bar(current_week_bids, x='Day', y='Bids', color_discrete_sequence=['#1f77b4'])
        fig_prev_year_same_week = px.bar(prev_year_week_bids, x='Day', y='Bids', color_discrete_sequence=['#ff7f0e'])
    
        fig_current_week.update_layout(
            paper_bgcolor='#ffffff',
            plot_bgcolor='#ffffff',
            font=dict(color='black'),
            xaxis=dict(title=dict(text='Day', font=dict(size=16)), tickfont=dict(size=14)),
            yaxis=dict(title=dict(text='Bids', font=dict(size=16)), tickfont=dict(size=14), gridcolor='#2d2d2d')
        )
    
        fig_prev_year_same_week.update_layout(
            paper_bgcolor='#ffffff',
            plot_bgcolor='#ffffff',
            font=dict(color='black'),
            xaxis=dict(title=dict(text='Day', font=dict(size=16)), tickfont=dict(size=14)),
            yaxis=dict(title=dict(text='Bids', font=dict(size=16)), tickfont=dict(size=14), gridcolor='#2d2d2d')
        )
    
        return html.Div([
                html.Link(
                    href='https://fonts.googleapis.com/css2?family=Roboto:wght@400;700&display=swap',
                    rel='stylesheet'
                ),
                html.H1(today_full, style={'text-align': 'center', 'font-family': 'Roboto', 'color': 'black'}),
                #Separate div for avg_bids_table and week_avg_bids_per_day_table
                html.Div([
                    html.Div([
                        html.H3("Average Bids", style={'text-align': 'center', 'font-family': 'Roboto', 'color': 'black'}),
                        avg_bids_table,
                    ], style={'width': '49%', 'display': 'inline-block', 'vertical-align': 'top'}),
                    html.Div([
                        html.H3("Week Average Bids per Day", style={'text-align': 'center', 'font-family': 'Roboto', 'color': 'black'}),
                        week_avg_bids_per_day_table,
                    ], style={'width': '49%', 'display': 'inline-block', 'vertical-align': 'top'}),
                ], style={'text-align': 'center', 'display': 'flex', 'justify-content': 'space-around'}),
                
                html.Div([
                    html.Div([
                        html.H3("Bids Over Time (Current Year)", style={'text-align': 'center', 'font-family': 'Roboto', 'color': 'black'}),
                        dcc.Graph(figure=fig_current_year)
                    ], style={'width': '49%', 'display': 'inline-block', 'vertical-align': 'top'}),
                    html.Div([
                        html.H3("Bids Over Time (Previous Year)", style={'text-align': 'center', 'font-family': 'Roboto', 'color': 'black'}),
                        dcc.Graph(figure=fig_previous_year)
                    ], style={'width': '49%', 'display': 'inline-block', 'vertical-align': 'top'}),
                ], style={'text-align': 'center', 'display': 'flex', 'justify-content': 'space-around'}),
                html.Div([
                    html.Div([
                        html.H3("Number of Bids per Day (Current Week)", style={'text-align': 'center', 'font-family': 'Roboto', 'color': 'black'}),
                        dcc.Graph(figure=fig_current_week.update_traces(text=current_week_bids['Bids'], textposition='outside')),
                        html.Div([
                            html.Hr(),
                            dash_table.DataTable(
                                data=current_week_bids.to_dict('records'),
                                columns=[{"name": i, "id": i} for i in current_week_bids.columns],
                                style_table={'overflowX': 'auto'},
                                style_header={'backgroundColor': 'white', 'color': 'black', 'fontFamily': 'Roboto'},
                                style_cell={'backgroundColor': 'white', 'color': 'black', 'fontFamily': 'Roboto', 'textAlign': 'center'},
                                page_size=10
                            )
                        ], style={'display': 'block', 'margin': '20px auto', 'text-align': 'center'})
                    ], style={'width': '49%', 'display': 'inline-block', 'vertical-align': 'top'}),
                    html.Div([
                        html.H3("Number of Bids per Day (Previous Year Same Week)", style={'text-align': 'center', 'font-family': 'Roboto', 'color': 'black'}),
                        dcc.Graph(figure=fig_prev_year_same_week.update_traces(text=prev_year_week_bids['Bids'], textposition='outside')),
                        html.Div([
                            html.Hr(),
                            dash_table.DataTable(
                                data=prev_year_week_bids.to_dict('records'),
                                columns=[{"name": i, "id": i} for i in prev_year_week_bids.columns],
                                style_table={'overflowX': 'auto'},
                                style_header={'backgroundColor': 'white', 'color': 'black', 'fontFamily': 'Roboto'},
                                style_cell={'backgroundColor': 'white', 'color': 'black', 'fontFamily': 'Roboto', 'textAlign': 'center'},
                                page_size=10
                            )
                        ], style={'display': 'block', 'margin': '20px auto', 'text-align': 'center'})
                    ], style={'width': '49%', 'display': 'inline-block', 'vertical-align': 'top'})
                ])
            ])


    def run_dash():
        app.run_server(port=8050)

    print("Starting application...")
    dash_thread = threading.Thread(target=run_dash)
    dash_thread.start()

    monitor = get_monitors()[0]
    width, height = monitor.width, monitor.height
    try:
        # Start the webview window with Dash app URL
        webview.create_window("Dealer Activity Analysis", "http://127.0.0.1:8050",  width=width, height=height, resizable=True)
        webview.start()
    except Exception as e:
        print(f"Error creating window: {str(e)}")
        raise

if __name__ == '__main__':
    dealer_dashboard()
