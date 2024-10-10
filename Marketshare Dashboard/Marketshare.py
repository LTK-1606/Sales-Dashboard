import os
import zipfile
import requests
from io import BytesIO
import pandas as pd
from datetime import datetime, timedelta
from dash import Dash, html, dcc, dash_table
import plotly.express as px
import webview  # Import pywebview
from screeninfo import get_monitors
import tempfile
import threading
from dash.dependencies import Input, Output, State
import plotly.graph_objects as go
import base64
import dash
import calendar
from salescalculation import salescalculation
import re
import textwrap
import sys

def customwrap(s,width=24):
    return "<br>".join(textwrap.wrap(s,width=width))


def main_marketshare():
    # URLs of the zip files
    urls = [
        "https://datamall.lta.gov.sg/content/dam/datamall/datasets/Facts_Figures/Vehicle%20Registration/Monthly%20New%20Registration%20of%20Motor%20Vehicles%20by%20Vehicle%20Quota%20Categories.zip",
        "https://datamall.lta.gov.sg/content/dam/datamall/datasets/Facts_Figures/Vehicle%20Registration/Monthly%20De-Registered%20Motor%20Vehicles%20under%20Vehicle%20Quota%20System%20(VQS).zip",
        "https://datamall.lta.gov.sg/content/dam/datamall/datasets/Facts_Figures/Vehicle%20Population/Monthly%20Motor%20Vehicle%20Population%20by%20Vehicle%20Type.zip",
        "https://datamall.lta.gov.sg/content/dam/datamall/datasets/Facts_Figures/Vehicle%20Ownership%20n%20Transfer/Monthly%20Type%20and%20Number%20of%20Vehicles%20Transferred.zip",
        "https://datamall.lta.gov.sg/content/dam/datamall/datasets/Facts_Figures/Vehicle%20Registration/Monthly%20Revalidation%20of%20COE%20of%20Existing%20Vehicles.zip"
    ]

    # Move extracted files to the current working directory and remove empty folders
    script_directory = os.getcwd()
    print(script_directory)

    # Define the Excel file name
    file_name = 'cleaned_consolidated_data.xlsx'

    # Construct the full file path
    file_path = os.path.join(script_directory, file_name)

    motorist_data = file_path


    # Create a temporary directory to save extracted data
    with tempfile.TemporaryDirectory() as temp_directory:
        print(f"Using temporary directory: {temp_directory}")

        # Function to download and extract zip files
        def download_and_extract_zip(url, dest_folder):
            try:
                response = requests.get(url)
                response.raise_for_status()
                print(f"Downloading and extracting {url}")
                with zipfile.ZipFile(BytesIO(response.content)) as z:
                    z.extractall(dest_folder)
            except requests.exceptions.RequestException as e:
                print(f"Error downloading {url}: {e}")
            except zipfile.BadZipFile as e:
                print(f"Error extracting {url}: {e}")

        # Download and extract each file
        for url in urls:
            download_and_extract_zip(url, temp_directory)


        # Check if the file exists and load it
        if os.path.exists(file_path):
            # Load the Excel file
            excel_sheets = pd.ExcelFile(motorist_data).sheet_names
            print(f"Excel sheets: {excel_sheets}")
        else:
            print(f"File not found: {file_path}")
        current_working_directory = os.getcwd()

        for root, dirs, files in os.walk(temp_directory):
            for dir in dirs:
                dir_path = os.path.join(root, dir)
                for file in os.listdir(dir_path):
                    print(f"Moving {file} from {dir_path} to {current_working_directory}")
                    os.rename(os.path.join(dir_path, file), os.path.join(current_working_directory, file))
                print(f"Removing empty directory: {dir_path}")
                os.rmdir(dir_path)

        # Load CSV files
        data_files = [f for f in os.listdir(current_working_directory) if f.endswith('.csv')]
        print(f"Data files found: {data_files}")

        all_dfs = {}
        for file in data_files:
            try:
                df = pd.read_csv(os.path.join(current_working_directory, file))
                print(f"Loaded {file}:")
                print(df.head())  # Print first few rows for verification
                df[df.columns[0]] = pd.to_datetime(df[df.columns[0]], errors='coerce')
                df['year'] = df[df.columns[0]].dt.year
                df['month'] = df[df.columns[0]].dt.month
                all_dfs[file] = df
            except Exception as e:
                print(f"Error processing {file}: {e}")

        # Define app layout
        app = Dash(__name__)
        app.layout = html.Div([
            dcc.Tabs(id="tabs", value='marketshare-tab', children=[
                dcc.Tab(label='Marketshare', value='marketshare-tab'),
                dcc.Tab(label='Sales', value='sales-tab'),
                dcc.Tab(label='Dealer', value='dealer-tab'),
            ]),
            html.Div(id='tabs-content')
        ])

        # Callback to control tab content
        @app.callback(
            Output('tabs-content', 'children'),
            Input('tabs', 'value')
        )

        def render_content(tab):
            if tab == 'marketshare-tab':
                return html.Div([
                    html.H1(
                        "Market Share Dashboard",
                        style={
                            'font-family': 'Roboto, sans-serif',
                            'font-weight': 'bold',
                            'font-size': '32px',
                            'text-align': 'center'
                        }
                    ),
                    html.Div([
                        html.Label('Select Category:'),
                        dcc.Dropdown(
                            id='info-dropdown',
                            options=[
                                {'label': 'Deregistration Information', 'value': 'deregistration'},
                                {'label': 'Revalidation Information', 'value': 'revalidation'},
                                {'label': 'New Registration Information', 'value': 'new_registration'},
                                {'label': 'Car Transfer Information', 'value': 'car_transfer'}
                            ],
                            value='deregistration'  # Default value
                        ),
                    ], style={'width': '100%', 'max-width': '1200px', 'margin': '0 auto'}),

                    html.Div(
                        id='tables-container',
                        style={
                            'display': 'grid',
                            'grid-template-columns': 'repeat(2, 1fr)',
                            'gap': '20px',
                            'width': '60%',  # Set the width for tables
                            'margin': '0 auto',  # Center the container within its parent
                            'max-width': '1200px',  # Set a maximum width for tables
                            'margin-top': '20px',  # Add margin-top to separate from the dropdown
                            'backgroundColor': '#f9f9f9',
                            'border': '2px solid #ddd',
                            'borderRadius': '10px',
                            'boxShadow': '2px 2px 5px rgba(0, 0, 0, 0.1)'
                        }
                    ),

                    html.Div(
                        id='graphs-container',
                        style={
                            'display': 'grid',
                            'grid-template-columns': 'repeat(2, 1fr)',
                            'gap': '20px',
                            'width': '100%',  # Adjust the width for graphs
                            'margin': '0 auto',  # Center the container horizontally
                            'margin-top': '20px',  # Add margin-top to separate from tables
                            'backgroundColor': '#f9f9f9',
                            'border': '2px solid #ddd',
                            'borderRadius': '10px',
                            'boxShadow': '2px 2px 5px rgba(0, 0, 0, 0.1)'
                        }
                    )
                ], style={'width': '100%', 'max-width': 'auto', 'margin': '0 auto'})

            elif tab == 'sales-tab':
                results_df = salescalculation()

                # Transpose the DataFrame
                transposed_df = results_df.T.reset_index()
                transposed_df.columns = ['Metric', 'Value']  # Rename columns for clarity

                # List of metrics that should be formatted as currency
                currency_metrics = [
                    'Scrap/Export Total Sum of Offers (Active New)',
                    'Scrap/Export Highest Offer (Active New)',
                    'Scrap/Export Total Sum of Offers (Active Requote)',
                    'Scrap/Export Highest Offer (Active Requote)',
                    'Quotation Total Sum of Offers (Active New)',
                    'Quotation Highest Offer (Active New)',
                    'Quotation Total Sum of Offers (Active Requote)',
                    'Quotation Highest Offer (Active Requote)',
                    'Quotation Total Sum of Offers (Followup)',
                    'Quotation Highest Offer (Followup)',
                    'Sold Total Sum of Price',
                    'Sold Highest Price Sold'
                ]

                # Create checkbox options
                sales_categories = [
                    'New', 'Scrap/Export', 'Quotation', 'Sold', 'Void'
                ]

                # Generate KPI cards for each metric in the DataFrame
                kpi_cards = []
                for _, row in transposed_df.iterrows():
                    metric = row['Metric']
                    value = row['Value']

                    # Debugging print statements
                    #print(f"Metric: {metric}, Value: {value}")

                    # Check if the metric is in the currency list and parse accordingly
                    if metric in currency_metrics:
                        value_format = ",.2f"
                        prefix = "$"
                    else:
                        value_format = ",.0f"
                        prefix = ""

                    # Create the KPI card
                    kpi_cards.append(
                        html.Div(
                            dcc.Graph(
                                figure=go.Figure(
                                    go.Indicator(
                                        value=value,
                                        mode="number",
                                        title={
                                            "text": customwrap(metric),
                                            "font": {"size": 16}
                                        },
                                        number={
                                            "font": {"size": 20, "color": "darkblue"},
                                            "valueformat": value_format,
                                            "prefix": prefix,
                                        }
                                    )
                                ),
                                style={
                                    'height': '100%',  # Full height of the parent div
                                    'width': '100%'    # Full width of the parent div
                                }
                            ),
                            style={
                                'display': 'inline-block',
                                'width': '300px',   # Fixed width for square shape
                                'height': '200px',  # Fixed height for square shape
                                'padding': '10px',
                                'margin': '10px',
                                'textAlign': 'center',
                                'backgroundColor': '#f9f9f9',
                                'border': '2px solid #ddd',
                                'borderRadius': '10px',
                                'boxShadow': '2px 2px 5px rgba(0, 0, 0, 0.1)',
                                'overflow': 'hidden',    # Hide overflow of content
                                'textOverflow': 'ellipsis',  # Add ellipsis for overflow text
                                'whiteSpace': 'normal',   # Allow text to wrap within the card
                                'position': 'relative'    # Ensure proper positioning
                            }
                        )
                    )

                return html.Div([
                    html.H1(
                        "Sales Dashboard",
                        style={
                            'font-family': 'Roboto, sans-serif',
                            'font-weight': 'bold',
                            'font-size': '32px',
                            'text-align': 'center'
                        }
                    ),html.Div(
                        [
                            dcc.Checklist(
                                id='category-filter',
                                options=[{'label': cat, 'value': cat} for cat in sales_categories],
                                value=['New'],  # Default selected categories
                                inline=True,
                                style={
                                    'font-size': '20px',        # Font size for the checklist options
                                    'text-align': 'center'      # Center text alignment within each option
                                }
                            )
                        ],
                        style={
                            'display': 'flex',                  # Use flexbox for alignment
                            'flex-direction': 'column',          # Stack items vertically
                            'align-items': 'center',             # Center items horizontally
                            'justify-content': 'center',         # Center items vertically if needed
                            'padding': '10px',                  # Optional padding around the checklist
                            'margin': '0 auto'                   # Center the div horizontally within its parent
                        }
                    ),
                    html.Div(
                        children=kpi_cards,
                        id = 'kpi-cards-container',
                        style={
                            'display': 'flex',
                            'flex-wrap': 'wrap',  # Wrap cards to the next line if they overflow
                            'justify-content': 'center',  # Center horizontally
                            'align-items': 'flex-start',  # Align items to the top
                            'padding': '20px',
                            'width': '100%'  # Full width of the container
                        }
                    )
                ])

            elif tab == 'dealer-tab':
                return html.Div([
                    html.H1(
                        "Dealer Activity Analysis",
                        style={
                            'font-family': 'Roboto, sans-serif',
                            'font-weight': 'bold',
                            'font-size': '32px',
                            'text-align': 'center'
                        }
                    ),

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
                    dcc.Dropdown(
                        id='sheet-selector',
                        placeholder='Select a sheet',
                        style={'width': '80%', 'margin': '20px auto'}
                    ),

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

                    dcc.Store(id='stored-data')
                ])



        @app.callback(
            Output('kpi-cards-container', 'children'),
            Input('category-filter', 'value'),
            Input('tabs', 'value')
        )



        def update_kpi_cards(selected_categories, tab):
            if tab == 'sales-tab':
                metric_to_category = {
                    'New Count New': 'New',
                    'New Count FollowUp': 'New',

                    'Scrap/Export Count (Active New)': 'Scrap/Export',
                    'Scrap/Export Total Number of Offers (Active New)': 'Scrap/Export',
                    'Scrap/Export Total Sum of Offers (Active New)': 'Scrap/Export',
                    'Scrap/Export Highest Offer (Active New)': 'Scrap/Export',
                    'Scrap/Export Count (Active Requote)': 'Scrap/Export',
                    'Scrap/Export Total Number of Offers (Active Requote)': 'Scrap/Export',
                    'Scrap/Export Total Sum of Offers (Active Requote)': 'Scrap/Export',
                    'Scrap/Export Highest Offer (Active Requote)': 'Scrap/Export',
                    'Scrap/Export Count (Followup)': 'Scrap/Export',
                    'Scrap/Export Count Overdue (Followup)': 'Scrap/Export',
                    'Scrap/Export Count (Appointment)': 'Scrap/Export',
                    'Scrap/Export Highest Number of Offers (Appointment)': 'Scrap/Export',
                    'Scrap/Export Average Number of Offers (Appointment)': 'Scrap/Export',

                    'Quotation Count (Active New)': 'Quotation',
                    'Quotation Total Number of Offers (Active New)': 'Quotation',
                    'Quotation Total Sum of Offers (Active New)': 'Quotation',
                    'Quotation Highest Offer (Active New)': 'Quotation',
                    'Quotation Count (Active Requote)': 'Quotation',
                    'Quotation Total Number of Offers (Active Requote)': 'Quotation',
                    'Quotation Total Sum of Offers (Active Requote)': 'Quotation',
                    'Quotation Highest Offer (Active Requote)': 'Quotation',
                    'Quotation Count (Followup)': 'Quotation',
                    'Quotation Count Overdue (Followup)': 'Quotation',
                    'Quotation Total Number of Offers (Followup)': 'Quotation',
                    'Quotation Highest Number of Offers (Followup)': 'Quotation',
                    'Quotation Total Sum of Offers (Followup)': 'Quotation',
                    'Quotation Highest Offer (Followup)': 'Quotation',
                    'Quotation Count (Appointment)': 'Quotation',
                    'Quotation Highest Number of Offers (Appointment)': 'Quotation',
                    'Quotation Average Number of Offers (Appointment)': 'Quotation',

                    'Sold Count of Sold': 'Sold',
                    'Sold Total Sum of Price': 'Sold',
                    'Sold Highest Price Sold': 'Sold',

                    'Void Count of Void': 'Void'
                }

                results_df = salescalculation()  # Fetch your data

                currency_metrics = [
                    'Scrap/Export Total Sum of Offers (Active New)',
                    'Scrap/Export Highest Offer (Active New)',
                    'Scrap/Export Total Sum of Offers (Active Requote)',
                    'Scrap/Export Highest Offer (Active Requote)',
                    'Quotation Total Sum of Offers (Active New)',
                    'Quotation Highest Offer (Active New)',
                    'Quotation Total Sum of Offers (Active Requote)',
                    'Quotation Highest Offer (Active Requote)',
                    'Quotation Total Sum of Offers (Followup)',
                    'Quotation Highest Offer (Followup)',
                    'Sold Total Sum of Price',
                    'Sold Highest Price Sold'
                ]

                # Transpose and format DataFrame
                transposed_df = results_df.T.reset_index()
                transposed_df.columns = ['Metric', 'Value']

                if selected_categories:
                    filtered_metrics = [metric for metric, category in metric_to_category.items() if category in selected_categories]
                    filtered_df = transposed_df[transposed_df['Metric'].isin(filtered_metrics)]
                else:
                    filtered_df = transposed_df

                # Create KPI cards
                kpi_cards = []
                for _, row in filtered_df.iterrows():
                    metric = row['Metric']
                    value = row['Value']

                    if metric in currency_metrics:
                        value_format = ",.2f"
                        prefix = "$"
                    else:
                        value_format = ",.0f"
                        prefix = ""

                    kpi_cards.append(
                        html.Div(
                            dcc.Graph(
                                figure=go.Figure(
                                    go.Indicator(
                                        value=value,
                                        mode="number",
                                        title={"text": customwrap(metric), "font": {"size": 16}},
                                        number={
                                            "font": {"size": 20, "color": "darkblue"},
                                            "valueformat": value_format,
                                            "prefix": prefix,
                                        }
                                    )
                                ),
                                style={'height': '100%', 'width': '100%'}
                            ),
                            style={
                                'display': 'inline-block',
                                'width': '300px',
                                'height': '200px',
                                'padding': '10px',
                                'margin': '10px',
                                'textAlign': 'center',
                                'backgroundColor': '#f9f9f9',
                                'border': '2px solid #ddd',
                                'borderRadius': '10px',
                                'boxShadow': '2px 2px 5px rgba(0, 0, 0, 0.1)',
                                'overflow': 'hidden',
                                'textOverflow': 'ellipsis',
                                'whiteSpace': 'normal',
                                'position': 'relative'
                            }
                        )
                    )

                return kpi_cards
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
                style_table={'maxHeight': '300px',
                             'overflowX': 'auto',
                             'backgroundColor': '#f9f9f9',
                             'border': '2px solid #ddd',
                             'borderRadius': '10px',
                             'boxShadow': '2px 2px 5px rgba(0, 0, 0, 0.1)'
                             },
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
                style_table={'maxHeight': '300px',
                             'overflowX': 'auto',
                             'backgroundColor': '#f9f9f9',
                             'border': '2px solid #ddd',
                             'borderRadius': '10px',
                             'boxShadow': '2px 2px 5px rgba(0, 0, 0, 0.1)'
                             },
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
                                style_table={'overflowX': 'auto',
                                             'backgroundColor': '#f9f9f9',
                                             'border': '2px solid #ddd',
                                             'borderRadius': '10px',
                                             'boxShadow': '2px 2px 5px rgba(0, 0, 0, 0.1)'
                                             },
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
                                style_table={'overflowX': 'auto',
                                             'backgroundColor': '#f9f9f9',
                                             'border': '2px solid #ddd',
                                             'borderRadius': '10px',
                                             'boxShadow': '2px 2px 5px rgba(0, 0, 0, 0.1)'
                                             },
                                style_header={'backgroundColor': 'white', 'color': 'black', 'fontFamily': 'Roboto'},
                                style_cell={'backgroundColor': 'white', 'color': 'black', 'fontFamily': 'Roboto', 'textAlign': 'center'},
                                page_size=10
                            )
                        ], style={'display': 'block', 'margin': '20px auto', 'text-align': 'center'})
                    ], style={'width': '49%', 'display': 'inline-block', 'vertical-align': 'top'})
                ])
            ])

        # Callback for updating graphs based on dropdown selection
        @app.callback(
            Output('graphs-container', 'children'),
            [Input('info-dropdown', 'value')]
        )

        def update_graph(selected_option):
            def clean_data(df, date_col, value_col):
                # Remove the word "Week " and convert date column to datetime
                df[date_col] = df[date_col].astype(str)
                df[date_col] = df[date_col].str.replace('Week ', '', regex=False)

                # Clean value column
                df[value_col] = df[value_col].replace({r'[\$,]': '', r'[\%]': ''}, regex=True)

                # Convert cleaned value column to numeric, setting errors to 'coerce' to handle invalid parsing
                df[value_col] = pd.to_numeric(df[value_col], errors='coerce')
                df[date_col] = pd.to_datetime(df[date_col], errors='coerce')

                return df

            def create_graph(df, x_col, y_col, title, x_axis_title, y_axis_title, color):
                df = df.sort_values(x_col)  # Sort by x-axis
                trace = go.Scatter(
                    x=df[x_col],
                    y=df[y_col],
                    mode='lines+markers+text',  # Show lines, markers, and text
                    line=dict(color=color),
                    marker=dict(size=4),  # Adjust marker size as needed
                    text=df[y_col].round(0),  # Display the rounded y values as text labels
                    textposition='top center',
                    textfont=dict(
                        size=10,  # Change text size here
                        family="Roboto"  # Set text to bold
                    ), # Position text labels
                    name=title
                )
                layout = go.Layout(
                    title=dict(
                        text=title,
                        x=0.5,  # Center title horizontally
                        xanchor='center',  # Align title to the center
                        font=dict(
                            family="Roboto",  # Use Roboto font or any other desired font
                            size=20,  # Title font size
                            color='black',  # Title font color
                        )
                    ),
                    xaxis=dict(title=x_axis_title, type='date'),  # Ensure x-axis is treated as date
                    yaxis=dict(title=y_axis_title)
                )
                fig = go.Figure(data=[trace], layout=layout)
                return fig

            def update_graphs(selected_option):
                graphs = []
                tables = []
                colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728']
                current_year = datetime.now().year

                # Load the Excel file
                if getattr(sys, 'frozen', False):
                    # When running as a bundled executable (e.g., PyInstaller)
                    script_dir = os.path.dirname(sys.executable)

                else:
                    # When running as a script
                    script_dir = os.path.dirname(os.path.abspath(__file__))

                file_name = 'cleaned_consolidated_data.xlsx'
                file_path = os.path.join(script_dir, file_name)

                if not os.path.exists(file_path):
                    print(f"File not found: {file_path}")
                    return []

                motorist_data = file_path
                excel_sheets = pd.ExcelFile(motorist_data).sheet_names
                print(f"Excel sheets: {excel_sheets}")

                if selected_option == 'deregistration':
                    sheets = ['Revenue', 'Conversion', 'Sold']
                    columns = ['Scrap', 'Scrap', 'Scrap']
                    lta_data = 'Deregistrations by Quota - Category A and B'

                elif selected_option == 'revalidation':
                    sheets = ['Revenue', 'Conversion', 'Sold']
                    columns = ['Coe Renewal', 'Coe Renewal', 'Coe Renewal']
                    lta_data = 'Monthly COE Revalidation'

                elif selected_option == 'new_registration':
                    sheets = ['Revenue', 'Conversion', 'Sold']
                    columns = ['Loan Paperwork', 'Loan Paperwork', 'Loan Paperwork']
                    lta_data = 'New Registrations by Quota - Category A and B'

                elif selected_option == 'car_transfer':
                    sheets = ['Revenue', 'Conversion', 'Sold']
                    columns = ['Quotation', 'Quotation', 'Quotation']
                    lta_data = 'Transfers by Type - Cars'

                for i, (sheet, column) in enumerate(zip(sheets, columns)):
                    if sheet in excel_sheets:
                        # Load and clean the data
                        df = pd.read_excel(motorist_data, sheet_name=sheet)
                        df = clean_data(df, 'Date', column)

                        # Create and add the graph
                        graphs.append(
                            dcc.Graph(
                                figure=create_graph(
                                    df,
                                    'Date',
                                    column,
                                    f'{sheet} - {column}',
                                    'Date',
                                    column,
                                    colors[i]
                                ),
                            )
                        )
                        # Filter columns to include only the selected ones
                        selected_columns = ['Date', column]  # Adjust if there are additional columns you want to include
                        df_filtered = df[selected_columns]
                        df_filtered = df_filtered.sort_values(by='Date', ascending=False)

                        # Format dates for tables
                        df_filtered['Date'] = pd.to_datetime(df_filtered['Date'], errors='coerce').dt.strftime('%d %B %Y')

                        if sheet == 'Conversion':
                            df_filtered[column] = df_filtered[column].apply(lambda x: f"{x:.0f}%")

                        if sheet == 'Revenue':
                            df_filtered[column] = df_filtered[column].apply(lambda x: f"${x:.0f}")

                        # Convert DataFrame to a table
                        table = dash_table.DataTable(
                            data=df_filtered.to_dict('records'),  # Convert filtered DataFrame to a list of dictionaries
                            columns=[{'name': col, 'id': col} for col in df_filtered.columns],  # Set table columns
                            style_table={
                                'overflowX': 'auto',  # Horizontal scroll
                                'overflowY': 'auto',  # Vertical scroll
                                'backgroundColor': '#f9f9f9',
                                'border': '2px solid #ddd',
                                'borderRadius': '10px',
                                'boxShadow': '2px 2px 5px rgba(0, 0, 0, 0.1)',
                                'maxHeight': '400px',  # Set a maximum height for vertical scroll
                                'width': '400px'  # Adjust table width
                            },
                            style_header={
                                'backgroundColor': 'rgb(230, 230, 230)',
                                'fontWeight': 'bold'
                            },
                            style_data={
                                'whiteSpace': 'normal',
                                'height': 'auto',
                                'padding': '2px',  # Reduce padding
                                'fontSize': '15px'  # Adjust font size
                            },
                            style_cell={
                                'minWidth': '10px',  # Minimum column width
                                'width': '50px',  # Fixed column width
                                'maxWidth': '50px',  # Maximum column width
                                'textAlign': 'center'  # Center-align text
                            }
                        )

                        # Add the table to the layout
                        tables.append(html.Div([
                            html.H3(f'{sheet} - {column}'),
                            table
                        ],  style={
                            'font-family': 'Roboto, sans-serif',
                            'font-size': '20px',
                            'text-align': 'center',
                            'margin': '20px',  # Margin around the container
                            'display': 'flex',  # Use flexbox for layout
                            'flexDirection': 'column',  # Arrange items in a column
                            'alignItems': 'center',  # Center items horizontally
                            'textAlign': 'center'  # Center text
                        }))

                # Load LTA data
                if selected_option == 'deregistration':
                    lta_df = pd.read_csv('M05-Dereg_by_Quota.csv')
                elif selected_option == 'revalidation':
                    lta_df = pd.read_csv('M10-Monthly_COE_Revalidation.csv')
                elif selected_option == 'new_registration':
                    lta_df = pd.read_csv('M02-New_Reg_by_Quota.csv')
                elif selected_option == 'car_transfer':
                    lta_df = pd.read_csv('M07-Trf_by_type.csv')

                if not lta_df.empty:
                    if selected_option in ['deregistration', 'new_registration']:
                        lta_df = lta_df[lta_df['category'].isin(['Category A', 'Category B'])]
                    #lta_df = lta_df[lta_df['category'].isin(['Category A', 'Category B'])] if selected_option == 'deregistration' or selected_option == 'new_registration' else lta_df == lta_df

                    # Clean data
                    lta_df = clean_data(lta_df, 'month', 'number' if selected_option != 'car_transfer' else 'numbers')

                    # Ensure 'month' column is in datetime format
                    lta_df['month'] = pd.to_datetime(lta_df['month'], errors='coerce')

                    # Filter for the current year
                    lta_df = lta_df[lta_df['month'].dt.year == current_year]

                    # Format dates for tables
                    lta_df_table = lta_df.copy()
                    lta_df_table['month'] = pd.to_datetime(lta_df_table['month'], format='%Y-%m-%d')

                    # Group by 'month' and aggregate 'number' or 'numbers' FOR TABLES
                    if selected_option != 'car_transfer':
                        lta_df_table = lta_df_table.groupby('month').agg({'number': 'sum'}).reset_index()
                    else:
                        lta_df_table = lta_df_table.groupby('month').agg({'numbers': 'sum'}).reset_index()

                    lta_df_table = lta_df_table.groupby('month').agg({'number': 'sum'}).reset_index() if selected_option != 'car_transfer' else lta_df_table.groupby('month').agg({'numbers': 'sum'}).reset_index()

                    lta_df_table = lta_df_table.sort_values(by='month', ascending=False)
                    lta_df_table['month'] = lta_df_table['month'].dt.strftime('%d %B %Y')
                    # Group by 'month' and aggregate 'number' or 'numbers'
                    if selected_option != 'car_transfer':
                        lta_df = lta_df.groupby('month').agg({'number': 'sum'}).reset_index()
                    else:
                        lta_df = lta_df.groupby('month').agg({'numbers': 'sum'}).reset_index()

                    y_col = 'number' if selected_option != 'car_transfer' else 'numbers'
                    y_axis_title = 'number' if selected_option != 'car_transfer' else 'numbers'

                    # Define the title and color
                    graph_title = lta_data  # Title of the graph
                    color = colors[len(sheets)]  # Color

                    # Create the graph
                    lta_graph = dcc.Graph(
                        figure=create_graph(
                            lta_df,
                            'month',
                            y_col,
                            graph_title,
                            'month',   # x-axis title
                            y_axis_title,  # y-axis title
                            color   # Color
                        ),
                    )

                    graphs.append(lta_graph)
                    # Convert the DataFrame to a table
                    table = dash_table.DataTable(
                        data=lta_df_table.to_dict('records'),  # Convert DataFrame to a list of dictionaries
                        columns=[{'name': col, 'id': col} for col in lta_df.columns],  # Set table columns
                        style_table={
                            'overflowX': 'auto',  # Horizontal scroll
                            'overflowY': 'auto',  # Vertical scroll
                            'maxHeight': '400px',  # Set a maximum height for vertical scroll
                            'width': '400px'
                        },
                        style_header={
                            'backgroundColor': 'rgb(230, 230, 230)',
                            'fontWeight': 'bold'
                        },
                        style_data={
                            'whiteSpace': 'normal',
                            'height': 'auto',
                            'padding': '2px',  # Reduce padding
                            'fontSize': '15px'  # Adjust font size
                        },
                        style_cell={
                            'minWidth': '10px',  # Minimum column width
                            'width': '50px',  # Fixed column width
                            'maxWidth': '50px',  # Maximum column width
                            'textAlign': 'center'  # Center-align text
                        }
                    )

                    # Add the table to the layout
                    tables.append(html.Div([
                        html.H3(lta_data),  # Set a descriptive header
                        table
                    ],  style={
                        'font-family': 'Roboto, sans-serif',
                        'font-size': '20px',
                        'text-align': 'center',
                        'margin': '20px',  # Margin around the container
                        'display': 'flex',  # Use flexbox for layout
                        'flexDirection': 'column',  # Arrange items in a column
                        'alignItems': 'center',  # Center items horizontally
                        'textAlign': 'center'  # Center text
                    }))

                return graphs + tables

            return update_graphs(selected_option)

        def run_dash():
            app.run_server(port=8050)

        # Run the app
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
            print(f"Error opening webview window: {e}")

if __name__ == '__main__':
    main_marketshare()
