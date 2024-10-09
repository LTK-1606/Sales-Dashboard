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
import sys

def main_marketshare():
    # URLs of the zip files
    urls = [
        "https://datamall.lta.gov.sg/content/dam/datamall/datasets/Facts_Figures/Vehicle%20Registration/Monthly%20New%20Registration%20of%20Motor%20Vehicles%20by%20Vehicle%20Quota%20Categories.zip",
        "https://datamall.lta.gov.sg/content/dam/datamall/datasets/Facts_Figures/Vehicle%20Registration/Monthly%20De-Registered%20Motor%20Vehicles%20under%20Vehicle%20Quota%20System%20(VQS).zip",
        "https://datamall.lta.gov.sg/content/dam/datamall/datasets/Facts_Figures/Vehicle%20Population/Monthly%20Motor%20Vehicle%20Population%20by%20Vehicle%20Type.zip",
        "https://datamall.lta.gov.sg/content/dam/datamall/datasets/Facts_Figures/Vehicle%20Ownership%20n%20Transfer/Monthly%20Type%20and%20Number%20of%20Vehicles%20Transferred.zip",
        "https://datamall.lta.gov.sg/content/dam/datamall/datasets/Facts_Figures/Vehicle%20Registration/Monthly%20Revalidation%20of%20COE%20of%20Existing%20Vehicles.zip"
    ]

    if getattr(sys, 'frozen', False):
        # When running as a bundled executable (e.g., PyInstaller)
        script_dir = os.path.dirname(sys.executable)
    else:
        # When running as a script
        script_dir = os.path.dirname(os.path.abspath(__file__))

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
        
        # Construct the full file path
        file_path = os.path.join(script_dir, "cleaned_consolidated_data.xlsx")

        # Check if the file exists and load it
        if os.path.exists(file_path):
            # Load the Excel file
            motorist_data = file_path
            excel_sheets = pd.ExcelFile(motorist_data).sheet_names
            print(f"Excel sheets: {excel_sheets}")
        else:
            print(f"File not found: {file_path}")
        current_working_directory = script_dir
        
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
                    'margin-top': '20px'  # Add margin-top to separate from the dropdown
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
                    'margin-top': '20px'  # Add margin-top to separate from tables
                }
            )
        ], style={'width': '100%', 'max-width': 'auto', 'margin': '0 auto'})


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
                        weight='bold',
                        family="Arial Bold"  # Set text to bold
                    ), # Position text labels
                    name=title
                )
                layout = go.Layout(
                    title=dict(
                        text=title,
                        x=0.5,  # Center title horizontally
                        xanchor='center',  # Align title to the center
                        font=dict(
                            family="Roboto, sans-serif",  # Use Roboto font or any other desired font
                            size=20,  # Title font size
                            color='black',  # Title font color
                            weight='bold'  # Make title bold
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

                file_path = os.path.join(script_dir, "cleaned_consolidated_data.xlsx")
                
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
                            'font-weight': 'bold',
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
                    csv_file = os.path.join(script_dir, "M05-Dereg_by_Quota.csv")
                    lta_df = pd.read_csv(csv_file)
                elif selected_option == 'revalidation':
                    csv_file = os.path.join(script_dir, "M10-Monthly_COE_Revalidation.csv")
                    lta_df = pd.read_csv(csv_file)
                elif selected_option == 'new_registration':
                    csv_file = os.path.join(script_dir, "M02-New_Reg_by_Quota.csv")
                    lta_df = pd.read_csv(csv_file)
                elif selected_option == 'car_transfer':
                    csv_file = os.path.join(script_dir, "M07-Trf_by_type.csv")
                    lta_df = pd.read_csv(csv_file)
            
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
                    lta_df_table['month'] = lta_df_table['month'].dt.strftime('%d %B %Y')
                    
                    # Group by 'month' and aggregate 'number' or 'numbers' FOR TABLES
                    if selected_option != 'car_transfer':
                        lta_df_table = lta_df_table.groupby('month').agg({'number': 'sum'}).reset_index()
                    else:
                        lta_df_table = lta_df_table.groupby('month').agg({'numbers': 'sum'}).reset_index()
                    
                    lta_df_table = lta_df_table.groupby('month').agg({'number': 'sum'}).reset_index() if selected_option != 'car_transfer' else lta_df_table.groupby('month').agg({'numbers': 'sum'}).reset_index()
                    
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
                        'font-weight': 'bold',
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