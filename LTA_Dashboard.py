import os
import zipfile
import requests
from io import BytesIO
import pandas as pd
from datetime import datetime, timedelta
from dash import Dash, html, dcc
import plotly.express as px
import webview  # Import pywebview
from screeninfo import get_monitors
import tempfile
import threading
from dash.dependencies import Input, Output, State
import plotly.graph_objects as go

# URLs of the zip files
urls = [
    "https://datamall.lta.gov.sg/content/dam/datamall/datasets/Facts_Figures/Vehicle%20Registration/Monthly%20New%20Registration%20of%20Motor%20Vehicles%20by%20Vehicle%20Quota%20Categories.zip",
    "https://datamall.lta.gov.sg/content/dam/datamall/datasets/Facts_Figures/Vehicle%20Registration/Monthly%20De-Registered%20Motor%20Vehicles%20under%20Vehicle%20Quota%20System%20(VQS).zip",
    "https://datamall.lta.gov.sg/content/dam/datamall/datasets/Facts_Figures/Vehicle%20Population/Monthly%20Motor%20Vehicle%20Population%20by%20Vehicle%20Type.zip",
    "https://datamall.lta.gov.sg/content/dam/datamall/datasets/Facts_Figures/Vehicle%20Ownership%20n%20Transfer/Monthly%20Type%20and%20Number%20of%20Vehicles%20Transferred.zip",
    "https://datamall.lta.gov.sg/content/dam/datamall/datasets/Facts_Figures/Vehicle%20Registration/Monthly%20Revalidation%20of%20COE%20of%20Existing%20Vehicles.zip"
]

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

    # Move extracted files to the current working directory and remove empty folders
    current_working_directory = os.getcwd()
    motorist_data = os.path.join(current_working_directory, 'cleaned_consolidated_data.xlsx')
    excel_sheets = pd.ExcelFile(motorist_data).sheet_names
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

    # Function to aggregate monthly data by year and calculate the average
    def aggregate_by_year(df, value_col, category_col):
        return df.groupby(['year', category_col]).agg({value_col: 'mean'}).reset_index()

    def generate_line_plots(df, title, x_col, y_col, category_col=None):
        if category_col:
            categories = df[category_col].unique()
            plots = []
            for category in categories:
                category_df = df[df[category_col] == category]
                if not category_df.empty:
                    category_df = aggregate_by_year(category_df, y_col, category_col)
                    category_df[y_col] = category_df[y_col].round().astype(int)
                    fig = px.line(category_df, x=x_col, y=y_col, title=f'{title} - {category}', markers=True)
                    fig.update_traces(
                        text=category_df[y_col],
                        textposition='top center',
                        mode='lines+markers+text',
                        textfont=dict(family="Roboto, sans-serif", size=8, color="Black")  # Data label font size
                    )
                    fig.update_layout(
                        title=dict(text=f'{title} - {category}', font=dict(family="Roboto, sans-serif", size=18, color="Black")),
                        xaxis_title=dict(text=x_col, font=dict(family="Roboto, sans-serif", size=14, color="Black")),
                        yaxis_title=dict(text=y_col, font=dict(family="Roboto, sans-serif", size=14, color="Black")),
                        xaxis=dict(tickfont=dict(family="Roboto, sans-serif", size=14, color="Black")),
                        yaxis=dict(tickfont=dict(family="Roboto, sans-serif", size=14, color="Black")),
                        legend_title=dict(text="Legend Title", font=dict(family="Roboto, sans-serif", size=14, color="Black")),
                        legend=dict(font=dict(family="Roboto, sans-serif", size=14, color="Black")),
                        hoverlabel=dict(font=dict(family="Roboto, sans-serif", size=14, color="Black")),
                        font=dict(family="Roboto, sans-serif", size=14, color="Black")  # Other text elements
                    )
                    plots.append(dcc.Graph(figure=fig))
            return plots
        else:
            if not df.empty:
                df = aggregate_by_year(df, y_col, 'category')
                df[y_col] = df[y_col].round().astype(int)
                fig = px.line(df, x=x_col, y=y_col, title=title, markers=True)
                fig.update_traces(
                    text=df[y_col],
                    textposition='top center',
                    mode='lines+markers+text',
                    textfont=dict(family="Roboto, sans-serif", size=8, color="Black")  # Data label font size
                )
                fig.update_layout(
                    title=dict(text=title, font=dict(family="Roboto, sans-serif", size=14, color="Black", weight="bold")),
                    xaxis_title=dict(text=x_col, font=dict(family="Roboto, sans-serif", size=14, color="Black")),
                    yaxis_title=dict(text=y_col, font=dict(family="Roboto, sans-serif", size=14, color="Black")),
                    xaxis=dict(tickfont=dict(family="Roboto, sans-serif", size=14, color="Black")),
                    yaxis=dict(tickfont=dict(family="Roboto, sans-serif", size=14, color="Black")),
                    legend_title=dict(text="Legend Title", font=dict(family="Roboto, sans-serif", size=14, color="Black")),
                    legend=dict(font=dict(family="Roboto, sans-serif", size=14, color="Black")),
                    hoverlabel=dict(font=dict(family="Roboto, sans-serif", size=14, color="Black")),
                    font=dict(family="Roboto, sans-serif", size=14, color="Black")  # Other text elements
                )
                return [dcc.Graph(figure=fig)]
            return []

    # Organize data by category/type
    plots_by_category = {}
    titles = {
        'M02-New_Reg_by_Quota.csv': "New Registrations by Quota Category",
        'M05-Dereg_by_Quota.csv': "Deregistrations by Quota",
        'M06-Vehs_by_Type.csv': "Vehicles by Type",
        'M07-Trf_by_type.csv': "Transfers by Type",
        'M10-Monthly_COE_Revalidation.csv': "Monthly COE Revalidation"
    }

    for file, df in all_dfs.items():
        title = titles.get(file, 'No Title')
        if 'category' in df.columns:
            category_col = 'category'
            y_col = 'number'
        elif 'vehicle_type' in df.columns:
            category_col = 'vehicle_type'
            y_col = 'number'
        elif 'type' in df.columns:
            category_col = 'type'
            y_col = 'numbers'
        else:
            continue

        # Generate plots for the current dataset
        plots = generate_line_plots(df, title, 'year', y_col, category_col)

        for plot in plots:
            category = plot.figure['layout']['title']['text'].split(' - ')[-1]
            if category not in plots_by_category:
                plots_by_category[category] = []
            plots_by_category[category].append(plot)

    # Define app layout with each category/type's graphs in a responsive row
    rows = []
    for category, plots in plots_by_category.items():
        row = html.Div([
            html.Div(plot, style={
                'flex': '1 1 30%',  # Adjust the flex-basis as needed
                'max-width': '90%',  # Make each graph 10% smaller
                'margin': '10px',  # Add margin to avoid overlap
                'box-sizing': 'border-box'  # Include padding and border in width
            }) for plot in plots
        ], style={
            'display': 'flex', 
            'flex-wrap': 'wrap',  # Allow wrapping to the next line
            'justify-content': 'space-around',  # Space graphs evenly
            'margin-bottom': '20px',
            'padding': '0',
            'box-sizing': 'border-box'
        })
        rows.append(row)
        
    def generate_single_graph(df, sheet_name, columns):
        # Check if 'Date' and selected columns exist in the DataFrame
        if 'Date' not in df.columns:
            raise KeyError("Column not found: Date")
        for col in columns:
            if col not in df.columns:
                raise KeyError(f"Column not found: {col}")
    
        # Create a copy of the DataFrame to avoid modifying the original
        df_copy = df.copy()
    
        # Ensure the 'Date' column in the copy is in datetime format
        df_copy['Date'] = pd.to_datetime(df_copy['Date'], errors='coerce')
    
        # Clean and convert the columns to numeric
        for col in columns:
            df_copy[col] = df_copy[col].replace({'\$': '', ',': '', '-': '0', '%': ''}, regex=True)
            df_copy[col] = pd.to_numeric(df_copy[col], errors='coerce').fillna(0)
    
        # Melt the DataFrame
        df_melted = df_copy.melt(id_vars='Date', value_vars=columns, var_name='Category', value_name='Value')
    
        # Drop rows where 'Value' is NaN
        df_melted = df_melted.dropna(subset=['Value'])
    
        if df_melted.empty:
            fig1 = go.Figure()
            fig1.update_layout(
                title="No Data Available",
                xaxis_title='Date',
                yaxis_title='Value'
            )
            return dcc.Graph(figure=fig1)
    
        fig1 = px.line(df_melted, x='Date', y='Value', color='Category', markers=True,
                       labels={'Date': 'Date', 'Value': 'Value'},
                       title='Category Trends Over Time')
    
        fig1.update_traces(
            text=df_melted['Value'],
            textposition='top center',
            mode='lines+markers+text'
        )
    
        fig1.update_layout(
            xaxis_title='Date',
            yaxis_title='Value',
            title={'text': "Category Trends Over Time", 'x': 0.5},
            plot_bgcolor='#E5ECF6'
        )
    
        return fig1
    category_defaults = {
        'New': ['New', 'Quotation', 'Consignment', 'Coe Renewal', 'Purchases', 'Insurances'],
        'Active': ['Scrap', 'Quotation'],
        'Follow Up': ['New', 'Quotation', 'Consignment', 'Coe Renewal', 'Purchases', 'Insurances'],
        'Appt Set': ['Quotation', 'Coe Renewal', 'Insurances'],
        'Consigned': ['Consignment', 'Floor'],
        'Loan Submission': ['Coe Renewal', 'Loan Paperwork', 'Floor', 'Purchases'],
        'Appt Today': [],
        'Sold': ['Quotation', 'Sales', 'Coe Renewal', 'Dealer Purchase', 'Floor', 'Purchases', 'Insurances'],
        'Conversion_1': [],
        'Revenue': ['New', 'Scrap', 'Quotation', 'Consignment', 'Sales', 'Coe Renewal', 'Loan Paperwork', 'Floor', 'Purchases', 'Insurances', 'Total'],
        'Void': ['Sales', 'Coe Renewal', 'Loan Paperwork', 'Consignment', 'Dealer Purchase', 'Floor', 'Purchases', 'Insurances'],
        'Void Sold': ['Sales', 'Coe Renewal', 'Insurances']
    }
    
    # Define app layout
    app = Dash(__name__)
    app.layout = html.Div([
        html.H1("Vehicle Data Dashboard", style={
            'font-family': 'Roboto, sans-serif',
            'font-weight': 'bold',
            'font-size': '32px',
            'text-align': 'center'
        }),
        
        html.Div([
            html.Label('Select Category:'),
               dcc.Dropdown(
                   id='sheet-dropdown',
                   options=[{'label': k, 'value': k} for k in category_defaults.keys()],
                   value='New'  # Default category
               ),
               html.Label('Select Columns:'),
               dcc.Checklist(
                   id='column-checklist',
                   options=[
                       {'label': 'New', 'value': 'New'},
                       {'label': 'Scrap', 'value': 'Scrap'},
                       {'label': 'Quotation', 'value': 'Quotation'},
                       {'label': 'Consignment', 'value': 'Consignment'},
                       {'label': 'Sales', 'value': 'Sales'},
                       {'label': 'Coe Renewal', 'value': 'Coe Renewal'},
                       {'label': 'Loan Paperwork', 'value': 'Loan Paperwork'},
                       {'label': 'Consignment Purchase', 'value': 'Consignment Purchase'},
                       {'label': 'Dealer Purchase', 'value': 'Dealer Purchase'},
                       {'label': 'Floor', 'value': 'Floor'},
                       {'label': 'Purchases', 'value': 'Purchases'},
                       {'label': 'Insurances', 'value': 'Insurances'},
                       {'label': 'Total', 'value': 'Total'}
                   ],
                   value=category_defaults['New'],  # Default values for the 'New' category
                   inline=True
               )
        ], style={'width': '48%', 'display': 'inline-block'}),
        
        dcc.Graph(id='fig1'),
        
        *rows
    ], style={
        'font-family': 'Roboto, sans-serif',
        'padding': '0 10px'
    })

    @app.callback(
    Output('fig1', 'figure'),
    [Input('sheet-dropdown', 'value'),
     Input('column-checklist', 'value')]
)
    def update_graph(selected_category, selected_columns):
        print(f"Selected Category: {selected_category}")
        print(f"Selected Columns: {selected_columns}")
    
        # Load the data from the selected sheet
        df = pd.read_excel(motorist_data, sheet_name=selected_category)
        print(f"DataFrame Columns: {df.columns}")
    
        # Ensure there are columns selected
        if not selected_columns:
            selected_columns = df.columns[1:]  # If no columns selected, use all except 'Date'
    
        # Filter the DataFrame based on selected columns
        df_filtered = df[['Date'] + selected_columns]
        print(f"Filtered DataFrame: {df_filtered.head()}")
    
        # Generate the graph with the filtered data
        return generate_single_graph(df_filtered, selected_category, selected_columns)    
        

    def run_dash():
        app.run_server(port=8050, debug=True)
        
    # Run the app
    if __name__ == '__main__':
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