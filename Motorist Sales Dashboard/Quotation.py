import requests
from bs4 import BeautifulSoup
import pandas as pd
import os
import sys

def scrape_quotation(output_dir):
    login_url = 'https://www.motorist.sg/admin-login'  # Replace with the actual login URL
    quotation_url = 'https://www.motorist.sg/enquiry/sales?filter=3&cso_id=&state_id='  # URL of the quotation page

    # Start a session
    with requests.Session() as session:
        # Get the login page to extract the authenticity token and any required cookies
        login_page_response = session.get(login_url)
        
        # Check if the request was successful
        if login_page_response.status_code == 200:
            print("Login page retrieved successfully.")
            
            # Parse the login page HTML to find the authenticity token
            login_page_soup = BeautifulSoup(login_page_response.content, 'html.parser')
            authenticity_token = login_page_soup.find('input', {'name': 'authenticity_token'}).get('value')
            
            # Credentials for login
            payload = {
                'user_admin[email]': 'limtzekang@motorist.sg',  # Replace with your actual email
                'user_admin[password]': '16062002',  # Replace with your actual password
                'authenticity_token': authenticity_token  # Include the authenticity token
                # Add any other hidden fields required by the login form
            }

            # Headers (if required)
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
                'Referer': login_url,  # Some sites check the referer header
                # Add any other headers that are required
            }

            # Post the login payload to the login URL
            login_response = session.post(login_url, data=payload, headers=headers)

            # Check if login was successful
            if login_response.status_code == 200 and "Logout" in login_response.text:
                print("Login successful!")

                # Get the quotation page content after logging in
                response = session.get(quotation_url)

                if response.status_code == 200:
                    print("Quotation page retrieved successfully.")

                    # Parse the HTML content using BeautifulSoup
                    soup = BeautifulSoup(response.content, 'html.parser')

                    # Define the tables and their headers
                    tables = {
                        "Active New": ["Seller", "Vehicle", "No of Offers", "Highest Offer", "Closing Date", "Link"],
                        "Active Requote": ["Seller", "Vehicle", "No of Offers", "Highest Offer", "Closing Date", "Link"],
                        "Followup": ["Seller", "Vehicle", "No of Offers", "Highest Offer", "Follow-Up Date", "Link"],
                        "Pending Agreement": ["Seller", "Vehicle", "No of Offers", "Highest Offer", "Follow-Up Date", "Link"],
                        "Appointment": ["Seller", "Vehicle", "No of Offers", "Highest Offer", "Appointment Date", "Link"]
                    }
                    
                    output_file = f"{output_dir}/quotation_data.xlsx"
                    writer = pd.ExcelWriter(output_file, engine='openpyxl')

                    # Print all headers found on the page
                    headers_on_page = [h2.string for h2 in soup.find_all('h2')]
                    print("Headers found on the page:", headers_on_page)
                    
                    table_found = False  # Flag to check if any table is found

                    # Loop through each table
                    for table_name, headers in tables.items():
                        print(f"Processing table: {table_name}")

                        # Find the table by its title
                        header = soup.find('h2', string=table_name)
                        if header:
                            table = header.find_next('table')
                            if table:
                                print(f"Table for {table_name} found.")

                                # Extract rows
                                rows = table.find_all('tr')[1:]  # Skip the header row

                                all_data = []
                                for row in rows:
                                    row_data = []
                                    cols = row.find_all('td')
                                    for col in cols:
                                        cell_data = [line.strip() for line in col.decode_contents().split('<br>')]
                                        row_data.extend(cell_data)
                                    all_data.append(row_data)

                                # Create a DataFrame and save it to a sheet
                                df = pd.DataFrame(all_data, columns=headers)
                                df.to_excel(writer, index=False, sheet_name=table_name)
                                print(f"Table {table_name} processed successfully.")
                                table_found = True
                            else:
                                print(f"Table for {table_name} not found.")
                        else:
                            print(f"Header {table_name} not found on the page.")

                    # Add a placeholder sheet if no table was found
                    if not table_found:
                        pd.DataFrame().to_excel(writer, index=False, sheet_name='NoData')

                    writer.close()
                    print("All tables scraped and saved successfully.")
                else:
                    print(f"Failed to retrieve the quotation page. Status code: {response.status_code}")
            else:
                print("Login failed. Check your credentials and try again.")
        else:
            print(f"Failed to retrieve the login page. Status code: {login_page_response.status_code}")
        return output_file

def extract_url(cell):
    soup = BeautifulSoup(cell, 'html.parser')
    a_tag = soup.find('a', href=True)
    if a_tag and not a_tag['href'].startswith(('tel:', 'mailto:')):
        return f'https://www.motorist.sg{a_tag["href"]}'
    return ''

def filter_quotation(filename, output_dir):
    df = pd.read_excel(filename, sheet_name=None)  # Read all sheets

    # Process each sheet
    for sheet_name, data in df.items():
        if sheet_name in ["Active New", "Active Requote", "Followup", "Pending Agreement", "Appointment"]:
            data['Link'] = data.iloc[:, -1].apply(extract_url)
            data['Seller'] = data['Seller'].str.replace(r'<br/>$', '', regex=True)
            vehicle_split = data['Vehicle'].str.split('<br/>', expand=True)
            data.insert(loc=2, column='Plate', value=vehicle_split[0])
            data.insert(loc=3, column='Model', value=vehicle_split[1])
            data.insert(loc=4, column='Manufacturing_date', value=vehicle_split[2])
            data.insert(loc=5, column='Details', value=vehicle_split[3])
            data.insert(loc=6, column='Country', value=vehicle_split[4])
            data.drop(columns=['Vehicle'], inplace=True)

    writer = pd.ExcelWriter(f"{output_dir}/filtered_quotation_data.xlsx", engine='openpyxl')
    for sheet_name, data in df.items():
        data.to_excel(writer, index=False, sheet_name=sheet_name)
    writer.close()
    print("Filtered data saved successfully.")


def main_quotation():
    if getattr(sys, 'frozen', False):
        #When running as a bundled executable (e.g., PyInstaller)
        script_dir = os.path.dirname(sys.executable)
    else:
        #When running as a script
        script_dir = os.path.dirname(os.path.abspath(__file__))
    filenames = scrape_quotation(script_dir)
    filter_quotation(filenames, script_dir)
