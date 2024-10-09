import requests
from bs4 import BeautifulSoup
import pandas as pd
import os
import sys

def scrape(output_dir):
    # URL of the login page and the base URL for scraping with pagination
    login_url = 'https://www.motorist.sg/admin-login'  # Replace with the actual login URL
    base_url = 'https://www.motorist.sg/enquiry/sales?cso_id=&filter={}&page={}&state_id='  # Base URL for scraping

    # List of filter values to scrape
    filters = [5]

    # Variable to set the limit of pages to scrape
    page_limit = 2  # Set the number of pages you want to scrape
        
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

                # Loop through each filter value
                for filter_value in filters:
                    print(f"Scraping data for filter={filter_value}...")
                    # List to store the table data from all pages for the current filter
                    all_data = []

                    # Loop through the pages up to the page limit
                    for page in range(1, page_limit + 1):
                        # Get the page content after logging in
                        scrape_url = base_url.format(filter_value, page)
                        response = session.get(scrape_url)

                        if response.status_code == 200:
                            print(f"Page {page} retrieved successfully for filter={filter_value}.")

                            # Parse the HTML content using BeautifulSoup
                            soup = BeautifulSoup(response.content, 'html.parser')
                            
                            # Find the table headers
                            headers = [header.text.strip() for header in soup.find_all('th')]

                            # Find the rows in the table
                            rows = soup.find_all('tr')[1:]  # Skip the header row

                            for row in rows:
                                # Initialize list to store row data
                                row_data = []

                                # Find all the columns in the row
                                cols = row.find_all('td')

                                # Extract text from each column and append to row_data
                                for col in cols:
                                    cell_data = [line.strip() for line in col.decode_contents().split('<br>')]
                                    row_data.extend(cell_data)

                                # Append the row data to the all_data list
                                all_data.append(row_data)

                            print(f"Page {page} scraped successfully for filter={filter_value}.")
                        else:
                            print(f"Failed to retrieve page {page} for filter={filter_value}. Status code: {response.status_code}")

                    # Create a DataFrame
                    df = pd.DataFrame(all_data, columns=headers)

                    # Save the DataFrame to a Excel file
                    excel_filename = f"{output_dir}/sold_data.xlsx"
                    df.to_excel(excel_filename, index=False, header=True)

                    print(f"Data scraped and saved to {excel_filename}")

                print("All data scraped and saved successfully.")
            else:
                print("Login failed. Check your credentials and try again.")
        else:
            print(f"Failed to retrieve the login page. Status code: {login_page_response.status_code}")

    return excel_filename

def extract_url(cell):
    soup = BeautifulSoup(cell, 'html.parser')
    a_tag = soup.find('a', href=True)
    if a_tag and not a_tag['href'].startswith(('tel:', 'mailto:')):
        return f'https://www.motorist.sg{a_tag["href"]}'
    return ''

def filter(filename, output_dir):
    df = pd.read_excel(filename)
    
    # Remove trailing </br> tag
    df['Seller'] = df['Seller'].str.replace(r'<br/>$', '', regex=True)
    # Split 'Vehicle' column into separate columns
    vehicle_split = df['Vehicle'].str.split('<br/>', expand=True)
    
    # Insert the split columns into the DataFrame
    df.insert(loc=2, column='Plate', value=vehicle_split[0])
    df.insert(loc=3, column='Model', value=vehicle_split[1])
    df.insert(loc=4, column='Manufacturing_date', value=vehicle_split[2])
    df.insert(loc=5, column='Details', value=vehicle_split[3])
    df.insert(loc=6, column='Country', value=vehicle_split[4])
    
    df.insert(loc=10, column='Dealer Name', value=df['Buyer'].apply(lambda x: BeautifulSoup(x, 'html.parser').find('br').previous_sibling.strip()))
    
    # Drop the original 'Buyer' column
    df.drop(columns=['Buyer'], inplace=True)
    
    # Apply extract_url function to the last column to create 'Link' column
    df['Link'] = df.iloc[:, -1].apply(extract_url)

    # Drop unnecessary columns (including the second last column)
    df.drop(columns=['Vehicle', df.columns[-2]], inplace=True)
    
    # Write the modified DataFrame to a new Excel file
    excel_output_filename = f"{output_dir}/filtered_sold_data.xlsx"
    writer = pd.ExcelWriter(excel_output_filename, engine='openpyxl')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    writer.close()
    print(f"Excel file filtered and saved: {excel_output_filename}")

def main_sold():
    if getattr(sys, 'frozen', False):
        #When running as a bundled executable (e.g., PyInstaller)
        script_dir = os.path.dirname(sys.executable)
    else:
        #When running as a script
        script_dir = os.path.dirname(os.path.abspath(__file__))
    filenames = scrape(script_dir)
    filter(filenames, script_dir)

if __name__ == '__main__':
    main_sold()