from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import os
import sys
import time
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
import pandas as pd
from datetime import datetime

def read_last_row_first_column(file_path, sheet_name=0):
    # Load the Excel file
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    
    # Access the last row and first column
    last_row_first_column = df.iloc[-1, 0]
    
    return last_row_first_column

def parse_week_string(week_str):
    # Extract the date part from the string
    date_str = week_str.split(' ')[1]
    
    # Convert the date part to a datetime object
    try:
        date_obj = datetime.strptime(date_str, '%Y-%m-%d')
        new_date_obj = date_obj + timedelta(days=7)
    except ValueError:
        new_date_obj = None
    
    return new_date_obj

def scrape(output_dir):
    # Setup Selenium with Chrome
    chrome_options = Options()
    chrome_options.add_argument("--headless")  # Run headless mode (no GUI)
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    
    try:
        # URL of the login page
        login_url = 'https://www.motorist.sg/admin-login'

        # Navigate to the login page
        print("Navigating to login page...")
        driver.get(login_url)

        # Find and fill login fields
        print("Filling login fields...")
        driver.find_element(By.ID, 'user_admin_email').send_keys('limtzekang@motorist.sg')
        driver.find_element(By.ID, 'user_admin_password').send_keys('16062002')
        submit_button = driver.find_element(By.CLASS_NAME, 'btn-success')
        submit_button.click()

        # Wait for login to complete
        print("Waiting for login to complete...")
        time.sleep(5)

        # Find Latest Pulled Data
        historic_data = os.path.join(output_dir, "consolidated_&_formatted_data (historical).xlsx")
        latest_date_str = read_last_row_first_column(historic_data, "New")
        latest_date = parse_week_string(latest_date_str)

        # Calculate dates
        end_date = datetime.now()
        if latest_date > end_date:
            latest_date -= timedelta(days=7)

        # For Pulling Historic Data
        # start_date = end_date - timedelta(weeks=52)
        # days_difference = 1

        # For Pulling New Data
        start_date = latest_date
        difference = end_date - start_date
        days_difference = difference.days

        current_date = start_date

        if days_difference > 0:
            # Initialize the Excel writer
            excel_filename = os.path.join(output_dir, "sales_dashboard (new).xlsx")
            writer = pd.ExcelWriter(excel_filename, engine='openpyxl')

            while current_date < end_date:
                week_start = current_date.strftime("%d/%m/%Y")
                week_end = (current_date + timedelta(days=6)).strftime("%d/%m/%Y")
                print(f"Scraping data for week from {week_start} to {week_end}...")
                
                base_url = f'https://www.motorist.sg/review/sales?filter=2&show_only_month=true&start={week_start}&end={week_end}&state_id='
                
                # Navigate to the base URL
                #print(f"Navigating to URL: {base_url}")
                driver.get(base_url)

                # Wait for the page to load
                #print("Waiting for page to load...")
                time.sleep(10)

                # Click the "Generate" button
                #print("Clicking the 'Generate' button...")
                generate_button = driver.find_element(By.ID, 'generatebutton2')
                generate_button.click()

                # Wait for data to load
                #print("Waiting for data to load...")
                time.sleep(20)

                # Execute JavaScript to ensure the page is fully loaded
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(5)

                # Print the page content for debugging
                page_source = driver.page_source

                # Parse the page content
                #print("Parsing the page content...")
                soup = BeautifulSoup(page_source, 'html.parser')
            
                tables = soup.find_all('table', class_='table table-striped table-condensed table-fixed-column table-no-bordered')
                if tables is None:
                    raise ValueError("Table with the specified class was not found.")        

                dataframes = []

                for table in tables:
                    row_headers = []
                    col_headers = []
                    t_head = table.find('thead')
                    t_body = table.find('tbody')
                    if t_head is None:
                        print("No thead")
                    else:
                        print("Found thead")
                        tr = t_head.find('tr')
                        for th in tr.find_all('th'):
                            col_header = th.find('div', class_='th-inner').get_text(strip=True)
                            col_headers.append(col_header)

                    if t_body is None:
                        print("No tbody")
                    else:
                        print("Found tbody")
                        for tr in table.find_all('tr'):
                            for td in tr.find_all('td'):
                                row_header = td.get_text(strip=True)
                                row_headers.append(row_header)
                    
                    # Function to separate the list based on entries with words
                    def separate_entries(data):
                        result = []
                        current_entry = []
                        for item in data:
                            if item.isalpha() or any(char.isalpha() for char in item):
                                if current_entry:
                                    result.append(current_entry)
                                current_entry = [item]
                            else:
                                current_entry.append(item)
                        if current_entry:
                            result.append(current_entry)
                        return result

                    # Separate the tbody list
                    if t_body and t_head:
                        separated_data = separate_entries(row_headers)
                        col_headers.pop(0)

                        # Create a list of DataFrames for each entry
                        for entry in separated_data:
                            name = entry[0]
                            values = entry[1:]
                            if len(values) == len(col_headers):
                                df = pd.DataFrame([values], columns=col_headers)
                                df.insert(0, 'Category', name)  # Insert category name in the first column
                                dataframes.append(df)

                if dataframes:
                    week_df = pd.concat(dataframes, ignore_index=True)
                    week_df.insert(0, 'Week Start', week_start)  # Add a column for the week start date
                    week_df.insert(1, 'Week End', week_end)  # Add a column for the week end date
                    sheet_name = f"Week {current_date.strftime('%Y-%m-%d')}"
                    week_df.to_excel(writer, sheet_name=sheet_name[:31], index=False)  # Ensure the sheet name is not too long
                    print(f"Scraped data from {week_start} to {week_end}")

                # Increment the current date by one week
                current_date += timedelta(weeks=1)

            # Save the Excel file
            writer.close()
            print(f"Data scraped and saved to {excel_filename}")
        else:
            excel_filename = None
            print("Data is already Updated!")
    
    except Exception as e:
        print(f"An error occurred: {e}")
        excel_filename = None
    finally:
        print("Quitting the browser...")
        driver.quit()

    return excel_filename

def main_salesdashboard():
    if getattr(sys, 'frozen', False):
        # When running as a bundled executable (e.g., PyInstaller)
        script_dir = os.path.dirname(sys.executable)
        excel_file = scrape(script_dir)
    else:
        # When running as a script
        script_dir = os.path.dirname(os.path.abspath(__file__))
        excel_file = scrape(script_dir)
    
    if excel_file:
        print(f"Excel file saved at {excel_file}")
    else:
        print("No file was saved.")

if __name__ == '__main__':
    main_salesdashboard()
