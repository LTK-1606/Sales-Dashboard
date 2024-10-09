import pandas as pd
import datetime
import os
import sys

currentdate = datetime.datetime.now()
currentmonth = currentdate.month
if getattr(sys, 'frozen', False):
    # When running as a bundled executable (e.g., PyInstaller)
    script_directory = os.path.dirname(sys.executable)
else:
    # When running as a script
    script_directory = os.path.dirname(os.path.abspath(__file__))
index = [
    'New Count New', 
    'New Count FollowUp', 
    'Scrap/Export Count (Active New)', 
    'Scrap/Export Total Number of Offers (Active New)', 
    'Scrap/Export Total Sum of Offers (Active New)', 
    'Scrap/Export Highest Offer (Active New)',
    'Scrap/Export Count (Active Requote)', 
    'Scrap/Export Total Number of Offers (Active Requote)', 
    'Scrap/Export Total Sum of Offers (Active Requote)', 
    'Scrap/Export Highest Offer (Active Requote)',
    'Scrap/Export Count (Followup)', 
    'Scrap/Export Count Overdue (Followup)',
    'Scrap/Export Count (Appointment)', 
    'Scrap/Export Highest Number of Offers (Appointment)', 
    'Scrap/Export Average Number of Offers (Appointment)',
    'Quotation Count (Active New)', 
    'Quotation Total Number of Offers (Active New)', 
    'Quotation Total Sum of Offers (Active New)', 
    'Quotation Highest Offer (Active New)',
    'Quotation Count (Active Requote)', 
    'Quotation Total Number of Offers (Active Requote)', 
    'Quotation Total Sum of Offers (Active Requote)', 
    'Quotation Highest Offer (Active Requote)',
    'Quotation Count (Followup)', 
    'Quotation Count Overdue (Followup)',
    'Quotation Total Number of Offers (Followup)',
    'Quotation Highest Number of Offers (Followup)',
    'Quotation Total Sum of Offers (Followup)',
    'Quotation Highest Offer (Followup)',
    'Quotation Count (Appointment)', 
    'Quotation Highest Number of Offers (Appointment)', 
    'Quotation Average Number of Offers (Appointment)',
    'Sold Count of Sold',
    'Sold Total Sum of Price',
    'Sold Highest Price Sold',
    'Void Count of Void'
]

def clean_and_convert(column):
    column = column.replace({'\$': '', ',': ''}, regex=True)
    return pd.to_numeric(column, errors='coerce')
    

def calculate_new():
    new_path =  os.path.join(script_directory,"filtered_new_data.xlsx") 
    new = pd.read_excel(new_path, "New")
    tabulated_results.append(new.shape[0])
    
    new_fu = pd.read_excel(new_path, "Followup")
    tabulated_results.append(new_fu.shape[0])
    
    return tabulated_results

def calculate_se():
    se_path =  os.path.join(script_directory,"filtered_scrapexport_data.xlsx") 
    se_an = pd.read_excel(se_path, "Active New")
    se_an['Highest Offer'] = clean_and_convert(se_an['Highest Offer'])

    tabulated_results.append(se_an.shape[0])
    tabulated_results.append(se_an['No of Offers'].sum())
    tabulated_results.append((se_an['Highest Offer'].sum()))
    tabulated_results.append((se_an['Highest Offer'].max()))

    se_ar = pd.read_excel(se_path, "Active Requote")
    se_ar['Highest Offer'] = clean_and_convert(se_ar['Highest Offer'])

    tabulated_results.append(se_ar.shape[0])
    tabulated_results.append(se_ar['No of Offers'].sum())
    tabulated_results.append((se_ar['Highest Offer'].sum()))
    tabulated_results.append((se_ar['Highest Offer'].max()))

    se_fu = pd.read_excel(se_path, "Followup")
    tabulated_results.append(se_fu.shape[0])
    se_fu['Follow-Up Date'] = pd.to_datetime(se_fu['Follow-Up Date'], errors='coerce')
    tabulated_results.append(se_fu[se_fu['Follow-Up Date'] < datetime.datetime.now()].shape[0])

    se_ap = pd.read_excel(se_path, "Appointment")
    tabulated_results.append(se_ap.shape[0])
    tabulated_results.append(round(se_ap['No of Offers'].max(), 0))
    tabulated_results.append(round(se_ap['No of Offers'].mean(), 0))
    
    return tabulated_results

def calculate_qn():
    qn_path =  os.path.join(script_directory,"filtered_quotation_data.xlsx") 
    qn_an = pd.read_excel(qn_path, "Active New")
    qn_an['Highest Offer'] = clean_and_convert(qn_an['Highest Offer'])

    tabulated_results.append(qn_an.shape[0])
    tabulated_results.append(qn_an['No of Offers'].sum())
    tabulated_results.append((qn_an['Highest Offer'].sum()))
    tabulated_results.append((qn_an['Highest Offer'].max()))

    qn_ar = pd.read_excel(qn_path, "Active Requote")
    qn_ar['Highest Offer'] = clean_and_convert(qn_ar['Highest Offer'])

    tabulated_results.append(qn_ar.shape[0])
    tabulated_results.append(qn_ar['No of Offers'].sum())
    tabulated_results.append((qn_ar['Highest Offer'].sum()))
    tabulated_results.append((qn_ar['Highest Offer'].max()))

    qn_fu = pd.read_excel(qn_path, "Followup")
    qn_fu['Highest Offer'] = clean_and_convert(qn_fu['Highest Offer'])
    qn_fu['Follow-Up Date'] = pd.to_datetime(qn_fu['Follow-Up Date'], errors='coerce')

    tabulated_results.append(qn_fu.shape[0])
    tabulated_results.append(qn_fu[qn_fu['Follow-Up Date'] < datetime.datetime.now()].shape[0])
    tabulated_results.append(qn_fu['No of Offers'].sum())
    tabulated_results.append(round(qn_fu['No of Offers'].max(), 0))
    tabulated_results.append((qn_fu['Highest Offer'].sum()))
    tabulated_results.append((qn_fu['Highest Offer'].max()))

    qn_ap = pd.read_excel(qn_path, "Appointment")
    tabulated_results.append(qn_ap.shape[0])
    tabulated_results.append(round(qn_ap['No of Offers'].max(), 0))
    tabulated_results.append(round(qn_ap['No of Offers'].mean(), 0))
    
    return tabulated_results

def calculate_sold():
    sold_path =  os.path.join(script_directory,"filtered_sold_data.xlsx") 
    sold = pd.read_excel(sold_path)
    sold['Price'] = clean_and_convert(sold['Price'])

    tabulated_results.append(sold.shape[0])
    tabulated_results.append((sold['Price'].sum()))
    tabulated_results.append((sold['Price'].max()))
    
    return tabulated_results

def calculate_void():
    void_path =  os.path.join(script_directory,"filtered_void_data.xlsx") 
    void = pd.read_excel(void_path)
    tabulated_results.append(void.shape[0])
    
    return tabulated_results

def salescalculation():
    global tabulated_results
    tabulated_results = []
    calculate_new()
    calculate_se()
    calculate_qn()
    calculate_sold()
    calculate_void()
    
    results_df = pd.DataFrame([tabulated_results], columns=index)
    sales_summary_path = os.path.join(script_directory,"sales_calculations_summary.xlsx") 
    results_df.to_excel(sales_summary_path, index=False)
    
    return results_df

if __name__ == '__main__':
    salescalculation()
