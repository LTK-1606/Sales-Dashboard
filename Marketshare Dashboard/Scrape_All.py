import Consignment
import Quotation
import ScrapExport
import Sold
import Void
import New
import SalesDashboard
import Consolidate_Format_Data
import Combine_Data
import Delete_Excel
import Marketshare

def main():
    New.main_new()
    ScrapExport.main_scrapexport()
    Quotation.main_quotation()
    Consignment.main_consignment()
    Sold.main_sold()
    Void.main_void()
    SalesDashboard.main_salesdashboard()
    Consolidate_Format_Data.main_consolidate_format_data()
    Combine_Data.main_combine_data()
    Delete_Excel.main_delete()
    Marketshare.main_marketshare()
    
if __name__ == '__main__':
    main()
