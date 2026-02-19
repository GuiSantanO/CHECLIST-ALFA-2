import pandas as pd
import os
import win32com.client

def test_pdf_export():
    df = pd.DataFrame({"Modelo": ["HP EliteBook", "Dell Latitude"], "RAM": ["8GB", "16GB"]})
    temp_excel = os.path.abspath("temp_test.xlsx")
    out_pdf = os.path.abspath("temp_test.pdf")
    
    df.to_excel(temp_excel, index=False)
    
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(temp_excel)
        wb.ExportAsFixedFormat(0, out_pdf)
        wb.Close(False)
        excel.Quit()
        print(f"Success! PDF created at {out_pdf}")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    test_pdf_export()
