import os
import win32com.client

def excel_to_pdf(input_excel_path, output_pdf_path=None):
    try:
        # Ensure the input file exists
        if not os.path.exists(input_excel_path):
            print(f"Error: The file '{input_excel_path}' does not exist.")
            return

        # Define output PDF path if not provided
        if output_pdf_path is None:
            output_pdf_path = os.path.splitext(input_excel_path)[0] + '.pdf'

        # Start an instance of Excel
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False  # Run Excel in the background
        
        # Open the Excel workbook
        workbook = excel.Workbooks.Open(input_excel_path)

        # Export the workbook as a PDF
        workbook.ExportAsFixedFormat(0, output_pdf_path)  # 0 refers to xlTypePDF
        
        # Close the workbook and quit Excel
        workbook.Close(False)
        excel.Quit()
        
        print(f"Successfully saved the PDF: {output_pdf_path}")
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # Ensure Excel application quits
        excel.Quit()

# Example usage
input_excel_path = r"C:\Users\risha\Desktop\Gourav Bhai\Bill.xlsx"  # Provide the path to your Excel file
output_pdf_path = r"C:\Users\risha\Desktop\Gourav Bhai\Bill.pdf"     # Optional: Provide a custom path for the PDF

excel_to_pdf(input_excel_path, output_pdf_path)
