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

        # Close the workbook
        workbook.Close(False)
        print(f"Successfully saved the PDF: {output_pdf_path}")
    except Exception as e:
        print(f"An error occurred while processing '{input_excel_path}': {e}")
    finally:
        # Ensure Excel application quits
        excel.Quit()

def convert_all_excel_to_pdf(directory_path):
    try:
        # Ensure the directory exists
        if not os.path.exists(directory_path):
            print(f"Error: The directory '{directory_path}' does not exist.")
            return

        # Iterate through all Excel files in the directory
        for filename in os.listdir(directory_path):
            if filename.endswith(('.xls', '.xlsx', '.xlsm')):  # Check for Excel files
                input_excel_path = os.path.join(directory_path, filename)
                output_pdf_path = os.path.splitext(input_excel_path)[0] + '.pdf'

                # Call the excel_to_pdf function
                excel_to_pdf(input_excel_path, output_pdf_path)

    except Exception as e:
        print(f"An error occurred: {e}")

# Example usage
directory_path = r"C:\Users\risha\Desktop\Gourav Bhai"  # Replace with the path to your directory containing Excel files
convert_all_excel_to_pdf(directory_path)
