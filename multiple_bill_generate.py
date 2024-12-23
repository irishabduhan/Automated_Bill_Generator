import openpyxl
import shutil
import os

def update_excel_file(file_path, updates):
    """
    Update the values in the INVOICE sheet of the given Excel file.

    :param file_path: Path to the Excel file.
    :param updates: A list of tuples, where each tuple contains the cell address and the new value.
                    Example: [("B2", "New Name"), ("B5", "New PAN")]
    """
    # Load the workbook and the 'INVOICE' sheet
    workbook = openpyxl.load_workbook(file_path)
    if 'INVOICE' not in workbook.sheetnames:
        raise ValueError("Sheet 'INVOICE' does not exist in the workbook.")

    sheet = workbook['INVOICE']

    # Apply updates
    for cell_address, new_value in updates:
        sheet[cell_address] = new_value

    # Save the workbook
    workbook.save(file_path)
    print(f"Updates applied and saved to {file_path}")

def display_menu():
    """
    Display a menu for the user to select which fields to modify.
    """
    print("Select the fields to update:")
    print("1. Name")
    print("2. PAN")
    print("3. Address")
    print("4. Date")
    print("5. Description")
    print("6. Amount")
    print("7. Bill To")
    print("8. Company Address")
    print("9. Invoice No.")
    print("10. Exit")

def main():
    initial_file_path = "C:\\Users\\risha\\Desktop\\Gourav Bhai\\Bill.xlsx"  # Path to the initial Excel file
    working_file_path = "C:\\Users\\risha\\Desktop\\Gourav Bhai\\Bill_Working.xlsx"  # Temporary working file

    # Create a backup of the initial file for repeated use
    if not os.path.exists(initial_file_path):
        print(f"Error: The file '{initial_file_path}' does not exist.")
        return

    shutil.copy(initial_file_path, working_file_path)

    while True:
        updates = []
        while True:
            display_menu()
            choice = input("Enter your choice: ")
            if choice == "1":
                new_value = input("Enter the new name: ")
                updates.append(("B3", new_value))
            elif choice == "2":
                new_value = "PAN - " + input("Enter the new PAN: ")
                updates.append(("B6", new_value))
            elif choice == "3":
                address_line_1 = input("Enter the first line of the address: ")
                updates.append(("B4", address_line_1))
                address_line_2 = input("Enter the second line of the address: ")
                updates.append(("B5", address_line_2))
            elif choice == "4":
                new_date = input("Enter the new date (DD-MM-YYYY): ")
                updates.append(("H11", new_date))
            elif choice == "5":
                new_description = input("Enter the new description: ")
                updates.append(("B14", new_description))
            elif choice == "6":
                new_amount = input("Enter the new amount: ")
                updates.append(("H14", new_amount))
            elif choice == "7":
                new_bill_to = "Bill to - " + input("Enter the new Bill To information: ")
                updates.append(("B8", new_bill_to))
            elif choice == "8":
                new_company_address = "ADD - " + input("Enter the new company address: ")
                updates.append(("B9", new_company_address))
            elif choice == "9":
                new_invoice_no = input("Enter the Invoice No: ")
                updates.append(("G11", new_invoice_no))
            elif choice == "10":
                break
            else:
                print("Invalid choice. Please try again.")

        if updates:
            update_excel_file(working_file_path, updates)
        else:
            print("No updates were made.")

        # Ask the user if they want to start again with the initial file
        restart = input("Do you want to restart with the initial file? (yes/no): ").strip().lower()
        if restart == "yes":
            shutil.copy(initial_file_path, working_file_path)  # Reset the working file
            print("File reset to initial state.")
        else:
            break

if __name__ == "__main__":
    main()
