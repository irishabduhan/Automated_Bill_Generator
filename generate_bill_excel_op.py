import openpyxl

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
    print("1. Fill deatils")
    # print("2. PAN")
    # print("3. Address")
    # print("4. Date")
    # print("5. Description")
    # print("6. Amount")
    # print("7. Bill To")
    # print("8. Company Address")
    # print("9. Invoice No.")
    print("10. Exit")

# Example usage
file_path = "C:\\Users\\risha\\Desktop\\Gourav Bhai\\Bill.xlsx"  # Path to your Excel file
updates = []

while True:
    display_menu()
    choice = input("Enter your choice: ")
    if choice == "1":
        new_value = input("Enter the new name: ")
        updates.append(("B3", new_value))
        new_value = "PAN - " + input("Enter the new PAN: ")
        updates.append(("B6", new_value))
        address_line_1 = input("Enter the first line of the address: ")
        updates.append(("B4", address_line_1))
        address_line_2 = input("Enter the second line of the address: ")
        updates.append(("B5", address_line_2))
        new_date = input("Enter the new date (DD-MM-YYYY): ")
        updates.append(("H11", new_date))
        new_description = input("Enter the new description: ")
        updates.append(("B14", new_description))
        new_amount = input("Enter the new amount: ")
        updates.append(("H14", new_amount))
        new_bill_to = "Bill to - " + input("Enter the new Bill To information: ")
        updates.append(("B8", new_bill_to))
        new_company_address = "ADD - " + input("Enter the new company address: ")
        updates.append(("B9", new_company_address))
        new_company_address = input("Enter the INVOCE no: ")
        updates.append(("G11", new_company_address))
    elif choice == "2":
        break
    else:
        print("Invalid choice. Please try again.")

if updates:
    update_excel_file(file_path, updates)
else:
    print("No updates were made.")
