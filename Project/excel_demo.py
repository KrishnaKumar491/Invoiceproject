# from openpyxl import Workbook

# # Create workbook
# workbook = Workbook()

# # Get active worksheet
# sheet = workbook.active

# # Rename worksheet
# sheet.title = "Invoices"

# # Add headers
# sheet["A1"] = "Invoice ID"
# sheet["B1"] = "Vendor"
# sheet["C1"] = "Amount"
# sheet["D1"] = "Status"

# # Add invoice details
# sheet["A2"] = 1001
# sheet["B2"] = "ABC Technologies"
# sheet["C2"] = 25000
# sheet["D2"] = "Approved"

# # Save Excel file
# workbook.save("invoice_report.xlsx")

# print("Excel file created successfully!")


# from openpyxl import load_workbook

# # Load existing Excel file
# workbook = load_workbook("invoice_report.xlsx")

# # Select worksheet
# sheet = workbook["Invoices"]

# # Read individual cells
# print("Invoice ID:", sheet["A2"].value)
# print("Vendor:", sheet["B2"].value)
# print("Amount:", sheet["C2"].value)
# print("Status:", sheet["D2"].value)


# from openpyxl import Workbook

# workbook = Workbook()
# sheet = workbook.active
# sheet.title = "Report"
# workbook.save("TestReport.xlsx")

# from openpyxl import load_workbook

# workbook = load_workbook("invoice_report.xlsx")
# sheet = workbook["Invoices"]

# for row in sheet.iter_rows(min_row=2, values_only=True):
#     print(row)


# from openpyxl import load_workbook

# workbook = load_workbook("invoice_report.xlsx")
# sheet = workbook["Invoices"]

# invoices = [
#     [1002, "XYZ Solutions", 35000, "Pending"],
#     [1003, "PQR Services", 45000, "Approved"],
#     [1004, "LMN Technologies", 15000, "Rejected"]
# ]

# for invoice in invoices:
#     sheet.append(invoice)

# workbook.save("invoice_report.xlsx")

# print("Invoice records added successfully!")

# from openpyxl import load_workbook

# workbook = load_workbook("invoice_report.xlsx")
# sheet = workbook["Invoices"]

# for row in sheet.iter_rows(min_row=2, values_only=True):
#     invoice_id, vendor, amount, status = row

#     if status == "Approved" and amount > 30000:
#         print(f"Invoice ID: {invoice_id}, Vendor: {vendor}, Amount: {amount}")

# from openpyxl import load_workbook

# workbook = load_workbook("invoice_report.xlsx")
# sheet = workbook["Invoices"]

# for row in range(2, sheet.max_row + 1):
#     invoice_id = sheet[f"A{row}"].value

#     if invoice_id == 1002:
#         sheet[f"D{row}"] = "Approved"
#         print("Invoice status updated successfully")

# workbook.save("invoice_report.xlsx")



# from openpyxl import load_workbook
# import time

# file_path = "invoice_report.xlsx"

# for attempt in range(3):
#     workbook = None

#     try:
#         workbook = load_workbook(file_path)

#         sheet = workbook["Invoices"]
#         sheet["D3"] = "Approved"

#         workbook.save(file_path)

#         print("Excel updated successfully")
#         break

#     except PermissionError:
#         print("Excel file is locked. Close it and retry.")

#         if attempt < 2:
#             time.sleep(5)

#     finally:
#         if workbook:
#             workbook.close()


from openpyxl import load_workbook, Workbook

# Open source Excel file
workbook = load_workbook("invoice_report.xlsx")
sheet = workbook["Invoices"]

# Create a new workbook for the report
report_workbook = Workbook()
report_sheet = report_workbook.active
report_sheet.title = "Approved Invoices"

# Add column headers
report_sheet.append(["Invoice ID", "Vendor", "Amount", "Status"])

# Read and filter invoices
for row in range(2, sheet.max_row + 1):

    invoice_id = sheet[f"A{row}"].value
    vendor = sheet[f"B{row}"].value
    amount = sheet[f"C{row}"].value
    status = sheet[f"D{row}"].value

    if status == "Approved" and amount > 30000:

        approved_invoice = [
            invoice_id,
            vendor,
            amount,
            status
        ]

        report_sheet.append(approved_invoice)

# Save the new report
report_workbook.save("approved_invoice_report.xlsx")

# Close both workbooks
workbook.close()
report_workbook.close()

print("Approved invoice report generated successfully!")





