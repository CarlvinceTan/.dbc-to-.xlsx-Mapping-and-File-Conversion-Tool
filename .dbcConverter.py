import os
import glob
import xlsxwriter
import re

# Set I/O directories
inputFile = glob.glob('Input/*')[0]

# Initialise .xlsx file
fileName = re.search(r"(?<=/)([\w]+)(?=.)", inputFile).group()
workbook = xlsxwriter.Workbook(os.path.join("Output", f"{fileName}.xlsx"))
worksheet = workbook.add_worksheet()

# Name Header
cell_format = workbook.add_format({
    "align": "center",
    "valign": "vcenter",
})
worksheet.merge_range("A1:B1", "Device", cell_format)
worksheet.write("A2", "From", cell_format)
worksheet.write("B2", "To", cell_format)
worksheet.merge_range("C1:C2", "Signal", cell_format)
worksheet.merge_range("D1:D2", "Message/Address ID", cell_format)
worksheet.merge_range("E1:F1", "Pin (High)", cell_format)
worksheet.write("E2", "From", cell_format)
worksheet.write("F2", "To", cell_format)
worksheet.write("G2", "From", cell_format)
worksheet.write("H2", "To", cell_format)
worksheet.merge_range("G1:H1", "Pin (Low)", cell_format)
    
# Setup arrays of string max lengths in each column
max_lengths = [len("From"), len("To"), len("Signal"), len("Message/Address ID"), len("From"), len("To"), len("From"), len("To")]
def update_max_lengths(string, col):
    if len(string) > max_lengths[col]:
        max_lengths[col] = len(string)

# Function to parse in pin configurations from pin_config.txt and .dbc
def write_row(device1, device2, signal, message, row): 
    lineCount = 1
    devices = [device1.lower(), device2.lower()]
    with open("pin_config.txt", 'r') as pinConfigFile:
        for line in pinConfigFile:
            if lineCount % 3 == 1:
                component1 = re.search(r'^\b(\w+)\b', line).group().lower()
                component2 = re.search(r'\b(\w+)\b$', line).group().lower()
            if lineCount % 3 == 2 and component1 in devices and component2 in devices:        
                pinHighFrom = re.search(r"(?<=High:\s)[\w]+(?=\s*->)", line).group()
                pinHighTo = re.search(r"(?<=->\s)[\w]+(?=,)", line).group()
                pinLowFrom = re.search(r"(?<=Low:\s)[\w]+(?=\s*->)", line).group()
                pinLowTo = re.search(r"\b(\w+)\b$", line).group()
                if component1 == deviceFrom:
                    worksheet.write(f"A{row}", deviceFrom)
                    worksheet.write(f"B{row}", deviceTo)
                    worksheet.write(f"C{row}", signal)
                    worksheet.write(f"D{row}", message)
                    worksheet.write(f"E{row}", pinHighFrom)
                    worksheet.write(f"F{row}", pinHighTo)
                    worksheet.write(f"G{row}", pinLowFrom)
                    worksheet.write(f"H{row}", pinLowTo)
                    row += 1
                else:   # Swap to make sure writing in correct from and to columns
                    worksheet.write(f"A{row}", deviceFrom)
                    worksheet.write(f"B{row}", deviceTo)
                    worksheet.write(f"C{row}", signal)
                    worksheet.write(f"D{row}", message)
                    worksheet.write(f"E{row}", pinHighTo)
                    worksheet.write(f"F{row}", pinHighFrom)
                    worksheet.write(f"G{row}", pinLowTo)
                    worksheet.write(f"H{row}", pinLowFrom)
                row += 1
            lineCount += 1

    pinConfigFile.close()
    return row

# Commence Main Parsing Logic!
with open(inputFile, 'r') as dbcFile:
    row = 3
    prev = False
    for line in dbcFile:
        if line.startswith("BO_"):
            deviceFrom = re.search(r"\b([\w]+)\b$", line).group()
            update_max_lengths(deviceFrom, 0)
            message = re.search(r"\b([\w]+)\b(?=:)", line).group()
            update_max_lengths(message, 3)
            prev = True
        elif line.startswith(" SG_"):
            deviceTo = re.search(r"\b([\w]+)\b$", line).group()
            update_max_lengths(deviceTo, 1)
            signal = re.search(r"\b([\w]+)\b(?= :)", line).group()
            update_max_lengths(signal, 2)

            row = write_row(deviceFrom, deviceTo, signal, message, row)  # Also increments row
            

# Resize width of columns
for col, max_len in enumerate(max_lengths):
    worksheet.set_column(col, col, max_len + 2)

workbook.close()
dbcFile.close()
