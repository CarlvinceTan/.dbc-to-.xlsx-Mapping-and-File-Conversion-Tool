import os
import glob
import xlsxwriter
import re

# Set I/O directories
inputFile = glob.glob('Input/*')[0]

# Initialise .xlsx file
fileName = os.path.splitext(os.path.basename(inputFile))[0]
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


def parse_pin_config(file_path="pin_config.txt"):
    pin_connections = []
    with open(file_path, "r") as pin_config_file:
        lines = [line.strip() for line in pin_config_file if line.strip()]

    for idx in range(0, len(lines), 2):
        device_line = lines[idx]
        pin_line = lines[idx + 1]

        component1 = re.search(r"^\s*(\w+)\s*->", device_line).group(1)
        component2 = re.search(r"->\s*(\w+)\s*$", device_line).group(1)
        pin_high_from = re.search(r"(?<=High:\s)[\w]+(?=\s*->)", pin_line).group()
        pin_high_to = re.search(r"(?<=->\s)[\w]+(?=,)", pin_line).group()
        pin_low_from = re.search(r"(?<=Low:\s)[\w]+(?=\s*->)", pin_line).group()
        pin_low_to = re.search(r"\b(\w+)\b$", pin_line).group()

        pin_connections.append(
            {
                "component1": component1,
                "component2": component2,
                "pin_high_from": pin_high_from,
                "pin_high_to": pin_high_to,
                "pin_low_from": pin_low_from,
                "pin_low_to": pin_low_to,
            }
        )

    return pin_connections


pin_connections = parse_pin_config()

# Function to parse in pin configurations from pin_config.txt and .dbc
def write_row(device1, device2, signal, message, row):
    sender = device1.lower()
    receiver = device2.lower()

    matched_connections = []
    for conn in pin_connections:
        endpoints = {conn["component1"].lower(), conn["component2"].lower()}
        if sender in endpoints and receiver in endpoints:
            matched_connections.append(conn)

    # Many DBC files use Vector__XXX as SG receiver placeholder.
    # In that case, use all known links for the sender from pin_config.txt.
    if not matched_connections and receiver in {"vector__xxx", "vector__independent_sig_msg"}:
        for conn in pin_connections:
            endpoints = {conn["component1"].lower(), conn["component2"].lower()}
            if sender in endpoints:
                matched_connections.append(conn)

    for conn in matched_connections:
        component1 = conn["component1"].lower()

        worksheet.write(f"A{row}", device1)
        worksheet.write(f"B{row}", device2)
        worksheet.write(f"C{row}", signal)
        worksheet.write(f"D{row}", message)

        if component1 == sender:
            worksheet.write(f"E{row}", conn["pin_high_from"])
            worksheet.write(f"F{row}", conn["pin_high_to"])
            worksheet.write(f"G{row}", conn["pin_low_from"])
            worksheet.write(f"H{row}", conn["pin_low_to"])
        else:
            worksheet.write(f"E{row}", conn["pin_high_to"])
            worksheet.write(f"F{row}", conn["pin_high_from"])
            worksheet.write(f"G{row}", conn["pin_low_to"])
            worksheet.write(f"H{row}", conn["pin_low_from"])

        row += 1

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
