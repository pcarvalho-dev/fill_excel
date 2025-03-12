import re
import openpyxl


def extract_data(text):
    """Extracts structured data from the input text."""
    data = {}

    # Example regex patterns (customize these based on your data structure)
    data['Name'] = re.search(r'Name:\s*(.*)', text)
    data['Email'] = re.search(r'Email:\s*([\w.-]+@[\w.-]+)', text)
    data['Phone'] = re.search(r'Phone:\s*(\+?\d[\d\s-]+)', text)
    data['Date'] = re.search(r'Date:\s*(\d{2}/\d{2}/\d{4})', text)

    # Extract matched values or set to empty string if not found
    for key in data:
        data[key] = data[key].group(1) if data[key] else ''

    return data


def write_to_excel(data, filename="output.xlsx"):
    """Writes extracted data to an Excel file."""
    try:
        wb = openpyxl.load_workbook(filename)
        sheet = wb.active
    except FileNotFoundError:
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.append(["Name", "Email", "Phone", "Date"])  # Header row

    sheet.append([data['Name'], data['Email'], data['Phone'], data['Date']])
    wb.save(filename)
    print(f"Data saved to {filename}")


if __name__ == "__main__":
    text = input("Enter the text to extract data from:\n")
    extracted_data = extract_data(text)
    write_to_excel(extracted_data)
