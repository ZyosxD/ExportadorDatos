import requests
from bs4 import BeautifulSoup
import re

def get_property_data(last_name, max_records=100):
    base_url = "https://www.utahcounty.gov/LandRecords/NameSearch.asp"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36"
    }

    records = []
    offset = 0

    while len(records) < max_records:
        params = {
            "av_name": last_name,
            "av_valid": "...",
            "offset": offset
        }

        response = requests.get(base_url, headers=headers, params=params)

        if response.status_code != 200:
            print(f"Error: Received status code {response.status_code}")
            break

        soup = BeautifulSoup(response.content, 'html.parser')

        results_table = soup.find('table', width="100%")

        if not results_table:
            print("No results table found.")
            break

        rows = results_table.find_all('tr')[1:] # Skip header row

        if not rows:
            print("No more records found.")
            break

        for row in rows:
            cols = row.find_all('td')
            if len(cols) < 6:
                continue

            owner_name = cols[0].text.strip()

            # Ownership status
            ownership_status = "Unknown"
            year_span = cols[4].find('span')
            if year_span and 'class' in year_span.attrs:
                if 'style1' in year_span['class']:
                    ownership_status = "Owner" # Green
                elif 'style2' in year_span['class']:
                    ownership_status = "Not Owner" # Red

            # Address and City
            address_full = cols[5].text.strip()
            address_parts = address_full.split(' - ')
            address = address_parts[0]
            city = address_parts[1] if len(address_parts) > 1 else ""

            # Get zip code from property page
            zip_code = ""
            property_link = cols[1].find('a')
            if property_link and 'href' in property_link.attrs:
                property_url = f"https://www.utahcounty.gov/LandRecords/{property_link['href']}"
                zip_code = get_zip_code(property_url, headers)

            records.append({
                "Name": owner_name,
                "Address": address,
                "City": city,
                "Zip Code": zip_code,
                "Ownership Status": ownership_status
            })

            if len(records) >= max_records:
                break

        # Check for "Next" link for pagination
        next_link = soup.find('a', string=re.compile(r'Next'))
        if not next_link:
            break

        offset += 100

    return records

def get_zip_code(property_url, headers):
    try:
        response = requests.get(property_url, headers=headers)
        if response.status_code == 200:
            soup = BeautifulSoup(response.content, 'html.parser')
            mailing_address_strong = soup.find('strong', string=re.compile(r'Mailing Address:'))
            if mailing_address_strong:
                mailing_address_text = mailing_address_strong.parent.text
                # Zip code is the last part of the address, usually 5 digits
                zip_match = re.search(r'\b\d{5}(?:-\d{4})?\b$', mailing_address_text)
                if zip_match:
                    return zip_match.group(0)
    except requests.exceptions.RequestException as e:
        print(f"Error fetching zip code from {property_url}: {e}")
    return ""

import csv

def save_to_csv(data, filename="utah_county_data.csv"):
    if not data:
        print("No data to save.")
        return

    keys = data[0].keys()
    with open(filename, 'w', newline='') as output_file:
        dict_writer = csv.DictWriter(output_file, keys)
        dict_writer.writeheader()
        dict_writer.writerows(data)
    print(f"Data saved to {filename}")

from openpyxl import Workbook
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors

def save_to_excel(data, filename="utah_county_data.xlsx"):
    if not data:
        print("No data to save to Excel.")
        return

    workbook = Workbook()
    sheet = workbook.active

    headers = list(data[0].keys())
    sheet.append(headers)

    for record in data:
        sheet.append(list(record.values()))

    workbook.save(filename)
    print(f"Data saved to {filename}")

def save_to_pdf(data, filename="utah_county_data.pdf"):
    if not data:
        print("No data to save to PDF.")
        return

    doc = SimpleDocTemplate(filename, pagesize=letter)
    elements = []

    headers = list(data[0].keys())
    table_data = [headers] + [list(item.values()) for item in data]

    table = Table(table_data)
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ])
    table.setStyle(style)

    elements.append(table)
    doc.build(elements)
    print(f"Data saved to {filename}")

LATINO_LAST_NAMES = [
    "Garcia", "Rodriguez", "Martinez", "Hernandez", "Lopez",
    "Gonzalez", "Perez", "Sanchez", "Ramirez", "Torres",
    "Flores", "Rivera", "Gomez", "Diaz", "Reyes",
    "Cruz", "Morales", "Ortiz", "Gutierrez", "Chavez"
]

if __name__ == "__main__":
    all_data = []
    try:
        for name in LATINO_LAST_NAMES:
            print(f"\nScraping data for last name: {name}")
            data = get_property_data(name, max_records=10)
            if data:
                print(f"Found {len(data)} records for {name}.")
                all_data.extend(data)
            else:
                print(f"No records found for {name}.")

        if all_data:
            print(f"\nTotal records scraped: {len(all_data)}")
            print("Saving data to files...")
            save_to_csv(all_data)
            save_to_excel(all_data)
            save_to_pdf(all_data)
            print("\nAll files saved successfully!")
        else:
            print("\nNo data was scraped. No files will be generated.")

    except KeyboardInterrupt:
        print("\n\nScraping process stopped by user.")
        if all_data:
            print(f"\n{len(all_data)} records were scraped before stopping.")
            save_choice = input("Do you want to save the scraped data? (y/n): ").lower()
            if save_choice == 'y':
                print("Saving data to files...")
                save_to_csv(all_data)
                save_to_excel(all_data)
                save_to_pdf(all_data)
                print("\nFiles saved successfully!")
