import openpyxl
from bs4 import BeautifulSoup

# Read the HTML file
with open("Amazon.html", "r", encoding="utf-8") as file:
    html_content = file.read()

# Parse HTML content with BeautifulSoup
soup = BeautifulSoup(html_content, "html.parser")

# Find all divs with specified class
divs = soup.find_all("div", class_="puis-card-container s-card-container s-overflow-hidden aok-relative puis-include-content-margin puis puis-v10dnrav6sitdx2esf8fqwtjtbs s-latency-cf-section puis-card-border")

# Create Excel workbook and sheet
workbook = openpyxl.Workbook()
sheet = workbook.active
sheet.title = "Amazon Products"

# Write headers
sheet["A1"] = "Product Name"
sheet["B1"] = "Product Price"
sheet["C1"] = "Product Reviews"

# Loop through divs to extract information
for index, div in enumerate(divs, start=2):  # Start from row 2 for data
    product_name = div.find("span", class_="a-size-medium a-color-base a-text-normal")
    if product_name:
        product_name = product_name.text.strip()
    else:
        product_name = " "
    
    product_price = div.find("span", class_="a-price-whole")
    if product_price:
        product_price = product_price.text.strip()
    else:
        product_price = " "
    
    product_reviews = div.find("span", class_="a-size-base")
    if product_reviews:
        product_reviews = product_reviews.text.strip()
    else:
        product_reviews = " "

    # Write data to Excel
    sheet[f"A{index}"] = product_name
    sheet[f"B{index}"] = product_price
    sheet[f"C{index}"] = product_reviews

# Save Excel file
workbook.save("Amazon_Products.xlsx")
