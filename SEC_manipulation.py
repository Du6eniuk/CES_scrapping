from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime

# Path to your local ChromeDriver
driver_path = 'D:/Chrome_Driver/chromedriver-win64/chromedriver.exe'

# Load the scraped data from the first Excel file
today = datetime.now().strftime('%Y-%m-%d')
input_file = f'sec_scraped_data_{today}.xlsx'
df = pd.read_excel(input_file)

# Create an empty list to store the transformed data
transformed_data = []

# Initialize Chrome WebDriver using Service
service = Service(driver_path)
driver = webdriver.Chrome(service=service)

# Loop through each row and process the link
for index, row in df.iterrows():
    url = row['link']  # Get the URL from the Link column
    name = row['name']  # Get the Name from the row
    note = row['note']  # Get the Note from the row
    
    print(f"Processing: {url}")

    # Visit the link and extract the required data
    driver.get(url)
    html_content = driver.page_source
    soup = BeautifulSoup(html_content, 'lxml')

    # Extract the date from the "Last Reviewed or Updated"
    try:
        date_updated = soup.find('div', class_='date-modified usa-prose').find('span', class_='nowrap').text.strip()
    except AttributeError:
        date_updated = ''  # Leave blank if not found

    # Try to find the website href or any valid href that is not an email
    try:
        website_section = soup.find_all('p')
        website_href = ''  # Default value in case no link is found

        for p in website_section:
            # Find all <a> tags in the <p> and check them one by one
            links = p.find_all('a', href=True)
            for link in links:
                possible_href = link['href'].strip()
                # Ignore emails and find a valid website href
                if not possible_href.startswith('mailto:'):
                    website_href = possible_href
                    break
            if website_href:  # Break outer loop if a valid website is found
                break
    except Exception as e:
        website_href = ''  # Leave blank if not found
        print(f"Error: {e}")

    # Append the transformed data in the correct format
    transformed_data.append({
        'Date Added': datetime.now().strftime('%d-%m-%Y'),
        'Country': 'United States of America',
        'Source': 'SEC',
        'Name': name,
        'PULLED LINKS': website_href,
        'Cleaning': '',  # Empty field
        'Notes': note,
        'Date added to website': date_updated,
        'CASE': url
    })

# Close the browser
driver.quit()

# Convert the transformed data into a DataFrame
transformed_df = pd.DataFrame(transformed_data)

# Save the transformed data to a new Excel file
output_file = f'sec_transformed_data_{today}.xlsx'
transformed_df.to_excel(output_file, index=False)

print(f"Transformation completed. Data saved to {output_file}")