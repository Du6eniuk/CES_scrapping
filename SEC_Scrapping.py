from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime

# Path to your local ChromeDriver
driver_path = 'D:/Chrome_Driver/chromedriver-win64/chromedriver.exe'

# Initialize Chrome WebDriver using Service
service = Service(driver_path)
driver = webdriver.Chrome(service=service)

# URL of the SEC page to scrape
url = 'The link was deleted for savety reasons'

# Open the webpage
driver.get(url)

# Get the page source after it loads
html_content = driver.page_source

# Parse the page content with BeautifulSoup
soup = BeautifulSoup(html_content, 'lxml')

# Close the browser
driver.quit()

# Create a list to store the data
data = []

# Find all the rows that contain the data (adjust the tag/class as needed)
rows = soup.find_all('tr')

# Loop through each row and extract the href, name, and note
for row in rows:
    try:
        # Extract the 'a' tag (link) and the name
        a_tag = row.find('a')
        link = 'The link was deleted for savety reasons' + a_tag['href']  # Full link
        name = a_tag.text.strip()  # Extract the name
        
        # Extract the note (adjust the class or tag as needed based on the HTML structure)
        note = row.find_all('td')[-1].text.strip()  # Last 'td' contains the note

        # Append the extracted data to the list
        data.append({
            'link': link,
            'name': name,
            'note': note
        })
    except Exception as e:
        # Skip rows where extraction fails
        print(f"Error processing row: {e}")
        continue

# Convert the list of dictionaries into a pandas DataFrame
df = pd.DataFrame(data)

# Save the DataFrame to an Excel file
today = datetime.now().strftime('%Y-%m-%d')
df.to_excel(f'sec_scraped_data_{today}.xlsx', index=False)

print(f"Scraping completed. Data saved to sec_scraped_data_{today}.xlsx")