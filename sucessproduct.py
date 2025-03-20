import requests
from bs4 import BeautifulSoup
import pandas as pd

def get_url():
    """Function to get URL input from the user"""
    return input("Enter the URL: ").strip()

def fetch_webpage(url):
    """Fetch webpage content and return BeautifulSoup object"""
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 Safari/537.36"
    }
    response = requests.get(url, headers=headers)
    
    if response.status_code == 200:
        return BeautifulSoup(response.content, 'html.parser')
    else:
        print(f"Failed to retrieve the webpage. Status code: {response.status_code}")
        return None

def extract_table_data(soup):
    """Extracts table data from the webpage and returns it as a DataFrame"""
    tables = soup.find_all('table')
    all_data = []

    for table in tables:
        rows = table.find_all('tr')
        headers = [th.get_text(strip=True) for th in rows[0].find_all(['th', 'td'])] if rows else []

        for row in rows[1:]:
            cells = row.find_all(['td', 'th'])
            row_data = [cell.get_text(strip=True) for cell in cells]

            if headers:
                row_dict = {headers[i]: row_data[i] for i in range(min(len(headers), len(row_data)))}
                all_data.append(row_dict)
            else:
                all_data.append(row_data)

    return pd.DataFrame(all_data)

def save_to_excel(df, filename="table_data.xlsx"):
    """Saves the DataFrame to an Excel file"""
    if not df.empty:
        df.to_excel(filename, index=False)
        print(f"All table data extracted and saved to {filename}")
    else:
        print("No table data found on the webpage.")

def main():
    """Main function to run the program"""
    url = get_url()
    soup = fetch_webpage(url)

    if soup:
        df = extract_table_data(soup)
        save_to_excel(df)

if __name__ == "__main__":
    main()
