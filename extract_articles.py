import pandas as pd
import requests
from bs4 import BeautifulSoup
import os

# Load the input Excel file
input_file_path = 'Input.xlsx'  # Adjust the path as necessary
input_df = pd.read_excel(input_file_path)

# Function to extract article title and text
def extract_article(url):
    response = requests.get(url)
    soup = BeautifulSoup(response.content, 'html.parser')
    
    # Assuming the article title is in a <h1> tag and the main content is in <p> tags
    # These selectors may need to be adjusted based on the actual HTML structure of the articles
    title = soup.find('h1').get_text() if soup.find('h1') else 'No Title Found'
    paragraphs = soup.find_all('p')
    text = '\n'.join([para.get_text() for para in paragraphs])
    
    return title, text

# Directory to save the extracted articles
output_dir = 'extracted_articles_1'
os.makedirs(output_dir, exist_ok=True)

# Extract and save articles
for index, row in input_df.iterrows():
    url_id = row['URL_ID']
    url = row['URL']
    try:
        title, text = extract_article(url)
        with open(f'{output_dir}/{url_id}.txt', 'w', encoding='utf-8') as file:
            file.write(f'{title}\n\n{text}')
        print(f'Successfully extracted article {url_id}')
    except Exception as e:
        print(f'Failed to extract article {url_id}: {e}')
