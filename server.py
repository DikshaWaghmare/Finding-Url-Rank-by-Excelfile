# finding rank of pages in given url by getting keywords using excel file and store in excel file

from flask import Flask, render_template, request
from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
import os

app = Flask(__name__)

def get_google_rankings(keywords, website_url):
    rankings = []

    try:
         # Configure the WebDriver to run in headless mode
        options = webdriver.ChromeOptions()
        options.add_argument('--headless')  # Run in headless mode
        options.add_argument('--disable-gpu')  # Disable GPU usage

        # Set up the Chrome WebDriver with the specified options
        driver = webdriver.Chrome(options=options)

        for keyword in keywords:
            # Construct the Google search URL
            query = keyword.replace(' ', '+')
            google_url = f"https://www.google.com/search?q={query}"

            # Open Google search in the browser
            driver.get(google_url)

            # Find and interact with search results
            search_results = driver.find_elements(By.CLASS_NAME, 'tF2Cxc')

            for index, result in enumerate(search_results):
                # Skip "People also ask" results
                if "related-question" in result.get_attribute('class'):
                    continue

                result_url = result.find_element(By.TAG_NAME, 'a').get_attribute('href')
                if website_url in result_url:
                    rankings.append({"Keywords": keyword, "Rank": index + 1, "URL": result_url})

    except Exception as e:
        print(f"An error occurred: {str(e)}")
    finally:
        driver.quit()  # Make sure to quit the browser when done

    return rankings

# display index page(form)
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/check_rankings', methods=['POST'])
def check_rankings():
    excel_file = request.files['excel_file']
    website_url = request.form['website_url']

    if not excel_file:
        return "No file uploaded."

    # Read keywords from the uploaded Excel file
    df = pd.read_excel(excel_file)

    # Check if the DataFrame is empty
    if df.empty:
        return render_template('index.html', empty_file_message="The uploaded Excel file is empty. Please upload a file with data.")

    keywords = df['keywords'].tolist()

    rankings = get_google_rankings(keywords, website_url)

    # Check if rankings are empty after processing
    if not rankings:
        return render_template('index.html', empty_data_message="No data found for the provided keywords and website URL.")

    # Create a DataFrame from the rankings
    output_df = pd.DataFrame(rankings)

    # Define the path for the output Excel file
    output_file_path = 'output_rankings.xlsx'

    # Check if the file exists, delete it to avoid appending to existing data
    if os.path.exists(output_file_path):
        os.remove(output_file_path)

    # Write the rankings to a new Excel file
    output_df.to_excel(output_file_path, index=False)

    return render_template('index.html', rankings=rankings)

if __name__ == '__main__':
    app.run(debug=True)