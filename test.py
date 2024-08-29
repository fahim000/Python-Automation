import openpyxl
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time


workbook = openpyxl.load_workbook('4BeatsQ1.xlsx')
today = datetime.now().strftime("%A")
sheet = workbook[today]


driver = webdriver.Chrome()
driver.get('http://www.google.com')

def get_suggestions(keyword):
    print(f"Searching for: {keyword}")  

    search_box = driver.find_element("name", "q")
    search_box.clear()
    search_box.send_keys(keyword)
    time.sleep(2) 

    
    try:
        suggestions_box = driver.find_element("xpath", "//ul[@role='listbox']")
        suggestions = suggestions_box.find_elements("xpath", ".//li")
        suggestion_texts = [s.text for s in suggestions if s.text]
    except Exception as e:
        print(f"Error loading autocomplete suggestions: {e}")
        suggestion_texts = []

    print(f"Suggestions found: {suggestion_texts}")  

    if suggestion_texts:
        longest = max(suggestion_texts, key=len)
        shortest = min(suggestion_texts, key=len)
        print(f"Longest: {longest}, Shortest: {shortest}")  
        return longest, shortest
    else:
        print("No suggestions found.")  
        return None, None


for row in sheet.iter_rows(min_row=3, max_row=sheet.max_row, min_col=3, max_col=3):
    keyword = row[0].value  

    print(f"Processing row: {row[0].row}, Keyword: {keyword}")  

    if keyword:
        longest, shortest = get_suggestions(keyword)
        if longest and shortest:
            sheet.cell(row=row[0].row, column=4).value = longest  
            sheet.cell(row=row[0].row, column=5).value = shortest  
        else:
            print(f"No suggestions were added for keyword: {keyword}")
    else:
        print(f"Empty cell found at row: {row[0].row}")


workbook.save('4BeatsQ1.xlsx')
driver.quit()
