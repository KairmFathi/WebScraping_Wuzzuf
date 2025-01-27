from selenium.webdriver import Chrome
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
import requests
import time
from datetime import date
from math import ceil
import pandas as pd
import os
from tkinter import messagebox
import sys


# Initializing and configuring a data frame to save the result
result_df = pd.DataFrame(columns=["No.", "Job title", "Company name", "Location", "Posting date", "Applying link"])
pd.set_option('display.max_rows', None)  # Show all rows
pd.set_option('display.max_columns', None)  # Show all columns


# Setup the hidden browser option
hide_browser = Options()
hide_browser.add_argument('--headless=new')


# Initiate the browser object and load the URL
browser= Chrome(options = hide_browser)
browser.get('https://wuzzuf.net/jobs/egypt')


# Load the same website with requests to use in the loading completion check
chech_load = requests.get('https://wuzzuf.net/jobs/egypt')
while True:
    if "complete" in chech_load.text:

        # Locate the search bar and ask user to enter a searching key
        search_bar = browser.find_element(By.XPATH, " //input[@class= 'css-ukkbbr e1n2h7jb1'] ")
        search_word = input("Enter your searching keyword :")
        search_bar.send_keys(search_word, Keys.ENTER)
        break
    else:
        messagebox.showinfo ("!", "Page not loaded yet, check your connection !")


# Identify the maximum results will be appeared per each page
try:
    time.sleep(5)
    results_per_page= browser.find_element(By.XPATH, "//li")
    results_per_page = ceil(int(str(results_per_page.text).split(" ")[3]))
    print("Results per Page: ", results_per_page)
except:
    messagebox.showinfo("Invalid search", "This is search has no result, try another keyword")
    sys.exit()


# Initialize selector object to activate posting date filter on the search result page
time.sleep(5)
first_click= browser.find_element(By.XPATH, "//div [@class='css-18uqayh']")
first_click.click()

# Ask user to select specific posting date with its realted exception handeling
while True:
    try:
        date_seletion= int(input("Enter preferred Posting Date:-  [ALL >>(1) .... Past 24 hours >>(2) .... Past Week >>(3) .... Past Month >>(4)]"))
        if date_seletion == 1:
            selection= browser.find_elements(By.XPATH, "//div [@class='css-bhwo3q e1kea1u61']")[0]
            selection.click()
            break
        elif date_seletion == 2:
            selection= browser.find_elements(By.XPATH, "//div [@class='css-bhwo3q e1kea1u61']")[1]
            selection.click()
            break
        elif date_seletion == 3:
            selection= browser.find_elements(By.XPATH, "//div [@class='css-bhwo3q e1kea1u61']")[2]
            selection.click()
            break
        elif date_seletion == 4:
            selection= browser.find_elements(By.XPATH, "//div [@class='css-bhwo3q e1kea1u61']")[3]
            selection.click()
            break
        else:
            messagebox.showerror("Invalid input", "Bad choice, you should choos from 1 to 4")
    except Exception as e:
        msg = str(e) + " please enter a less number"
        messagebox.showinfo("Invalid input", msg)
  
  
#Identify number of pages in this filtered search
try:
    time.sleep(5)
    search_result= browser.find_element(By.XPATH, "//span/strong")
    search_result = int(str(search_result.text))
    num_pages= ceil(search_result / results_per_page)
    print("Number of results in this search: ", search_result)
    print("Number of pages in this search = ",num_pages, "\n", "_".center(80,"_"), "\n")
except:
    messagebox.showinfo("Invalid search", "This is search has no result, try another keyword")
    sys.exit()
       

# Initializing "next page" object if there is more than 1 paged in the search result
if num_pages == 1:
    pass
else:
    next_page = browser.find_element(By.XPATH, "//button[@class='css-zye1os ezfki8j0']")


# Parsing jobs from the first page and load details to the data frame
serial = 1
jobs= browser.find_elements(By.XPATH, "//div/h2/a [@class= 'css-o171kl']")
company= browser.find_elements(By.XPATH, "//div/div/div/a [@class = 'css-17s97q8']")
location= browser.find_elements(By.XPATH, "//div/div/div/span [@class = 'css-5wys0k']")
post_date= browser.find_elements(By.XPATH, "//div/div/div [contains (@class,'css-4c4ojb') or contains (@class, 'css-do6t5g')]")
for j, c, l, d in zip(jobs,company,location, post_date):
    result_df.loc[len(result_df)] = [serial, j.text, c.text, l.text, d.text, j.get_attribute('href')]
    serial+=1
         

# navigating through the other Pages and load jobs' details to the df
counter = 1
while counter < num_pages:
    if next_page:
         next_page.click()
         time.sleep(5)
         jobs= browser.find_elements(By.XPATH, "//div/h2/a [@class= 'css-o171kl']")
         company= browser.find_elements(By.XPATH, "//div/div/div/a [@class = 'css-17s97q8']")
         location= browser.find_elements(By.XPATH, "//div/div/div/span [@class = 'css-5wys0k']")
         post_date= browser.find_elements(By.XPATH, "//div/div/div [contains (@class,'css-4c4ojb') or contains (@class, 'css-do6t5g')]")
         for j, c, l, d in zip(jobs,company,location, post_date):
            result_df.loc[len(result_df)] = [serial, j.text, c.text, l.text, d.text, j.get_attribute('href')]
            serial+=1
    else:
         break
    next_page = browser.find_elements(By.XPATH, "//button[@class='css-zye1os ezfki8j0']")[-1]
    counter+=1


# Displaying result data frame 
print(result_df.to_string(index=False))
# print(serial, "-", j.text, "\n", c.text, "\n", l.text, "\n", d.text,"\n", j.get_attribute('href'),"\n", "_".center(80,"_"))


# Saving the results as an excel sheet on the desktop and, if not permitted, will be at the same location of the notebook
file_name = search_word.upper() + " -Wuzzuf_Search_Result - " + str(date.today()) + ".xlsx"
# result_df.to_excel(file_name, index=False)
try:
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    # file_name = "Wuzzuf_Search_Result.xlsx"
    file_path = os.path.join(desktop_path, file_name)
    result_df.to_excel(file_path, index=False)
except OSError:
    onedrive_path = r"{}\OneDrive\Desktop\{}".format(os.path.expanduser("~"), file_name)
    result_df.to_excel(onedrive_path, index=False)
except PermissionError:
    D_path = r"C:\{}".format(file_name)
    result_df.to_excel(D_path, index=False) 
except:
    result_df.to_excel(file_name, index=False)


browser.close()