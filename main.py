import sys

import selenium
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import pandas as pd
from re import search


print("Ivan's License Plate Checker Program")

# variables

choice = input("Enter 1 to set a custom filepath or enter 0 to use the default filepath (/Users/ivan/Desktop/10kwords.xlsx) \n")
if (choice == "1"):
    filepath = input("Enter filepath: ")
else:
    filepath = "/Users/ivan/Desktop/10kwords.xlsx"

# loads excel sheet
DN = pd.read_excel(filepath)

# prints column name
print(DN.columns)

# list of words to try
wordslist = DN['the'].tolist()

# index of word in list
index = 0

# starts up chrome driver and goes to license plate website
browser = webdriver.Chrome()
browser.get('https://transactions.dmv.virginia.gov/dmvnet/plate_purchase/select_plate.asp')

# checks if license plate type frame is loaded and switches to it
element2 = WebDriverWait(browser, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, 's2plttype')))

# finds form for license plate type
Box1 = browser.find_element(By.XPATH, '//select[@NAME="plates"]')

# finds the option "Scenic Autumn" and selects it
Box1.find_element(By.XPATH, "//select[@NAME='plates']/option[text()='Scenic Autumn']").click()

print("Available Plates: \n")
# repeats program for each word
for x in wordslist:

    # variable denoting confirmed(1) or unconfirmed(0) word length requirements
    i = 0

    # while loop used to skip words longer than 7 characters
    while i < 1:

        # loads current word
        word = wordslist[index]

        # gets length of the word and changes up index to try the next word
        if len(str(word)) > 7:
            index += 1

        # if the word is less than 7 characters it passes and the program continues
        else:
            i += 1

    # initiaties list of letters and adds each letter of word to it
    list1 = list
    list1 = list(str(word))

    # exits all frames back to home page (needed after every new frame change)
    browser.switch_to.default_content()

    # switches to the frame with the text boxes so it can locate
    browser.switch_to.frame('s2end')

    # makes sure all at least 1 of the input boxes is loaded
    element = WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, '//input[@NAME = "Let1"]')))

    # checks if theres a letter in the index and finds letter input boxes and sends designated letter
    if len(list1) > 0:
        Box1 = browser.find_element(By.XPATH, '//input[@NAME = "Let1"]')
        Box1.send_keys(list1[0])
        if len(list1) > 1:
            Box1 = browser.find_element(By.XPATH, '//input[@NAME = "Let2"]')
            Box1.send_keys(list1[1])
            if len(list1) > 2:
                Box1 = browser.find_element(By.XPATH, '//input[@NAME = "Let3"]')
                Box1.send_keys(list1[2])
                if len(list1) > 3:
                    Box1 = browser.find_element(By.XPATH, '//input[@NAME = "Let4"]')
                    Box1.send_keys(list1[3])
                    if len(list1) > 4:
                        Box1 = browser.find_element(By.XPATH, '//input[@NAME = "Let5"]')
                        Box1.send_keys(list1[4])
                        if len(list1) > 5:
                            Box1 = browser.find_element(By.XPATH, '//input[@NAME = "Let6"]')
                            Box1.send_keys(list1[5])
                            if len(list1) > 6:
                                Box1 = browser.find_element(By.XPATH, '//input[@NAME = "Let7"]')
                                Box1.send_keys(list1[6])

    # clicks the view plate button
    browser.find_element(By.XPATH, "//input[@value='View Plate']").click()

    # checks if plate is good and returns the name if yes
    Box1 = browser.find_element(By.NAME, 'PltNo')
    if "Congratulations" in browser.page_source:
        print(Box1.get_attribute('value'))

    # clear form
    browser.find_element(By.XPATH, "//input[@value='Clear Form']").click()

    index += 1
