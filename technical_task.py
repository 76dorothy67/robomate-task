import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
import requests
import time


class RPA_Challenge:

    def __init__(self):
        self.driver = webdriver.Chrome()

    def download_document(self, link, document_name):
        response = requests.get(link)
        with open(document_name, "wb") as file:
            file.write(response.content)

    def get_data_frame(self, document_name):
        data_frame = pd.read_excel(document_name)
        return data_frame

    def button_start(self,button_name):
        button_start = self.driver.find_element(By.XPATH, f"//button[contains(text(), '{button_name}')]")
        return button_start.click()

    def input_data(self, label_name, input_data):
        label = self.driver.find_element(By.XPATH, f"//label[text()='{label_name}']")
        input = label.find_element(By.XPATH, "following-sibling::input")
        input.send_keys(input_data)


    def button_submit(self,button_name):
        button_submit = self.driver.find_element(By.XPATH, f"//input[@value='{button_name}']")
        return button_submit.click()


if __name__ == '__main__':

    bot = RPA_Challenge()

    url = 'https://rpachallenge.com'
    bot.driver.get(url)

    download_link = "https://rpachallenge.com/assets/downloadFiles/challenge.xlsx"
    document = 'data.xlsx'

    bot.download_document(download_link, document)

    df = bot.get_data_frame(document)

    bot.button_start('Start')

    df.columns = df.columns.str.strip()
    labels_name = df.columns.tolist()

    for i in range(len(df)):
        for j in labels_name:
            if j == "Phone Number":
                bot.input_data(j, str(df.at[i, j]))  # Конвертувати у рядок
            else:
                bot.input_data(j, df.at[i, j])
        bot.button_submit('Submit')

    time.sleep(5)

    bot.driver.quit()
