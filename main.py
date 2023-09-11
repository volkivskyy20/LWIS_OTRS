import os
import sys
import time

import pandas as pd
from PySide6 import QtWidgets
from PySide6.QtWidgets import QFileDialog
from element_manager import *
from pynput.keyboard import Key, Controller
from selenium import webdriver
from selenium.webdriver.common.by import By

from ui import Ui_MainWindow


class MyWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.file_path = None
        self.file_path1 = None
        self.directory = None
        self.file_path3 = None
        self.directorymain = None

        # Create instance variables to store file paths
        self.selected_week_file = None
        self.selected_email_file = None
        self.selected_1151_file = None

        self.ui.commandLinkButton.clicked.connect(self.run_clicked)
        self.ui.btn_week.clicked.connect(self.select_week_file)
        self.ui.btn_snd.clicked.connect(self.select_email_file)
        self.ui.pushButton.clicked.connect(self.prepare)
        self.ui.pushButton_5.clicked.connect(self.prepare1)
        self.ui.pushButton_2.clicked.connect(self.snd_1151)
        self.ui.btn_dir.clicked.connect(self.dir1151)
        self.ui.btn_dir_2.clicked.connect(self.dirmain)
        self.ui.btn_prt_1151.clicked.connect(self.select_1151)
        self.ui.pushButton_3.clicked.connect(self.crtmaintkt)
        self.ui.pushButton_4.clicked.connect(self.crt1151tkt)

    def prepare1(self):
        # Your existing code for reading and processing the Excel file
        input_excel_path = self.file_path
        output_excel_path = 'prepared_emails.xlsx'
        columns_to_extract = ['Relation', 'Email']
        data = pd.read_excel(input_excel_path, usecols=columns_to_extract)
        data = data.drop_duplicates()

        # Split the 'Email' column by ';' or ',' and keep only the first part
        data['Email'] = data['Email'].str.split('[;,]').str[0].str.strip()

        data = data.query(
            'Relation != 1156 and Relation != 1018 and Relation != 1293 and Relation != 1294 and Relation != 1295 and '
            'Relation != 1296 and Relation != 1522 and Relation != 1151 and Relation != 1383')

        # Save the processed data to a new Excel file
        data.to_excel(output_excel_path, index=False)

        # Your code for creating separate Excel files from the second Excel file
        input_excel = self.selected_1151_file
        df = pd.read_excel(input_excel)
        for i, row in df.iterrows():
            split_df = pd.DataFrame([row], columns=df.columns)
            output_file = f'Partner_1151_{i + 1}.xlsx'
            split_df.to_excel(output_file, index=False)

        print(f"{len(df)} Excel files created by splitting each row.")

    def crt1151tkt(self):
        folder_path = self.directory
        file_list = os.listdir(folder_path)

        oklog = self.ui.okta_login.text()  # Get text from okta_login QLineEdit
        okpass = self.ui.okta_pass.text()  # Get text from okta_pass QLineEdit
        otrs_login = self.ui.otrs_login.text()  # Get text from otrs_login QLineEdit
        otrs_pass = self.ui.otrs_pass.text()  # Get text from otrs_pass QLineEdit
        driver = webdriver.Chrome()
        excel_file = self.file_path1
        # to open the url in browser
        options = webdriver.ChromeOptions()
        driver.get('https://rhenus.okta.com/')
        options.add_experimental_option('detach', True)
        # to type content in input field
        driver.find_element(By.XPATH, get_xpath(driver, '126YJaa68GerZcM')).send_keys(oklog)

        # to click on input field
        driver.find_element(By.XPATH, get_xpath(driver, 'f1GMPJOh2_4aGTs')).click()

        # to type content in input field
        driver.find_element(By.XPATH, get_xpath(driver, 'dVdU1X5TlD1mXYy')).send_keys(okpass)
        # to click on input field
        driver.find_element(By.XPATH, get_xpath(driver, 'GtyGHu4AyVcmHk0')).click()
        time.sleep(25)
        driver.get('https://servicecentre.fl-app.rhenus.com/otrs/index.pl')

        # to click on input field
        driver.find_element(By.XPATH, get_xpath(driver, 'J0umjEiaBJtThVN')).click()

        # to type content in input field
        driver.find_element(By.XPATH, get_xpath(driver, 'SJW8GnK3RynhwLa')).send_keys(otrs_login)

        # to click on input field
        driver.find_element(By.XPATH, get_xpath(driver, 'Xg7h4qepocPSEYa')).click()

        # to type content in input field
        driver.find_element(By.XPATH, get_xpath(driver, 'WB0VHg1ua_feRSb')).send_keys(otrs_pass)

        # to click on the element(Login) found
        driver.find_element(By.XPATH, get_xpath(driver, 'wdz8eE4RTiLW0dT')).click()
        time.sleep(5)
        numbers = []
        field_values = []

        for filename in file_list:
            driver.get('https://servicecentre.fl-app.rhenus.com/otrs/index.pl?Action=AgentTicketPhone')
            time.sleep(5)
            driver.find_element(By.XPATH,
                                '/html/body/div[1]/div[3]/div[1]/div[2]/form/fieldset/div[2]/div[1]/div/input').click()
            time.sleep(2)
            driver.find_element(By.PARTIAL_LINK_TEXT, "Unclassified").click()
            time.sleep(2)
            driver.find_element(By.XPATH, '//*[@id="Dest_Search"]').send_keys("sc")
            time.sleep(1)
            driver.find_element(By.PARTIAL_LINK_TEXT, "SC Common").click()
            time.sleep(1)
            driver.find_element(By.XPATH, '//*[@id="FromCustomer"]').send_keys("barcelona@es.rhenus.com")

            time.sleep(1)
            driver.find_element(By.XPATH, '//*[@id="Subject"]').click()

            driver.find_element(By.XPATH, '//*[@id="Subject"]').send_keys(filename)
            time.sleep(1)
            field_value_element = driver.find_element(By.XPATH, '//*[@id="Subject"]')
            field_value = field_value_element.text
            field_values.append(field_value)
            iframe = driver.find_element(By.XPATH, "//iframe[@title='Rich Text Editor, RichText']")
            time.sleep(1)
            text_to_input = "**"
            driver.switch_to.frame(iframe)
            contenteditable_field = driver.find_element(By.XPATH, "//body[@contenteditable='true']")
            contenteditable_field.send_keys(text_to_input)
            time.sleep(10)

            current_url = driver.current_url
            keyword = "TicketID="
            if keyword in current_url:
                ticket_id = current_url.split(keyword)[1]
                print(ticket_id)

                numbers.append(ticket_id)

        datas = {
            "TicketID": numbers,
            "Field_Value": field_values
        }
        dfs = pd.DataFrame(datas)

        # Save DataFrame to an Excel file
        excel_file_path = "Created tickets 1151.xlsx"
        dfs.to_excel(excel_file_path, index=False)
        sys.exit()

    def crtmaintkt(self):
        oklog = self.ui.okta_login.text()  # Get text from okta_login QLineEdit
        okpass = self.ui.okta_pass.text()  # Get text from okta_pass QLineEdit
        otrs_login = self.ui.otrs_login.text()  # Get text from otrs_login QLineEdit
        otrs_pass = self.ui.otrs_pass.text()  # Get text from otrs_pass QLineEdit
        driver = webdriver.Chrome()
        excel_file = self.file_path1
        # to open the url in browser
        options = webdriver.ChromeOptions()
        driver.get('https://rhenus.okta.com/')
        options.add_experimental_option('detach', True)
        # to type content in input field
        driver.find_element(By.XPATH, get_xpath(driver, '126YJaa68GerZcM')).send_keys(oklog)

        # to click on input field
        driver.find_element(By.XPATH, get_xpath(driver, 'f1GMPJOh2_4aGTs')).click()

        # to type content in input field
        driver.find_element(By.XPATH, get_xpath(
            driver, 'dVdU1X5TlD1mXYy')).send_keys(okpass)

        # to click on input field
        driver.find_element(By.XPATH, get_xpath(driver, 'GtyGHu4AyVcmHk0')).click()
        time.sleep(25)
        driver.get('https://servicecentre.fl-app.rhenus.com/otrs/index.pl')

        # to click on input field
        driver.find_element(By.XPATH, get_xpath(driver, 'J0umjEiaBJtThVN')).click()

        # to type content in input field
        driver.find_element(By.XPATH, get_xpath(driver, 'SJW8GnK3RynhwLa')).send_keys(otrs_login)

        # to click on input field
        driver.find_element(By.XPATH, get_xpath(driver, 'Xg7h4qepocPSEYa')).click()

        # to type content in input field
        driver.find_element(By.XPATH, get_xpath(driver, 'WB0VHg1ua_feRSb')).send_keys(otrs_pass)

        # to click on the element(Login) found
        driver.find_element(By.XPATH, get_xpath(driver, 'wdz8eE4RTiLW0dT')).click()
        driver.get('https://servicecentre.fl-app.rhenus.com/otrs/index.pl?Action=AgentTicketPhone')
        time.sleep(5)
        numbers = []
        field_values = []

        data = pd.read_excel(excel_file)

        for index, row in data.iterrows():
            driver.get('https://servicecentre.fl-app.rhenus.com/otrs/index.pl?Action=AgentTicketPhone')
            time.sleep(5)
            driver.find_element(By.XPATH,
                                '/html/body/div[1]/div[3]/div[1]/div[2]/form/fieldset/div[2]/div[1]/div/input').click()
            time.sleep(2)
            driver.find_element(By.PARTIAL_LINK_TEXT, "Unclassified").click()
            time.sleep(2)
            driver.find_element(By.XPATH, '//*[@id="Dest_Search"]').send_keys("sc")
            time.sleep(2)
            driver.find_element(By.PARTIAL_LINK_TEXT, "SC Common").click()
            time.sleep(2)
            driver.find_element(By.XPATH, '//*[@id="FromCustomer"]').send_keys(row['Email'])
            time.sleep(1)
            driver.find_element(By.XPATH, '//*[@id="Subject"]').click()
            rel = str(row['Relation'])
            driver.find_element(By.XPATH, '//*[@id="Subject"]').send_keys('Partner ' + rel)
            time.sleep(1)
            field_value_element = driver.find_element(By.XPATH, '//*[@id="Subject"]')
            field_value = field_value_element.text
            field_values.append(field_value)
            iframe = driver.find_element(By.XPATH, "//iframe[@title='Rich Text Editor, RichText']")
            text_to_input = "**"
            driver.switch_to.frame(iframe)
            contenteditable_field = driver.find_element(By.XPATH, "//body[@contenteditable='true']")
            contenteditable_field.send_keys(text_to_input)
            time.sleep(10)
            current_url = driver.current_url
            keyword = "TicketID="
            if keyword in current_url:
                ticket_id = current_url.split(keyword)[1]
                print(ticket_id)

                numbers.append(ticket_id)

        datas = {
            "TicketID": numbers,
            "Field_Value": field_values
        }
        dfs = pd.DataFrame(datas)

        # Save DataFrame to an Excel file
        excel_file_path = "Created tickets.xlsx"
        dfs.to_excel(excel_file_path, index=False)
        sys.exit()

    def snd_1151(self):
        oklog = self.ui.okta_login.text()  # Get text from okta_login QLineEdit
        okpass = self.ui.okta_pass.text()  # Get text from okta_pass QLineEdit
        otrs_login = self.ui.otrs_login.text()  # Get text from otrs_login QLineEdit
        otrs_pass = self.ui.otrs_pass.text()  # Get text from otrs_pass QLineEdit
        driver = webdriver.Chrome()
        excel_file = self.file_path1
        # to open the url in browser
        options = webdriver.ChromeOptions()
        driver.get('https://rhenus.okta.com/')
        options.add_experimental_option('detach', True)
        # to type content in input field
        driver.find_element(By.XPATH, get_xpath(driver, '126YJaa68GerZcM')).send_keys(oklog)

        # to click on input field
        driver.find_element(By.XPATH, get_xpath(driver, 'f1GMPJOh2_4aGTs')).click()

        # to type content in input field
        driver.find_element(By.XPATH, get_xpath(
            driver, 'dVdU1X5TlD1mXYy')).send_keys(okpass)

        # to click on input field
        driver.find_element(By.XPATH, get_xpath(driver, 'GtyGHu4AyVcmHk0')).click()
        time.sleep(25)
        driver.get('https://servicecentre.fl-app.rhenus.com/otrs/index.pl')

        # to click on input field
        driver.find_element(By.XPATH, get_xpath(driver, 'J0umjEiaBJtThVN')).click()

        # to type content in input field
        driver.find_element(By.XPATH, get_xpath(driver, 'SJW8GnK3RynhwLa')).send_keys(otrs_login)

        # to click on input field
        driver.find_element(By.XPATH, get_xpath(driver, 'Xg7h4qepocPSEYa')).click()

        # to type content in input field
        driver.find_element(By.XPATH, get_xpath(driver, 'WB0VHg1ua_feRSb')).send_keys(otrs_pass)

        # to click on the element(Login) found
        driver.find_element(By.XPATH, get_xpath(driver, 'wdz8eE4RTiLW0dT')).click()
        # Read the Excel file
        variables_df = pd.read_excel('Created tickets 1151.xlsx')

        # Base URL
        base_url = "https://servicecentre.fl-app.rhenus.com/otrs/index.pl?Action=AgentTicketZoom;TicketID={}"

        # Apply the URL formatting to each TicketID value
        variables_df['ModifiedURL'] = variables_df['TicketID'].apply(lambda ticket_id: base_url.format(ticket_id))

        # Specify the path to the output Excel file
        excel_file_path = 'Modified_tickets 1151.xlsx'

        # Save the DataFrame to an Excel file
        variables_df.to_excel(excel_file_path, index=False)

        mod_df = pd.read_excel('Modified_tickets 1151.xlsx')
        for index, row in mod_df.iterrows():
            ids = str(row['TicketID'])
            driver.get(
                "https://servicecentre.fl-app.rhenus.com/otrs/index.pl?ChallengeToken=MPIganck4JmxWn5bpF6Fwz2qDMQiNwtF&Action=AgentTicketCompose&TicketID=" + ids)  # time.sleep(5)
            time.sleep(5)
            iframe = driver.find_element(By.XPATH, "//iframe[@title='Rich Text Editor, RichText']")
            driver.switch_to.frame(iframe)

            # Find the contenteditable field within the iframe
            contenteditable_field = driver.find_element(By.XPATH, "//body[@contenteditable='true']")

            # Clear any existing content and send new text
            okta_pass_2_value = self.ui.okta_pass_2.text()

            text_to_input = f'''
        Good afternoon,

 

        Could you please fill in delivery date to the attached file?
        Thank you

        Best regards,
        {okta_pass_2_value}
        Service Centre Freight

        Rhenus Freight Network GmbH, Rhenus-Platz 1, 59439 Holzwickede, Deutschland
        Sitz: Rhenus-Platz 1, 59439 Holzwickede; AG Hamm, HRB 7788; Geschäftsführer: Petra Finke, Karin Peschel, Carolin Yilmaz,  St.-Nr. 316/5950/1308;UST-ID-Nr.: DE 288 342 692
        Soweit wir als Dienstleister beauftragt werden, arbeiten wir ausschließlich auf Grundlage der Allgemeinen Deutschen Spediteurbedingungen 2017 – (ADSp 2017, abrufbar unter: http://www.de.rhenus.com/adsp/) – und – soweit diese für die Erbringung logistischer Leistungen nicht gelten – nach den Logistik-AGB, Stand März 2006 (abrufbar unter: http://www.de.rhenus.com/adsp/). Hinweis: Die ADSp 2017 weichen in Ziffer 23 hinsichtlich des Haftungshöchstbetrages für Güterschäden (§ 431 HGB) vom Gesetz ab, indem sie die Haftung bei multimodalen Transporten unter Einschluss einer Seebeförderung und bei unbekanntem Schadenort auf 2 SZR/kg und im Übrigen die Regelhaftung von 8,33 SZR/kg zusätzlich auf 1,25 Millionen Euro je Schadenfall sowie 2,5 Millionen Euro je Schadenereignis, mindestens aber 2 SZR/kg, beschränken. Für die von uns eingesetzten Subunternehmer finden die ADSp, die Logistik-AGB, sowie sonstige AGB keine Anwendung. Erfüllungsort und Gerichtstand für beide Teile ist Holzwickede.

                    '''
            contenteditable_field.clear()
            contenteditable_field.send_keys(text_to_input)

            # Switch back to the main frame
            driver.switch_to.default_content()
            time.sleep(1)
            keyboard = Controller()
            driver.find_element(By.XPATH, get_xpath(driver, '4lf0OHPurAbzIZo')).click()
            time.sleep(1)
            # Find the h1 element using the specified text
            h1_element = driver.find_element(By.XPATH, '//div[@class="Header"]/h1')
            h1_text = h1_element.text
            # Now you can extract the variable value using string manipulation or regular expressions
            variable_value = h1_text.split(" — ")[1]
            keyboard.type(self.directory + "\\" + variable_value)
            time.sleep(1)
            keyboard.press(Key.enter)
            keyboard.release(Key.enter)
            time.sleep(1)
            driver.find_element(By.XPATH, '//*[@id="StateID_Search"]').send_keys("closed successful")
            time.sleep(1)
            driver.find_element(By.PARTIAL_LINK_TEXT, "closed successful").click()
            time.sleep(1)
            driver.find_element(By.XPATH, '//*[@id="DynamicField_sctask_Search"]').send_keys("LWIS")
            time.sleep(1)
            driver.find_element(By.XPATH, '//*[@id="DynamicField_scwork"]').send_keys("1")
            time.sleep(1)
            driver.find_element(By.XPATH, '//*[@id="TimeUnits"]').send_keys("1")
            time.sleep(1)
            driver.find_element(By.PARTIAL_LINK_TEXT, "LWIS check delivery").click()
            time.sleep(10)
        sys.exit()

    def dirmain(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ShowDirsOnly
        directorymain = QFileDialog.getExistingDirectory(self, 'Select Directory', options=options)
        if directorymain:
            directorymain = directorymain.replace('/', '\\')
            self.ui.email_file_4.setText(directorymain)
            self.directorymain = directorymain

    def dir1151(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ShowDirsOnly
        directory = QFileDialog.getExistingDirectory(self, 'Select Directory', options=options)
        if directory:
            directory = directory.replace('/', '\\')
            self.ui.email_file_3.setText(directory)
            self.directory = directory

    def prepare(self):
        input_excel_path = self.file_path
        output_excel_path = 'prepared_emails.xlsx'

        # Assuming your columns are named 'Column1' and 'Column2', replace them with your actual column names
        columns_to_extract = ['Relation', 'Email']

        # Load the Excel file into a DataFrame
        data = pd.read_excel(input_excel_path, usecols=columns_to_extract)

        # Remove duplicate rows
        data = data.drop_duplicates()

        data = data.query(
            'Relation != 1156 and Relation != 1018 and Relation != 1293 and Relation != 1294 and Relation != 1295 and '
            'Relation != 1296 and Relation != 1522 and Relation != 1151 and Relation != 1383')

        # Save the processed data to a new Excel file
        data.to_excel(output_excel_path, index=False)
        # # Read the original Excel file
        input_excel = self.selected_1151_file

        df = pd.read_excel(input_excel)

        # Iterate through each row and create a separate Excel file
        for i, row in df.iterrows():
            # Create a DataFrame with the current row as the data
            split_df = pd.DataFrame([row], columns=df.columns)

            # Generate the output file name based on the index of the row
            output_file = f'Partner_1151_{i + 1}.xlsx'

            # Save the DataFrame as an Excel file with headers
            split_df.to_excel(output_file, index=False)

        print(f"{len(df)} Excel files created by splitting each row.")

    def run_clicked(self):
        oklog = self.ui.okta_login.text()  # Get text from okta_login QLineEdit
        okpass = self.ui.okta_pass.text()  # Get text from okta_pass QLineEdit
        otrs_login = self.ui.otrs_login.text()  # Get text from otrs_login QLineEdit
        otrs_pass = self.ui.otrs_pass.text()  # Get text from otrs_pass QLineEdit
        driver = webdriver.Chrome()
        excel_file = self.file_path1
        # to open the url in browser
        options = webdriver.ChromeOptions()
        driver.get('https://rhenus.okta.com/')
        options.add_experimental_option('detach', True)
        # to type content in input field
        driver.find_element(By.XPATH, get_xpath(driver, '126YJaa68GerZcM')).send_keys(oklog)

        # to click on input field
        driver.find_element(By.XPATH, get_xpath(driver, 'f1GMPJOh2_4aGTs')).click()

        # to type content in input field
        driver.find_element(By.XPATH, get_xpath(
            driver, 'dVdU1X5TlD1mXYy')).send_keys(okpass)

        # to click on input field
        driver.find_element(By.XPATH, get_xpath(driver, 'GtyGHu4AyVcmHk0')).click()
        time.sleep(25)
        driver.get('https://servicecentre.fl-app.rhenus.com/otrs/index.pl')

        # to click on input field
        driver.find_element(By.XPATH, get_xpath(driver, 'J0umjEiaBJtThVN')).click()

        # to type content in input field
        driver.find_element(By.XPATH, get_xpath(driver, 'SJW8GnK3RynhwLa')).send_keys(otrs_login)

        # to click on input field
        driver.find_element(By.XPATH, get_xpath(driver, 'Xg7h4qepocPSEYa')).click()

        # to type content in input field
        driver.find_element(By.XPATH, get_xpath(driver, 'WB0VHg1ua_feRSb')).send_keys(otrs_pass)

        # to click on the element(Login) found
        driver.find_element(By.XPATH, get_xpath(driver, 'wdz8eE4RTiLW0dT')).click()
        # Read the Excel file
        variables_df = pd.read_excel('Created tickets.xlsx')

        # Base URL
        base_url = "https://servicecentre.fl-app.rhenus.com/otrs/index.pl?Action=AgentTicketZoom;TicketID={}"

        # Apply the URL formatting to each TicketID value
        variables_df['ModifiedURL'] = variables_df['TicketID'].apply(lambda ticket_id: base_url.format(ticket_id))

        # Specify the path to the output Excel file
        excel_file_path = 'Modified_tickets.xlsx'

        # Save the DataFrame to an Excel file
        variables_df.to_excel(excel_file_path, index=False)

        mod_df = pd.read_excel('Modified_tickets.xlsx')
        for index, row in mod_df.iterrows():
            ids = str(row['TicketID'])
            driver.get(
                "https://servicecentre.fl-app.rhenus.com/otrs/index.pl?ChallengeToken=MPIganck4JmxWn5bpF6Fwz2qDMQiNwtF&Action=AgentTicketCompose&TicketID=" + ids)  # time.sleep(5)
            time.sleep(5)
            iframe = driver.find_element(By.XPATH, "//iframe[@title='Rich Text Editor, RichText']")
            driver.switch_to.frame(iframe)

            # Find the contenteditable field within the iframe
            contenteditable_field = driver.find_element(By.XPATH, "//body[@contenteditable='true']")
            okta_pass_2_value = self.ui.okta_pass_2.text()
            # Clear any existing content and send new text
            text_to_input = f'''
Good afternoon,
Could you please fill in delivery date to the attached file?
Thank you
            
Best regards,
{okta_pass_2_value}
Service Centre Freight

Rhenus Freight Network GmbH, Rhenus-Platz 1, 59439 Holzwickede, Deutschland
Sitz: Rhenus-Platz 1, 59439 Holzwickede; AG Hamm, HRB 7788; Geschäftsführer: Petra Finke, Karin Peschel, Carolin Yilmaz,  St.-Nr. 316/5950/1308;UST-ID-Nr.: DE 288 342 692
Soweit wir als Dienstleister beauftragt werden, arbeiten wir ausschließlich auf Grundlage der Allgemeinen Deutschen Spediteurbedingungen 2017 – (ADSp 2017, abrufbar unter: http://www.de.rhenus.com/adsp/) – und – soweit diese für die Erbringung logistischer Leistungen nicht gelten – nach den Logistik-AGB, Stand März 2006 (abrufbar unter: http://www.de.rhenus.com/adsp/). Hinweis: Die ADSp 2017 weichen in Ziffer 23 hinsichtlich des Haftungshöchstbetrages für Güterschäden (§ 431 HGB) vom Gesetz ab, indem sie die Haftung bei multimodalen Transporten unter Einschluss einer Seebeförderung und bei unbekanntem Schadenort auf 2 SZR/kg und im Übrigen die Regelhaftung von 8,33 SZR/kg zusätzlich auf 1,25 Millionen Euro je Schadenfall sowie 2,5 Millionen Euro je Schadenereignis, mindestens aber 2 SZR/kg, beschränken. Für die von uns eingesetzten Subunternehmer finden die ADSp, die Logistik-AGB, sowie sonstige AGB keine Anwendung. Erfüllungsort und Gerichtstand für beide Teile ist Holzwickede.

            '''
            contenteditable_field.clear()
            contenteditable_field.send_keys(text_to_input)

            # Switch back to the main frame
            driver.switch_to.default_content()
            time.sleep(1)
            keyboard = Controller()
            driver.find_element(By.XPATH, get_xpath(driver, '4lf0OHPurAbzIZo')).click()
            time.sleep(1)
            h1_element = driver.find_element(By.XPATH, '//h1[contains(text(), "Partner ")]')
            partner_number_text = h1_element.text
            partner_number = partner_number_text.split("Partner ")[1]
            keyboard.type(self.directorymain + "\\" + partner_number + '.xlsx')
            time.sleep(1)
            keyboard.press(Key.enter)
            keyboard.release(Key.enter)
            time.sleep(1)
            driver.find_element(By.XPATH, '//*[@id="StateID_Search"]').send_keys("closed successful")
            time.sleep(1)
            driver.find_element(By.PARTIAL_LINK_TEXT, "closed successful").click()
            time.sleep(1)
            driver.find_element(By.XPATH, '//*[@id="DynamicField_sctask_Search"]').send_keys("LWIS")
            time.sleep(1)
            driver.find_element(By.XPATH, '//*[@id="DynamicField_scwork"]').send_keys("1")
            time.sleep(1)
            driver.find_element(By.XPATH, '//*[@id="TimeUnits"]').send_keys("1")
            time.sleep(1)
            driver.find_element(By.PARTIAL_LINK_TEXT, "LWIS check delivery").click()
            time.sleep(10)
        sys.exit()

    def select_week_file(self):
        options = QtWidgets.QFileDialog.Options()
        options |= QtWidgets.QFileDialog.ReadOnly
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Select Week File", "", "All Files (*)",
                                                             options=options)
        if file_path:
            self.selected_week_file = file_path
            self.ui.week_addr.setText(file_path)  # Update the label text
            self.file_path = file_path

    def select_1151(self):
        options = QtWidgets.QFileDialog.Options()
        options |= QtWidgets.QFileDialog.ReadOnly
        file_path3, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Select Week File", "", "All Files (*)",
                                                              options=options)
        if file_path3:
            self.selected_1151_file = file_path3
            self.ui.email_file_2.setText(file_path3)  # Update the label text
            self.file_path3 = file_path3

    def select_email_file(self):
        options = QtWidgets.QFileDialog.Options()
        options |= QtWidgets.QFileDialog.ReadOnly
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Select Email File", "", "All Files (*)",
                                                             options=options)
        if file_path:
            self.selected_email_file = file_path
            self.ui.email_file.setText(file_path)
            self.file_path1 = file_path


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    window = MyWindow()
    window.show()
    sys.exit(app.exec_())
