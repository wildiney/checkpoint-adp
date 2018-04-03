from selenium import webdriver
from bs4 import BeautifulSoup
from datetime import datetime,timedelta
import time
import pprint
import pandas as pd
import xlsxwriter

class Folha_de_Ponto:
    def __init__(self, login, password, debug=True):
        self.login = login
        self.password = password
        self.debug = debug
        self.adp_login()
    
    def dprint(self, alert):
        if self.debug is True:
            print(alert)

    def adp_login(self):
        self.dprint("Accessing")
        phantomjs_path = r"./resources/phantomjs.exe"
        browser = webdriver.PhantomJS(phantomjs_path)
        browser.get('https://www.adpweb.com.br/expert/')
        browser.switch_to.frame('ADP')
        time.sleep(5)

        self.dprint("Entering Credentials")
        user = browser.find_element_by_xpath("""//*[@id="login"]""")
        user.send_keys(self.login)

        password = browser.find_element_by_xpath("""//*[@id="login-pw"]""")
        password.send_keys(self.password)

        login = browser.find_element_by_xpath("""//*[@id="revit_form_Button_0"]""")
        login.click()
        time.sleep(5)

        self.dprint("Accessing Meu Painel")
        meuPainel = browser.find_element_by_xpath("""//*[@id="Meu Painel_navItem"]/span""")
        meuPainel.click()
        time.sleep(5)

        self.dprint("Accessing Ponto")
        ponto = browser.find_element_by_xpath("""//*[@id="painel_icones"]/span[6]/a""")
        ponto.click()
        time.sleep(5)

        browser._switch_to.frame('BMPROG')

        self.dprint("Accessing Folha")
        folha = browser.find_element_by_xpath("""/html/body/a/form/table[5]""")
        source_code = folha.get_attribute("outerHTML")

        self.parse(source_code)


    def parse(self,html):
        l = []
        soup = BeautifulSoup(html, 'html.parser')
        all = soup.select('table tbody tr')
        for item in all:
            try:
                data = item.find_all('td', class_='listaval2')[0].text.replace('\n', '')
            except:
                data = None
            try:
                entrada = item.find_all('td', class_='listaval2')[6].find_all('font')[0].text
            except:
                entrada = None
            try:
                almoco = item.find_all('td', class_='listaval2')[6].find_all('font')[4].text
            except:
                almoco = None
            try:
                retorno = item.find_all('td', class_='listaval2')[6].find_all('font')[8].text
            except:
                retorno = None
            try:
                saida = item.find_all('td', class_='listaval2')[6].find_all('font')[12].text
            except:
                saida = None

            if data is not None:
                d = {}
                d["lbl01-data"] = data
            if entrada is not None:
                if entrada != ".":
                    d["lbl02-entrada"] = datetime.strptime(data+entrada[-5:], '%d/%m/%Y%H:%M')
                else:
                    d["lbl02-entrada"] = datetime.strptime(data+str("00:00"), '%d/%m/%Y%H:%M')

            if almoco is not None:
                if almoco != ".":
                    d["lbl03-almoco"] = datetime.strptime(data+almoco[-5:], '%d/%m/%Y%H:%M')
                else:
                    d["lbl03-almoco"] = datetime.strptime(data+str("00:00"), '%d/%m/%Y%H:%M')
            if retorno is not None:
                if retorno != ".":
                    d["lbl04-retorno"] = datetime.strptime(data+retorno[-5:], '%d/%m/%Y%H:%M')
                else:
                    d["lbl04-retorno"] = datetime.strptime(data+str("00:00"), '%d/%m/%Y%H:%M')
            if saida is not None:
                if saida != ".":
                    d["lbl05-saida"] = datetime.strptime(data+saida[-5:], '%d/%m/%Y%H:%M')
                else:
                    d["lbl05-saida"] = datetime.strptime(data+str("00:00"), '%d/%m/%Y%H:%M')
            if data is not None:
                l.append(d)

        self.dprint("Writing file")
        df = pd.DataFrame(l)
        writer = pd.ExcelWriter("folha-de-ponto.xlsx", engine="xlsxwriter")
        df.to_excel(writer, sheet_name="Ponto")
        workbook = writer.book
        worksheet = writer.sheets['Ponto']
        format1 = workbook.add_format({'num_format': 'hh:mm'})
        worksheet.set_column('C:F', None, format1)
        writer.save()
        self.dprint("Finished")