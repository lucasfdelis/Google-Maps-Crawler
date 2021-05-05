from selenium import webdriver
import xlsxwriter
import pandas as pd

workbook = xlsxwriter.Workbook('pesquisa_maps.xlsx')
sheet1 = workbook.add_worksheet()
style = workbook.add_format({'bold': True})
sheet1.write(0, 0, 'CNPJ',style)
sheet1.write(0, 1, 'RAZÃO SOCIAL',style)
sheet1.write(0, 2, 'CONTATO',style)
sheet1.write(0, 3, 'ENDEREÇO',style)

driver = webdriver.Chrome()
driver.get('https://www.google.com.br/maps/')
n = input("APERTE QUALQUER BOTÃO QUANDO O SITE TIVER SIDO ABERTO")
k = 1
nome_ant = numero_ant = endereco_ant = ''
while True:
        try:
                nome = driver.find_element_by_xpath('//*[@id="pane"]/div/div[1]/div/div/div[2]/div[1]/div[1]/div[1]/h1/span[1]')
                numero = driver.find_element_by_css_selector("[data-tooltip='Copiar número de telefone']")
                endereco = driver.find_element_by_css_selector("[data-item-id='address']")
                if(len(nome.text)!='' and len(numero.text)!='' and len(endereco.text)!=''):
                        if(nome_ant != nome.text and numero_ant != numero.text and endereco_ant != endereco.text):
                                print(nome.text)
                                print(numero.text)
                                print(endereco.text)
                                sheet1.write(k, 1, nome.text)
                                sheet1.write(k, 2, numero.text)
                                sheet1.write(k, 3, endereco.text)
                                nome_ant = nome.text
                                numero_ant = numero.text
                                endereco_ant = endereco.text
                                k = k+1
        except:
                pass
workbook.close()

