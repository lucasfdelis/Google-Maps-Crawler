from selenium import webdriver
import xlwings as xw

wb = xw.Book('teste.xlsx')
sht1 = wb.sheets['Plan1']
sht1.range('A1').value = 'NOME'
sht1.range('B1').value = 'ENDEREÇO'
sht1.range('C1').value = 'NUMERO'

driver = webdriver.Chrome()
driver.get('https://www.google.com.br/maps/')
n = input("APERTE QUALQUER BOTÃO QUANDO O SITE TIVER SIDO ABERTO")
k = 2
nome_ant = numero_ant = endereco_ant = ''
while True:
      try:
              nome = driver.find_element_by_xpath('//*[@id="pane"]/div/div[1]/div/div/div[2]/div[1]/div[1]/div[1]/h1/span[1]')
              numero = driver.find_element_by_css_selector("[data-tooltip='Copiar número de telefone']")
              endereco = driver.find_element_by_css_selector("[data-item-id='address']")
              if(len(nome.text)>0 and len(numero.text)>0 and len(endereco.text)>0):
                      if(nome_ant != nome.text and numero_ant != numero.text and endereco_ant != endereco.text):
                              print(nome.text)
                              print(numero.text)
                              print(endereco.text)
                              print('')
                              sht1.range('A'+str(k)).value = nome.text
                              sht1.range('B'+str(k)).value = numero.text
                              sht1.range('C'+str(k)).value = endereco.text
                              nome_ant = nome.text
                              numero_ant = numero.text
                              endereco_ant = endereco.text
                              k = k+1
      except Exception as e:
              pass

