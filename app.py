from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.alert import Alert
import openpyxl
from time import sleep
import pyperclip
#carregar planilha
workbook =  openpyxl.load_workbook('produtos_ficticios.xlsx')
pag_prod = workbook['Produtos']
driver= webdriver.Chrome()



#acessar o site 
driver.get('https://cadastro-produtos-devaprender.netlify.app/')

#maximiza a janela
driver.maximize_window()



#Laço de repetiçaõ para funcionar o programa
for linha in pag_prod.iter_rows(min_row=2):

    # Rolar a página para baixo um pouco (ex: 500 pixels)
    driver.execute_script("window.scrollBy(0, 500);")
    #Nome do Produto
    nome_prod=linha[0].value
    
    #Oque vai ser copiado
    pyperclip.copy(nome_prod)
    
   
    driver.find_element(By.NAME, 'product_name').send_keys(nome_prod)
    
    

#------------------------------------------------------------
#PAGINA 1

    #Descrição do Produto
    descricao=linha[1].value
    pyperclip.copy(descricao)
    
    driver.find_element(By.NAME, 'description').send_keys(descricao)
    
    sleep(1)



    #Categoria
    categoria=linha[2].value
    pyperclip.copy(categoria)
    driver.find_element(By.NAME, 'category').send_keys(categoria)
    sleep(1)

    #Codigo do Produto
    codigo_prod=linha[3].value
    pyperclip.copy(codigo_prod)
    driver.find_element(By.NAME, 'product_code').send_keys(codigo_prod)
    sleep(1)


    #Peso Em Kilos
    peso_kg=linha[4].value
    pyperclip.copy(peso_kg)
    driver.find_element(By.NAME, 'weight').send_keys(peso_kg)
    sleep(1)


    #Dimensões do Produto
    dimensoes_lap=linha[5].value
    driver.find_element(By.NAME, 'dimensions').send_keys(dimensoes_lap)
    pyperclip.copy(dimensoes_lap)
    sleep(1)


    
    driver.find_element(By.CLASS_NAME, "btn-primary").click()

    sleep(3)


    
#--------------------------------------------------------
#PAG 2
 
    #Preço
    preco=linha[6].value
    pyperclip.copy(preco)

    driver.find_element(By.NAME, 'price').send_keys(preco)
    
    sleep(1)


    #Quantidade em Estoque
    qnt_estoque=linha[7].value
    pyperclip.copy(qnt_estoque)
    driver.find_element(By.NAME, 'stock').send_keys(qnt_estoque)
    sleep(1)

    #Data de Validade
    data_validade=linha[8].value
    pyperclip.copy(data_validade)
    driver.find_element(By.NAME, 'expiry_date').send_keys(data_validade)
    sleep(1)

    #Cor
    cor=linha[9].value
    pyperclip.copy(cor)
    driver.find_element(By.NAME, 'color').send_keys(cor)
    sleep(1)

    #Tamanho
    tamanho=linha[10].value
    driver.find_element(By.NAME, "size").send_keys(tamanho)
    sleep(1)
        
    
 
    #Material
    material=linha[11].value
    driver.find_element(By.NAME, "material").send_keys(material)
    sleep(1)


    driver.find_element(By.CLASS_NAME, "btn-primary").click()
    


    
 #--------------------------------------------------------------------------

 #PAG 3

    #Fabricante
    fabricante=linha[12].value
    driver.find_element(By.NAME, "manufacturer").send_keys(fabricante)
    sleep(1)


    #Pais de Origem
    pais_origem=linha[13].value
    driver.find_element(By.NAME, "country").send_keys(pais_origem)
    sleep(1)


    #Observações
    obs=linha[14].value
    driver.find_element(By.NAME, "remarks").send_keys(obs)
    sleep(1)


    #Codigo de Barras do Produto
    codigo_barras=linha[15].value
    driver.find_element(By.NAME, "barcode").send_keys(codigo_barras)
    sleep(1)

    
    #Local no Armazem
    loc_armazem=linha[16].value
    driver.find_element(By.NAME, "warehouse_location").send_keys(loc_armazem)
    sleep(1)
    
    driver.find_element(By.CLASS_NAME, "btn-primary").click()

    alert= Alert(driver)
    alert.accept()

    sleep(2)
     
    driver.find_element(By.XPATH, "//button[@onclick=\"window.location.href='index.html'\"]").click()


    
    

  
