from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import time 

estados = ['AM', 'BA', 'CE', 'DF', 'ES', 'GO', 'MA', 'MT', 'PB', 'PE', 'RN', 'RJ', 'SC']
valor_mes = int(input("Digite o mês:"))


driver = webdriver.Chrome()
url = f'http://www.cub.org.br/cub-m2-estadual/MG/'
driver.get(url)
try:
    select_element = driver.find_element(By.ID, "mes")  
    select = Select(select_element)
    
    value_to_select = f"{valor_mes - 1}" 
    time.sleep(1)
    select.select_by_value(value_to_select)
    
    select_element = driver.find_element(By.ID, "sinduscon")  
    select = Select(select_element)
    select.select_by_value("1")
    
    button = driver.find_element(By.XPATH, f"//input[@value='Gerar Relatório em PDF']")
    button.click()
    
    time.sleep(1)


finally:
        driver.quit()
        

driver = webdriver.Chrome()
url = f'http://www.cub.org.br/cub-m2-estadual/PR/'
driver.get(url)
try:
    select_element = driver.find_element(By.ID, "mes")  
    select = Select(select_element)
    value_to_select = f"{valor_mes}" 
    select.select_by_value(value_to_select)
    
    select_element = driver.find_element(By.ID, "sinduscon")  
    select = Select(select_element)
    select.select_by_value("18")
    


    button = driver.find_element(By.XPATH, f"//input[@value='Gerar Relatório em PDF']")
    button.click()
    
    time.sleep(1)

finally:
        driver.quit()


driver = webdriver.Chrome()
url = f'http://www.cub.org.br/cub-m2-estadual/PI/'
driver.get(url)
try:
    select_element = driver.find_element(By.ID, "mes")  
    select = Select(select_element)
    
    value_to_select = f"{valor_mes - 2}" 
    select.select_by_value(value_to_select)

    button = driver.find_element(By.XPATH, f"//input[@value='Gerar Relatório em PDF']")
    button.click()
    
    time.sleep(1)

finally:
        driver.quit()


for estado in estados:

    driver = webdriver.Chrome()

    url = f'http://www.cub.org.br/cub-m2-estadual/{estado}/'
    driver.get(url)

    try:
        select_element = driver.find_element(By.ID, "mes")  
        select = Select(select_element)

        value_to_select = f"{valor_mes}" 
        select.select_by_value(value_to_select)

        button = driver.find_element(By.XPATH, f"//input[@value='Gerar Relatório em PDF']")
        button.click()
        
        time.sleep(1)

    
    finally:
            driver.quit()