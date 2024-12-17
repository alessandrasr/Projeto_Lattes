from selenium import webdriver # type: ignore
from selenium.webdriver.chrome.service import Service # type: ignore
from webdriver_manager.chrome import ChromeDriverManager # type: ignore
import pandas as pd # type: ignore
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os

# Opções do Chrome
options = webdriver.ChromeOptions()
options.add_argument("--headless")  # Executar em modo invisível

# Inicializar o driver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# Planilha com os CPFs
arquivo_excel = "funcionarios.xlsx"
df = pd.read_excel(arquivo_excel)

# Garantir que o CPF esteja no formato correto
df['CPF'] = df['CPF'].astype(str).str.zfill(11)

# URL base
url_base = "https://buscatextual.cnpq.br/buscatextual/busca.do?metodo=forwardPaginaResultados&registros=0%3B10&query=idx_cpf%3A"

# Lista para armazenar resultados
resultados = []


## Extração de Dados
for cpf in df['CPF']:
    url_completa = f"{url_base}{cpf}"
    print(f"Consultando CPF: {cpf}")
    
    driver.get(url_completa)
    driver.implicitly_wait(time_to_wait=10) # Aguardar carregamento da página
    
    driver.implicitly_wait(time_to_wait=10) 
    try:
        # Tenta encontrar o elemento do nome
        elemento_nome = driver.find_element(By.XPATH, "/html/body/form/div/div[4]/div/div/div/div[3]/div/div[3]/ol/li/b/a")
        
        # Extrair o texto do elemento e remover espaços extras
        nome = elemento_nome.text.strip()
        elemento_nome.click()
    except:
        nome = ""
        continue

    time.sleep(2) 

    # Currículo
    curriculo = driver.find_element(By.XPATH, "//*[@id='idbtnabrircurriculo']")
    curriculo.click() 
    


    # Trocando de página
    WebDriverWait(driver, 10).until(
    lambda d: len(d.window_handles) > 1  # Aguarda até que haja mais de uma janela aberta
    )
    

    driver.switch_to.window(window_name=driver.window_handles[1]) # Alternar para a nova janela
   
    #Informações
    elemento_informacoes = driver.find_element(By.XPATH, "/html/body/div[1]/div[3]/div/div/div/div[1]/ul")
    informacoes = elemento_informacoes.text

    #Identificação
    elemento_identificacao =driver.find_element(By.XPATH, "/html/body/div[1]/div[3]/div/div/div/div[3]/div")
    identificacao = elemento_identificacao.text

    #Resumo
    elementos_resumo = WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.XPATH, "/html/body/div[1]/div[3]/div/div/div/div[2]/p"))
    )
    resumos = [resumo.text.strip() for resumo in elementos_resumo]


   
    driver.implicitly_wait(time_to_wait=10)


    # Adicionar resultados
    resultados.append({
        "CPF": cpf,
        "Nome": nome,
        "URL": url_completa,
        "Informações": informacoes,
        "Identificação": identificacao,
        "Resumo": "\n".join(resumos)
        })

    # Fechar a janela atual e voltar para a original
    driver.close()
    driver.switch_to.window(driver.window_handles[0])

# Criar o DataFrame e salvar no Excel
df_resultados = pd.DataFrame(resultados)
df_resultados.to_excel("resultados_lattes.xlsx", index=False)
os.startfile("resultados_lattes.xlsx")  # Abre o Excel

print("Resultados salvos com sucesso no arquivo 'resultados.xlsx'.")

# Fechar o navegador
driver.quit()