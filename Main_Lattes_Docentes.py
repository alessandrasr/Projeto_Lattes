from selenium import webdriver  # type: ignore
from selenium.webdriver.chrome.service import Service  # type: ignore
from webdriver_manager.chrome import ChromeDriverManager  # type: ignore
import pandas as pd  # type: ignore
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import re

options = webdriver.ChromeOptions()
options.add_argument("--headless")

# Inicializar o driver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

df = pd.read_excel("funcionarios.xlsx")
df['CPF'] = df['CPF'].astype(str).str.zfill(11)

url_base = "https://buscatextual.cnpq.br/buscatextual/busca.do?metodo=forwardPaginaResultados&registros=0%3B10&query=idx_cpf%3A"

resultados = []

# Extração
for index, cpf in enumerate(df['CPF']):
    numero_linha = index + 2
    url_completa = f"{url_base}{cpf}"
    print(f"Consultando CPF (Linha {numero_linha}): {cpf}")

    tentativas = 0
    max_tentativas = 3
    sucesso = False

    while tentativas < max_tentativas and not sucesso:
        try:
            driver.get(url_completa)
            driver.implicitly_wait(10)

            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//ol/li/b/a")))
            elemento_nome = driver.find_element(By.XPATH, "//ol/li/b/a")
            nome = elemento_nome.text.strip()
            elemento_nome.click()

            time.sleep(2)

            # Clicar no botão para abrir currículo
            driver.find_element(By.ID, "idbtnabrircurriculo").click()

            # Esperar a nova aba abrir
            WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > 1)
            abas = driver.window_handles

            if len(abas) > 1:
                driver.switch_to.window(abas[1])  # Muda para aba do currículo

                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "ul.informacoes-autor li:last-child"))
                )
                elemento_data = driver.find_element(By.CSS_SELECTOR, "ul.informacoes-autor li:last-child")
                informacoes = elemento_data.text.strip()

                match = re.search(r"\d{2}/\d{2}/\d{4}", informacoes)
                data_atualizacao = match.group() if match else ""

                elemento_titulacao = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//a[@name='FormacaoAcademicaTitulacao']/following::div[contains(@class,'data-cell')][1]"))
                )
                titulacao = elemento_titulacao.text.strip()

                # Adiciona aos resultados
                resultados.append({
                    "CPF": cpf,
                    "Nome": nome,
                    "Data": data_atualizacao,
                    "Formação/Titulação": titulacao,
                    "URL": url_completa
                })

                # Fechar a aba do currículo e voltar para a anterior
                driver.close()
                driver.switch_to.window(abas[0])
                sucesso = True
            else:
                print(f"[Linha {numero_linha}] A aba do currículo não foi aberta.")

        except Exception as e:
            tentativas += 1
            print(f"[Linha {numero_linha}] Tentativa {tentativas} falhou: {e}")
            if len(driver.window_handles) > 1:
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
            time.sleep(1)

    if not sucesso:
        print(f"[Linha {numero_linha}] Não foi possível obter os dados após {max_tentativas} tentativas.")

# Exportar para Excel
df_resultados = pd.DataFrame(resultados)
df_resultados.to_excel("resultados_lattes.xlsx", index=False)
os.startfile("resultados_lattes.xlsx")

print("Resultados salvos com sucesso no arquivo 'resultados_lattes.xlsx'.")
driver.quit()
