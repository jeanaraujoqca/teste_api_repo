from flask import Flask, request, jsonify
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import os

app = Flask(__name__)

# Função principal para executar a automação Selenium
def submit_form(df: pd.DataFrame, email: str, senha: str):
    # Configuração do Chrome
    chrome_options = Options()
    chrome_options.add_argument("--headless")  # Executa o Chrome em modo headless
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")

    # Inicialização do driver do Chrome
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

    try:
        url_sharepoint = 'https://queirozcavalcanti.sharepoint.com/sites/qca360/Lists/treinamentos_qca/AllItems.aspx'
        driver.get(url_sharepoint)
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "i0116")))

        # Login
        driver.find_element(By.ID, 'i0116').send_keys(email)
        driver.find_element(By.ID, 'idSIButton9').click()
        WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.ID, 'i0118')))
        driver.find_element(By.ID, 'i0118').send_keys(senha)
        driver.find_element(By.ID, 'idSIButton9').click()
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, 'idSIButton9'))).click()

        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//button[text()="Novo"]'))).click()
        WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR, "iframe")))

        # Realizar a automação de preenchimento de formulário
        casos_sucesso = []
        casos_fracasso = []

        for index, id in enumerate(df['ID']):
            try:
                colaborador = df.loc[index, 'Nome']
                email_colaborador = df.loc[index, 'Email']
                unidade = df.loc[index, 'UNIDADE']
                treinamento = df.loc[index, 'TREINAMENTO']
                tipo_de_treinamento = df.loc[index, 'TIPO DO TREINAMENTO']
                categoria = df.loc[index, 'CATEGORIA']
                instituicao_instrutor = df.loc[index, 'INSTITUIÇÃO/INSTRUTOR']
                carga_horaria = df.loc[index, 'CARGA HORÁRIA']
                inicio_do_treinamento = df.loc[index, 'INICIO DO TREINAMENTO']
                termino_do_treinamento = df.loc[index, 'TERMINO DO TREINAMENTO']

                # Funções para preencher os campos e selecionar as opções
                def clica_seleciona_informacao(selector, valor, selecionar_xpath):
                    driver.find_element(By.CSS_SELECTOR, selector).click()
                    driver.find_element(By.XPATH, selecionar_xpath).send_keys(valor)
                    driver.find_element(By.XPATH, f'//li[text()="{valor}"]').click()

                clica_seleciona_informacao('div[title="NOME DO INTEGRANTE"]', colaborador,
                                             '//*[@id="powerapps-flyout-react-combobox-view-0"]/div/div/div/div/input')
                clica_seleciona_informacao('div[title="E-MAIL"]', email_colaborador,
                                             '//*[@id="powerapps-flyout-react-combobox-view-1"]/div/div/div/div/input')
                clica_seleciona_informacao('div[title="UNIDADE"]', unidade,
                                             '//*[@id="powerapps-flyout-react-combobox-view-2"]/div/div/div/div/input')
                driver.find_element(By.CSS_SELECTOR, 'input[title="TREINAMENTO"]').send_keys(treinamento)
                clica_seleciona_informacao('div[title="TIPO DO TREINAMENTO."]', tipo_de_treinamento,
                                             '//*[@id="powerapps-flyout-react-combobox-view-3"]/div/div/div/div/input')
                driver.find_element(By.CSS_SELECTOR, 'input[title="INSTITUIÇÃO/INSTRUTOR"]').send_keys(instituicao_instrutor)
                clica_seleciona_informacao('div[title="CATEGORIA"]', categoria,
                                             '//*[@id="powerapps-flyout-react-combobox-view-4"]/div/div/div/div/input')
                driver.find_element(By.CSS_SELECTOR, 'input[title="INICIO DO TREINAMENTO"]').send_keys(inicio_do_treinamento)
                driver.find_element(By.CSS_SELECTOR, 'input[title="TERMINO DO TREINAMENTO"]').send_keys(termino_do_treinamento)

                driver.find_element(By.XPATH, '//*[@id="appRoot"]/div[3]/div/div[4]/div[2]/div/div[2]/div[3]/div/div/div/div[1]/div/div/div/div/div/div/div/div/div/div[1]/button/span').click()

                casos_sucesso.append({'Caso': id, 'Status': 'Sucesso'})
                WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, '//*[contains(text(), "Caso criado com sucesso")]')))

            except Exception as e:
                casos_fracasso.append({'Treinamento': id, 'Status': f'Erro inesperado: {str(e)}'})

        # Salvar os resultados em arquivos
        df_sucesso = pd.DataFrame(casos_sucesso)
        df_fracasso = pd.DataFrame(casos_fracasso)
        df_sucesso.to_excel('casos_sucesso.xlsx', index=False)
        df_fracasso.to_excel('casos_fracasso.xlsx', index=False)

        # enviar_relatorio() # Caso queira usar a função de enviar e-mail

        return {'status': 'sucesso'}

    finally:
        driver.quit()

# Endpoint Flask para rodar a automação
@app.route("/automacao-horas/", methods=['POST'])
def run_automation():
    if 'file' not in request.files:
        return jsonify({'status': 'erro', 'mensagem': 'Nenhum arquivo enviado'}), 400

    file = request.files['file']
    email = request.form.get('email')
    senha = request.form.get('senha')

    if not file or not email or not senha:
        return jsonify({'status': 'erro', 'mensagem': 'Faltam parâmetros obrigatórios'}), 400

    try:
        df = pd.read_excel(file)
        result = submit_form(df, email, senha)
        return jsonify(result)
    
    except Exception as e:
        return jsonify({'status': 'erro', 'mensagem': f'Internal Server Error: {str(e)}'}), 500

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5000, debug=False)
