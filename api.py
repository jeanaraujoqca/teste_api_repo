from flask import Flask, request, jsonify
import pandas as pd
import asyncio
from playwright.async_api import async_playwright
# import win32com.client as win32

app = Flask(__name__)

# Função para enviar o relatório por email após a execução
# def enviar_relatorio():
#     try:
#         outlook = win32.Dispatch('outlook.application')
#         namespace = outlook.GetNamespace('MAPI')

#         def achar_pasta_por_nome(nome_pasta, parent_folder=None):
#             if parent_folder is None:
#                 parent_folder = namespace.Folders
                
#             for folder in parent_folder:
#                 if folder.Name == nome_pasta:
#                     return folder
#                 sub_folder = achar_pasta_por_nome(nome_pasta, folder.Folders)
#                 if sub_folder:
#                     return sub_folder
#             return None 
        
#         sent_items_folder = achar_pasta_por_nome("Itens Enviados")
#         nome_remetente = 'Desconhecido' if not sent_items_folder else sent_items_folder.Items.GetLast().SenderName
        
#         mail = outlook.CreateItem(0)
#         mail.Subject = 'Relatório de Uso da Automação de Lançamento de Horas'
#         mail.Body = f'{nome_remetente} utilizou a automação de lançamento de horas.'
#         mail.To = 'daniellerodrigues@queirozcavalcanti.adv.br'
#         mail.Send()
#     except Exception as e:
#         print(f"Erro ao enviar o relatório: {str(e)}")

# Função principal para executar a automação Playwright
async def submit_form_async(df: pd.DataFrame, email: str, senha: str):
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        page = await browser.new_page()

        url_sharepoint = 'https://queirozcavalcanti.sharepoint.com/sites/qca360/Lists/treinamentos_qca/AllItems.aspx'
        await page.goto(url_sharepoint)
        await page.wait_for_timeout(5000)

        try:
            await page.fill('#i0116', email)
            await page.click('#idSIButton9')
            await page.wait_for_timeout(2000)
            await page.fill('#i0118', senha)
            await page.click('#idSIButton9')
            await page.wait_for_timeout(2000)
            await page.click('#idSIButton9')  # Botão "Sim"
        except Exception as e:
            return {'status': 'erro', 'mensagem': f'Erro ao fazer login: {str(e)}'}
        
        await page.wait_for_timeout(10000)

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

                await page.click('button:has-text("Novo")')
                await page.wait_for_timeout(5000)
                
                iframe = page.frame_locator("iframe").nth(0)
                iframe2 = iframe.frame_locator("iframe.player-app-frame")

                async def clica_seleciona_informacao(iframe, endereco1, endereco2, valor2, endereco3):
                    await iframe.locator(endereco1).click()  
                    await iframe.locator(endereco2).fill(valor2)
                    await iframe.locator(endereco3).nth(0).click()

                await clica_seleciona_informacao(iframe2, 'div[title="NOME DO INTEGRANTE"]', 
                                                 '//*[@id="powerapps-flyout-react-combobox-view-0"]/div/div/div/div/input', colaborador, 
                                                 f'li:has-text("{colaborador}")')
                
                await clica_seleciona_informacao(iframe2, 'div[title="E-MAIL"]', 
                                                 '//*[@id="powerapps-flyout-react-combobox-view-1"]/div/div/div/div/input', email_colaborador, 
                                                 f'li:has-text("{email_colaborador}")')
                
                await clica_seleciona_informacao(iframe2, 'div[title="UNIDADE"]', 
                                                 '//*[@id="powerapps-flyout-react-combobox-view-2"]/div/div/div/div/input', unidade, 
                                                 f'li:has-text("{unidade}")')
                
                await iframe2.locator('input[title="TREINAMENTO"]').fill(treinamento)

                await clica_seleciona_informacao(iframe2, 'div[title="TIPO DO TREINAMENTO."]', 
                                                 '//*[@id="powerapps-flyout-react-combobox-view-3"]/div/div/div/div/input', tipo_de_treinamento, 
                                                 f'li:has-text("{tipo_de_treinamento}")')

                await iframe2.locator('input[title="INSTITUIÇÃO/INSTRUTOR"]').fill(instituicao_instrutor)
                await clica_seleciona_informacao(iframe2, 'div[title="CATEGORIA"]', 
                                                 '//*[@id="powerapps-flyout-react-combobox-view-4"]/div/div/div/div/input', categoria, 
                                                 f'li:has-text("{categoria}")')
                
                await iframe2.locator('input[title="INICIO DO TREINAMENTO"]').fill(inicio_do_treinamento)
                await iframe2.locator('input[title="TERMINO DO TREINAMENTO"]').fill(termino_do_treinamento)

                await page.locator('//*[@id="appRoot"]/div[3]/div/div[4]/div[2]/div/div[2]/div[3]/div/div/div/div[1]/div/div/div/div/div/div/div/div/div/div[1]/button/span').click()
                
                casos_sucesso.append({'Caso': id, 'Status': 'Sucesso'})
                await page.wait_for_timeout(3000)

            except Exception as e:
                casos_fracasso.append({'Treinamento': id, 'Status': f'Erro inesperado: {str(e)}'})

        df_sucesso = pd.DataFrame(casos_sucesso)
        df_fracasso = pd.DataFrame(casos_fracasso)
        df_sucesso.to_excel('casos_sucesso.xlsx', index=False)
        df_fracasso.to_excel('casos_fracasso.xlsx', index=False)

        # enviar_relatorio()

        await browser.close()
        return {'status': 'sucesso'}

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
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        result = loop.run_until_complete(submit_form_async(df, email, senha))
        return jsonify(result)
    
    except Exception as e:
        return jsonify({'status': 'erro', 'mensagem': f'Internal Server Error: {str(e)}'}), 500

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5000, debug=False)
