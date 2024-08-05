import re
import time
from playwright.sync_api import Playwright, sync_playwright, expect

def run(playwright: Playwright) -> None:#Inicie o navegador

    browser = playwright.chromium.launch(headless=False)#Chama o playwright e busca abrir com ele o Chromium com o cabeçaljo visivel

    context = browser.new_context(accept_downloads=True)#Cria um novo contexto dentro do chromium permitindo criar uma pagina em seguida, e passando um gerenciador de Dowload
    page = context.new_page()#Cria a pagina
    page.goto("http://153.le-pv.com/renac_web/login")#Acessa com o metodo GOTO o link do site RENAC
    page.get_by_placeholder("Please enter user name").click()#Seleciona e clica na pagina o input com o text especificado
    page.get_by_placeholder("Please enter user name").fill("vertysgroup")#Escreve na pagina o o valor especificado
    page.get_by_placeholder("Please enter password").click()#Seleciona e clica  na pagina o input com o text especificado
    page.get_by_placeholder("Please enter password").fill("vertysvip!@#")#Escreve o texto no campo selecionado
    page.get_by_role("button", name="Log In").click()#Clica em entrar

    time.sleep(5)#Espera 5 segundos para continuar

    page.locator("li").filter(has_text="Station Station ListDevice").locator("div").click()#Seleciona o elemento LI do HTML com o texto X dentro da div
    page.locator("#app li").filter(has_text=re.compile(r"^Device List$")).locator("span").click()#Filtra os elementos que correspondem ao ID X e text X

    page.locator("#app form div").filter(has_text="Device SN").get_by_role("textbox").click()#Acessa o vmapo especificado pelo elemnento
    page.locator("#app form div").filter(has_text="Device SN").get_by_role("textbox").fill("8701A31221034146")#Descreve o SN dentor do input selecionado
    page.get_by_role("button", name="Search").click()#Clica em buscar
    page.wait_for_load_state("networkidle")#Espera a pagibna carregar

    time.sleep(7)#Espera 7 segunbdos pára continuar


    page.get_by_role("cell", name="8701A31221034146").click()#Seleciona a cell que corresponda ao SN

    # Aguardando o estado da página
    page.wait_for_load_state("networkidle")

    #For para pular as paginas que aparecem para conseguir chegfar no dia que comça a geração
    for i in range(14):

        print(i)
        page.get_by_role("button", name="").nth(1).click()

    #For para apartir do dia  com geração ir coletnado os dados e exportando apr ao excel
    for i in range(25):

        time.sleep(3)#Espera 3 segundos

        #Seleciona as caixas de busca de cada valor desejado no excel
        page.locator(".once-item").click()
        page.get_by_label("Voltage R").check()
        page.get_by_text("Frequency R").click()
        page.get_by_label("PV2 Voltage").check()
        page.get_by_label("PV1 Power").check()
        page.get_by_label("PV2 Power").check()
        page.get_by_label("PV2 Current").check()
        page.get_by_label("PV2 Current").check()
        page.get_by_label("Yield", exact=True).check()
        page.get_by_label("PV power", exact=True).check()
        page.get_by_label("Current R").check()
        page.get_by_text("PV power", exact=True).click()

        time.sleep(2)#Espera 2 segundos

        #Um tentativa de realizar o dowloand encontradno o buttun de Export
        try:
            with page.expect_download() as download1_info:#Com o metodo de Dowload acessa e busca o button

                page.get_by_role("button", name="Export").nth(1).click()


            download1 = download1_info.value#Coleta o vlaor dos dados
            download1.save_as(f"C://Users//Vertys//Desktop//{i}.xls")  # Salva os arquivos e nomea eles

        except Exception as e:
            print(f'Erro {e}')

        print(i)#Printa o vlaor de I
        page.get_by_role("button", name="").nth(1).click()#Pula para o proximo dia

    # Apos o processo fecha a pagina
    context.close()
    browser.close()

#Chama o def que iniciou com o playwright com sincronização
with sync_playwright() as playwright:
    run(playwright)
