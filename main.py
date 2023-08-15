from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import time
import os
import shutil
import json


accounts = [
    {
        "user": "",
        "pass": "",
        "dir": "",
        "dt_first_nf_e_3": ""
    },
]

for acc in accounts:
# acc = accounts[1]

    folder_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), acc["dir"])
    if os.path.exists(folder_path):
        shutil.rmtree(folder_path)
    os.mkdir(folder_path)

    # Cria uma instância do navegador Google Chrome
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')

    prefs = {
        "download.default_directory": folder_path,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True
    }

    chrome_options.add_experimental_option("prefs", prefs)

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)

    def login(driver, username, password):
        url = "https://erp.tiny.com.br/login/"
        
        driver.get(url)

        driver.implicitly_wait(10)

        input_username = driver.find_element(By.NAME, 'username')
        input_username.send_keys(username)
        
        input_password = driver.find_element(By.NAME, 'senha')
        input_password.send_keys(password)

        input_password.send_keys(Keys.ENTER)

        try:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "bs-modal-ui-popup")))
            button = driver.find_element(By.XPATH, "//button[text()='Continuar']")
            time.sleep(2)
            button.click()
            time.sleep(3)
            write_to_logs(f"Logged as: {username}")
        except:
            write_to_logs(f"Logged as: {username}")

    def months_between(date1, date2):
        date_format = "%Y-%m-%d"
        d1 = datetime.strptime(date1, date_format)
        d2 = datetime.strptime(date2, date_format)
        return abs((d2.year - d1.year) * 12 + d2.month - d1.month)

    def write_to_logs(text):
        with open(folder_path + "/logs.txt", "a") as log_file:
            now = datetime.now()
            formatted_time = now.strftime('%d/%m/%Y %H:%M:%S')
            log_file.write(formatted_time + ": " + text + "\n")

    def download_contas_receber_e_pagar(driver):
        write_to_logs('download_contas_receber_e_pagar')
        links = [
            # "https://erp.tiny.com.br/exportacao_contas_receber?criterio=periodo&tipoPeriodo=todasEmAberto",
            "https://erp.tiny.com.br/exportacao_contas_receber?criterio=periodo&tipoPeriodo=emissao",
            # "https://erp.tiny.com.br/exportacao_contas_receber?criterio=periodo&tipoPeriodo=pagamento",
            # "https://erp.tiny.com.br/exportacao_contas_receber?criterio=periodo&tipoPeriodo=atrasadas",
            # "https://erp.tiny.com.br/exportacao_contas_receber?criterio=periodo&tipoPeriodo=canceladas",
            
            # "https://erp.tiny.com.br/exportacao_contas_pagar?criterio=periodo&tipoPeriodo=todasEmAberto",
            "https://erp.tiny.com.br/exportacao_contas_pagar?criterio=periodo&tipoPeriodo=emissao",
            # "https://erp.tiny.com.br/exportacao_contas_pagar?criterio=periodo&tipoPeriodo=pagamento",
            # "https://erp.tiny.com.br/exportacao_contas_pagar?criterio=periodo&tipoPeriodo=atrasadas",
            # "https://erp.tiny.com.br/exportacao_contas_pagar?criterio=periodo&tipoPeriodo=canceladas"
        ]

        for link in links:
            driver.get(link)

            try:
                wait = WebDriverWait(driver, 10)
                table = wait.until(EC.presence_of_element_located((By.ID, 'lista')))

                script = """
                    var tipoExportacao = document.querySelector('input[name="tipoExportacao"]');
                    tipoExportacao.value = "CSV";
                """
                driver.execute_script(script)
                
                rows = table.find_elements(By.TAG_NAME, 'tr')
                for i in range(1, len(rows) + 1):
                    write_to_logs(f'Downloading: {i} / {len(rows)}')
                    driver.execute_script(f"baixarArquivo({i})")
                    time.sleep(1)
            except:
                continue

    def download_caixa(driver):
        write_to_logs("download_caixa")
        driver.get('https://erp.tiny.com.br/caixa')
        time.sleep(5)
        driver.execute_script('document.querySelector("#opc-todas").click()')
        driver.execute_script('document.querySelector("#item-conta-todas").click()')
        time.sleep(2)
        driver.execute_script('baixarArquivo()')
        time.sleep(3)

    def download_notas_fiscais_saida(driver):
        write_to_logs("download_notas_fiscais_saida")
        current_date = datetime.now().strftime('%d/%m/%Y')
        driver.get(f'https://erp.tiny.com.br/exportacao_xml_nfe?tipo=S&dataIni=01/01/2015&dataFim={current_date}&modelo=55')
        time.sleep(3)
        driver.execute_script('baixarArquivo()')
        time.sleep(5)

    def download_notas_fiscais_entrada(driver):
        write_to_logs("download_notas_fiscais_entrada")
        current_date = datetime.now().strftime('%d/%m/%Y')
        driver.get(f'https://erp.tiny.com.br/exportacao_xml_nfe?tipo=E&dataIni=01/01/2015&dataFim={current_date}&modelo=55')
        time.sleep(3)
        driver.execute_script('baixarArquivo()')
        time.sleep(5)

    def download_notas_fiscais_entrada_emitidas_terceiros(driver, dt_first_nf_e_3):
        write_to_logs("download_notas_fiscais_entrada_emitidas_terceiros")
        url = 'https://erp.tiny.com.br/exportacao_xml_terceiros?tipo=E&modelo=55'
        
        driver.get(url)
        time.sleep(3)
            
        driver.execute_script('baixarArquivo()')
        
        current_date = datetime.now().strftime("%Y-%m-%d")
        secure_month = 1
        num_iterations = months_between(dt_first_nf_e_3, current_date) + secure_month
        for i in range(num_iterations):
            write_to_logs("Downloading " + current_date + f" - {i + 1} month")
            time.sleep(1)
            driver.execute_script('incrementaMeses(-1)')
            time.sleep(2)
            driver.execute_script('baixarArquivo()')

    def download_ordens_compra(driver, dt_first_oc):
        write_to_logs("download_notas_fiscais_entrada_emitidas_terceiros")
        url = 'https://erp.tiny.com.br/exportacao_pedidos_compra'
        
        driver.get(url)
        time.sleep(3)
        click_csv_opc = """document.querySelector('input[name="tipoExportacao"][value="CSV"]').click()""";
        driver.execute_script(click_csv_opc)
        time.sleep(1)
        
        driver.execute_script('baixarArquivo()')
        
        current_date = datetime.now().strftime("%Y-%m-%d")
        secure_month = 1
        num_iterations = months_between(dt_first_oc, current_date) + secure_month
        for i in range(num_iterations):
            write_to_logs("Downloading " + current_date + f" - {i + 1} month")
            time.sleep(1)
            driver.execute_script('previousMonth()')
            time.sleep(2)
            driver.execute_script('baixarArquivo()')

    def get_relation_idPedido_idNfe_vendas(driver, folder_path):
        write_to_logs("get_relation_idPedido_idNfe_vendas")
        driver.get('https://erp.tiny.com.br/vendas#list')
        time.sleep(3)
        
        aba_todos = """document.querySelector('a[data-filter="x"]').click()"""
        driver.execute_script(aba_todos)
        periodo_sem_filtro = """document.getElementById("opc-per-todas").click()"""
        driver.execute_script(periodo_sem_filtro)
        time.sleep(3)
        
        pegar_script = """
        function pegar() {
            let result = {};
            let trs = document.querySelectorAll('tr[id][idnota]');
            trs.forEach(tr => {
                let id = tr.getAttribute('id');
                let idnota = tr.getAttribute('idnota');
                if (idnota != '0') {
                    result[id] = idnota;
                }
            });
            return result;
        }
        window.pegar = pegar;
        """

        driver.execute_script(pegar_script)

        result = {}
        while True:
            result.update(driver.execute_script("return pegar()"))
            try:
                next_page = """document.querySelector('li.pnext a.link-pg').click()"""
                driver.execute_script(next_page)
                time.sleep(3)
            except:
                break

        with open(f"{folder_path}/idPedido-idNfe-saida.json", "w") as f:
            json.dump(result, f)

    def get_relation_idPedido_idNfe_compras(driver, folder_path):
        write_to_logs("get_relation_idPedido_idNfe_compras")
        driver.get('https://erp.tiny.com.br/pedidos_compra#list')
        time.sleep(3)
        
        # aba_todos = """document.querySelector('a[data-filter="x"]').click()"""
        # driver.execute_script(aba_todos)
        # periodo_sem_filtro = """document.getElementById("opc-per-todas").click()"""
        # driver.execute_script(periodo_sem_filtro)
        # time.sleep(3)
        
        pegar_script = """
        function pegar_dicionario_get_relation_idPedido_idNfe_compras() {
            let dicionario_get_relation_idPedido_idNfe_compras = {};
            var tabela = document.getElementById('tabelaListagem');
            var linhas = tabela.getElementsByTagName('tr');
            for (var i = 0; i < linhas.length; i++) {
                var linha = linhas[i];
                var ancora = linha.querySelector('a[idnota]');
                if (ancora) {
                    var idLinha = linha.getAttribute('id');
                    var hrefAncora = ancora.getAttribute('href');
                    dicionario_get_relation_idPedido_idNfe_compras[idLinha] = Number(hrefAncora.replace('notas_entrada#edit/', ''));
                }
            }

            return dicionario_get_relation_idPedido_idNfe_compras;
        }
        window.pegar_dicionario_get_relation_idPedido_idNfe_compras = pegar_dicionario_get_relation_idPedido_idNfe_compras;
        """

        driver.execute_script(pegar_script)

        result = {}
        while True:
            result.update(driver.execute_script("return pegar_dicionario_get_relation_idPedido_idNfe_compras()"))
            try:
                next_page = """document.querySelector('li.pnext a.link-pg').click()"""
                driver.execute_script(next_page)
                time.sleep(3)
            except:
                break

        with open(f"{folder_path}/idPedidoCompra-idNfe-entrada.json", "w") as f:
            json.dump(result, f)

    def export_planilha_nf_s(driver):
        write_to_logs("export_planilha_nf_s")
        driver.get("https://erp.tiny.com.br/exportacao_planilha_nf?tipo=S&modelo=55")
        time.sleep(5)
        driver.find_element(By.ID, "opc-todas").click()
        
        script_config = """
        var inputs = document.querySelectorAll('input[type="checkbox"]');
        inputs.forEach(function(input) {
            input.checked = true;
        });
        var inputCSV = document.querySelector('input[value="CSV"]');
        inputCSV.click()
        """
        
        driver.execute_script(script_config)
        time.sleep(1)
        
        driver.execute_script("baixarArquivo()")
        time.sleep(4)

    def logout(driver):
        print("logout")
        driver.get("https://erp.tiny.com.br/logout")
        time.sleep(5)

    def every_downloads_chrome(driver):
        if not driver.current_url.startswith("chrome://downloads"):
            driver.get("chrome://downloads/")
        return driver.execute_script("""
            var items = document.querySelector('downloads-manager')
                .shadowRoot.getElementById('downloadsList').items;
            if (items.every(e => e.state === "COMPLETE"))
                return items.map(e => e.fileUrl || e.file_url);
            """)

    try:
        login(driver, acc["user"], acc["pass"])
        
        get_relation_idPedido_idNfe_vendas(driver, folder_path)

        export_planilha_nf_s(driver)
        
        get_relation_idPedido_idNfe_compras(driver, folder_path)
        
        # download_ordens_compra(driver, acc["dt_first_oc"])

        download_notas_fiscais_entrada_emitidas_terceiros(driver, acc["dt_first_nf_e_3"])

        download_notas_fiscais_entrada(driver)

        download_notas_fiscais_saida(driver)

        download_contas_receber_e_pagar(driver)

        download_caixa(driver)

        logout(driver)

    except Exception as e:
        print(f"Ocorreu um erro: {e}")

    finally:
        # paths = WebDriverWait(driver, 120, 1).until(every_downloads_chrome)
        time.sleep(100)
        driver.quit()