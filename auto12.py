from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep

placas = ["RIO02A1"] * 10

def main():
    print("🚀 Iniciando o script...")

    driver = webdriver.Chrome()
    wait = WebDriverWait(driver, 20)

    try:
        driver.get('https://mdfe-beta.hivecloud.com.br/')
        
        # Login
        wait.until(EC.presence_of_element_located((By.XPATH, '//lib-form-control[1]//input'))).send_keys('Omar.Teixeira@jti.com')
        driver.find_element(By.XPATH, '//lib-form-control[2]//input').send_keys('17318208')
        driver.find_element(By.XPATH, '//div[2]/lib-button/button/span').click()

        # Selecionar empresa
        wait.until(EC.element_to_be_clickable((By.XPATH, '//lib-await-panel/div/div/div/div[2]/button/span'))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, '//lib-company-selection-card[14]/div'))).click()

        # Abrir menu MDF-e
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="menuLateral"]/div[2]/lib-sidenav-menu-item[2]/a'))).click()

        for placa in placas:
            print(f"\n🔍 Processando placa: {placa}")
            
            sleep(5)
            # Pesquisar placa
            search_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Pesquisar MDFe']")))
            search_input.clear()
            sleep(2)
            search_input.send_keys(placa)
            sleep(2)
            search_input.send_keys(Keys.ENTER)

            wait.until(EC.presence_of_element_located((By.XPATH, "//table//tr[1]")))
            sleep(3)

            try:
                # Marcar checkbox
                checkbox = wait.until(EC.element_to_be_clickable((By.XPATH, "//table/tbody/tr[1]//p-checkbox//div[@class='p-checkbox-box']")))
                driver.execute_script("arguments[0].click();", checkbox)
                print("✅ Checkbox marcado.")
            except Exception as e:
                print("❌ Falha ao marcar checkbox:", e)
                continue

            # Clicar em Encerrar
            try:
                sleep(3)
                botao_encerrar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button/span[contains(text(),'Encerrar')]")))
                driver.execute_script("arguments[0].click();", botao_encerrar)
                print("✅ Botão de Encerrar clicado.")
            except Exception as e:
                print("❌ Falha ao clicar em Encerrar:", e)
                continue
             
            # Confirmar encerramento no modal
            try:
                sleep(3)
                confirmar_modal = wait.until(EC.element_to_be_clickable((By.XPATH, '//div[contains(@class,"cdk-overlay-pane")]//lib-button[2]/button')))
                driver.execute_script("arguments[0].click();", confirmar_modal)
                print("✅ Encerramento confirmado.")
            except Exception as e:
                print("❌ Erro ao confirmar encerramento:", e)
                continue
           

            # Clicar no botão "Fechar" do modal de confirmação
            try:
                sleep(3)
                botao_fechar_modal = wait.until(EC.element_to_be_clickable((By.XPATH,
                    '//app-mdfe-close-modal-response//button[span[text()="Fechar"]]')))
                driver.execute_script("arguments[0].click();", botao_fechar_modal)
                print("✅ Modal de encerramento fechado.")
            except Exception as e:
                print("❌ Erro ao fechar modal de encerramento:", e)
                sleep(5)

            print(f"🏁 Placa processada com sucesso: {placa}")
            sleep(4.5)
            

    except Exception as e:
        print("🚨 Erro geral durante execução:", e)

    finally:
        print("\n🧹 Finalizando navegador.")
        driver.quit()

if __name__ == "__main__":
    main()
