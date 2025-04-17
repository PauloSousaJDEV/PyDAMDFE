from common_imports import *

# Instancia o objeto Navegador
atualNavegador = Navegador

class Damdfe:
    
    def __init__(self, nome, placa):
        self.nome = nome
        self.placa = placa

    # Função para exibir a placa
    def exibirPlaca(self):
        print(f"Placa: {self.placa}")
        return f"Placa: {self.placa}"

    # Função que executa a automação para cancelar o DAMDFE
    def cancelarDamdfe(self):
        atualNavegador.abrirNavegador()
        atualNavegador.Localidade
        try:
            search_input = atualNavegador.wait.until(EC.presence_of_element_located(
                (By.XPATH, "//input[@placeholder='Pesquisar MDFe']")))

            search_input.clear()
            search_input.send_keys(self.placa)
            sleep(1)
            search_input.send_keys(Keys.ENTER)

            sleep(2)
            checkbox = atualNavegador.wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//table/tbody/tr[1]//p-checkbox//div[@class='p-checkbox-box']")))
            atualNavegador.driver.execute_script("arguments[0].click();", checkbox)

            sleep(2)
            botao_encerrar = atualNavegador.wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//button/span[contains(text(),'Encerrar')]")))
            atualNavegador.driver.execute_script("arguments[0].click();", botao_encerrar)

            sleep(2)
            confirmar_modal = atualNavegador.wait.until(EC.element_to_be_clickable(
                (By.XPATH, '//div[contains(@class,"cdk-overlay-pane")]//lib-button[2]/button')))
            atualNavegador.driver.execute_script("arguments[0].click();", confirmar_modal)

            sleep(3.5)
            botao_fechar_modal = atualNavegador.wait.until(EC.element_to_be_clickable(
                (By.XPATH, '//app-mdfe-close-modal-response//button[span[text()="Fechar"]]')))
            atualNavegador.driver.execute_script("arguments[0].click();", botao_fechar_modal)

        except Exception as e:
            print(f"❌ Erro durante o encerramento do MDFe: {e}")
        finally:
            if atualNavegador.driver:
                atualNavegador.driver.quit()
                print("🛑 Driver do Selenium finalizado com sucesso.")

# Função para selecionar ambiente conforme localidade
def Selecionar_ambiente(localidade):
    localidade = localidade.lower()
    if localidade == "duque de caxias":
        atualNavegador.wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//lib-await-panel/div/div/div/div[2]/button/span'))).click()
        atualNavegador.wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//lib-company-selection/lib-await-panel/div/div/div/lib-company-selection-card[15]/div'))).click()
        atualNavegador.wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="menuLateral"]/div[2]/lib-sidenav-menu-item[2]/a'))).click()
        
    elif localidade == "porto alegre":
        atualNavegador.wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//lib-await-panel/div/div/div/div[2]/button/span'))).click()
        atualNavegador.wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//lib-company-selection-card[14]/div'))).click()
        atualNavegador.wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="menuLateral"]/div[2]/lib-sidenav-menu-item[2]/a'))).click()
    else:
        raise ValueError("Ambiente inválido. Escolha 'Duque de Caxias' ou 'Porto Alegre'.")

