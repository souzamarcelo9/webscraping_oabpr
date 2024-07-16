import time
from selenium.webdriver.common.by import By

from driver.driver import Driver
from iterator.iteration import Interation
from ext.functions import *
from ext.elements import XPATH, CSS

class Bot(Interation):
    """Classe que define um bot para interação automatizada com páginas da web."""

    def __init__(self, log_file=True):
        """
        Inicializa um objeto Bot.

        Args:
            log_file (bool): Define se os registros serão salvos em um arquivo de log (padrão: True).
        """
        self.logger = setup_logging(to_file=True)
        
        self.driver = Driver(
            browser='chrome',
            headless=False,
            incognito=False,
            download_path='',
            desabilitar_carregamento_imagem=False
        ).driver

        self.base_url = 'https://www.oabpr.org.br/servicos-consulta-de-advogados/lista-de-advogados/?nr_inscricao=0&nome=0&cidade=0&especialidade=0&situacao=A'

        super().__init__(self.driver)

    def open_oab(self):
        """Abre a página da OAB."""

        try:
            self.load_page(self.base_url)
            self.wait_for(XPATH['dado'])
            return True
        except Exception as e:
            self.logger.error(f"Erro ao abrir página inicial da OAB.")
            self.quit()

    def get_last_page(self) -> int:
        """Retorna a última página da lista de advogados."""

        try:
            last_page = self.find(XPATH['ultima_pagina']).get_attribute('href')
            last_page = last_page.split('pg=')[-1]
            return int(last_page)
        except Exception as e:
            self.logger.error(f"Erro ao obter ultima pagina. {e}")

    def process_page(self, page: int):
        """
        Processa uma página da lista de advogados.

        Args:
            page (int): Número da página a ser processada.
        """

        try:
            if page != 1:
                self.load_page(self.base_url + f'&pg={page}')
            advogados_elements = self.find_all(XPATH['dado'])
            advogados = {advogado.text: advogado.get_attribute('href') for advogado in advogados_elements}
            
            for nome, link in advogados.items():
                self.process_advogado(nome, link)

        except Exception as e:
            self.logger.error(f"Erro ao processar página {page}: {e}")
            self.quit()

    def is_solved(self):
        """Verifica se o captcha já foi resolvido."""

        time.sleep(2)
        solved = self.find(CSS['solved'], metodo='css').get_attribute('aria-checked')
        if solved == 'true':
            return True
        else:
            return False

    def process_advogado(self, advogado: str, link: str):
        """
        Processa um advogado.

        Args:
            link (str): Link do advogado.
        """

        self.logger.info(f'Processando o advogado {advogado}')
        try:
            self.load_page(link)            
            self.resolve_captcha()
            dados = self.get_values()
            insert_values(dados)

        except Exception as e:
            self.logger.error(f"Erro ao processar advogado {advogado}: {e}")
            self.quit()

    def resolve_captcha(self):
        """Resolução do captcha."""
        
        try:
            iframe = self.find(CSS['reCAPTCHA'], metodo='css')
            self.driver.switch_to.frame(iframe)

            self.find(CSS['body'], metodo='css').click()

            if not self.is_solved():
                self.logger.debug('Resolvendo captcha...')
                self.driver.switch_to.default_content()

                iframe_img = self.find(XPATH['iframe'])
                self.driver.switch_to.frame(iframe_img)

                audio_button = self.find(CSS['audio'], metodo='css')
                audio_button.click()

                audio_link = self.find(CSS['audio-source'], metodo='css').get_attribute('href')
                path_mp3, path_wav = get_temps_files()

                text = convert_audio_to_string(audio_link, path_mp3, path_wav)
                input_audio = self.find(CSS['input'], metodo='css')
                input_audio.send_keys(text.lower())
                
                self.find(CSS['verify_btn'], metodo='css').click()

            self.driver.switch_to.default_content()
            self.find(CSS['enviar_btn'], metodo='css').click()
            self.wait_for(CSS['imprimir_btn'], metodo='css')
        except Exception as e:
            self.logger.error(f"Erro ao resolver captcha: {e}")
            self.quit()

    def get_values(self):
        """Retorna os dados do advogado."""

        labels = [
            'Número de Inscrição', 'Advogado', 'Impedimentos', 'Situação', 
            'Subseção', 'Data da Inscrição', 'Endereço Comercial', 'Telefone Comercial'
        ]
        dados = {label: '' for label in labels}

        try:
            self.wait_for(CSS['rows'], metodo='css')
            rows = self.find_all(CSS['rows'], metodo='css')
            
            for row in rows:
                cells = row.find_elements(By.TAG_NAME, 'td')
                if len(cells) > 1:
                    label = substituir_ultima_letra(cells[0].text.strip())
                    value = cells[1].text.strip()
                    if label in dados:
                        dados[label] = value
            
            # Retorna os dados em uma ordem consistente
            return [dados[label] for label in labels]
        except Exception as e:
            self.logger.error(f"Erro ao obter dados do advogado: {e}")
            self.quit()
        