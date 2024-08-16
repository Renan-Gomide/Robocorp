
# class NewsScraper:
#     SEARCH_BUTTON_XPATH = '/html/body/nav/form/button'
#     SEARCH_BOX_XPATH = '/html/body/nav/form/input'
#     FILTER_TAB_XPATH = '//*[@id="refine"]'
    
#     CATEGORY_XPATH_MAP = {
#         'article': '//*[@id="custom-facet-0"]/div/div[2]/div/ul/li[2]/label',
#         'audio': '//*[@id="custom-facet-0"]/div/div[2]/div/ul/li[3]/label',
#         'gallery': '//*[@id="custom-facet-0"]/div/div[2]/div/ul/li[4]/label',
#         'video': '//*[@id="custom-facet-0"]/div/div[2]/div/ul/li[5]/label',
#         'default': '//*[@id="custom-facet-0"]/div/div[2]/div/ul/li[1]/label'
#     }
    
#     TIME_XPATH_MAP = {
#         'day': '//*[@id="custom-facet-1"]/div/div[2]/div/ul/li[2]/label',
#         'week': '//*[@id="custom-facet-1"]/div/div[2]/div/ul/li[3]/label',
#         'month': '//*[@id="custom-facet-1"]/div/div[2]/div/ul/li[4]/label',
#         'year': '//*[@id="custom-facet-1"]/div/div[2]/div/ul/li[5]/label',
#         'default': '//*[@id="custom-facet-1"]/div/div/ul/li[1]/label'
#     }

#     def __init__(self, config_file='config.ini'):
#         self.config = configparser.ConfigParser()
#         self.config.read(config_file)
#         self.query = self.config.get('search', 'query')
#         self.category = self.config.get('filters', 'category', fallback='default')
#         self.time = self.config.get('filters', 'time', fallback='default')
#         self.driver = webdriver.Chrome()
#         self.driver.implicitly_wait(15)

#     def search_news(self):
#         self.driver.get('https://www.news.com.au/')
#         try:
#             self._click_element(self.SEARCH_BUTTON_XPATH)
#             search_box = self._find_element(self.SEARCH_BOX_XPATH)
#             search_box.send_keys(self.query)
#             search_box.send_keys(Keys.RETURN)
#             logging.info("Search submitted.")
#             WebDriverWait(self.driver, 10).until(
#                 EC.presence_of_element_located((By.XPATH, self.FILTER_TAB_XPATH))
#             )
#             self._apply_filters()
#         except Exception as e:
#             logging.error(f"Error applying the filters: {e}")
#             self.driver.quit()
#             raise

#     def _apply_filters(self):
#         self._click_element(self.FILTER_TAB_XPATH)
#         category_xpath = self.CATEGORY_XPATH_MAP.get(self.category, self.CATEGORY_XPATH_MAP['default'])
#         time_xpath = self.TIME_XPATH_MAP.get(self.time, self.TIME_XPATH_MAP['default'])
#         self._click_element(category_xpath)
#         logging.info(f"Category filter '{self.category}' applied.")
#         self._click_element(time_xpath)
#         logging.info(f"Time filter applied for the last {self.time}.")
#         self._click_element(self.FILTER_TAB_XPATH)

#     def _click_element(self, xpath):
#         WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH, xpath))).click()

#     def _find_element(self, xpath):
#         return WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, xpath)))

#     def download_and_save_image(self, image_url, filename, format='JPEG'):
#         try:
#             with urlopen(image_url) as response:
#                 image_data = response.read()
#             image = Image.open(BytesIO(image_data))
#             os.makedirs(os.path.dirname(filename), exist_ok=True)
#             image.save(filename, format=format)
#             logging.info(f"Image saved as '{filename}'.")
#         except Exception as e:
#             logging.error(f"Error downloading or saving the image: {e}")

#     def process_articles(self):
#         articles = self.driver.find_elements(By.CLASS_NAME, 'storyblock')
#         output_dir = 'output/news_images'
#         os.makedirs(output_dir, exist_ok=True)
#         articles_data = []
#         for article in articles:
#             try:
#                 title = article.find_element(By.CLASS_NAME, 'storyblock_title').text
#                 date = article.find_element(By.CLASS_NAME, 'storyblock_datetime').get_attribute('datetime')
#                 description = self._get_element_text(article, By.CLASS_NAME, 'storyblock_standfirst')
#                 image_url = article.find_element(By.CLASS_NAME, 'responsive-img_img').get_attribute('src')
#                 image_filename = self._get_image_filename(output_dir, image_url)
#                 self.download_and_save_image(image_url, image_filename)
#                 articles_data.append(self._extract_article_data(title, date, description, image_filename))
#                 sleep(1)

                


#             except Exception as e:
#                 logging.warning(f"Error processing the article: {e}")
#         return articles_data

#     def _get_element_text(self, parent, by, value):
#         elements = parent.find_elements(by, value)
#         return elements[0].text if elements else None

#     def _get_image_filename(self, output_dir, image_url):
#         parsed_url = urlparse(image_url)
#         return os.path.join(output_dir, os.path.basename(parsed_url.path) + ".jpeg")

#     def _extract_article_data(self, title, date, description, image_filename):
#         search_phrases = [self.query]
#         phrase_count = sum(title.lower().count(phrase.lower()) for phrase in search_phrases) + \
#                        (sum(description.lower().count(phrase.lower()) for phrase in search_phrases) if description else 0)
#         money_pattern = re.compile(r'(\$\d+(\.\d{1,2})?)|(\d+\s?(dollars|USD))')
#         contains_money = bool(money_pattern.search(title) or (description and money_pattern.search(description)))
#         return {
#             "Title": title,
#             "Date": date,
#             "Description": description,
#             "Picture Filename": os.path.relpath(image_filename, 'output'),
#             "Count of Search Phrases": phrase_count,
#             "Contains Money": contains_money
#         }

#     def save_to_excel(self, articles_data):
#         output_dir = 'output'
#         os.makedirs(output_dir, exist_ok=True)
#         file_path = os.path.join(output_dir, "news_data.xlsx")
#         workbook = Workbook()
#         worksheet = workbook.active
#         worksheet.title = "News Data"
#         if articles_data:
#             worksheet.append(list(articles_data[0].keys()))
#             for article in articles_data:
#                 worksheet.append(list(article.values()))
#         workbook.save(file_path)
#         logging.info(f"Data saved in '{file_path}'.")

#     def run(self):
#         try:
#             self.search_news()
#             articles_data = self.process_articles()
#             self.save_to_excel(articles_data)
#         except Exception as e:
#             logging.error(f"Error in the scraping process: {e}")
#         finally:
#             self.driver.quit()

# @task
# def execute_scraper():
#     scraper = NewsScraper()
#     scraper.run()

# if __name__ == "__main__":
#     execute_scraper()

from robocorp.tasks import task
import os
import re
import configparser
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from urllib.parse import urlparse
from time import sleep
import urllib.request
from PIL import Image
from io import BytesIO
from openpyxl import Workbook
import logging


logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class NewsScraper:
    def __init__(self, config_file='config.ini'):
        self.config = configparser.ConfigParser()
        self.config.read(config_file)
        self.query = self.config.get('search', 'query')
        self.category = self.config.get('filters', 'category')
        self.time = self.config.get('filters', 'time')
        self.driver = webdriver.Chrome()
        self.driver.implicitly_wait(15)

    def search_news(self):
        url = 'https://www.news.com.au/'
        self.driver.get(url)

        try:
            search_button = self.driver.find_element(By.XPATH, '/html/body/nav/form/button')
            search_button.click()
            logging.info("Search button clicked.")

            search_box = self.driver.find_element(By.XPATH, '/html/body/nav/form/input')
            search_box.send_keys(self.query)
            search_box.send_keys(Keys.RETURN)
            logging.info("Search submitted.")
            sleep(5)

            filter_tab = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="refine"]'))
            )
            filter_tab.click()

            category_xpath_map = {
                'article': '//*[@id="custom-facet-0"]/div/div[2]/div/ul/li[2]/label',
                'audio': '//*[@id="custom-facet-0"]/div/div[2]/div/ul/li[3]/label',
                'gallery': '//*[@id="custom-facet-0"]/div/div[2]/div/ul/li[4]/label',
                'video': '//*[@id="custom-facet-0"]/div/div[2]/div/ul/li[5]/label',
            }
            category_xpath = category_xpath_map.get(self.category, '//*[@id="custom-facet-0"]/div/div[2]/div/ul/li[1]/label')

            category_filter = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, category_xpath))
            )
            category_filter.click()
            logging.info(f"Category filter '{self.category}' applied.")

            time_xpath_map = {
                'day': '//*[@id="custom-facet-1"]/div/div[2]/div/ul/li[2]/label',
                'week': '//*[@id="custom-facet-1"]/div/div[2]/div/ul/li[3]/label',
                'month': '//*[@id="custom-facet-1"]/div/div[2]/div/ul/li[4]/label',
                'year': '//*[@id="custom-facet-1"]/div/div[2]/div/ul/li[5]/label',
            }
            time_xpath = time_xpath_map.get(self.time, '//*[@id="custom-facet-1"]/div/div[2]/div/ul/li[1]/label')

            time_filter = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, time_xpath))
            )
            time_filter.click()
            logging.info(f"Time filter applied for the last {self.time}.")
            sleep(3)

            filter_tab.click()
            sleep(2)

        except Exception as e:
            logging.error(f"Error applying the filters: {e}")
            self.driver.quit()
            raise

    def download_and_save_image(self, image_url, filename, format='JPEG'):
        try:
            with urllib.request.urlopen(image_url) as response:
                image_data = response.read()
            image = Image.open(BytesIO(image_data))
            os.makedirs(os.path.dirname(filename), exist_ok=True)
            image.save(filename, format=format)
            logging.info(f"Image saved as '{filename}'.")
        except Exception as e:
            logging.error(f"Error downloading or saving the image: {e}")

    def process_articles(self):
        articles = self.driver.find_elements(By.CLASS_NAME, 'storyblock')
        
        output_dir = 'output/news_images'
        os.makedirs(output_dir, exist_ok=True)
        
        articles_data = []

        for article in articles:
            try:
                title_element = article.find_element(By.CLASS_NAME, 'storyblock_title')
                title = title_element.text

                date_element = article.find_element(By.CLASS_NAME, 'storyblock_datetime')
                date = date_element.get_attribute('datetime')

                description_element = article.find_elements(By.CLASS_NAME, 'storyblock_standfirst')
                description = description_element[0].text if description_element else None

                image_element = article.find_element(By.CLASS_NAME, 'responsive-img_img')
                image_url = image_element.get_attribute('src')

                parsed_url = urlparse(image_url)
                image_filename = os.path.join(output_dir, os.path.basename(parsed_url.path) + ".jpeg")

                self.download_and_save_image(image_url, image_filename)

                # Caminho relativo para salvar no Excel
                relative_image_path = os.path.join('news_images', os.path.basename(parsed_url.path) + ".jpeg")

                search_phrases = [self.query]
                phrase_count = sum(title.lower().count(phrase.lower()) for phrase in search_phrases) + \
                               (sum(description.lower().count(phrase.lower()) for phrase in search_phrases) if description else 0)

                money_pattern = re.compile(r'(\$\d+(\.\d{1,2})?)|(\d+\s?(dollars|USD))')
                contains_money = bool(money_pattern.search(title) or (description and money_pattern.search(description)))

                articles_data.append({
                    "Title": title,
                    "Date": date,
                    "Description": description,
                    "Picture Filename": relative_image_path,
                    "Count of Search Phrases": phrase_count,
                    "Contains Money": contains_money
                })

                sleep(1)

            except Exception as e:
                logging.warning(f"Error processing the article: {e}")
        return articles_data

    def save_to_excel(self, articles_data):

        output_dir = 'output'
        os.makedirs(output_dir, exist_ok=True)
        file_path = os.path.join(output_dir, "news_data.xlsx")

        # Cria uma nova workbook e ativa a primeira worksheet
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "News Data"

        # Se a estrutura de dados for uma lista de dicionários, converta-a para uma lista de listas.
        if articles_data:
            # Adiciona os cabeçalhos
            headers = list(articles_data[0].keys())
            worksheet.append(headers)

            # Adiciona os dados
            for article in articles_data:
                worksheet.append(list(article.values()))

        # Salva o workbook no arquivo Excel
        workbook.save(file_path)
        logging.info("Data saved in 'news_data.xlsx'.")
    
    def run(self):
        try:
            self.search_news()
            articles_data = self.process_articles()
            self.save_to_excel(articles_data)
        except Exception as e:
            logging.error(f"Error in the scraping process: {e}")
        finally:
            self.driver.quit()


@task
def run():
    scraper = NewsScraper()
    scraper.run()

if __name__ == "__main__":
    run()
