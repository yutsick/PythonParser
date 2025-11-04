"""
Парсер для сайту ascale.es - Керамограніт
Збирає дані з колекцій бренду Ascale
Результат: ascale_ceramic.xlsx
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from bs4 import BeautifulSoup
import pandas as pd
import time
from tqdm import tqdm
import os
import json

class AscaleParser:
    def __init__(self, output_file='ascale_ceramic.xlsx'):
        """Ініціалізація парсера"""
        self.output_file = output_file
        self.progress_file = 'progress_ascale.json'
        self.driver = None
        self.wait = None
        self.processed_urls = set()
        self.brand = 'Ascale'
        self.category = 'Керамограніт'
        self.init_driver()
        self.load_progress()
        
    def init_driver(self):
        """Ініціалізація драйвера браузера"""
        options = webdriver.ChromeOptions()
        # Закоментуйте наступний рядок, якщо хочете бачити браузер
        options.add_argument('--headless')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36')
        options.add_argument('--disable-gpu')
        options.add_argument('--window-size=1920,1080')
        options.add_argument('--lang=en-US')
        
        try:
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=options)
            self.wait = WebDriverWait(self.driver, 15)
            print("✅ Браузер успішно запущено")
        except Exception as e:
            print(f"❌ Помилка запуску браузера: {e}")
            raise
    
    def load_progress(self):
        """Завантажує прогрес з файлу"""
        if os.path.exists(self.progress_file):
            try:
                with open(self.progress_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.processed_urls = set(data.get('processed_urls', []))
                print(f"📂 Завантажено прогрес: {len(self.processed_urls)} товарів вже оброблено")
            except Exception as e:
                print(f"⚠️ Не вдалось завантажити прогрес: {e}")
                self.processed_urls = set()
        else:
            print("🆕 Початок нового парсингу")
    
    def save_progress(self):
        """Зберігає прогрес у файл"""
        try:
            with open(self.progress_file, 'w', encoding='utf-8') as f:
                json.dump({
                    'processed_urls': list(self.processed_urls)
                }, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"⚠️ Помилка збереження прогресу: {e}")
    
    def restart_driver(self):
        """Перезапуск драйвера при помилках"""
        print("🔄 Перезапуск браузера...")
        try:
            if self.driver:
                self.driver.quit()
        except:
            pass
        
        time.sleep(3)
        self.init_driver()
    
    def get_collection_urls(self, main_url):
        """Отримує URL всіх колекцій"""
        print("🌐 Завантаження головної сторінки колекцій...")
        
        try:
            self.driver.get(main_url)
            time.sleep(3)
            
            # Прокручуємо сторінку, щоб завантажити всі елементи
            self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)
            
            html = self.driver.page_source
            soup = BeautifulSoup(html, 'html.parser')
            
            # Шукаємо всі блоки колекцій
            collection_blocks = soup.find_all('div', class_='jet-listing-grid__item')
            
            collections = []
            for block in collection_blocks:
                try:
                    # Шукаємо посилання на колекцію
                    link = block.find('a', {'data-element_type': 'container'})
                    if link and link.get('href'):
                        collection_url = link['href']
                        
                        # Назва колекції
                        heading = block.find('h3', class_='elementor-heading-title')
                        collection_name = heading.text.strip() if heading else ''
                        
                        collections.append({
                            'name': collection_name,
                            'url': collection_url
                        })
                        
                except Exception as e:
                    print(f"⚠️ Помилка обробки блоку колекції: {e}")
                    continue
            
            print(f"✅ Знайдено колекцій: {len(collections)}")
            return collections
            
        except Exception as e:
            print(f"❌ Помилка завантаження колекцій: {e}")
            return []
    
    def parse_collection_page(self, collection_url, collection_name):
        """Парсить сторінку колекції та отримує всі товари"""
        print(f"\n📂 Обробка колекції: {collection_name}")
        
        try:
            self.driver.get(collection_url)
            time.sleep(3)
            
            # Прокручуємо сторінку
            self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)
            
            html = self.driver.page_source
            soup = BeautifulSoup(html, 'html.parser')
            
            # Шукаємо всі картки товарів
            product_cards = soup.find_all('div', class_='jet-listing-grid__item')
            
            products = []
            for card in product_cards:
                try:
                    # Назва товару
                    title_elem = card.find('h3', class_='elementor-heading-title')
                    if not title_elem:
                        continue
                    
                    title_link = title_elem.find('a')
                    title = title_link.text.strip() if title_link else title_elem.text.strip()
                    product_url = title_link['href'] if title_link and title_link.get('href') else None
                    
                    if not product_url or product_url in self.processed_urls:
                        continue
                    
                    # Опис товару
                    description_elem = card.find('div', class_='description')
                    description = ''
                    if description_elem:
                        desc_container = description_elem.find('div', class_='elementor-widget-container')
                        if desc_container:
                            # Збираємо весь текст з параграфів
                            paragraphs = desc_container.find_all('p')
                            description = ' '.join([p.get_text(strip=True) for p in paragraphs])
                    
                    # Картинка товару (превью)
                    img_elem = card.find('img', class_='lazyloaded')
                    if not img_elem:
                        img_elem = card.find('img')
                    
                    feature_photo = ''
                    if img_elem:
                        feature_photo = img_elem.get('src') or img_elem.get('data-lazy-src', '')
                    
                    products.append({
                        'url': product_url,
                        'title': title,
                        'description': description,
                        'feature_photo': feature_photo,
                        'collection': collection_name
                    })
                    
                except Exception as e:
                    print(f"⚠️ Помилка обробки картки товару: {e}")
                    continue
            
            print(f"✨ Знайдено товарів у колекції: {len(products)}")
            return products
            
        except Exception as e:
            print(f"❌ Помилка обробки колекції: {e}")
            return []
    
    def parse_product_detail(self, url):
        """Парсить детальну сторінку товару та отримує галерею і тип поверхні"""
        max_retries = 3
        
        for attempt in range(max_retries):
            try:
                self.driver.get(url)
                time.sleep(3)
                
                # Прокручуємо до галереї
                self.driver.execute_script("window.scrollTo(0, 800);")
                time.sleep(1)
                
                html = self.driver.page_source
                soup = BeautifulSoup(html, 'html.parser')
                
                # Шукаємо всі слайди у свайпері
                gallery_images = []
                swiper_slides = soup.find_all('div', class_='swiper-slide')
                
                for slide in swiper_slides[:3]:  # Беремо максимум 3 зображення
                    # Пропускаємо дублікати
                    if 'swiper-slide-duplicate' in slide.get('class', []):
                        continue
                    
                    img = slide.find('img', class_='swiper-slide-image')
                    if img:
                        img_url = img.get('data-lazy-src') or img.get('src', '')
                        if img_url and img_url.startswith('http'):
                            gallery_images.append(img_url)
                
                # Доповнюємо до 3 елементів
                while len(gallery_images) < 3:
                    gallery_images.append('')
                
                # Шукаємо тип поверхні
                surface_type = ''
                format_rows = soup.find_all('div', class_='jedv-enabled--yes')
                
                for row in format_rows:
                    # Шукаємо всі heading елементи в рядку
                    headings = row.find_all('div', class_='elementor-widget-heading')
                    
                    # Третій елемент - це тип поверхні
                    if len(headings) >= 3:
                        surface_span = headings[2].find('span', class_='elementor-heading-title')
                        if surface_span:
                            surface_type = surface_span.text.strip()
                            break
                
                # Переклад типів поверхонь
                surface_translations = {
                    'Polished': 'Полірована',
                    'Matt': 'Матова',
                    'Lappato': 'Лаппатована',
                    'Feel': 'Натуральна',
                    'Natural': 'Натуральна',
                    'Velvet': 'Оксамитова',
                    'Structured': 'Структурована'
                }
                
                # Перекладаємо якщо знайдено переклад
                translated_surfaces = []
                if surface_type:
                    for surf in surface_type.split(','):
                        surf = surf.strip()
                        # Перевіряємо чи є переклад
                        translated = surface_translations.get(surf, surf)
                        translated_surfaces.append(translated)
                    
                    surface_type = ', '.join(translated_surfaces)
                
                return {
                    'gallery': gallery_images[:3],
                    'surface_type': surface_type
                }
                
            except WebDriverException as e:
                if attempt < max_retries - 1:
                    print(f"⚠️ Помилка з'єднання, спроба {attempt + 2}/{max_retries}...")
                    self.restart_driver()
                    time.sleep(2)
                else:
                    print(f"❌ Не вдалось обробити {url}: {e}")
                    return {
                        'gallery': ['', '', ''],
                        'surface_type': ''
                    }
            except Exception as e:
                print(f"⚠️ Помилка обробки {url}: {e}")
                return {
                    'gallery': ['', '', ''],
                    'surface_type': ''
                }
    
    def save_product_to_excel(self, product_data):
        """Додає один товар до Excel файлу"""
        try:
            print(f"💾 Збереження: {product_data.get('Title', 'Без назви')}")
            
            # Якщо файл існує, дописуємо
            if os.path.exists(self.output_file):
                from openpyxl import load_workbook
                
                wb = load_workbook(self.output_file)
                ws = wb['Products']
                
                ws.append([
                    product_data['Brand'],
                    product_data['Category'],
                    product_data['Collection'],
                    product_data['Title'],
                    product_data['Description'],
                    product_data['Feature photo'],
                    product_data['Type'],
                    product_data['Gallery1'],
                    product_data['Gallery2'],
                    product_data['Gallery3']
                ])
                
                wb.save(self.output_file)
                wb.close()
                print(f"   ✅ Збережено!")
            else:
                # Створюємо новий файл
                df = pd.DataFrame([product_data])
                with pd.ExcelWriter(self.output_file, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Products')
                    
                    worksheet = writer.sheets['Products']
                    column_widths = {
                        'A': 15,  # Brand
                        'B': 25,  # Category
                        'C': 20,  # Collection
                        'D': 30,  # Title
                        'E': 60,  # Description
                        'F': 50,  # Feature photo
                        'G': 20,  # Type
                        'H': 50,  # Gallery1
                        'I': 50,  # Gallery2
                        'J': 50,  # Gallery3
                    }
                    for col, width in column_widths.items():
                        worksheet.column_dimensions[col].width = width
                        
                print(f"   ✅ Файл створено: {self.output_file}")
                        
        except Exception as e:
            print(f"❌ Помилка збереження: {e}")
            import traceback
            traceback.print_exc()
    
    def parse_all(self, main_url):
        """Основна функція парсингу"""
        print("🚀 Початок парсингу Ascale керамограніту\n")
        
        # Отримуємо всі колекції
        collections = self.get_collection_urls(main_url)
        
        if not collections:
            print("❌ Не знайдено жодної колекції!")
            return
        
        # Обробляємо кожну колекцію
        for collection in collections:
            print(f"\n{'='*60}")
            print(f"📂 Колекція: {collection['name']}")
            print(f"🔗 URL: {collection['url']}")
            print(f"{'='*60}")
            
            # Отримуємо товари з колекції
            products = self.parse_collection_page(collection['url'], collection['name'])
            
            if not products:
                print(f"⚠️ Немає товарів у колекції '{collection['name']}'")
                continue
            
            # Обробляємо кожен товар
            print(f"\n📦 Обробка товарів...")
            
            for product in tqdm(products, desc=collection['name']):
                if product['url'] and product['url'] not in self.processed_urls:
                    try:
                        print(f"\n🔍 Обробка: {product['title']}")
                        
                        # Отримуємо галерею та тип поверхні з детальної сторінки
                        details = self.parse_product_detail(product['url'])
                        
                        # Формуємо дані товару
                        product_data = {
                            'Brand': self.brand,
                            'Category': self.category,
                            'Collection': product['collection'],
                            'Title': product['title'],
                            'Description': product['description'],
                            'Feature photo': product['feature_photo'],
                            'Type': details['surface_type'],
                            'Gallery1': details['gallery'][0],
                            'Gallery2': details['gallery'][1],
                            'Gallery3': details['gallery'][2]
                        }
                        
                        # Зберігаємо товар
                        self.save_product_to_excel(product_data)
                        
                        # Додаємо до оброблених
                        self.processed_urls.add(product['url'])
                        
                        # Зберігаємо прогрес
                        if len(self.processed_urls) % 5 == 0:
                            self.save_progress()
                        
                        time.sleep(0.5)
                        
                    except Exception as e:
                        print(f"\n❌ Помилка обробки товару: {e}")
                        self.save_progress()
                        continue
        
        # Фінальне збереження
        self.save_progress()
        
        print(f"\n✅ Парсинг завершено!")
        print(f"📊 Всього оброблено товарів: {len(self.processed_urls)}")
        print(f"💾 Файл збережено: {self.output_file}")
    
    def close(self):
        """Закриває браузер"""
        if self.driver:
            self.driver.quit()


def main():
    """Головна функція"""
    parser = AscaleParser(output_file='ascale_ceramic.xlsx')
    
    try:
        # URL головної сторінки колекцій
        main_url = 'https://www.ascale.es/en/collections/'
        
        # Запускаємо парсинг
        parser.parse_all(main_url)
        
    except KeyboardInterrupt:
        print("\n\n⏸️ Парсинг зупинено")
        print("💾 Прогрес збережено")
    except Exception as e:
        print(f"\n❌ Критична помилка: {e}")
        print("💾 Прогрес збережено")
    finally:
        parser.save_progress()
        parser.close()


if __name__ == "__main__":
    main()