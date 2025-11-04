"""
Парсер для сайту sapienstone.com - Керамограніт
Збирає дані з колекцій бренду Sapienstone
Результат: sapienstone_ceramic.xlsx
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

class SapienstoneParser:
    def __init__(self, output_file='sapienstone_ceramic.xlsx'):
        """Ініціалізація парсера"""
        self.output_file = output_file
        self.progress_file = 'progress_sapienstone.json'
        self.driver = None
        self.wait = None
        self.processed_urls = set()
        self.brand = 'Sapienstone'
        self.category = 'Керамограніт'
        self.base_url = 'https://www.sapienstone.com'
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
    
    def parse_catalog_page(self, catalog_url):
        """Парсить сторінку каталогу та отримує всі товари"""
        print("🌐 Завантаження каталогу...")
        
        try:
            self.driver.get(catalog_url)
            time.sleep(3)
            
            # Прокручуємо сторінку для завантаження всіх елементів
            self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)
            
            html = self.driver.page_source
            soup = BeautifulSoup(html, 'html.parser')
            
            # Шукаємо всі контейнери товарів
            product_containers = soup.find_all('div', class_='product-container')
            
            print(f"✅ Знайдено товарів: {len(product_containers)}")
            
            products = []
            for container in product_containers:
                try:
                    # Посилання на товар
                    link = container.find('a')
                    if not link or not link.get('href'):
                        continue
                    
                    product_url = self.base_url + link['href']
                    
                    # Пропускаємо вже оброблені
                    if product_url in self.processed_urls:
                        continue
                    
                    # Назва товару
                    title = ''
                    p_tag = container.find('p')
                    if p_tag:
                        strong = p_tag.find('strong')
                        if strong:
                            title = strong.text.strip()
                    
                    # Тип поверхні (Cashmere, тощо)
                    surface_type = ''
                    if p_tag:
                        i_tag = p_tag.find('i')
                        if i_tag:
                            surface_type = i_tag.text.strip()
                    
                    # Картинка (превью)
                    feature_photo = ''
                    img = container.find('img')
                    if img and img.get('src'):
                        feature_photo = self.base_url + img['src']
                    
                    products.append({
                        'url': product_url,
                        'title': title,
                        'surface_type': surface_type,
                        'feature_photo': feature_photo
                    })
                    
                except Exception as e:
                    print(f"⚠️ Помилка обробки контейнера: {e}")
                    continue
            
            print(f"✨ Нових товарів для обробки: {len(products)}")
            return products
            
        except Exception as e:
            print(f"❌ Помилка завантаження каталогу: {e}")
            return []
    
    def parse_product_detail(self, url):
        """Парсить детальну сторінку товару та отримує галерею"""
        max_retries = 3
        
        for attempt in range(max_retries):
            try:
                self.driver.get(url)
                time.sleep(3)
                
                # Прокручуємо до слайдера
                self.driver.execute_script("window.scrollTo(0, 500);")
                time.sleep(1)
                
                html = self.driver.page_source
                soup = BeautifulSoup(html, 'html.parser')
                
                # Шукаємо slick-slider
                gallery_images = []
                slick_track = soup.find('div', class_='slick-track')
                
                if slick_track:
                    slides = slick_track.find_all('div', class_='slick-slide')
                    
                    for slide in slides[:3]:  # Беремо перші 3 слайди
                        # Шукаємо посилання на велике зображення
                        link = slide.find('a')
                        if link and link.get('href'):
                            # Беремо big зображення, а не thumb
                            img_url = self.base_url + link['href']
                            gallery_images.append(img_url)
                
                # Доповнюємо до 3 елементів
                while len(gallery_images) < 3:
                    gallery_images.append('')
                
                return gallery_images[:3]
                
            except WebDriverException as e:
                if attempt < max_retries - 1:
                    print(f"⚠️ Помилка з'єднання, спроба {attempt + 2}/{max_retries}...")
                    self.restart_driver()
                    time.sleep(2)
                else:
                    print(f"❌ Не вдалось обробити {url}: {e}")
                    return ['', '', '']
            except Exception as e:
                print(f"⚠️ Помилка обробки {url}: {e}")
                return ['', '', '']
    
    def translate_surface_type(self, surface_type):
        """Перекладає тип поверхні на українську"""
        translations = {
            'Cashmere': 'Кашемір',
            'Polished': 'Полірована',
            'Matt': 'Матова',
            'Silk': 'Шовкова',
            'Natural': 'Натуральна',
            'Honed': 'Шліфована',
            'Structured': 'Структурована'
        }
        
        return translations.get(surface_type, surface_type)
    
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
                    product_data['Title'],
                    product_data['Type'],
                    product_data['Feature photo'],
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
                        'A': 20,  # Brand
                        'B': 25,  # Category
                        'C': 30,  # Title
                        'D': 20,  # Type
                        'E': 50,  # Feature photo
                        'F': 50,  # Gallery1
                        'G': 50,  # Gallery2
                        'H': 50,  # Gallery3
                    }
                    for col, width in column_widths.items():
                        worksheet.column_dimensions[col].width = width
                        
                print(f"   ✅ Файл створено: {self.output_file}")
                        
        except Exception as e:
            print(f"❌ Помилка збереження: {e}")
            import traceback
            traceback.print_exc()
    
    def parse_all(self, catalog_url):
        """Основна функція парсингу"""
        print("🚀 Початок парсингу Sapienstone керамограніту\n")
        
        # Отримуємо всі товари з каталогу
        products = self.parse_catalog_page(catalog_url)
        
        if not products:
            print("❌ Не знайдено товарів!")
            return
        
        # Обробляємо кожен товар
        print(f"\n📦 Обробка товарів...")
        
        for product in tqdm(products, desc="Прогрес"):
            if product['url'] and product['url'] not in self.processed_urls:
                try:
                    print(f"\n🔍 Обробка: {product['title']}")
                    
                    # Отримуємо галерею з детальної сторінки
                    gallery = self.parse_product_detail(product['url'])
                    
                    # Перекладаємо тип поверхні
                    surface_type_ua = self.translate_surface_type(product['surface_type'])
                    
                    # Формуємо дані товару
                    product_data = {
                        'Brand': self.brand,
                        'Category': self.category,
                        'Title': product['title'],
                        'Type': surface_type_ua,
                        'Feature photo': product['feature_photo'],
                        'Gallery1': gallery[0],
                        'Gallery2': gallery[1],
                        'Gallery3': gallery[2]
                    }
                    
                    print(f"   📋 Дані зібрано: Type={surface_type_ua}")
                    
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
    parser = SapienstoneParser(output_file='sapienstone_ceramic.xlsx')
    
    try:
        # URL каталогу
        catalog_url = 'https://www.sapienstone.com/collections'
        
        # Запускаємо парсинг
        parser.parse_all(catalog_url)
        
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