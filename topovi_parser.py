"""
Парсер для сайту topovi.com.ua
Версія 3.0 - з записом у файл та продовженням з місця зупинки
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

class TopoviParser:
    def __init__(self, output_file='topovi_products.xlsx'):
        """Ініціалізація парсера"""
        self.output_file = output_file
        self.progress_file = 'progress.json'
        self.driver = None
        self.wait = None
        self.processed_urls = set()
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
        
    def load_all_products(self, url):
        """Завантажує всі товари, натискаючи кнопку 'Load more'"""
        print("🌐 Завантаження сторінки категорії...")
        
        max_retries = 3
        for attempt in range(max_retries):
            try:
                self.driver.get(url)
                time.sleep(3)
                break
            except WebDriverException as e:
                if attempt < max_retries - 1:
                    print(f"⚠️ Помилка завантаження сторінки, спроба {attempt + 2}/{max_retries}...")
                    self.restart_driver()
                else:
                    raise
        
        # Натискаємо кнопку "Load more" доки вона є
        click_count = 0
        consecutive_errors = 0
        
        while consecutive_errors < 3:
            try:
                # Шукаємо кнопку load-more
                load_more_btn = self.driver.find_element(By.CSS_SELECTOR, '.btn.load-more')
                
                # Перевіряємо чи кнопка видима і активна
                if load_more_btn.is_displayed():
                    # Скролимо до кнопки
                    self.driver.execute_script("arguments[0].scrollIntoView(true);", load_more_btn)
                    time.sleep(0.5)
                    
                    # Клікаємо
                    load_more_btn.click()
                    click_count += 1
                    print(f"📥 Завантажено блок #{click_count}...")
                    time.sleep(2)
                    consecutive_errors = 0  # Скидаємо лічильник помилок
                else:
                    break
                    
            except NoSuchElementException:
                print("✅ Всі товари завантажено!")
                break
            except Exception as e:
                consecutive_errors += 1
                print(f"⚠️ Помилка при натисканні ({consecutive_errors}/3): {e}")
                time.sleep(2)
        
        # Отримуємо HTML після завантаження всіх товарів
        return self.driver.page_source
    
    def parse_product_list(self, html):
        """Парсить список товарів з категорії"""
        soup = BeautifulSoup(html, 'html.parser')
        cards = soup.find_all('div', class_='stone_card')
        
        print(f"\n🔍 Знайдено товарів на сторінці: {len(cards)}")
        
        products = []
        for card in cards:
            try:
                link = card.find('a', class_='info')
                product_url = link['href'] if link else None
                
                # Пропускаємо вже оброблені товари
                if product_url in self.processed_urls:
                    continue
                
                # Назва товару
                title = card.find('p', class_='stone_name')
                title_text = title.get('title', '') if title else ''
                
                # Бренд
                brand = card.find('p', class_='stone_company')
                brand_text = brand.text.strip() if brand else ''
                
                # Картинка
                img = card.find('img', class_='stone_cover')
                img_url = img['src'] if img else ''
                
                # Тип поверхні
                surface_type = card.find('div', class_='additional-info__title')
                surface_text = surface_type.find('span').text.strip() if surface_type and surface_type.find('span') else ''
                
                products.append({
                    'url': product_url,
                    'title': title_text,
                    'brand': brand_text,
                    'feature_photo': img_url,
                    'type': surface_text
                })
                
            except Exception as e:
                print(f"⚠️ Помилка обробки картки: {e}")
                continue
        
        new_products = len(products)
        print(f"✨ Нових товарів для обробки: {new_products}")
        
        return products
    
    def parse_product_detail(self, url, category_name):
        """Парсить детальну сторінку товару"""
        max_retries = 3
        
        for attempt in range(max_retries):
            try:
                self.driver.get(url)
                time.sleep(2)
                
                html = self.driver.page_source
                soup = BeautifulSoup(html, 'html.parser')
                
                # Код товару з h1
                h1 = soup.find('h1')
                code = h1.text.strip() if h1 else ''
                
                # Використовуємо передану категорію
                category = category_name
                
                # Галерея зображень
                gallery_images = []
                gallery = soup.find('div', class_='gellery_for')
                
                if gallery:
                    # Шукаємо всі зображення в слайдері
                    images = gallery.find_all('img', {'data-fancybox': 'gallery'})
                    
                    for img in images[:5]:  # Максимум 5 зображень
                        img_url = img.get('href') or img.get('src', '')
                        # Беремо великі зображення (1280)
                        if img_url and '1280' in img_url:
                            gallery_images.append(img_url)
                        elif img_url:
                            # Якщо немає 1280, намагаємось замінити розмір
                            img_url = img_url.replace('/320/', '/1280/').replace('/540/', '/1280/')
                            gallery_images.append(img_url)
                
                # Доповнюємо до 5 елементів порожніми значеннями
                while len(gallery_images) < 5:
                    gallery_images.append('')
                
                return {
                    'code': code,
                    'category': category,
                    'gallery': gallery_images[:5]
                }
                
            except WebDriverException as e:
                if attempt < max_retries - 1:
                    print(f"⚠️ Помилка з'єднання, спроба {attempt + 2}/{max_retries}...")
                    self.restart_driver()
                    time.sleep(2)
                else:
                    print(f"❌ Не вдалось обробити {url}: {e}")
                    return {
                        'code': '',
                        'category': '',
                        'gallery': ['', '', '', '', '']
                    }
            except Exception as e:
                print(f"⚠️ Помилка обробки {url}: {e}")
                return {
                    'code': '',
                    'category': '',
                    'gallery': ['', '', '', '', '']
                }
    
    def save_product_to_excel(self, product_data):
        """Додає один товар до Excel файлу"""
        try:
            print(f"💾 Збереження товару: {product_data.get('Title', 'Без назви')}")
            
            # Якщо файл існує, дописуємо до нього
            if os.path.exists(self.output_file):
                print(f"   📂 Файл існує, дописуємо...")
                # Використовуємо openpyxl для швидкого дописування
                from openpyxl import load_workbook
                
                wb = load_workbook(self.output_file)
                ws = wb['Products']
                
                # Додаємо новий рядок
                ws.append([
                    product_data['Brand'],
                    product_data['Category'],
                    product_data['Title'],
                    product_data['Code'],
                    product_data['Feature photo'],
                    product_data['Type'],
                    product_data['Gallery1'],
                    product_data['Gallery2'],
                    product_data['Gallery3'],
                    product_data['Gallery4'],
                    product_data['Gallery5']
                ])
                
                wb.save(self.output_file)
                wb.close()
                print(f"   ✅ Збережено!")
            else:
                print(f"   🆕 Створюємо новий файл...")
                # Створюємо новий файл з заголовками
                df = pd.DataFrame([product_data])
                with pd.ExcelWriter(self.output_file, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Products')
                    
                    # Налаштовуємо ширину колонок
                    worksheet = writer.sheets['Products']
                    column_widths = {
                        'A': 20, 'B': 25, 'C': 30, 'D': 20, 'E': 50,
                        'F': 15, 'G': 50, 'H': 50, 'I': 50, 'J': 50, 'K': 50
                    }
                    for col, width in column_widths.items():
                        worksheet.column_dimensions[col].width = width
                print(f"   ✅ Файл створено: {self.output_file}")
                        
        except Exception as e:
            print(f"❌ Помилка збереження в Excel: {e}")
            import traceback
            traceback.print_exc()
    
    def parse_all(self, categories):
        """Основна функція парсингу"""
        print("🚀 Початок парсингу topovi.com.ua\n")
        
        # Обробляємо кожну категорію
        for category_name, category_url in categories.items():
            print(f"\n{'='*60}")
            print(f"📂 Обробка категорії: {category_name}")
            print(f"🔗 URL: {category_url}")
            print(f"{'='*60}\n")
            
            # Завантажуємо всі товари з категорії
            html = self.load_all_products(category_url)
            
            # Парсимо список товарів
            products = self.parse_product_list(html)
            
            if not products:
                print(f"✅ Немає нових товарів у категорії '{category_name}'")
                continue
            
            # Обробляємо кожен товар
            print(f"\n📦 Обробка детальних сторінок товарів з категорії '{category_name}'...")
            
            for product in tqdm(products, desc=f"{category_name}"):
                if product['url'] and product['url'] not in self.processed_urls:
                    try:
                        print(f"\n🔍 Обробка: {product['title']}")
                        details = self.parse_product_detail(product['url'], category_name)
                        
                        # Формуємо дані товару
                        product_data = {
                            'Brand': product['brand'],
                            'Category': details['category'],
                            'Title': product['title'],
                            'Code': details['code'],
                            'Feature photo': product['feature_photo'],
                            'Type': product['type'],
                            'Gallery1': details['gallery'][0],
                            'Gallery2': details['gallery'][1],
                            'Gallery3': details['gallery'][2],
                            'Gallery4': details['gallery'][3],
                            'Gallery5': details['gallery'][4],
                        }
                        
                        print(f"   📋 Дані зібрано: Brand={product_data['Brand']}, Code={product_data['Code']}")
                        
                        # Зберігаємо товар у файл
                        self.save_product_to_excel(product_data)
                        
                        # Додаємо URL до оброблених
                        self.processed_urls.add(product['url'])
                        
                        # Зберігаємо прогрес кожні 10 товарів
                        if len(self.processed_urls) % 10 == 0:
                            self.save_progress()
                        
                        # Невелика затримка між запитами
                        time.sleep(0.5)
                        
                    except Exception as e:
                        print(f"\n❌ Критична помилка при обробці товару: {e}")
                        self.save_progress()
                        continue
        
        # Фінальне збереження прогресу
        self.save_progress()
        
        print(f"\n✅ Парсинг завершено!")
        print(f"📊 Всього оброблено товарів: {len(self.processed_urls)}")
        print(f"💾 Файл збережено: {self.output_file}")
    
    def close(self):
        """Закриває браузер"""
        if self.driver:
            self.driver.quit()
    
    def reset_progress(self):
        """Скидає прогрес (для повторного парсингу)"""
        if os.path.exists(self.progress_file):
            os.remove(self.progress_file)
        if os.path.exists(self.output_file):
            os.remove(self.output_file)
        self.processed_urls = set()
        print("🔄 Прогрес скинуто")


def main():
    """Головна функція"""
    parser = TopoviParser(output_file='topovi_products.xlsx')
    
    try:
        # Категорії для парсингу
        categories = {
            'Кварцовий камінь': 'https://topovi.com.ua/stones/types=kvarcevyy-kamen',
            'Натуральний камінь': 'https://topovi.com.ua/stones/types=naturalniy-kamin',
            'Акриловий камінь': 'https://topovi.com.ua/stones/types=akrilovyy-kamen'
        }
        
        # Запускаємо парсинг
        parser.parse_all(categories)
        
    except KeyboardInterrupt:
        print("\n\n⏸️ Парсинг зупинено користувачем")
        print("💾 Прогрес збережено. Для продовження запустіть скрипт знову")
    except Exception as e:
        print(f"\n❌ Критична помилка: {e}")
        print("💾 Прогрес збережено. Для продовження запустіть скрипт знову")
    finally:
        parser.save_progress()
        parser.close()


if __name__ == "__main__":
    main()