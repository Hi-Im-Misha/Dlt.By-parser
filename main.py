import requests
from bs4 import BeautifulSoup
import xlwt
import os
from urllib.parse import urlparse
import re
import time
from openpyxl import Workbook
from openpyxl.utils import get_column_letter


BASE_URL = "https://dlt.by"

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
}



def get_all_categories():
    response = requests.get(BASE_URL, headers=headers)
    soup = BeautifulSoup(response.text, "html.parser")
    category_links = []
    
    for a_tag in soup.select(".menu-body a[href]"):
        link = a_tag["href"]
        if link.startswith("/"):
            link = BASE_URL + link
        category_links.append(link)
    
    return category_links

def get_product_links(category_url):
    response = requests.get(category_url, headers=headers)
    soup = BeautifulSoup(response.text, "html.parser")
    product_links = []
    
    for div in soup.find_all("div", class_="product-thumb"):
        a_tag = div.find("a", href=True)
        if a_tag:
            product_links.append(a_tag["href"])
    
    return product_links

def parse_product_page(url):
    response = requests.get(url, headers=headers)
    soup = BeautifulSoup(response.text, "html.parser")

    title = soup.select_one("h1.product-header").text.strip() if soup.select_one("h1.product-header") else "Нет заголовка"
    
    price_tag = soup.select_one(".autocalc-product-price")
    price = price_tag.text.strip() if price_tag else "Нет цены"
    
    manufacturer = soup.select_one("li.product-info-li")
    manufacturer = manufacturer.text.replace("Производитель:", "").strip() if manufacturer else "Нет производителя"
   
    article = soup.select_one("li.main-product-model strong#main-product-model")
    article = article.text.strip() if article else "Нет артикула"

    description_div = soup.select_one("div#tab-description")
    description = description_div.get_text(separator=" ", strip=True) if description_div else "Нет описания"

    image_urls = []
    
    image_tags = soup.select('div#image-additional a[href]')
    image_urls = [a['href'] for a in image_tags if a['href'].endswith('.jpg')]

    return {
        "title": title,
        "price": price,
        "manufacturer": manufacturer,
        "article": article,
        "description": description,
        "image_urls": image_urls,
        "url": url
    }




def clean_folder_name(name):
    return re.sub(r'[\\/:"*?<>|]+', '_', name)

def save_to_xlsx(products, file_path):
    folder = os.path.dirname(file_path)
    photos_root_dir = os.path.join(folder, "photos")
    os.makedirs(photos_root_dir, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "Products"

    excel_headers = [
        "Название", "Цена", "Производитель", "Артикул", "Описание", "Ссылка",
        "Фото (ссылки)", "Фото (локально)"
    ]
    ws.append(excel_headers)

    for product in products:
        image_urls = product.get("image_urls", [])

        clean_title = clean_folder_name(product["title"])
        product_photo_dir = os.path.join(photos_root_dir, clean_title)
        os.makedirs(product_photo_dir, exist_ok=True)

        saved_image_paths = []
        for i, img_url in enumerate(image_urls):
            try:
                img_data = requests.get(img_url, headers=headers).content
                img_name = f"{product['article'] or 'product'}_{i + 1}.jpg"
                img_path = os.path.join(product_photo_dir, img_name)
                with open(img_path, "wb") as f:
                    f.write(img_data)
                saved_image_paths.append(os.path.join(clean_title, img_name))
            except Exception as e:
                print(f"Не удалось сохранить изображение {img_url}: {e}")

        row_data = [
            product["title"],
            product["price"],
            product["manufacturer"],
            product["article"],
            product["description"],
            product["url"],
            ", ".join(image_urls) if image_urls else "Нет ссылок",
            ", ".join(saved_image_paths) if saved_image_paths else "Нет фото"
        ]
        ws.append(row_data)

    wb.save(file_path)
    print(f"Данные и фото сохранены в {file_path}")





def generate_filename_from_url(url):
    path = urlparse(url).path
    name = path.strip("/").split("/")[-1] or "products"
    return name



def main():
    url = 'https://dlt.by/tiling-tool/plitkorezi-dlja-krupnoformatnoj-p/'
    folder_name = generate_filename_from_url(url)
    output_dir = os.path.join("products", folder_name)
    output_file = os.path.join(output_dir, f"{folder_name}.xlsx")

    all_products = []

    try:
        if url == BASE_URL:
            categories = get_all_categories()
        else:
            categories = [url]

        for category in categories:
            product_links = get_product_links(category)
            for link in product_links:
                full_link = link if link.startswith("http") else BASE_URL + link
                product_data = parse_product_page(full_link)
                all_products.append(product_data)

    except Exception as e:
        print(f"Произошла ошибка: {e}")

    finally:
        if all_products:
            time.sleep(1)
            save_to_xlsx(all_products, output_file)
        else:
            print("Нет данных для сохранения.")



if __name__ == "__main__":
    main()
