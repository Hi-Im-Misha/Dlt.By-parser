import requests
from bs4 import BeautifulSoup
import xlwt
import os

BASE_URL = "https://dlt.by"

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
}

OUTPUT_FILE = "products.xls"

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
    price = soup.select_one("h2#main-product-price .autocalc-product-price").text.strip() if soup.select_one("h2#main-product-price .autocalc-product-price") else "Нет цены"
    manufacturer = soup.select_one("li.product-info-li")
    manufacturer = manufacturer.text.replace("Производитель:", "").strip() if manufacturer else "Нет производителя"
    article = soup.select_one("li.main-product-model strong#main-product-model")
    article = article.text.strip() if article else "Нет артикула"
    description = soup.select_one("div#tab-description")
    description = description.text.strip() if description else "Нет описания"
    
    return {
        "title": title,
        "price": price,
        "manufacturer": manufacturer,
        "article": article,
        "description": description,
        "url": url
    }

def save_to_xls(products):
    wb = xlwt.Workbook()
    ws = wb.add_sheet("Products")

    headers = ["Название", "Цена", "Производитель", "Артикул", "Описание", "Ссылка"]
    for col, header in enumerate(headers):
        ws.write(0, col, header)

    for row, product in enumerate(products, start=1):
        ws.write(row, 0, product["title"])
        ws.write(row, 1, product["price"])
        ws.write(row, 2, product["manufacturer"])
        ws.write(row, 3, product["article"])
        ws.write(row, 4, product["description"])
        ws.write(row, 5, product["url"])

    wb.save(OUTPUT_FILE)
    print(f"Данные сохранены в {OUTPUT_FILE}")

def main():
    all_products = []
    
    try:
        categories = get_all_categories()
        for category in categories:
            product_links = get_product_links(category)
            for link in product_links:
                full_link = link if link.startswith("http") else BASE_URL + link
                product_data = parse_product_page(full_link)
                all_products.append(product_data)
                print(product_data)

    except Exception as e:
        print(f"Произошла ошибка: {e}")
    
    finally:
        save_to_xls(all_products)

if __name__ == "__main__":
    main()
