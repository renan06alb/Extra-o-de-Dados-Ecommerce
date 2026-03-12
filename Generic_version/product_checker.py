import requests
import urllib.parse
from openpyxl import Workbook


def buscar_produto(sku):
    
    sku_formatado = urllib.parse.quote(sku)

    url = f"https://www.example.com/api/catalog_system/pub/products/search?ft={sku_formatado}"

    try:
        response = requests.get(url)

        if response.status_code != 200:
            return None

        data = response.json()

        if not data:
            return None

        produto = data[0]
        item = produto["items"][0]
        seller = item["sellers"][0]["commertialOffer"]

        nome = produto["productName"]
        preco = seller["Price"]
        preco_original = seller["ListPrice"]
        estoque = seller["AvailableQuantity"]

        status_estoque = "In Stock" if estoque > 0 else "Out of Stock"

        return {
            "name": nome,
            "price": preco,
            "list_price": preco_original,
            "stock": status_estoque
        }

    except:
        return None


def gerar_relatorio(skus):

    wb = Workbook()
    ws = wb.active
    ws.title = "Products"

    ws.append([
        "SKU",
        "Product",
        "Price",
        "Original Price",
        "Stock Status"
    ])

    for sku in skus:

        produto = buscar_produto(sku)

        if produto:

            ws.append([
                sku,
                produto["name"],
                produto["price"],
                produto["list_price"],
                produto["stock"]
            ])

        else:

            ws.append([
                sku,
                "Not Found",
                "-",
                "-",
                "-"
            ])

    wb.save("product_report.xlsx")


def main():

    with open("skus.txt", "r") as file:
        skus = [line.strip() for line in file.readlines()]

    gerar_relatorio(skus)

    print("Report generated: product_report.xlsx")


if __name__ == "__main__":
    main()