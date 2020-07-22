import requests
import xlsxwriter
import re

from bs4 import BeautifulSoup

def parse_page(file_name):

    workbook = xlsxwriter.Workbook(f'storage/{file_name}.xlsx')
    worksheet = workbook.add_worksheet()

    worksheet.write('A1', 'Наименование')
    worksheet.write('B1', 'Цена')
    worksheet.write('C1', 'URL фото 1')
    worksheet.write('D1', 'URL фото 2')
    worksheet.write('E1', 'URL фото 3')
    worksheet.write('F1', 'URL фото 4')
    worksheet.write('G1', 'URL фото 5')
    worksheet.write('H1', 'URL фото 6')
    worksheet.write('I1', 'Размеры')
    worksheet.write('J1', 'Вес')
    worksheet.write('K1', 'Описание')

    row_marker = 2

    for i in range(1, 12):

        url = f'https://gazoncity.ru/catalog/?PAGEN_1={i}'

        base_response = requests.get(url)

        base_soup = BeautifulSoup(base_response.text, 'html.parser')

        items = base_soup.find_all('div', {'class': 'item'})

        links = [item.find('div', {'class': 'title'}).find('a').get('href') for item in items \
                if item.find('div', {'class': 'title'}).find('a').get('href') \
                not in ['/company/', '/catalog/', '/services/', '/info/', '/contacts/']]

        for link in links:

            response = requests.get(f'https://gazoncity.ru{link}')

            soup = BeautifulSoup(response.text, 'html.parser')

            name = soup.find('h1', {'id': 'pagetitle'})

            if name != None:
                worksheet.write(f'A{row_marker}', name.text)
                weight = re.search(r"\(([А-Яа-я0-9 ]+)\)", name.text)
                if weight != None and 'кг' in weight.group(1):
                    worksheet.write(f'J{row_marker}', weight.group(1))

            price = soup.find('span', {'itemprop': 'price'})

            if price != None:
                worksheet.write(f'B{row_marker}', price.text)

            related_images = soup.find_all('img', {'title': name.text}, src=True)

            images_urls = [image.get('src') for image in related_images]

            i = 67
            for url in images_urls:
                if i <= 72:
                    worksheet.write(f'{chr(i)}{row_marker}', url)
                    i += 1

            props_table = soup.find('table', {'class': 'props_table'})

            dim_res = [0, 0, 0]
            weight_value = 0

            for item in props_table.find_all('span'):
                
                if 'Длина' in item.text:
                    parent = item.find_parent('tr')
                    length = parent.find('td', {'class': 'char_value'})
                    dim_res[0] = length.text.strip('\n\n\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t')
                if 'Ширина' in item.text:
                    parent = item.find_parent('tr')
                    width = parent.find('td', {'class': 'char_value'})
                    dim_res[1] = width.text.strip('\n\n\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t')
                if 'Высота' in item.text:
                    parent = item.find_parent('tr')
                    heigth = parent.find('td', {'class': 'char_value'})
                    dim_res[2] = heigth.text.strip('\n\n\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t')

            worksheet.write(f'I{row_marker}', f'{dim_res[0]} * {dim_res[1]} * {dim_res[2]}')

            description = soup.find('div', {'class': 'previewtext', 'itemprop': 'description'})
            description_text = soup.find('p')

            if description_text != None:

                worksheet.write(f'K{row_marker}', description_text.text.strip('\r\n\t'))

            row_marker += 1

    workbook.close()