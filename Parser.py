import time

import openpyxl
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# Список всех категорий

categories = [
    # Жінки
    {'id': '2353713', 'text': 'шкіра', 'url': 'https://www.zara.com/ua/ua/uawoman-leather-l1174.html',
     'parent': '1248446'},
    {'id': '2353511', 'text': 'swimwear', 'url': 'https://www.zara.com/ua/ua/woman-beachwear-l1052.html',
     'parent': '1248446'},
    {'id': '2352722', 'text': 'куртки', 'url': 'https://www.zara.com/ua/ua/woman-jackets-l1114.html',
     'parent': '1248446'},
    {'id': '2352686', 'text': 'блейзери', 'url': 'https://www.zara.com/ua/ua/woman-blazers-l1055.html',
     'parent': '1248446'},
    {'id': '2352822', 'text': 'сукні', 'url': 'https://www.zara.com/ua/ua/woman-dresses-l1066.html',
     'parent': '1248446'},
    {'id': '2353252', 'text': 'спідниці', 'url': 'https://www.zara.com/ua/ua/woman-skirts-l1299.html',
     'parent': '1248446'},
    {'id': '2353213', 'text': 'джинси', 'url': 'https://www.zara.com/ua/ua/woman-jeans-l1119.html',
     'parent': '1248446'},
    {'id': '2353142', 'text': 'штани', 'url': 'https://www.zara.com/ua/ua/woman-trousers-l1335.html',
     'parent': '1248446'},
    {'id': '2352909', 'text': 'сорочки', 'url': 'https://www.zara.com/ua/ua/woman-shirts-l1217.html',
     'parent': '1248446'},
    {'id': '2353010', 'text': 'топи | боді', 'url': 'https://www.zara.com/ua/ua/woman-tops-l1322.html',
     'parent': '1248446'},
    {'id': '2352954', 'text': 'футболки', 'url': 'https://www.zara.com/ua/ua/woman-tshirts-l1362.html',
     'parent': '1248446'},
    {'id': '2353417', 'text': 'взуття', 'url': 'https://www.zara.com/ua/ua/woman-shoes-l1251.html',
     'parent': '1248446'},
    {'id': '2352648', 'text': 'пальта-тренчі | пальта', 'url': 'https://www.zara.com/ua/ua/woman-outerwear-l1184.html',
     'parent': '1248446'},
    {'id': '2352851', 'text': 'кардигани | светри', 'url': 'https://www.zara.com/ua/ua/woman-knitwear-l1152.html',
     'parent': '1248446'},
    {'id': '2352737', 'text': 'жилетки', 'url': 'https://www.zara.com/ua/ua/woman-outerwear-vests-l1204.html',
     'parent': '1248446'},
    {'id': '2353281', 'text': 'шорти', 'url': 'https://www.zara.com/ua/ua/woman-trousers-shorts-l1355.html',
     'parent': '1248446'},
    {'id': '2353038', 'text': 'трикотаж', 'url': 'https://www.zara.com/ua/ua/woman-knitwear-l1152.html',
     'parent': '1248446'},
    {'id': '2353444', 'text': 'сумки', 'url': 'https://www.zara.com/ua/ua/woman-bags-l1024.html?v1=2353495',
     'parent': '1248446'},
    {'id': '2353545', 'text': 'аксесуари | біжутерія', 'url': 'https://www.zara.com/ua/ua/woman-accessories-l1003.html',
     'parent': '1248446'},
    {'id': '2353556', 'text': 'спідня білизна | піжами', 'url': 'https://www.zara.com/ua/ua/woman-lingerie-l4021.html',
     'parent': '1248446'},
    {'id': '2353578', 'text': 'beauty', 'url': 'https://www.zara.com/ua/ua/woman-beauty-makeup-l4414.html',
     'parent': '1248446'},
    {'id': '2353301', 'text': 'co-ord sets', 'url': 'https://www.zara.com/ua/ua/woman-co-ords-l1061.html',
     'parent': '1248446'},
    {'id': '2353284', 'text': 'базові моделі', 'url': 'https://www.zara.com/ua/ua/woman-basics-l1050.html',
     'parent': '1248446'},
    {'id': '2354009', 'text': 'домашній одяг', 'url': 'https://www.zara.com/ua/ua/woman-loungewear-l3519.html',
     'parent': '1248446'},
    {'id': '2351428', 'text': 'худi', 'url': 'https://www.zara.com/ua/ua/woman-sweatshirts-l1320.html',
     'parent': '1248446'},
    # Чоловіки
    {'id': '2379238', 'text': 'верхній одяг', 'url': 'https://www.zara.com/ua/ua/man-outerwear-l715.html',
     'parent': '1248450'},
    {'id': '2351277', 'text': 'штани', 'url': 'https://www.zara.com/ua/ua/man-trousers-l838.html',
     'parent': '1248450'},
    {'id': '2351396', 'text': 'джинси', 'url': 'https://www.zara.com/ua/ua/man-jeans-l659.html',
     'parent': '1248450'},
    {'id': '2351498', 'text': 'светри | кардигани', 'url': 'https://www.zara.com/ua/ua/man-knitwear-l681.html',
     'parent': '1248450'},
    {'id': '2351428', 'text': 'худi', 'url': 'https://www.zara.com/ua/ua/man-sweatshirts-l821.html',
     'parent': '1248450'},
    {'id': '2351541', 'text': 'футболки', 'url': 'https://www.zara.com/ua/ua/man-tshirts-l855.html',
     'parent': '1248450'},
    {'id': '2351641', 'text': 'верхні сорочки', 'url': 'https://www.zara.com/ua/ua/man-overshirts-l3174.html',
     'parent': '1248450'},
    {'id': '2351463', 'text': 'сорочки', 'url': 'https://www.zara.com/ua/ua/man-shirts-l737.html',
     'parent': '1248450'},
    {'id': '2351623', 'text': 'футболки-поло', 'url': 'https://www.zara.com/ua/ua/man-polos-l733.html',
     'parent': '1248450'},
    {'id': '2351785', 'text': 'шорти-бермуди', 'url': 'https://www.zara.com/ua/ua/man-bermudas-l592.html',
     'parent': '1248450'},
    {'id': '2351607', 'text': 'блейзери', 'url': 'https://www.zara.com/ua/ua/man-blazers-l608.html',
     'parent': '1248450'},
    {'id': '2351760', 'text': 'cargo', 'url': 'https://www.zara.com/ua/ua/man-trousers-cargo-l1780.html',
     'parent': '1248450'},
    {'id': '2351573', 'text': 'спортивні костюми', 'url': 'https://www.zara.com/ua/ua/man-jogging-l679.html',
     'parent': '1248450'},
    {'id': '2352261', 'text': 'взуття', 'url': 'https://www.zara.com/ua/uan/man-shoes-l769.html',
     'parent': '1248450'},
    {'id': '2352295', 'text': 'сумки | рюкзаки', 'url': 'https://www.zara.com/ua/ua/man-bags-l563.html',
     'parent': '1248450'},
    {'id': '2352353', 'text': 'аксесуари', 'url': 'https://www.zara.com/ua/ua/man-accessories-l537.html',
     'parent': '1248450'},
    {'id': '2352496', 'text': 'макіяж', 'url': 'https://www.zara.com/ua/ua/man-beauty-l4622.html',
     'parent': '1248450'},
    # Діти 1
    {'id': '12484531', 'text': 'Пальта',
     'url': 'https://www.zara.com/ua/uk/dity-maliuky-divchata-verkhnii-odiah-l131.html',
     'parent': '1248461'},
    {'id': '12484532', 'text': 'Сукні та комбінезони'
        , 'url': 'https://www.zara.com/ua/uk/dity-maliuky-divchata-sukni-l108.html',
     'parent': '1248461'},
    {'id': '12484533', 'text': 'Футболки',
     'url': 'https://www.zara.com/ua/uk/dity-maliuky-divchata-futbolky-l162.html',
     'parent': '1248461'},
    {'id': '12484534', 'text': 'Сорочки',
     'url': 'https://www.zara.com/ua/uk/dity-maliuky-divchata-sorochky-l133.html',
     'parent': '1248461'},
    {'id': '12484535', 'text': 'Трикотаж',
     'url': 'https://www.zara.com/ua/uk/dity-maliuky-divchata-trykotazh-l122.html?v1=2356664',
     'parent': '1248461'},
    {'id': '12484536', 'text': 'Взуття',
     'url': 'https://www.zara.com/ua/uk/dity-maliuky-divchata-vzuttia-l136.html?v1=2358279',
     'parent': '1248461'},
    # Діти 2
    {'id': '12484551', 'text': 'Пальта',
     'url': 'https://www.zara.com/ua/uk/dity-divchata-verkhnii-odiah-l394.html?v1=2357257',
     'parent': '1248462'},
    {'id': '12484552', 'text': 'Сукні',
     'url': 'https://www.zara.com/ua/uk/dity-divchata-sukni-l360.html?v1=2357826',
     'parent': '1248462'},
    {'id': '12484553', 'text': 'футболки',
     'url': 'https://www.zara.com/ua/uk/dity-divchata-futbolky-l450.html?v1=2357483',
     'parent': '1248462'},
    {'id': '12484554', 'text': 'Сорочки',
     'url': 'https://www.zara.com/ua/uk/dity-divchata-sorochky-l401.html?v1=2357462',
     'parent': '1248462'},
    {'id': '12484555', 'text': 'Трикотаж',
     'url': 'https://www.zara.com/ua/uk/dity-divchata-trykotazh-l385.html?v1=2357704',
     'parent': '1248462'},
    {'id': '12484556', 'text': 'взуття',
     'url': 'https://www.zara.com/ua/uk/dity-divchata-vzuttia-l404.html?v1=2357876',
     'parent': '1248462'},
    # Макіяж
    {'id': '23717321', 'text': 'Макіяж для жінок',
     'url': 'https://www.zara.com/ua/uk/dity-divchata-verkhnii-odiah-l394.html?v1=2357257',
     'parent': '2371732'},
    {'id': '23717322', 'text': 'Макіяж для чоловіків',
     'url': 'https://www.zara.com/ua/uk/man-beauty-l4622.html?v1=2352520',
     'parent': '2371732'},
]

# Создаем новый файл Excel и добавляем рабочий лист
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Export Products Sheet"

# Указываем заголовки для столбцов
headers = ['Код товара',  # 1
           'Url',  # 2_a
           'Url_ua_site',  # 2_b
           'Название_позиции',  # 3
           'Название_позиции_укр',  # 4
           'Поисковые_запросы',  # 5
           'Поисковые_запросы_укр',  # 6
           'Описание',  # 7
           'Описание_укр',  # 8
           'Тип_товара',  # 9
           'Цена',  # 10
           'Валюта',  # 11
           'Единица_измерения',  # 12
           'Минимальный_объем_заказа',  # 13
           'Оптовая_цена',  # 14
           'Минимальный_заказ_опт',  # 15
           'Ссылка_изображения',  # 16
           'Наличие',  # 17
           'Количество',  # 18
           'Номер_группы',  # 19
           'Название_группы',  # 20
           'Адрес_подраздела',  # 21
           'Возможность_поставки',  # 22
           'Срок_поставки',  # 23
           'Способ_упаковки',  # 24
           'Способ_упаковки_укр',  # 25
           'Уникальный_идентификатор',  # 26
           'Идентификатор_товара',  # 27
           'Идентификатор_подраздела',  # 28
           'Идентификатор_группы',  # 29
           'Производитель',  # 30
           'Страна_производитель',  # 31
           'Скидка',  # 32
           'ID_группы_разновидностей',  # 33
           'Личные_заметки',  # 34
           'Продукт_на_сайте',  # 35
           'Cрок действия скидки от',  # 36
           'Cрок действия скидки до',  # 37
           'Цена от',  # 38
           'Ярлык',  # 39
           'HTML_заголовок',  # 40
           'HTML_заголовок_укр',  # 41
           'HTML_описание',  # 42
           'HTML_описание_укр',  # 43
           'HTML_ключевые_слова',  # 44
           'HTML_ключевые_слова_укр',  # 45
           'Вес,кг',  # 46
           'Ширина,см',  # 47
           'Высота,см',  # 48
           'Длина,см',  # 49
           'Где_находится_товар',  # 50
           'Код_маркировки_(GTIN)',  # 51
           'Номер_устройства_(MPN)',  # 52
           'Название_Характеристики',  # 53
           'Измерение_Характеристики',  # 54
           'Значение_Характеристики',  # 55
           'Название_Характеристики',  # 56
           'Измерение_Характеристики',  # 57
           'Значение_Характеристики',  # 58
           'Название_Характеристики',  # 59
           'Измерение_Характеристики',  # 60
           'Значение_Характеристики',  # 61
           ]

ws.append(headers)

ws2 = wb.create_sheet(title="Export Groups Sheet")

# Указываем заголовки для столбцов
headers = ['Номер_группы',  # 1
           'Название_группы',  # 2
           'Название_группы_укр',  # 3
           'Идентификатор_группы',  # 4
           'Номер_родителя',  # 5
           'Идентификатор_родителя',  # 6
           'HTML_заголовок_группы',  # 7
           'HTML_заголовок_группы_укр',  # 8
           'HTML_описание_группы',  # 9
           'HTML_описание_группы_укр',  # 10
           'HTML_ключевые_слова_группы',  # 11
           'HTML_ключевые_слова_группы_укр',  # 12
           ]

ws2.append(headers)

ws2.append(['', '', 'Zara', '107107107', '', ''])
ws2.append(['', '', 'Жінки', '1248446', '', '107107107'])
ws2.append(['', '', 'Чоловіки', '1248450', '', '107107107'])
ws2.append(['', '', 'Діти', '1248452', '', '107107107'])
ws2.append(['', '', 'Дівчатка', '1248453', '', '1248452'])
ws2.append(['', '', '1,5-6 років', '1248461', '', '1248453'])
ws2.append(['', '', '6-14 років', '1248462', '', '1248453'])
ws2.append(['', '', 'Парфуми', '2371732', '', '107107107'])

# Настройки веб-драйвера
service = Service('chromedriver.exe')
driver = webdriver.Chrome(service=service)

for category in categories:

    category_id = category['id']
    category_text = category['text']
    category_url = category['url']
    parent_category_id = category['parent']

    product_url = ''
    changefreq = ''

    # Загружаем страницу категорий
    print("Parsing category: ", category_text)
    print("Parsing category name: ", category_text)
    driver.get(category_url)

    time.sleep(2)

    buttons = driver.find_elements(By.CSS_SELECTOR, '.view-option-selector .view-option-selector-button')

    if len(buttons) >= 3:
        buttons[2].click()

    time.sleep(2)

    # Прокручиваем страницу до самого низа
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

    time.sleep(2)
    # Прокручиваем страницу до самого низа 2
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

    time.sleep(2)
    # Прокручиваем страницу до самого низа 2
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

    time.sleep(2)

    products = driver.find_elements(By.CLASS_NAME, 'product-link')
    product_urls = set()  # Используем множество для хранения уникальных URL-адресов

    for product in products:
        try:
            product_url = product.get_attribute('href')
            if product_url:  # Проверяем, что ссылка не пустая
                product_urls.add(product_url)  # Добавляем только уникальные ссылки в множество
        except StaleElementReferenceException:
            print("Stale element exception occurred. Skipping this product.")
            continue

    product_urls = list(product_urls)  # Преобразуем множество обратно в список, если это необходимо

    # Счетчики строк
    current_row = 2
    current_category_row = 10
    added_categories = []

    product_counter = 0

    for product_url_ua in product_urls:

        product_counter += 1

        print("Pause", product_url_ua)
        time.sleep(1)

        # Переходим на страницу товара
        driver.get(product_url_ua)

        try:
            name_element = driver.find_element(By.CLASS_NAME, 'product-detail-info__header-name')
            name = name_element.text.strip()
            print("Name is ", name)

            # Получаем id товара
            product_id = driver.execute_script('''
                const elementWithId = document.querySelector('[id^=product-]');
                if (elementWithId) {
                    return elementWithId.getAttribute('id').split('-')[1];
                } else {
                    return null;
                }
            ''')
            print("Product ID is ", product_id)

            # Получаем описание товара
            description_element = driver.find_element(By.CSS_SELECTOR, '.product-detail-description p')
            description_text = description_element.text.strip()

            # Получаем изображения товара
            images = driver.execute_script('''
                 const scriptElementImages = document.querySelector('script[type="application/ld+json"]');

                    let images = '';
                    if (scriptElementImages) {
                        const jsonData = JSON.parse(scriptElementImages.textContent);
                        const uniqueImages = new Set();

                        jsonData.forEach(product => {
                            const images = product.image;
                            images.slice(0, 10).forEach(image => {
                                uniqueImages.add(image);
                            });
                        });

                        const imagesArray = Array.from(uniqueImages); // Преобразуем Set в массив
                        images = imagesArray.join(', '); // Соединяем элементы массива через запятую
                        return images;
                    }

                ''')
            print("Images list ", images)

            # Получаем Brand name
            brand_name = 'ZARA'

            # Получаем размеры
            sizes = driver.execute_script('''
                    const sizeList = document.querySelectorAll('.size-selector-list__item .product-size-info__main-label');
                    const sizes = [];

                    sizeList.forEach(item => {
                        sizes.push(item.textContent.trim());
                    });

                    const sizesString = sizes.join(',');
                    return sizesString;
                  ''')

            wait = WebDriverWait(driver, 10)
            categories = wait.until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, '.layout-categories-category--highlighted')))

            en_product_url = product_url_ua.replace("ua/uk/", "es/en/")

            driver.get(en_product_url)

            try:
                price_element = driver.find_element(By.CLASS_NAME, 'money-amount__main')
                price = price_element.text.strip().replace('EUR', '')
                print("Product is ", product_url)
                print("Price is ", price)
            except NoSuchElementException:
                print("Element price not found for. Skip this product ", product_url)
                continue  # Пропускаем этот товар и переходим к следующему

            # PAGE 2

            # Если есть категории, записываем данные во второй лист Excel
            # if categories:
            #     # Запись данных на второй лист Excel
            #     ws2.append([categories[0]['id'], categories[1]['text']])

        except NoSuchElementException:
            print("Some field not found", product_url_ua)
            continue  # Пропускаем этот товар и переходим к следующему

        # sys.exit()

        # Разбиваем строку размеров по запятым, чтобы получить отдельные размеры
        individual_sizes = sizes.split(',')

        # Проверяем количество размеров
        if len(individual_sizes) <= 1:  # Если размеров нет или только один
            ws.append([
                product_id,  # 1
                en_product_url,  # 2_a
                product_url_ua,  # 2_b
                f'=GOOGLETRANSLATE(E{current_row},"UK","RU")',  # 3
                name,  # 4
                '',  # 5
                '',  # 6
                f'=GOOGLETRANSLATE(I{current_row},"UK","RU")',  # 7 description
                description_text,  # 8
                '',  # 9
                price,  # 10
                'EUR',  # 11
                'шт.',  # 12
                '',  # 13
                '',  # 14
                '',  # 15
                images,  # 16
                '+',  # 17'
                '',  # 18
                '',  # 19
                '',  # 20
                '',  # 21
                '',  # 22
                '',  # 23
                '',  # 24
                '',  # 25
                product_id,  # 26
                '',  # 26
                '',  # 26
                category_id,  # 29 ИДЕНТИФИКАТОР ГРУППЫ
                brand_name,  # 30
                '',  # 31
                '',  # 32
                '',  # 33
                '',  # 34
                '',  # 35
                '',  # 36
                '',  # 37
                '',  # 38
                '',  # 39
                '',  # 40
                '',  # 41
                '',  # 42
                '',  # 43
                '',  # 44
                '',  # 45
                '',  # 46
                '',  # 47
                '',  # 48
                '',  # 49
                '',  # 50
                '',  # 51
                '',  # 52
                'Размер',  # 53,
                '',  # 54
                sizes,  # 55
            ])

            current_row += 1  # увеличиваем current_row для следующего товара
        else:  # Если размеров больше одного

            groupPrefix = 1

            # Добавляем каждый размер в отдельной строке
            for individual_size in individual_sizes:
                ws.append([
                    ''.join([str(product_id), str(groupPrefix)]),  # 1
                    en_product_url,  # 2_a
                    product_url_ua,  # 2_b
                    f'=GOOGLETRANSLATE(E{current_row},"UK","RU")',  # 3
                    name,  # 4
                    '',  # 5
                    '',  # 6
                    f'=GOOGLETRANSLATE(I{current_row},"UK","RU")',  # 7 description
                    description_text,  # 8
                    '',  # 9
                    price,  # 10
                    'EUR',  # 11
                    'шт.',  # 12
                    '',  # 13
                    '',  # 14
                    '',  # 15
                    images,  # 16
                    '+',  # 17'
                    '',  # 18
                    '',  # 19
                    '',  # 20
                    '',  # 21
                    '',  # 22
                    '',  # 23
                    '',  # 24
                    '',  # 25
                    ''.join([str(product_id), str(groupPrefix)]),  # 26
                    '',  # 26
                    '',  # 26
                    category_id,  # 29 ИДЕНТИФИКАТОР ГРУППЫ
                    brand_name,  # 30
                    '',  # 31
                    '',  # 32
                    product_id,  # 33 id группы разновидностей
                    '',  # 34
                    '',  # 35
                    '',  # 36
                    '',  # 37
                    '',  # 38
                    '',  # 39
                    '',  # 40
                    '',  # 41
                    '',  # 42
                    '',  # 43
                    '',  # 44
                    '',  # 45
                    '',  # 46
                    '',  # 47
                    '',  # 48
                    '',  # 49
                    '',  # 50
                    '',  # 51
                    '',  # 52
                    'Размер',  # 53
                    '',  # 54
                    individual_size,  # 55
                ])

                groupPrefix += 1

        # Проверяем, есть ли текущая категория в списке уже добавленных категорий
        if category_text not in added_categories:
            # Если категория ещё не добавлена, добавляем её в лист Excel и в список добавленных категорий
            ws2.append([
                '',  # 1
                f'=GOOGLETRANSLATE(C{current_row},"UK","RU")',  # 2
                category_text,  # 3
                category_id,  # 4
                '',  # 5
                parent_category_id,  # 6
            ])
            added_categories.append(category_text)  # Добавляем категорию в список добавленных категорий
            current_category_row += 1

        current_row += 1

        print("Data written for", product_url)

        if product_counter % 10 == 0:  # Save the Excel file every 10 products
            wb.save('output.xlsx')
            print(f"Saved data for {product_counter} products.")

# Save the Excel file after processing all products
wb.save('output.xlsx')

# Закрываем веб-драйвер и сохраняем файл Excel
driver.quit()
wb.close()

print("Data saved to output.xlsx")
