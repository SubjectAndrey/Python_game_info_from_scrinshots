import os
from datetime import datetime
from PIL import Image
import pytesseract
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import difflib

def extract_text_from_area(image, coordinates):  # Функция для извлечения текста из области скриншота
    cropped_image = image.crop(coordinates)
    text = pytesseract.image_to_string(cropped_image, lang='rus')  # Извлекает текст в строку на русском
    return text

# Указываем путь к JSON-файлу с учетными данными для доступа к Google Sheets API
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('buoyant-habitat-280609-7f79ebdef4ec.json', scope)
client = gspread.authorize(credentials)

# Открываем таблицу Google Sheets по ее ID
sheet = client.open_by_key('1y8l6BnzMtoAnqEmPIz5gSRtL_cZah1YmML-YvEesBCk')

# Получаем названия листов в таблице
sheet_names = [s.title for s in sheet.worksheets()]

# Координаты и размеры области для выделения текста
city_name = (24, 133, 200, 156)  # Левая часть экрана
all_money = (105, 185, 162, 204)
money = (166, 186, 213, 206)
culture = (124,233, 197, 256)
all_cityzens = (101, 282, 157, 302)
cityzens = (165, 281, 212, 303)
all_power = (103, 330, 151, 351)
power = (161, 331, 212, 352)
happynes = (164, 378, 213, 399)
all_science = (113, 428, 167, 447)
science = (172, 427, 213, 449)
trade = (84, 475, 121, 499)
speciality = (51, 532, 196, 551)
# Правая часть экрана
storage = (1485, 48, 1586, 66)
all_wood = (1477, 96, 1537, 116)
wood = (1533, 96, 1592, 116)
all_iron = (1477, 136, 1537, 155)
iron = (1533, 136, 1592, 157)
all_stone = (1477, 176, 1537, 196)
stone = (1533, 176, 1592, 196)
all_food = (1477, 216, 1537, 236)
food = (1533, 216, 1592, 236)
all_meat = (1477, 256, 1537, 276)
meat = (1533, 256, 1592, 276)
all_fish = (1477, 296, 1537, 316)
fish = (1533, 296, 1592, 316)
all_wheat = (1477, 336, 1537, 356)
wheat = (1533, 336, 1592, 356)
all_rice = (1477, 376, 1537, 396)
rice = (1533, 376, 1592, 396)
all_corn = (1477, 416, 1537, 436)
corn = (1533, 416, 1592, 436)
all_spice = (1477, 456, 1537, 476)
spice = (1533, 456, 1592, 476)
all_shirts = (1477, 496, 1537, 516)
shirts = (1533, 496, 1592, 516)
all_paintings = (1477, 536, 1537, 556)
paintings = (1533, 536, 1592, 556)
all_cakes = (1477, 576, 1537, 596)
cakes = (1533, 576, 1592, 596)
all_diamonds = (1477, 616, 1537, 636)
diamonds = (1533, 616, 1592, 636)
all_oxen = (1477, 656, 1537, 676)
oxen = (1533, 656, 1592, 676)

screenshots_folder = 'screenshots/'  # Указываем путь к папке со скриншотами

files = os.listdir(screenshots_folder)  # Получаем список файлов в папке

# Фильтруем только скриншоты, сделанные сегодня
today = datetime.now().strftime("%Y-%m-%d")
screenshots_today = [file for file in files if file.endswith('.png') and datetime.fromtimestamp(
    os.path.getctime(screenshots_folder + file)).strftime("%Y-%m-%d") == today]

screenshots_today.sort(key=lambda x: os.path.getctime(screenshots_folder + x),
                       reverse=True)  # Сортируем скриншоты по времени создания

# Обрабатываем только 3 последних скриншота, если их больше 4
for screenshot in screenshots_today[:4]:
    screenshot_path = os.path.join(screenshots_folder, screenshot)
    screenshot_image = Image.open(screenshot_path)

    # Извлечение текста из каждой области
    city_name_text = extract_text_from_area(screenshot_image, city_name)
    all_money_text = extract_text_from_area(screenshot_image, all_money)
    money_text = extract_text_from_area(screenshot_image, money)
    culture_text = extract_text_from_area(screenshot_image, culture)
    all_cityzens_text = extract_text_from_area(screenshot_image, all_cityzens)
    cityzens_text = extract_text_from_area(screenshot_image, cityzens)
    all_power_text = extract_text_from_area(screenshot_image, all_power)
    power_text = extract_text_from_area(screenshot_image, power)
    happynes_text = extract_text_from_area(screenshot_image, happynes)
    all_science_text = extract_text_from_area(screenshot_image, all_science)
    science_text = extract_text_from_area(screenshot_image, science)
    trade_text = extract_text_from_area(screenshot_image, trade)
    speciality_text = extract_text_from_area(screenshot_image, speciality)
    storage_text = extract_text_from_area(screenshot_image, storage)
    all_wood_text = extract_text_from_area(screenshot_image, all_wood)
    wood_text = extract_text_from_area(screenshot_image, wood)
    all_iron_text = extract_text_from_area(screenshot_image, all_iron)
    iron_text = extract_text_from_area(screenshot_image, iron)
    all_stone_text = extract_text_from_area(screenshot_image, all_stone)
    stone_text = extract_text_from_area(screenshot_image, stone)
    all_food_text = extract_text_from_area(screenshot_image, all_food)
    food_text = extract_text_from_area(screenshot_image, food)
    all_meat_text = extract_text_from_area(screenshot_image, all_meat)
    meat_text = extract_text_from_area(screenshot_image, meat)
    all_fish_text = extract_text_from_area(screenshot_image, all_fish)
    fish_text = extract_text_from_area(screenshot_image, fish)
    all_wheat_text = extract_text_from_area(screenshot_image, all_wheat)
    wheat_text = extract_text_from_area(screenshot_image, wheat)
    all_rice_text = extract_text_from_area(screenshot_image, all_rice)
    rice_text = extract_text_from_area(screenshot_image, rice)
    all_corn_text = extract_text_from_area(screenshot_image, all_corn)
    corn_text = extract_text_from_area(screenshot_image, corn)
    all_spice_text = extract_text_from_area(screenshot_image, all_spice)
    spice_text = extract_text_from_area(screenshot_image, spice)
    all_shirts_text = extract_text_from_area(screenshot_image, all_shirts)
    shirts_text = extract_text_from_area(screenshot_image, shirts)
    all_paintings_text = extract_text_from_area(screenshot_image, all_paintings)
    paintings_text = extract_text_from_area(screenshot_image, paintings)
    all_cakes_text = extract_text_from_area(screenshot_image, all_cakes)
    cakes_text = extract_text_from_area(screenshot_image, cakes)
    all_diamonds_text = extract_text_from_area(screenshot_image, all_diamonds)
    diamonds_text = extract_text_from_area(screenshot_image, diamonds)
    all_oxen_text = extract_text_from_area(screenshot_image, all_oxen)
    oxen_text = extract_text_from_area(screenshot_image, oxen)

    # Создание словаря с извлеченным текстом
    text_dict = {
        'city_name': city_name_text,
        'all_money': all_money_text,
        'money': money_text,
        'culture': culture_text,
        'all_cityzens': all_cityzens_text,
        'cityzens': cityzens_text,
        'all_power': all_power_text,
        'power': power_text,
        'happynes': happynes_text,
        'all_science': all_science_text,
        'science': science_text,
        'trade': trade_text,
        'speciality': speciality_text,
        'storage': storage_text,
        'all_wood': all_wood_text,
        'wood': wood_text,
        'all_iron': all_iron_text,
        'iron': iron_text,
        'all_stone': all_stone_text,
        'stone': stone_text,
        'all_food': all_food_text,
        'food': food_text,
        'all_meat': all_meat_text,
        'meat': meat_text,
        'all_fish': all_fish_text,
        'fish': fish_text,
        'all_wheat': all_wheat_text,
        'wheat': wheat_text,
        'all_rice': all_rice_text,
        'rice': rice_text,
        'all_corn': all_corn_text,
        'corn': corn_text,
        'all_spice': all_spice_text,
        'spice': spice_text,
        'all_shirts': all_shirts_text,
        'shirts': shirts_text,
        'all_paintings': all_paintings_text,
        'paintings': paintings_text,
        'all_cakes': all_cakes_text,
        'cakes': cakes_text,
        'all_diamonds': all_diamonds_text,
        'diamonds': diamonds_text,
        'all_oxen': all_oxen_text,
        'oxen': oxen_text
    }

    update_requests = []  # Создаем пустой список для запросов на обновление

    for sheet_name in sheet_names:
        similarity_ratio = difflib.SequenceMatcher(None, text_dict['city_name'], sheet_name).ratio()
        if similarity_ratio >= 0.4:
            worksheet = sheet.worksheet(sheet_name)
            update_request = {
                "updateCells": {
                    "rows": [
                        {
                            "values": [
                                {"userEnteredValue": {"stringValue": text_dict['city_name']}},
                                {"userEnteredValue": {"stringValue": ""}},
                            ]
                        },
                        {
                            "values": [
                                {"userEnteredValue": {"stringValue": text_dict['all_money']}},
                                {"userEnteredValue": {"stringValue": text_dict['money']}},
                            ]
                        },
                        {
                            "values": [
                                {"userEnteredValue": {"stringValue": text_dict['culture']}},
                                {"userEnteredValue": {"stringValue": ""}},
                            ]
                        },
                        {
                            "values": [
                                {"userEnteredValue": {"stringValue": text_dict['all_cityzens']}},
                                {"userEnteredValue": {"stringValue": text_dict['cityzens']}},
                            ]
                        },
                        {
                            "values": [
                                {"userEnteredValue": {"stringValue": text_dict['all_power']}},
                                {"userEnteredValue": {"stringValue": text_dict['power']}},
                            ]
                        },
                        {
                            "values": [
                                {"userEnteredValue": {"stringValue": text_dict['happynes']}},
                                {"userEnteredValue": {"stringValue": ""}},
                            ]
                        },
                        {
                            "values": [
                                {"userEnteredValue": {"stringValue": text_dict['all_science']}},
                                {"userEnteredValue": {"stringValue": text_dict['science']}},
                            ]
                        },
                        # Добавьте еще необходимые строки в таком же формате
                    ],
                    "fields": "*",
                    "range": {
                        "sheetId": worksheet.id,
                        "startRowIndex": 0,
                        "endRowIndex": 9,
                        "startColumnIndex": 1,
                        "endColumnIndex": 3
                    }
                }
            }
            update_requests.append(update_request)
            update_request = {
                "updateCells": {
                    "rows": [
                        {
                            "values": [
                                {"userEnteredValue": {"stringValue": text_dict['storage']}},
                                {"userEnteredValue": {"stringValue": ""}},
                            ]
                        },
                        {
                            "values": [
                                {"userEnteredValue": {"stringValue": text_dict['all_wood']}},
                                {"userEnteredValue": {"stringValue": text_dict['wood']}},
                            ]
                        },
                        {
                            "values": [
                                {"userEnteredValue": {"stringValue": text_dict['all_iron']}},
                                {"userEnteredValue": {"stringValue": text_dict['iron']}},
                            ]
                        },
                        {
                            "values": [
                                {"userEnteredValue": {"stringValue": text_dict['all_stone']}},
                                {"userEnteredValue": {"stringValue": text_dict['stone']}},
                            ]
                        },
                        {
                            "values": [
                                {"userEnteredValue": {"stringValue": text_dict['all_food']}},
                                {"userEnteredValue": {"stringValue": text_dict['food']}},
                            ]
                        },
                        {
                            "values": [
                                {"userEnteredValue": {"stringValue": text_dict['all_meat']}},
                                {"userEnteredValue": {"stringValue": text_dict['meat']}},
                            ]
                        },
                        {
                            "values": [
                                {"userEnteredValue": {"stringValue": text_dict['all_fish']}},
                                {"userEnteredValue": {"stringValue": text_dict['fish']}},
                            ]
                        },
                            {
                            "values": [
                                {"userEnteredValue": {"stringValue": text_dict['all_wheat']}},
                                {"userEnteredValue": {"stringValue": text_dict['wheat']}},
                            ]
                        },
                        {
                                "values": [
                                    {"userEnteredValue": {"stringValue": text_dict['all_rice']}},
                                    {"userEnteredValue": {"stringValue": text_dict['rice']}},
                                ]
                        },
                        {
                            "values": [
                                {"userEnteredValue": {"stringValue": text_dict['all_corn']}},
                                {"userEnteredValue": {"stringValue": text_dict['corn']}},
                            ]
                        },
                        {
                            "values": [
                                {"userEnteredValue": {"stringValue": text_dict['all_spice']}},
                                {"userEnteredValue": {"stringValue": text_dict['spice']}},
                            ]
                        },
                        {
                            "values": [
                                {"userEnteredValue": {"stringValue": text_dict['all_shirts']}},
                                {"userEnteredValue": {"stringValue": text_dict['shirts']}},
                            ]
                        },
                        {
                            "values": [
                                {"userEnteredValue": {"stringValue": text_dict['all_paintings']}},
                                {"userEnteredValue": {"stringValue": text_dict['paintings']}},
                            ]
                        },
                        {
                            "values": [
                                {"userEnteredValue": {"stringValue": text_dict['all_cakes']}},
                                {"userEnteredValue": {"stringValue": text_dict['cakes']}},
                            ]
                        },
                        {
                            "values": [
                                {"userEnteredValue": {"stringValue": text_dict['all_diamonds']}},
                                {"userEnteredValue": {"stringValue": text_dict['diamonds']}},
                            ]
                        },
                        {
                            "values": [
                                {"userEnteredValue": {"stringValue": text_dict['all_oxen']}},
                                {"userEnteredValue": {"stringValue": text_dict['oxen']}},
                            ]
                        },
                        # Добавьте еще необходимые строки в таком же формате
                    ],
                    "fields": "*",
                    "range": {
                        "sheetId": worksheet.id,
                        "startRowIndex": 1,
                        "endRowIndex": 17,
                        "startColumnIndex": 4,
                        "endColumnIndex": 6
                    }
                }
            }
            update_requests.append(update_request)  # Добавляем запрос на обновление в список

    print(f'Text from {screenshot}: {text_dict}')  # Вывод извлеченного текста для проверки

    response = sheet.batch_update({"requests": update_requests})  # Отправляем запрос на batch_update



