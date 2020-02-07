import time
from selenium.webdriver import Chrome
from selenium.webdriver.support.ui import Select
import requests
import base64

# Локальный путь к веб драйверу хрома
path_webdriver = r'D:\chromedriver.exe'

# апи кей
API_KEY = ''

# Функция получения данных с сайта http://fssprus.ru/
def start_selenium_for_fssp(last_name, first_name, middle_name, birthday):
    browser = Chrome(path_webdriver)
    # URL сайта с которого будет брать информацию
    browser.get('https://fssprus.ru')
    # У меня сразу открывлась окно, которое закрывала основной сайт, это окно я закрываю
    browser.find_element_by_class_name('tingle-modal__close').click()
    # нажимаю кнопку расширенный поиск, чтобы появились поля для ввода данных
    browser.find_element_by_class_name('main-form__toggle-open').click()
    # ввожу данные
    browser.find_element_by_name('is[last_name]').send_keys(last_name)
    browser.find_element_by_name('is[first_name]').send_keys(first_name)
    browser.find_element_by_name('is[patronymic]').send_keys(middle_name)
    browser.find_element_by_name('is[date]').send_keys(birthday)
    # Нажимаю поиск
    browser.find_element_by_class_name('main-form__btns').find_element_by_class_name('btn-primary').click()
    # Делаю паузу, чтобы сайт успел прогрузиться
    time.sleep(3)

    # функция отправки изображения капчи стороннему сервис для распознования и ввод полученных данных
    def capcha():
        time.sleep(3)
        # находим изображение капчи
        x = browser.find_element_by_id('capchaVisual')
        # получаем кодировку base64 изображения
        image_code_base64 = x.get_attribute('src')
        image_code_base64 = image_code_base64.replace('data:image/jpeg;base64,','')
        # декодируем изображение
        dec = base64.b64decode(image_code_base64)
        # Сохраняем изображение
        filename = 'some_image.jpg'
        with open(filename, 'wb') as f:
            f.write(dec)

        # создем цикл, который будет отправялть запрос, до тех пор пока не получим подтвержения получения изображения
        get_task_id = False
        while not get_task_id:

            r = requests.post('https://rucaptcha.com/in.php',
                              data={'key': API_KEY, 'method': 'post', 'lang': 'ru', 'json': 1},
                              files={'file': open(filename, 'rb')}
                              )
            # Если получили подтвержение, то выходим из цикла
            if r.json()['status'] == 1:
                get_task_id = True
            # Если сервис перегружен, ждем 5 сек и пробуем заново
            else:
                time.sleep(5)

        capcha_ready = False
        # создаем цикл, который будет получать ответ
        while not capcha_ready:
            r1 = requests.get(
                'https://rucaptcha.com/res.php?key={}&action={}&id={}&json=1'.format(
                    API_KEY, 'get',
                    r.json()['request'])
            )
            # Если капча еще не разгадана, ждем 3 сек и пробуем еще раз
            if r1.json()['request'] == 'CAPCHA_NOT_READY':
                time.sleep(3)
            # если разгадана, то выходим из цикла
            else:
                capcha_ready = True

        # получаем текст капчи
        text_for_capcha = r1.json()['request']
        # вводим текст
        browser.find_element_by_id("captcha-popup-code").send_keys(text_for_capcha)
        # нажимаем отправить капчу
        browser.find_element_by_id("ncapcha-submit").click()

        time.sleep(3)

    capcha_passed = False
    # создаю цикл, который запускает функцию отправки капчи и ввод ответа
    while not capcha_passed:
        # Если капча по какой то причине не пройдена, то функция запускается заново
        try:
            if browser.find_element_by_class_name('popup-wrapper'):
                capcha()
        # Если капча успешно пройдена, цикл заканчивается
        except:
            capcha_passed = True

    # Я создал список с заголовками исполнительных производств
    # и в этот же список буду добавять найденые исполнительные производства
    list_all_enforcements_proceeding = [
        ['Должник (физ. лицо: ФИО, дата и место рождения; юр. лицо: наименование, юр. адрес)',
         'Исполнительное производство (номер, дата возбуждения)',
         'Реквизиты исполнительного документа (вид, дата принятия органом, номер, наименование органа, '
         'выдавшего исполнительный документ)',
         'Дата, причина окончания или прекращения ИП (статья, часть, пункт основания)',
         'Сервис',
         'Предмет исполнения, сумма непогашенной задолженности',
         'Отдел судебных приставов (наименование, адрес)',
         'Судебный пристав-исполнитель'
         ]
    ]

    # Если исполнительных производств нет то завершаем проверку данного физ.лица
    try:
        browser.find_element_by_class_name('b-search-message__text').text
        # закрываем браузер
        browser.close()
        return True
    # Если исполнительные производства есть, то запишеи их в список и вернем его
    except:
        # Обработка таблицы и получение строк с исполнительными производствами
        table = browser.find_element_by_class_name('iss')
        all_rows = table.find_elements_by_tag_name('tr')
        for row in all_rows:
            list_enforcement_proceeding = []
            for block in row.find_elements_by_tag_name('td'):
                list_enforcement_proceeding.append(block.text)
            list_all_enforcements_proceeding.append(list_enforcement_proceeding)
        browser.close()
        return list_all_enforcements_proceeding

# Функция получения данных с сайта http://sudrf.ru/
def start_selenium_for_sudrf(last_name, first_name, middle_name):
    # Формирую ФИО
    fio = last_name + ' ' + first_name[0] + '.' + middle_name[0] + '.'
    browser = Chrome(path_webdriver)
    browser.get('https://sudrf.ru/index.php?id=300#sp')
    # Выбираю Субъект Российской Федерации
    select_russian_region = browser.find_element_by_id('spSearchArea')
    select_russian_region = select_russian_region.find_element_by_id('court_subj')
    # 77 это москва
    Select(select_russian_region).select_by_value('77')
    # ввожу ФИО
    browser.find_element_by_id('f_name').send_keys(fio)
    # Нажимаю кнопку поиска
    main_area = browser.find_element_by_id('spSearchArea')
    main_area.find_element_by_class_name('form-button').click()
    # Жду 10 секунд, чтобы сайт успел загрузить данные
    time.sleep(10)
    # Если судебные дела найдены, то добавлю их в список и возвращаю его
    try:
        table = browser.find_element_by_id('resultTable')
        all_rows = table.find_elements_by_tag_name('tr')
        list_all_court_case = []

        for row in all_rows:
            list_court_case = []
            for block in row.find_elements_by_tag_name('td'):
                list_court_case.append(block.text)
            list_all_court_case.append(list_court_case)
        browser.close()
        return list_all_court_case
    # Если судебных дел нет, то возвращаю True
    except:
        block_result = main_area.find_element_by_id('resulfs')
        block_result.find_elements_by_tag_name('b')
        browser.close()
        return True
