# Документация библиотеки vk_api: https://github.com/python273/vk_api
# Официальная документация VK API по разделу сообщений: https://vk.com/dev/messages
# Получить токен: https://vkhost.github.io/
#
# import vk_api  # использование VK API
# from vk_api.utils import get_random_id  # снижение количества повторных отправок сообщения
# from dotenv import load_dotenv  # загрузка информации из .env-файла
# import os  # работа с файловой системой
# from openpyxl import Workbook
# import requests
# from PIL import Image
# from openpyxl.drawing.image import Image as Im
# import random


class Bot:
    """
    Базовый класс бота ВКонтакте
    """

    # текущая сессия ВКонтакте
    vk_session = None

    # доступ к API ВКонтакте
    vk_api_access = None

    # пометка авторизованности
    authorized = False

    # id пользователя ВКонтакте (например, 1234567890) в виде строки
    # можно использовать, если диалог будет вестись только с конкретным человеком
    default_user_id = None

    def __init__(self):
        """
        Инициализация бота при помощи получения доступа к API ВКонтакте
        """
        # загрузка информации из .env-файла
        load_dotenv()

        # авторизация
        self.vk_api_access = self.do_auth()

        if self.vk_api_access is not None:
            self.authorized = True

        # получение id пользователя из файла настроек окружения .env в виде строки USER_ID="1234567890"
        self.default_user_id = os.getenv("USER_ID")

    def do_auth(self):
        """
        Авторизация за пользователя (не за группу или приложение)
        Использует переменную, хранящуюся в файле настроек окружения .env в виде строки ACCESS_TOKEN="1q2w3e4r5t6y7u8i9o..."
        :return: возможность работать с API
        """
        token = os.getenv("ACCESS_TOKEN")
        try:
            self.vk_session = vk_api.VkApi(token=token)
            return self.vk_session.get_api()
        except Exception as error:
            print(error)
            return None

    def send_message(self, receiver_user_id=None, message_text="тестовое сообщение"):
        """
        Отправка сообщения от лица авторизованного пользователя
        :param receiver_user_id: уникальный идентификатор получателя сообщения
        :param message_text: текст отправляемого сообщения
        """
        if not self.authorized:
            print("Unauthorized. Check if ACCESS_TOKEN is valid")
            return

        # если не указан ID - берём значение по умолчанию, если таковое указано в .env-файле
        if receiver_user_id is None:
            receiver_user_id = self.default_user_id

        try:
            a = self.vk_api_access.messages.getHistoryAttachments(peer_id=receiver_user_id, media_type='photo',
                                                                  count=100)
            a = a['items']
            wb = Workbook()
            ws = wb.active
            count = 1
            count2 = 14

            for i in range(len(a)):
                max_screen = 500
                min_screen = 256
                for j in range(len(a[i]['attachment']['photo']['sizes'])):
                    if max_screen > a[i]['attachment']['photo']['sizes'][j]['height'] > min_screen:
                        screen = a[i]['attachment']['photo']['sizes'][j]['height']
                        screen_url = a[i]['attachment']['photo']['sizes'][j]['url']
                if screen_url == None:
                    screen_url = 'https://media-cldnry.s-nbcnews.com/image/upload/rockcms/2021-11/211116-harry-potter-al-1232-b41548.jpg'
                url_from_vk = requests.get(screen_url).content
                with open(f'image_from_vk{i}.jpg', 'wb') as handler:
                    handler.write(url_from_vk)

                img = Image.open(f"image_from_vk{i}.jpg")
                img_small = img.resize((256, 256))
                img_small.save(f'image_from_vk{i}.jpg')

                list_for_kama = ['1234', '12345']
                random_photo = random.choice(list_for_kama)
                random_photo1 = random.choice(list_for_kama)
                img = Im(f'image_from_vk{i}.jpg')
                img2 = Im(f'{random_photo}.jpg')
                img3 = Im(f'{random_photo1}.jpg')
                img4 = Im(f'image_from_vk{i}.jpg')

                ws.add_image(img, f'A{count}')
                ws[f'A{count2}'] = f'СУХТП говно[{i}]'
                ws.add_image(img2, f'Q{count}')
                ws[f'Q{count2}'] = f'А где здесь карш??[{i}]'

                count += 13
                count2 += 13
                ws.add_image(img3, f'I{count}')
                ws[f'I{count2}'] = f'Еду в москоу на чилле[{i}]'
                ws.add_image(img4, f'Y{count}')
                ws[f'Y{count2}'] = f'СУХТП говно[{i}]'

                count += 13
                count2 += 13

            wb.save('logoKama.xlsx')

            # self.vk_api_access.messages.send(user_id=receiver_user_id, message=message_text, random_id=get_random_id())
            print(f"Сообщение отправлено для ID {receiver_user_id} с текстом: {message_text}")
        except Exception as error:
            print(error)
