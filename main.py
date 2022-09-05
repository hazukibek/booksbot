import os
import telebot
import logging
import openpyxl
from config import *
from flask import Flask, request
from telebot import types

bot = telebot.TeleBot(BOT_TOKEN)
server = Flask(__name__)
logger = telebot.logger
logger.setLevel(logging.DEBUG)

path = 'C:\Users\Asus\PycharmProjects\myplate\users.xlsx'
wb_obj = openpyxl.load_workbook(path)
sheet_obj = wb_obj.active


@bot.message_handler(commands=['start'])
def send_message(message):
    bot.send_message(message.chat.id, "Привет!\n /juzkitap- книги 100 китап \n /NISbooks -учебники НИШа ")


@bot.message_handler(commands=['juzkitap'])
def send_message(message):
    bot.send_message(message.chat.id,
                     "_Зияткерлік мектеп оқушылары оқуы тиіс 100 кітап аясында қазақ әдебиетінен ұсынылатын шығармалар тізімі_ \n_Негізгі мектеп_ \n1. Жетім \n2. Қорғансыздың  күні\n3. Сіз бен Біз\n4. Ақан сері\n5. Ғажайып планета\n6. Аманай мен Заманай\n7. Алғашқы айлар \n8. Тағдырдың кейбір кездері\n9. Бақбақ басы толған күн\n10. Жау тылындағы бала\n11. Мен  апамның баласымын\n12. Мен қарапайым қарттарды сүйем\n13. Үш бақытым\n14. Жар жағалаған қыз\n15. «Мен қазақпын» поэмасы\n16. Интернатта болған жағдаят\n17. Ұшқан ұя\n18. Жабайы алма\n19. Соғыстың соңғы жесірі\n20. «Еңбек бірлігі» әңгімесі\n21. Жапон балладасы\n22. Ананың анасы\n23. Қазақ солдаты\n24. Қызыл кітап\n25. Бір өкініш, бір үміт\n26. Ақын өлімі туралы аңыз\n27. Ақиқат пен аңыз\n28. Он бес жыл өткен соң\n 29. Балалық шағың\n30. Бір атаның балалары\n31. Менің атым Қожа\n32. Балалық шаққа саяхат\n33. Ауыл шетіндегі үй\n34. Өмір-өзен\n35. Өркениеттің адасуы\n36. Құз басындағы аңшының зары\n37. Ләйлі-Мәжнүн\n38. Қоңыр күз еді \n39. Жақсы мен жаман туралы\n40. Жаяу Мұса\n_Жоғарғы мектеп_\n41. Қара сөздер\n 42. Бір тойым бар\n43. Ақбілек \n44. Ғасырдан да ұзақ күн\n45. Қаһарлы күндер\n46. Ақбоз ат\n47. Өліара\n48. «Қырық мысал»\n49. «Елім-ай» трилогиясы\n50. Атау кере\n51. Өз отыңды өшірме\n52. Жаңғақ\n53. Ақ боз үй\n54. Өмір мектебі\n55. Сасырдың сүті\n56. Сәйгүліктер\n57. Аңыздың ақыры\n58. Сары қазақ\n59. Ай мен Айша\n60. Махаббат қызық мол жылдар\n_Список рекомендуемой художественной литературы для чтения в основной и старшей школах  Назарбаев Интеллектуальных школ на русском языке в рамках проекта «100 книг»_\n_Основная школа_\n61. «Ночевала тучка золотая» \n62. «Оруженосец Кашка» \n63. «Матерь человеческая» \n64. «Чучело» \n65. «Человек амфибия» \n66. «Белый Бим Черное ухо» \n67. «Парадокс» \n68. «Емшан» \n69. «Мой зеленоглазый аруах» \n70. Стихотворения о природе, родине: «Вечер», «Последний шмель», «Полевые цветы» и другие\n71. «Не позволяй душе лениться», «Некрасивая девочка» \n_Старшая школа_\n72. «Пиковая дама» \n73. «Отцы и дети» \n74. «Доктор Живаго» \n75. «Волоколамское шоссе» \n76. «Плаха» \n77. «Земля, поклонись человеку!» \n78. «Баллада о времени» \n79. «Хроника Великого джута» \n80. «Матренин двор» \n_Список рекомендуемой художественной литературы для чтения в основной и старшей школах Назарбаев Интеллектуальных школ на английском языке в рамках проекта «100 книг»_\n_Основная школа (рекомендуемые уровни А1, А2 и В1)_ \n81. Большие надежды/Great Expectations\n82. Приключения Тома Сойера/Adventures of Tom Sawyer\n83. Трое в лодке, не считая собаки/Three Men in a Boat (To Say Nothing of the Dog) \n84. Машина времени/The Time Machine\n85. Приключения Алисы в стране чудес/Alice in Wonderland\n86. Удивительный Волшебник из страны Оз / The Wonderful Wizard of Oz\n87. 20 тысяч лье под водой/20,000 Leagues Under the Sea \n88. Маленький принц/The Little Prince\n89. Зов предков/The Call of the Wild\n90. Последний из могикан\n_Старшая школа (рекомендуемые уровни В2, С1, С2, в оригинале)_ \n91. Джейн Эйр/Jane Eyre\n92. Три товарища/Three Comrades\n93. Гордость и предубеждение/Pride and Prejudice\n94. Дэвид Копперфильд/David Copperfield\n95. Пигмалион/Pygmalion  \n96. Старик и море/The Old Man and the Sea\n97. Ромео и Джульетта/Romeo and Juliette \n98. Портрет Дориана Грея / The Picture of Dorian Gray \n99. Оливер Твист/Oliver Twist\n100. Повелитель мух/ Lord of the Flies\n  ",
                     parse_mode="Markdown")


@bot.message_handler(commands=['NISbooks'])
def get_grade(message):
    markup_grade = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True, row_width=1)
    item_seven = types.KeyboardButton('7 класс')
    item_eight = types.KeyboardButton('8 класс')
    item_nine = types.KeyboardButton('9 класс')
    item_ten = types.KeyboardButton('10 класс')
    item_eleven = types.KeyboardButton('11 класс')
    item_twelve = types.KeyboardButton('12 класс')
    markup_grade.add(item_seven, item_eight, item_nine, item_ten, item_eleven, item_twelve)
    bot.send_message(message.chat.id, "Хорошо! Сначала выбери класс", reply_markup=markup_grade)


@bot.message_handler(content_types=['text'])
def get_book(message):
    global grade
    grade = message.text
    bot.reply_to(message, "Окей")
    markup_close = types.ReplyKeyboardRemove()
    bot.send_message(message.chat.id,
                     "Теперь напиши нужный учебник в таком формате: Қазақстан тарихы \n Математика-казахский \n Биология-русский \n То есть, _Название учебника-язык_, БЕЗ ПРОБЕЛОВ!!! (если учебник на одном языке напишите только название предмета) ",
                     reply_markup=markup_close)
    bot.register_next_step_handler(message, get_user_book)


@bot.message_handler(content_types=['text'])
def get_user_book(message):
    book = message.text.lower()
    path = open("/content/drive/MyDrive/NIS учебники/" + book + ".pdf", "rb")
    bot.send_message(message.chat.id, "Ищем книгу...")
    bot.send_document(message.chat.id, path)
    path.close()


@bot.message_handler(content_types=['text'])
def get_user_text(message):
    int(message.text)
    if message.text.lower() == "3":
        bot.send_message(message.chat.id, "Извините, этой книги нет в базе данных.")

    elif int(message.text) > 100 or int(message.text) < 1:
        bot.send_message(message.chat.id, "Выберите существующий номер книги (от 1 до 100)")

    else:
        for i in range(1, 101):
            if message.text.lower() == str(i):
                cell_obj = sheet_obj.cell(row=i, column=1)
                file_book = open("/content/drive/MyDrive/100-кітап/" + cell_obj.value + ".pdf", "rb")
                bot.send_message(message.chat.id, "Ищем книгу...")
                bot.send_document(message.chat.id, file_book)
                file_book.close()


@server.route(f"/{BOT_TOKEN}", methods=["POST"])
def redirect_message():
    json_string = request.get_data().decode("utf-8")
    update = telebot.types.Update.de_json(json_string)
    bot.process_new_updates([update])
    return "!", 200


if __name__ == "__main__":
    bot.remove_webhook()
    bot.set_webhook(url=APP_URL)
    server.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))