from aiogram import types
from aiogram.utils.callback_data import CallbackData

callback_numbers = CallbackData("fabnum", "action")

menu_markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
balance = types.KeyboardButton("Баланс")
kod = types.KeyboardButton("Заказать прогноз")
about = types.KeyboardButton("О боте")
menu_markup.add(balance, kod, about)

back_cancel_markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
back = types.KeyboardButton("Назад")
cancel = types.KeyboardButton("Отмена")
back_cancel_markup.add(back, cancel)

















customer_menu_markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
balance = types.KeyboardButton("Баланс")
create_order = types.KeyboardButton("Создать заказ")
find_talent = types.KeyboardButton("Поиск исполнителей")
executor_menu = types.KeyboardButton("Перейти в меню исполнителя")
watch_order_review = types.KeyboardButton("Посмотреть отклики на заказ")
customer_menu_markup.row(balance, create_order, find_talent)
customer_menu_markup.row(watch_order_review)
customer_menu_markup.row(executor_menu)

executor_menu_markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
create_questionnaire = types.KeyboardButton("Создать анкету")
find_job = types.KeyboardButton("Поиск заказов")
customer_menu = types.KeyboardButton("Перейти в меню заказчика")
watch_tz = types.KeyboardButton("Посмотреть отклики на анкету")
executor_menu_markup.row(balance, create_questionnaire, find_job)
executor_menu_markup.row(watch_tz)
executor_menu_markup.row(customer_menu)

registration_markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
customer_register = types.KeyboardButton("Зарегистрироваться как заказчик")
executor_register = types.KeyboardButton("Зарегистрироваться как исполнитель")
registration_markup.add(customer_register, executor_register)

balance_markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
add_balance = types.KeyboardButton("Пополнить баланс")
withdraw = types.KeyboardButton("Вывод средств")

balance_markup.add(add_balance, withdraw, back)

back_cancel_markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
cancel = types.KeyboardButton("Отмена")
back_cancel_markup.add(back, cancel)

back_cancel_remove_markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
remove = types.KeyboardButton("Удалить")
back_cancel_remove_markup.add(back, remove, cancel)

tz_markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
skip = types.KeyboardButton("Пропустить")
tz_markup.add(back, skip, cancel)

amount_balance = types.ReplyKeyboardMarkup(resize_keyboard=True)
item5 = types.KeyboardButton("5")
item10 = types.KeyboardButton("10")
item25 = types.KeyboardButton("25")
item50 = types.KeyboardButton("50")
item100 = types.KeyboardButton("100")
amount_balance.add(item5, item10, item25, item50, item100, cancel)

confirm_payment_markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
confirm_payment = types.KeyboardButton("Проверить оплату")
confirm_payment_markup.add(cancel, confirm_payment)

switch_orders = types.InlineKeyboardMarkup()
back_inline = types.InlineKeyboardButton(text="<", callback_data=callback_numbers.new(action="back"))
approve = types.InlineKeyboardButton(text="Подтвердить", callback_data=callback_numbers.new(action="approve"))
forward_inline = types.InlineKeyboardButton(text=">", callback_data=callback_numbers.new(action="forvard"))
switch_orders.add(back_inline, approve, forward_inline)

category_markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
design = types.KeyboardButton("Дизайн")
it = types.KeyboardButton("Разработка и IT")
text = types.KeyboardButton("Работа с текстом и перевод")
seo = types.KeyboardButton("SEO и трафик")
marketing = types.KeyboardButton("Соцсети и реклама")
audio = types.KeyboardButton("Аудио, видео, съёмка")
business = types.KeyboardButton("Бизнесс и быт")
gdz = types.KeyboardButton("Решебник")
category_markup.add(design, it, text, seo, marketing, audio, business,gdz)
category_markup.row(back,cancel)

tz_markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
skip = types.KeyboardButton("Пропустить")
tz_markup.add(back, skip, cancel)
