from aiogram import types
from aiogram.dispatcher import FSMContext
from aiogram.dispatcher.filters import Text
from aiogram.dispatcher.filters.state import StatesGroup, State
from aiogram.types import chat

from main import bot, dp, Database


class Balance(StatesGroup):
    Check = State()
    Amount = State()

@dp.message_handler(Text(equels="Баланс"))
async def check_balance(message: types.Message, state: FSMContext):
    async with state.proxy() as data_storage:
        money = Database.get_balance(user_id=message.from_user.id)
        data_storage["money"] = money
        await bot.send_message(message.chat.id, 'Ваш балланс = ' + str(money) + " USD")
        await Balance.Check.set()

