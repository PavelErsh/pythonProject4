from telebot import types
from aiogram import Bot, Dispatcher, executor, types
from aiogram.dispatcher.filters import Text
import main
import os
import reader
import csv_writer


#t.me/slkjfs_bot
bot = Bot(token='5098342584:AAHvrg9gcKUCrnSl8qpW4fIPC21TkYVtzgY')

TOKEN = '5098342584:AAHvrg9gcKUCrnSl8qpW4fIPC21TkYVtzgY'


dp = Dispatcher(bot)



@dp.message_handler(commands="start")
async def cmd_start(message: types.Message):
    keyboard = types.ReplyKeyboardMarkup()
    button_1 = types.KeyboardButton(text="Скачать Excel документ")
    keyboard.add(button_1)
    await message.answer("Хочешь скачать Excel документ c текущими играми?", reply_markup=keyboard)

@dp.message_handler(Text(equals="Скачать Excel документ"))
async def with_puree(message: types.Message):
    data = main.get_data()
    print(os.getcwd())
    reader.crate_book()
    csv_writer.games_create_csv()
    cid = message.from_user.id
    duck = open('my_book.xlsx', 'rb')
    await bot.send_document(cid, duck)




if __name__ == "__main__":
    # Запуск бота
    executor.start_polling(dp, skip_updates=True)