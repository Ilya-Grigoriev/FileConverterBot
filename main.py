from telegram.ext import MessageHandler, Updater, Filters, CommandHandler, ConversationHandler
from telegram import ReplyKeyboardMarkup, ReplyKeyboardRemove, InputFile
import xlrd
import csv
import os


def start_keyboard():
    reply_keyboard = [['Перевод из Excel в CSV', 'Перевод из CSV в Excel']]
    markup = ReplyKeyboardMarkup(reply_keyboard, one_time_keyboard=False)
    return markup


def start(update, context):
    update.message.reply_text(f'Здравствуйте, {update.message["chat"]["first_name"]}! Я умею конвертировать документы.')
    update.message.reply_text('Вот мои возможности:', reply_markup=start_keyboard())


def response(update, context):
    text = update.message.text
    if text == 'Перевод из Excel в CSV':
        update.message.reply_text('Отправьте файл для преобразования:', reply_markup=ReplyKeyboardRemove())
        return 'EXCEL_TO_CSV'
    elif text == 'Перевод из CSV в Excel':
        return 'CSV_TO_EXCEL'


def excel_to_csv(update, context):
    chat_id = update.message['chat']['id']
    format_file = update.message['document']['file_name'].split('.')[1]
    initial_name = update.message['document']['file_name'].split('.')[0]
    result_excel_file = f'data/{chat_id}.{format_file}'
    with open(result_excel_file, 'wb') as file:
        context.bot.get_file(update.message['document']['file_id']).download(out=file)
        file.close()
    result_csv_file = f'data/{chat_id}.csv'
    with open(result_csv_file, mode='w', newline='') as csvfile:
        writer = csv.writer(
            csvfile, delimiter=';', quotechar='"', quoting=csv.QUOTE_MINIMAL)
        book = xlrd.open_workbook(result_excel_file)
        for sheet_number in range(book.nsheets):
            sh = book.sheet_by_index(sheet_number)
            for rx in range(sh.nrows):
                writer.writerow(sh.row_values(rx))
        csvfile.close()
    result_file = f'{initial_name}.csv'
    update.message.reply_text('Конвертация завершена', reply_markup=start_keyboard())
    context.bot.send_document(chat_id=chat_id, document=open(result_csv_file, mode='rb'), filename=result_file)
    os.remove(result_excel_file)
    os.remove(result_csv_file)


# def csv_to_excel(update, context):
#     chat_id = update.message['chat']['id']
#     format_file = update.message['document']['file_name'].split('.')[1]
#     initial_name = update.message['document']['file_name'].split('.')[0]
#     result_csv_file = f'data/{chat_id}.{format_file}'


def stop(update, context):
    update.message.reply_text('Программа завершена', reply_markup=start_keyboard())
    return ConversationHandler.END


def main():
    token = '677970032:AAEJifhRsPjJG2luEgAvQ7Q9pwX8IG9VQ8I'
    updater = Updater(token)
    dp = updater.dispatcher
    dp.add_handler(CommandHandler('start', start, pass_user_data=True))
    conv_handler = ConversationHandler(
        entry_points=[MessageHandler(Filters.text & (~ Filters.command), response)],
        states={
            'EXCEL_TO_CSV': [MessageHandler(Filters.document, excel_to_csv)],
            'CSV_TO_EXCEL': [MessageHandler(Filters.document, csv_to_excel)]
        },
        fallbacks=[CommandHandler('stop', stop)]
    )

    dp.add_handler(conv_handler)
    updater.start_polling()
    updater.idle()


if __name__ == '__main__':
    main()
