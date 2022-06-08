from telegram.ext import MessageHandler, Updater, Filters, CommandHandler, ConversationHandler
from telegram import ReplyKeyboardMarkup, ReplyKeyboardRemove, InputFile
import xlrd
from openpyxl import Workbook
import csv
import os


def start_keyboard():
    reply_keyboard = [['Excel->CSV', 'CSV->Excel'],
                      ['PDF->WORD', 'WORD->PDF']]
    markup = ReplyKeyboardMarkup(reply_keyboard, one_time_keyboard=False)
    return markup


def find_delimiter(path):
    sniffer = csv.Sniffer()
    with open(path, encoding='utf8') as fp:
        delimiter = sniffer.sniff(fp.read(5000)).delimiter
    return delimiter


def get_file_info(update):
    format_file = update.message['document']['file_name'].split('.')[1]
    initial_name = update.message['document']['file_name'].split('.')[0]
    return format_file, initial_name


def start(update, context):
    update.message.reply_text(f'Здравствуйте, {update.message["chat"]["first_name"]}! Я умею конвертировать документы.')
    update.message.reply_text('Вот мои возможности:', reply_markup=start_keyboard())


def response(update, context):
    text = update.message.text
    if text == 'Excel->CSV':
        update.message.reply_text('Отправьте файл для преобразования:', reply_markup=ReplyKeyboardRemove())
        return 'EXCEL_TO_CSV'
    elif text == 'CSV->Excel':
        update.message.reply_text('Отправьте файл для преобразования:', reply_markup=ReplyKeyboardRemove())
        return 'CSV_TO_EXCEL'
    elif text == 'PDF->WORD':
        update.message.reply_text('Отправьте файл для преобразования:', reply_markup=ReplyKeyboardRemove())
        return 'PDF_TO_WORD'
    elif text == 'WORD->PDF':
        update.message.reply_text('Отправьте файл для преобразования:', reply_markup=ReplyKeyboardRemove())
        return 'WORD_TO_PDF'


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
        writer = csv.writer(csvfile, delimiter=';', quotechar='"', quoting=csv.QUOTE_MINIMAL)
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


def csv_to_excel(update, context):
    chat_id = update.message['chat']['id']
    format_file = update.message['document']['file_name'].split('.')[1]
    initial_name = update.message['document']['file_name'].split('.')[0]
    result_csv_file = f'data/{chat_id}.csv'
    with open(result_csv_file, 'wb') as file:
        context.bot.get_file(update.message['document']['file_id']).download(out=file)
        file.close()
    delimiter = find_delimiter(result_csv_file)
    result_excel_file = f'data/{chat_id}.xlsx'
    with open(result_csv_file, encoding='utf8') as csvfile:
        reader = csv.reader(csvfile, delimiter=delimiter, quotechar='"')
        wb = Workbook()
        ws = wb.active
        for row in reader:
            ws.append(row)
        wb.save(result_excel_file)
        csvfile.close()
    result_file = f'{initial_name}.xlsx'
    update.message.reply_text('Конвертация завершена', reply_markup=start_keyboard())
    context.bot.send_document(chat_id=chat_id, document=open(result_excel_file, mode='rb'), filename=result_file)
    os.remove(result_excel_file)
    os.remove(result_csv_file)


def pdf_to_word(update, context):
    chat_id = update.message['chat']['id']
    format_file = update.message['document']['file_name'].split('.')[1]
    initial_name = update.message['document']['file_name'].split('.')[0]
    result_pdf_file = f'data/{chat_id}.csv'
    with open(result_pdf_file, 'wb') as file:
        context.bot.get_file(update.message['document']['file_id']).download(out=file)
        file.close()


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
            'EXCEL_TO_CSV': [MessageHandler(Filters.document, excel_to_csv, pass_update_queue=True)],
            'CSV_TO_EXCEL': [MessageHandler(Filters.document, csv_to_excel, pass_user_data=True)],
            'PDF_TO_WORD': [MessageHandler(Filters.document, pdf_to_word, pass_user_data=True)],
            'WORD_TO_PDF': [MessageHandler(Filters.document, word_to_pdf, pass_user_data=True)]
        },
        fallbacks=[CommandHandler('stop', stop)]
    )
    dp.add_handler(conv_handler)
    updater.start_polling()
    updater.idle()


if __name__ == '__main__':
    main()
