from telegram.ext import MessageHandler, Updater, Filters, CommandHandler, ConversationHandler
from telegram import ReplyKeyboardMarkup, ReplyKeyboardRemove, InputFile
from pywintypes import com_error
import pythoncom
import traceback
import openpyxl
from pdf2docx import Converter
from docx2pdf import convert
import csv
import os


def start_keyboard():
    reply_keyboard = [['Excel->CSV', 'CSV->Excel'],
                      ['PDF->DOCX', 'DOCX->PDF']]
    markup = ReplyKeyboardMarkup(reply_keyboard, one_time_keyboard=False)
    return markup


def find_delimiter(path: str) -> str:
    sniffer = csv.Sniffer()
    with open(path, encoding='utf8') as fp:
        delimiter = sniffer.sniff(fp.read(5000)).delimiter
    return delimiter


def get_file_info(update) -> tuple:
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
    elif text == 'PDF->DOCX':
        update.message.reply_text('Отправьте файл для преобразования:', reply_markup=ReplyKeyboardRemove())
        return 'PDF_TO_DOCX'
    elif text == 'DOCX->PDF':
        update.message.reply_text('Отправьте файл для преобразования:', reply_markup=ReplyKeyboardRemove())
        return 'DOCX_TO_PDF'


def excel_to_csv(update, context):
    result_csv_file = ''
    result_excel_file = ''
    try:
        expan = update.message['document']['file_name'].split('.')[1]
        if expan not in ['xlsx', 'xlsm', 'xltx', 'xltm']:
            update.message.reply_text('Неверный формат для конвертации', reply_markup=start_keyboard())
            return ConversationHandler.END
        chat_id = update.message['chat']['id']
        format_file, initial_name = get_file_info(update)
        result_excel_file = f'data/{chat_id}.{format_file}'
        with open(result_excel_file, 'wb') as file:
            context.bot.get_file(update.message['document']['file_id']).download(out=file)
            file.close()
        result_csv_file = f'data/{chat_id}.csv'
        wb = openpyxl.load_workbook(result_excel_file)
        sheet = wb.active
        with open(result_csv_file, mode='w', newline='', encoding='utf8') as csvfile:
            writer = csv.writer(csvfile, delimiter=';', quotechar='"', quoting=csv.QUOTE_MINIMAL)
            for row in sheet.rows:
                writer.writerow([col.value for col in row])
            csvfile.close()
            wb.close()
        result_file = f'{initial_name}.csv'
        update.message.reply_text('Конвертация завершена', reply_markup=start_keyboard())
        context.bot.send_document(chat_id=chat_id, document=open(result_csv_file, mode='rb'), filename=result_file)
    except Exception:
        print(traceback.format_exc())
        update.message.reply_text('Не удалось обработать Ваш запрос. Попробуйте позже', reply_markup=start_keyboard())
    finally:
        if os.path.isfile(result_excel_file):
            os.remove(result_excel_file)
        if os.path.isfile(result_csv_file):
            os.remove(result_csv_file)
        return ConversationHandler.END


def csv_to_excel(update, context):
    result_csv_file = ''
    result_excel_file = ''
    try:
        expan = update.message['document']['file_name'].split('.')[1]
        if expan != 'csv':
            update.message.reply_text('Неверный формат для конвертации', reply_markup=start_keyboard())
            return ConversationHandler.END
        chat_id = update.message['chat']['id']
        format_file, initial_name = get_file_info(update)
        result_csv_file = f'data/{chat_id}.csv'
        with open(result_csv_file, 'wb') as file:
            context.bot.get_file(update.message['document']['file_id']).download(out=file)
            file.close()
        delimiter = find_delimiter(result_csv_file)
        result_excel_file = f'data/{chat_id}.xlsx'
        with open(result_csv_file, encoding='utf8') as csvfile:
            reader = csv.reader(csvfile, delimiter=delimiter, quotechar='"')
            wb = openpyxl.Workbook()
            ws = wb.active
            for row in reader:
                ws.append(row)
            wb.save(result_excel_file)
            csvfile.close()
        result_file = f'{initial_name}.xlsx'
        update.message.reply_text('Конвертация завершена', reply_markup=start_keyboard())
        context.bot.send_document(chat_id=chat_id, document=open(result_excel_file, mode='rb'), filename=result_file)
    except Exception:
        print(traceback.format_exc())
        update.message.reply_text('Не удалось обработать Ваш запрос. Попробуйте позже', reply_markup=start_keyboard())
    finally:
        if os.path.isfile(result_csv_file):
            os.remove(result_csv_file)
        if os.path.isfile(result_excel_file):
            os.remove(result_excel_file)
        return ConversationHandler.END


def pdf_to_docx(update, context):
    result_pdf_file = ''
    result_docx_file = ''
    try:
        expan = update.message['document']['file_name'].split('.')[1]
        if expan != 'pdf':
            update.message.reply_text('Неверный формат для конвертации', reply_markup=start_keyboard())
            return ConversationHandler.END
        chat_id = update.message['chat']['id']
        format_file, initial_name = get_file_info(update)
        result_pdf_file = f'data/{chat_id}.pdf'
        with open(result_pdf_file, 'wb') as file:
            context.bot.get_file(update.message['document']['file_id']).download(out=file)
            file.close()
        result_docx_file = f'data/{chat_id}.docx'
        result_file = f'{initial_name}.docx'
        cv = Converter(result_pdf_file)
        cv.convert(result_docx_file)
        cv.close()
        update.message.reply_text('Конвертация завершена', reply_markup=start_keyboard())
        context.bot.send_document(chat_id=chat_id, document=open(result_docx_file, mode='rb'), filename=result_file)
    except Exception:
        print(traceback.format_exc())
        update.message.reply_text('Не удалось обработать Ваш запрос. Попробуйте позже', reply_markup=start_keyboard())
    finally:
        if os.path.isfile(result_pdf_file):
            os.remove(result_pdf_file)
        if os.path.isfile(result_docx_file):
            os.remove(result_docx_file)
        return ConversationHandler.END


def docx_to_pdf(update, context):
    result_pdf_file = ''
    result_docx_file = ''
    try:
        expan = update.message['document']['file_name'].split('.')[1]
        if expan != 'docx':
            update.message.reply_text('Неверный формат для конвертации', reply_markup=start_keyboard())
            return ConversationHandler.END
        pythoncom.CoInitialize()
        chat_id = update.message['chat']['id']
        format_file, initial_name = get_file_info(update)
        result_docx_file = f'data/{chat_id}.docx'
        with open(result_docx_file, 'wb') as file:
            context.bot.get_file(update.message['document']['file_id']).download(out=file)
            file.close()
        print(update.message)
        result_pdf_file = f'data/{chat_id}.pdf'
        result_file = f'{initial_name}.pdf'
        convert(result_docx_file)
        update.message.reply_text('Конвертация завершена', reply_markup=start_keyboard())
        context.bot.send_document(chat_id=chat_id, document=open(result_docx_file, mode='rb'), filename=result_file)
    except com_error as ce:
        error = ce.excepinfo
        update.message.reply_text(f'Произошла ошибка: "{error[2]}"', reply_markup=start_keyboard())
    except Exception:
        print(traceback.format_exc())
        update.message.reply_text('Не удалось обработать Ваш запрос. Попробуйте позже', reply_markup=start_keyboard())
    finally:
        if os.path.isfile(result_docx_file):
            os.remove(result_docx_file)
        if os.path.isfile(result_pdf_file):
            os.remove(result_pdf_file)
        return ConversationHandler.END


def stop(update, context):
    update.message.reply_text('Программа завершена', reply_markup=start_keyboard())
    return ConversationHandler.END


def main():
    token = '677970032:AAEJifhRsPjJG2luEgAvQ7Q9pwX8IG9VQ8I'
    if os.path.isdir('data') is not True:
        os.mkdir('data')
    updater = Updater(token)
    dp = updater.dispatcher
    dp.add_handler(CommandHandler('start', start, pass_user_data=True))
    conv_handler = ConversationHandler(
        entry_points=[MessageHandler(Filters.text & (~ Filters.command), response)],
        states={
            'EXCEL_TO_CSV': [MessageHandler(Filters.document, excel_to_csv, pass_update_queue=True)],
            'CSV_TO_EXCEL': [MessageHandler(Filters.document, csv_to_excel, pass_user_data=True)],
            'PDF_TO_DOCX': [MessageHandler(Filters.document, pdf_to_docx, pass_user_data=True)],
            'DOCX_TO_PDF': [MessageHandler(Filters.document, docx_to_pdf, pass_user_data=True)]
        },
        fallbacks=[CommandHandler('stop', stop)]
    )
    dp.add_handler(conv_handler)
    updater.start_polling()
    updater.idle()


if __name__ == '__main__':
    main()
