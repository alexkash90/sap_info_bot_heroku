import telebot, xlrd, openpyxl, datetime
from openpyxl import load_workbook

bot = telebot.TeleBot('836575673:AAHBC2xJziuDCDPB2yYfZlCgifET8pa9_jU')


with open('users.txt', 'r') as fp:
    user_ids = [int(l) for l in fp.readlines()]

@bot.message_handler(func=lambda message: message.chat.id not in user_ids)
def access_msg(message):
    bot.send_message(message.chat.id, 'No Access! откройте телеграм бота @myidbot, и отправьте свой идентификтатор разработчику @alexkash90')



@bot.message_handler(commands=['help'])
def get_text_messages(message):
	bot.send_message(message.from_user.id, "Введите номер заказа, чтобы посмотреть его статус. Дата вначале означает <Дату поставки>")


@bot.message_handler(content_types=['text'])
def get_text_messages(message):
	rb = xlrd.open_workbook(r'data.xlsx')
	sheet = rb.sheet_by_index(0)
	rows = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
	orders = [item for item in rows if item[0] == message.text]

	if len(orders) == 0:
		bot.send_message(message.from_user.id, "Заказ: {} не найден".format(message.text))
	else:
		for order in orders:
			date = datetime.datetime(*xlrd.xldate_as_tuple(order[2], rb.datemode)).strftime("%d/%m/%Y")
			bot.send_message(message.from_user.id, "{} {} {}".format(date, order[3], order[4]))


bot.polling()

