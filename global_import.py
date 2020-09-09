import time
import gspread
import prefill_spreadsheet as prefiller
import batch_preparation as bp
from oauth2client.service_account import ServiceAccountCredentials
from read_config import read_config

config = read_config()

def run():
	scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
	creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret2.json', scope)
	gc = gspread.authorize(creds)
	spreadsheet = gc.open_by_key(config['gforms_spreadsheet_id'])
	worksheet = spreadsheet.get_worksheet(0)
	form_data = worksheet.get_all_records()
	global_spreadsheet_name = 'Global Export Pohoda'
	global_spreadsheet = None
	processed_emails = set()
	try:
		global_spreadsheet = gc.open(global_spreadsheet_name)
	except gspread.exceptions.SpreadsheetNotFound:
		global_spreadsheet = gc.create(global_spreadsheet_name)
		global_spreadsheet.share(config['admin_gmail'], perm_type='user', role='writer')
	writing_worksheet = global_spreadsheet.get_worksheet(0)
	starting_line = 1
	batches = []
	
	# debugging
	# entry = parse_entry_from_form(form_data[0])
	# email = entry['email']	
	# starting_line, batch = collect_global_batch(gc, entry, writing_worksheet, starting_line)
	# batches = batches + batch
	# processed_emails.add(email)
	# return
	# debugging end
	
	for _ in range(1):
		for row in form_data:
			entry = parse_entry_from_form(row)
			email = entry['email']
			if email not in processed_emails:
				starting_line, batch = collect_global_batch(gc, entry, writing_worksheet, starting_line)
				batches = batches + batch
				processed_emails.add(email)
	writing_worksheet.batch_update(batches)

def parse_customer_list_line(line):
	item_list = list(line.values())
	return {
		'id':item_list[1],
		'type':item_list[2],
		'color':item_list[3],
		'brand':item_list[4],
		'short description':item_list[5],
		'size':item_list[6],
		'price':item_list[7]
	}

def parse_customer_list_line_v2(item_list):
	return {
		'id':               get_list_value_or_default(item_list, 1 - 1),
		'type':             get_list_value_or_default(item_list, 2 - 1),
		'color':            get_list_value_or_default(item_list, 3 - 1),
		'brand':            get_list_value_or_default(item_list, 4 - 1),
		'short description':get_list_value_or_default(item_list, 5 - 1),
		'size':             get_list_value_or_default(item_list, 6 - 1),
		'price':            get_list_value_or_default(item_list, 7 - 1)
	}

def parse_entry_from_form(form_row):
	row_list = list(form_row.values())
	return {
		'timestamp': row_list[0],
		'name': row_list[1],
		'email': row_list[2],
		'surname': row_list[3],
		'phone': row_list[4],
		'residence': row_list[5],
		'GDPR confirmation': row_list[6],
		'PSC': row_list[7],
		'vendor id': row_list[8]
	}

def get_list_value_or_default(lst, index):
	if index < len(lst):
		return lst[index]
	return ''

def collect_global_batch(gc, entry, writing_worksheet, starting_line):
	email = entry['email']
	name = entry['name'] + ' ' + entry['surname']
	vendor_id = entry['vendor id']
	try:
		spreadsheet = gc.open(email)
		
		lines = spreadsheet.values_batch_get('Sheet1!B5:H34')['valueRanges'][0]
		if 'values' not in lines:
			return starting_line, []
		read_lines = lines['values']
		# read_lines = spreadsheet.get_worksheet(0).get_all_records(head=4)[0:30]
		lines_to_write = []
		for line in read_lines:
			entry = parse_customer_list_line_v2(line)
			item_id = entry['id']
			item_combined_description = entry['type'] + '; ' + entry['color'] + '; ' + entry['short description']
			item_size = entry['size']
			item_price = entry['price']
			vendor_name = name
			line_to_write = [
				'1', 'Karta', '2', 'BUR', '2', 'BUR', '1', 'B',
				item_id, item_combined_description, item_size, 
				'ks', '0', item_price, item_price, vendor_id, 
				vendor_name, item_price
			]
			lines_to_write.append(line_to_write)
		col_length = len(lines_to_write)
		if col_length > 0:
			row_length = len(lines_to_write[0])
			writing_range_start = gspread.utils.rowcol_to_a1(starting_line, 1)
			writing_range_end = gspread.utils.rowcol_to_a1(starting_line + col_length - 1, row_length)
			writing_range = writing_range_start + ':' + writing_range_end
			batch = bp.get_block_batch(lines_to_write, writing_range)
			# NOTE(Andy): 100 reads per 100 seconds, 3 reads per sheet
			time.sleep(3)
			return col_length + starting_line, [batch]
	except gspread.exceptions.SpreadsheetNotFound:
		print("Unable to access spreadsheet {}".format(email))
	return starting_line, []

if __name__ == '__main__':
	run()
