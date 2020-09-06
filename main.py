import gspread
import prefill_spreadsheet as prefiller
from oauth2client.service_account import ServiceAccountCredentials
from read_config import read_config

config = read_config()

def run():
	scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
	creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
	gc = gspread.authorize(creds)

	spreadsheet = gc.open_by_key(config['gforms_spreadsheet_id'])

	all_data = spreadsheet.sheet1.get_all_records()
	
	# debugging
	# entry = parse_entry_from_form(all_data[1])
	# handle_entry(gc, entry)
	# return
	# end debugging
	
	for row in all_data:
		entry = parse_entry_from_form(row)
		handle_entry(gc, entry)


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


def handle_entry(gc, entry):
	email = entry['email']

	if not is_valid(entry):
		print("Entry {} not valid, ignoring".format(email))
		return

	should_write = False

	try:
		spreadsheet = gc.open(email)
		print("Found {}".format(email))
		
		# Uncomment whenever a major change is done
		# should_write = True
	except gspread.exceptions.SpreadsheetNotFound:
		print("Unable to access spreadsheet {}, creating".format(email))
		spreadsheet = gc.create(email)

		# Sharing (we also need to share with admin, since the sheet is created by a service account)
		spreadsheet.share(config['admin_gmail'], perm_type='user', role='reader')
		spreadsheet.share(email, perm_type='user', role='writer')
		should_write = True

	if should_write:
		prefiller.run(entry, spreadsheet)


def is_valid(entry):
	return str(entry['email']).endswith("@gmail.com")

if __name__ == '__main__':
	run()
