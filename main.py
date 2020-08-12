import gspread
from oauth2client.service_account import ServiceAccountCredentials
from read_config import read_config


config = read_config()


def run():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
    gc = gspread.authorize(creds)

    spreadsheet = gc.open_by_key(config['gforms_spreadsheet_id'])

    all_data = spreadsheet.sheet1.get_all_records()

    for row in all_data:
        entry = parse_entry_from_form(row)

        handle_entry(gc, entry)


def parse_entry_from_form(form_row):
    row_list = list(form_row.values())

    return {
        'name': row_list[1],
        'email': row_list[2]
    }


def handle_entry(gc, entry):
    email = entry['email']

    should_write = False

    try:
        spreadsheet = gc.open(email)
        print("Found {}".format(email))
        should_write = True
    except gspread.exceptions.SpreadsheetNotFound:
        print("Unable to access spreadsheet {}, creating".format(email))
        spreadsheet = gc.create(email)

        # Sharing (we also need to share with admin, since the sheet is created by a service account)
        spreadsheet.share(config['admin_gmail'], perm_type='user', role='reader')
        spreadsheet.share(email, perm_type='user', role='writer')

        # Uncomment whenever a major change is done
        should_write = True

    if should_write:
        sheet = spreadsheet.sheet1
        sheet.update('A1', 'Name')
        sheet.update('A2', entry['name'])


if __name__ == '__main__':
    run()
