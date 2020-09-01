import gspread
import numpy as np
import batch_preparation as bp
from gspread_formatting import *
from oauth2client.service_account import ServiceAccountCredentials
from read_config import read_config

def run():
	row_infos = [
		{ 'start': 'B1', 'data': ['č. prodejce:', 'příjmení, jméno:', 'tel', 'e-mail:']},
		{ 'start': 'B4', 'data': [
			'Číslo prodejce a pořadové číslo zboží', 'Druh zboží (vyber ze seznamu)', 
			'Barva (vyber ze seznamu)', 'Značka', 
			'Bližší specifikace - tři slova (stručný popis, např. motiv na zboží, název knihy, atd.) -max. 30 znaků-', 
			'Velikost', 'Cena (zaokrouhleno na 10 Kč)\nnapř.10, 20, 50'
		]}
	]
	col_infos = [
		{ 'start': 'C35', 'data': [
		'Celkem cena za zisk z prodeje:', 'Provize organizátorům 20%:', 'Celková vyplacená cena za prodej na burze:'
		]}
	]
	colors = [
		'Béžová','Bílá','Černá','Červená','Fialová','Hnědá','Modrá','Oranžová','Růžová','Stříbrná',
		'Šedá','Tyrkysová','Vínová','Zelená','Zlatá','Žlutá','jiná'
	]
	selection_types = [
		'Autosedačky','Boty','Brusle','Bunda','Čepice, klobouk','Dresy','Dupačky, overal',
		'Džíny','Helma','Hračka','Kabát','Kabelky, tašky, batohy','Kalhotky, slipy, boxerky',
		'Kalhoty','Karnevalový kostým','Kniha','Kočárek','Kojenecká výbava -mimo oblečení',
		'Kojenecké body','Kolo','Komplet, Souprava','Košile, Halenka, Blůzka','Kraťasy, Trenýrky',
		'Legíny','Mikina','Nábytek','Pláštěnka','Plavky','Plyšová hračka','Podprsenka','Ponožky',
		'Punčocháče','Pyžamo','Společenské hry','Společenské oblečení','Sportovní potřeba','Sukně',
		'Svetr','Šaty','Těhotenské oblečení','Tepláky','Tílko','Tričko dl. rukáv','Tričko kr. rukáv',
		'Župan','Ostatní zboží'
	]
	col_validations = [
		{ 'start':'C5', 'data':selection_types, 'count':30 },
		{ 'start':'D5', 'data':colors, 'count': 30}
	]
	non_editable_ranges = ['A:B', 'I:Z34', 'C35:Z']
	writing_spreadsheet_key = "1vN72A2jyyP7u2Yo9SJdI11NTChalIMfnMvoRrDVKyCM"
	email = "test@gmail.com"
	phone_number = "123456789"
	full_name = "Linus Tuxin"
	vendor_id = "A001"
	vendor_info = [vendor_id, full_name, phone_number, email]
	config = read_config()
	scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
	creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
	gc = gspread.authorize(creds)
	spreadsheet = gc.open_by_key(writing_spreadsheet_key)
	worksheet = spreadsheet.get_worksheet(0)
	prefill_spreadsheet(worksheet, row_infos, col_infos, vendor_info)
	set_spreadsheet_validation(worksheet, [], col_validations, non_editable_ranges)

def set_spreadsheet_validation(worksheet, row_validations, col_validations, non_editable_ranges):
	for row in row_validations:
		bp.one_of_condition_row_from_a1(worksheet, row['data'], row['start'], row['count'])
	for col in col_validations:
		bp.one_of_condition_column_from_a1(worksheet, col['data'], col['start'], col['count'])
	for non_editable_range in non_editable_ranges:
		worksheet.add_protected_range(non_editable_range)

def generate_item_ids(vendor_id, count):
	item_ids = []
	for i in range(count):
		item_ids.append(vendor_id + str(i + 1).zfill(2))
	return item_ids

def prefill_spreadsheet(worksheet, row_infos, col_infos, vendor_info):
	batch = []
	for row in row_infos:
		batch.append(bp.get_row_batch_from_a1(row['data'], row['start']))
	for col in col_infos:
		batch.append(bp.get_column_batch_from_a1(col['data'], col['start']))
	batch.append(bp.get_row_batch_from_a1(vendor_info, 'B2'))
	batch.append(bp.get_column_batch_from_a1(generate_item_ids(vendor_info[0], 30), 'B5'))
	worksheet.batch_update(batch)
	
if __name__ == '__main__':
	run()
