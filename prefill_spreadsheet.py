import gspread
import numpy as np
from gspread_formatting import *
from oauth2client.service_account import ServiceAccountCredentials
from read_config import read_config

def run():
	config = read_config()
	scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
	creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
	gc = gspread.authorize(creds)
	spreadsheet = gc.open_by_key("1vN72A2jyyP7u2Yo9SJdI11NTChalIMfnMvoRrDVKyCM")
	worksheet = spreadsheet.get_worksheet(0)
	prefill_spreadsheet(worksheet)
	set_spreadsheet_validation(worksheet)

def set_spreadsheet_validation(worksheet):
	pass

def one_of_condition_column_from_a1(worksheet, categories, start_cell, count):
	validation_rule = DataValidationRule(
		BooleanCondition('ONE_OF_LIST', categories),
		showCustomUi=True
	)
	start_row, start_col = gspread.utils.a1_to_rowcol(start_cell)
	end_cell = gspread.utils.rowcol_to_a1(start_row + count - 1, start_col)
	cell_range = start_cell + ":" + end_cell
	set_data_validation_for_cell_range(worksheet, cell_range, validation_rule)

def add_colors(worksheet):
	colors = [
		'Béžová','Bílá','Černá','Červená','Fialová','Hnědá','Modrá','Oranžová','Růžová','Stříbrná',
		'Šedá','Tyrkysová','Vínová','Zelená','Zlatá','Žlutá','jiná'
	]
	start_cell = 'N6'
	#add_column_from_a1(worksheet, colors, start_cell)
	one_of_condition_column_from_a1(worksheet, colors, 'D5', 30)
	
def get_selection_types_batch():
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
	start_cell = 'L6'
	return get_column_batch_from_a1(selection_types, start_cell)
	
def get_colors_batch():
	colors = [
		'Béžová','Bílá','Černá','Červená','Fialová','Hnědá','Modrá','Oranžová','Růžová','Stříbrná',
		'Šedá','Tyrkysová','Vínová','Zelená','Zlatá','Žlutá','jiná'
	]
	start_cell = 'N6'
	return get_column_batch_from_a1(colors, start_cell)
	
def add_selection_types(worksheet):
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
	start_cell = 'L6'
	#add_column_from_a1(worksheet, selection_types, start_cell)
	one_of_condition_column_from_a1(worksheet, selection_types, 'C5', 30)
	
def get_column_batch_from_a1(data, start_cell): 
	start_row, start_col = gspread.utils.a1_to_rowcol(start_cell)
	end_cell = gspread.utils.rowcol_to_a1(start_row + len(data) - 1, start_col)
	cell_range = start_cell + ":" + end_cell
	return { 
		'range': cell_range,
		'values': np.array([data]).transpose().tolist(),
	}
	
def get_row_batch_from_a1(data, start_cell):
	start_row, start_col = gspread.utils.a1_to_rowcol(start_cell)
	end_cell = gspread.utils.rowcol_to_a1(start_row, start_col + len(data) - 1)
	cell_range = start_cell + ":" + end_cell
	return {
		'range': cell_range,
		'values':[data]
	}
	
def get_header_batches():
	header1 = ['č. prodejce:', 'příjmení, jméno:', 'tel', 'e-mail:']
	header2 = [
		'Číslo prodejce a pořadové číslo zboží', 'Druh zboží (vyber ze seznamu)', 
		'Barva (vyber ze seznamu)', 'Značka', 
		'Bližší specifikace - tři slova (stručný popis, např. motiv na zboží, název knihy, atd.) -max. 30 znaků-', 
		'Velikost', 'Cena (zaokrouhleno na 10 Kč)\nnapř.10, 20, 50'
	]
	start_cell1 = 'B1'
	start_cell2 = 'B4'
	batch1 = get_row_batch_from_a1(header1, start_cell1)
	batch2 = get_row_batch_from_a1(header2, start_cell2)
	return [batch1, batch2]
	
	for x, arr in enumerate(header):
		for y, value in enumerate(arr):
			worksheet.update_cell(x + 1, y + 1, value)
	
def prefill_spreadsheet(worksheet):
	batch = []
	batch.append(get_selection_types_batch())
	batch.append(get_colors_batch())
	batch = batch + get_header_batches()
	print(batch)
	#return
	worksheet.batch_update(batch)
	
if __name__ == '__main__':
	run()
