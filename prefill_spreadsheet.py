import gspread
import numpy as np
import batch_preparation as bp
import re
import gspread_formatting as gf
from oauth2client.service_account import ServiceAccountCredentials
from read_config import read_config

def run(entry_dict, spreadsheet):
	row_infos = [
		{'start': 'B1', 'data': ['č. prodejce:', 'příjmení, jméno:', '', 'tel:', 'e-mail:']},
		{'start': 'B4', 'data': [
			'Číslo prodejce a pořadové číslo zboží', 'Druh zboží\n(vyber ze seznamu)', 
			'Barva\n(vyber ze seznamu)', 'Značka', 
			'Bližší specifikace - tři slova\n(stručný popis, např.motiv na zboží, název knihy, atd.......)\n-max.30 znaků-', 
			'Velikost', 'Cena\n(zaokrouhleno na 10 Kč)\nnapř.10, 20, 50'
		]},
		{'start': 'B40', 'data': ['!!!!', 'Registrační poplatek Strančické burzy je 20 Kč na jeden prodejní seznam.']}
	]
	col_infos = [
		{'start': 'C35', 'data': [
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
		{'start': 'C5', 'data': selection_types, 'count': 30},
		{'start': 'D5', 'data': colors, 'count': 30}
	]
	non_editable_ranges = ['A:B', 'I:Z34', 'C35:Z']
	bold_ranges = ['B2:H34', 'B37:H40']
	underline_ranges = ['F2']
	background_color_infos = [
		{'range': 'B4:H4' , 'color': gf.color(195/255,219/255,226/255)},
		{'range': 'C5:H34', 'color': gf.color(243/255,212/255,213/255)}
	]
	text_color_infos = [
		{'range': 'F2', 'color': gf.color(165/255,153/255,174/255)}
	]
	grid_inner_ranges = ['B1:B2', 'E1:E2', 'F2', 'B4:H34', 'H2']
	grid_outer_ranges = ['C1:D1', 'C2:D2', 'B3:D3', 'E3:F3', 'B1:H37', 'B35:H35', 'B36:H36', 'B37:H37']
	column_widths = [13, 130, 213, 100, 132, 286, 85, 147, 20, 51]
	row_heights = [
		32, 47, 46, 115, 
		40, 40, 40, 40, 40, 
		40, 40, 40, 40, 40, 
		40, 40, 40, 40, 40, 
		40, 40, 40, 40, 40, 
		40, 40, 40, 40, 40, 
		40, 40, 40, 40, 40, 
		40, 40, 40, 34, 40, 38
	]
	font_infos = [
		{'range': 'B1:F1', 'size': 11},
		{'range': 'B2', 'size': 28},
		{'range': 'C2', 'size': 16},
		{'range': 'B2', 'size': 28},
		{'range': 'B3:H3', 'size': 24},
		{'range': 'B4:H34', 'size': 13},
		{'range': 'E2:F2', 'size': 14}
	]
	merged_cells = ['E3:F3', 'C40:J40'] 
	writing_spreadsheet_key = '1vN72A2jyyP7u2Yo9SJdI11NTChalIMfnMvoRrDVKyCM'
	email = entry_dict['email']
	phone_number = entry_dict['phone']
	full_name = entry_dict['name'] + ' ' + entry_dict['surname']
	vendor_id = entry_dict['vendor id']
	vendor_info = [vendor_id, full_name, '', phone_number, email]
	config = read_config()
	scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
	creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
	gc = gspread.authorize(creds)
	worksheet = spreadsheet.get_worksheet(0)
	batch_text_fill(worksheet, row_infos, col_infos, vendor_info)
	batch_format(worksheet, bold_ranges, background_color_infos, column_widths, row_heights, font_infos, text_color_infos, underline_ranges)
	set_spreadsheet_validation(worksheet, [], col_validations, non_editable_ranges)
	set_merged_cells(worksheet, merged_cells)
	
def set_merged_cells(worksheet, cell_ranges):
	for cell_range in cell_ranges:
		worksheet.merge_cells(cell_range)

def set_bold(batch, worksheet, cell_ranges):
	text_format = gf.textFormat(bold=True)
	cell_format = gf.cellFormat(textFormat = text_format)
	for cell_range in cell_ranges:
		batch.format_cell_ranges(worksheet, [(cell_range, cell_format)])

def set_underline(batch, worksheet, cell_ranges):
	text_format = gf.textFormat(underline=True)
	cell_format = gf.cellFormat(textFormat = text_format)
	for cell_range in cell_ranges:
		batch.format_cell_ranges(worksheet, [(cell_range, cell_format)])

def set_text_color(batch, worksheet, cell_infos):
	for cell_info in cell_infos:
		color = cell_info['color']
		cell_range = cell_info['range']
		text_format = gf.textFormat(foregroundColor = color)
		cell_format = gf.cellFormat(textFormat = text_format)
		batch.format_cell_ranges(worksheet, [(cell_range, cell_format)])

def set_colors(batch, worksheet, color_infos):
	for color_info in color_infos:
		color = color_info['color']
		cell_range = color_info['range']
		cell_format = gf.cellFormat(backgroundColor = color)
		batch.format_cell_ranges(worksheet, [(cell_range, cell_format)])

def set_text_alignment(batch, worksheet, cell_range):
	cell_format = gf.cellFormat(horizontalAlignment = 'CENTER', verticalAlignment = 'MIDDLE')
	batch.format_cell_ranges(worksheet, [(cell_range, cell_format)])

def set_wrap_strategy(batch, worksheet, cell_range):
	cell_format = gf.cellFormat(wrapStrategy = 'WRAP')
	batch.format_cell_ranges(worksheet, [(cell_range, cell_format)])

def set_custom_column_widths(batch, worksheet, column_widths):
	updates = []
	col = 1
	for column_width in column_widths:
		column_a1 = gspread.utils.rowcol_to_a1(1, col)
		column_a1 = re.sub('[0-9]+', '', column_a1)
		updates.append((column_a1, column_width))
		col += 1
	batch.set_column_widths(worksheet, updates)
	
def set_custom_row_heights(batch, worksheet, row_heights):
	updates = []
	row = 1
	for row_height in row_heights:
		row_a1 = str(row)
		updates.append((row_a1, row_height))
		row += 1
	batch.set_row_heights(worksheet, updates)

def set_fonts(batch, worksheet, font_infos):
	for font_info in font_infos:
		cell_range = font_info['range']
		text_format = gf.textFormat(fontSize = font_info['size'])
		cell_format = gf.cellFormat(textFormat = text_format)
		batch.format_cell_ranges(worksheet, [(cell_range, cell_format)])

def set_spreadsheet_validation(worksheet, row_validations, col_validations, non_editable_ranges):
	for row in row_validations:
		one_of_condition_row_from_a1(worksheet, row['data'], row['start'], row['count'])
	for col in col_validations:
		one_of_condition_column_from_a1(worksheet, col['data'], col['start'], col['count'])
	for non_editable_range in non_editable_ranges:
		worksheet.add_protected_range(non_editable_range)

def generate_item_ids(vendor_id, count):
	item_ids = []
	for i in range(count):
		item_ids.append(vendor_id + str(i + 1).zfill(2))
	return item_ids

def one_of_condition_row_from_a1(worksheet, categories, start_cell, count):
	validation_rule = gf.DataValidationRule(
		gf.BooleanCondition('ONE_OF_LIST', categories),
		showCustomUi=True
	)
	start_row, start_col = gspread.utils.a1_to_rowcol(start_cell)
	end_cell = gspread.utils.rowcol_to_a1(start_row, start_col + count - 1)
	cell_range = start_cell + ":" + end_cell
	gf.set_data_validation_for_cell_range(worksheet, cell_range, validation_rule)

def one_of_condition_column_from_a1(worksheet, categories, start_cell, count):
	validation_rule = gf.DataValidationRule(
		gf.BooleanCondition('ONE_OF_LIST', categories),
		showCustomUi=True
	)
	start_row, start_col = gspread.utils.a1_to_rowcol(start_cell)
	end_cell = gspread.utils.rowcol_to_a1(start_row + count - 1, start_col)
	cell_range = start_cell + ":" + end_cell
	gf.set_data_validation_for_cell_range(worksheet, cell_range, validation_rule)

def batch_text_fill(worksheet, row_infos, col_infos, vendor_info):
	batch = []
	for row in row_infos:
		batch.append(bp.get_row_batch_from_a1(row['data'], row['start']))
	for col in col_infos:
		batch.append(bp.get_column_batch_from_a1(col['data'], col['start']))
	batch.append(bp.get_row_batch_from_a1(vendor_info, 'B2'))
	batch.append(bp.get_column_batch_from_a1(generate_item_ids(vendor_info[0], 30), 'B5'))
	worksheet.batch_update(batch)

def batch_format(worksheet, bold_ranges, background_color_infos, column_widths, row_heights, font_infos, text_color_infos, underline_ranges):
	whole_sheet_cell_range = 'A1:J40'
	with gf.batch_updater(worksheet.spreadsheet) as formatting_batch:
		set_bold(formatting_batch, worksheet, bold_ranges)
		set_colors(formatting_batch, worksheet, background_color_infos)
		set_custom_column_widths(formatting_batch, worksheet, column_widths)
		set_custom_row_heights(formatting_batch, worksheet, row_heights)
		set_fonts(formatting_batch, worksheet, font_infos)
		set_text_color(formatting_batch, worksheet, text_color_infos)
		set_underline(formatting_batch, worksheet, underline_ranges)
		set_text_alignment(formatting_batch, worksheet, whole_sheet_cell_range)
		set_wrap_strategy(formatting_batch, worksheet, whole_sheet_cell_range)
