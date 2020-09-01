import gspread
import numpy as np
from gspread_formatting import *

def one_of_condition_row_from_a1(worksheet, categories, start_cell, count):
	validation_rule = DataValidationRule(
		BooleanCondition('ONE_OF_LIST', categories),
		showCustomUi=True
	)
	start_row, start_col = gspread.utils.a1_to_rowcol(start_cell)
	end_cell = gspread.utils.rowcol_to_a1(start_row, start_col + count - 1)
	cell_range = start_cell + ":" + end_cell
	set_data_validation_for_cell_range(worksheet, cell_range, validation_rule)

def one_of_condition_column_from_a1(worksheet, categories, start_cell, count):
	validation_rule = DataValidationRule(
		BooleanCondition('ONE_OF_LIST', categories),
		showCustomUi=True
	)
	start_row, start_col = gspread.utils.a1_to_rowcol(start_cell)
	end_cell = gspread.utils.rowcol_to_a1(start_row + count - 1, start_col)
	cell_range = start_cell + ":" + end_cell
	set_data_validation_for_cell_range(worksheet, cell_range, validation_rule)
	
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