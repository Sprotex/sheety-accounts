import gspread
import numpy as np

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