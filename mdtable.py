#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Program to convert .xlsx to grid_tables for markdown
"""

import sys, signal
sys.dont_write_bytecode = True
signal.signal(signal.SIGINT, signal.SIG_DFL)

import argparse
import openpyxl
import pyperclip

from mods.func_prompt_io import check_exist



# =============== Constant =============== #
BORDER_VERTICAL = "|"
BORDER_HORIZONTAL = "-"
BORDER_CROSS = "+"
BORDER_HEADER = "="



#%% ==============================
class Cell:
	def __init__(self, coordinate, value):
		self._coordinate = coordinate
		self._value = None
		self._merged_cells = []

		self.set_value(value)


	@property
	def value(self):
		return self._value

	@property
	def coordinate(self):
		return self._coordinate

	@property
	def length(self):
		return len(self._value)

	@property
	def is_merged(self):
		if len(self._merged_cells) == 0:
			return False
		else:
			return True

	@property
	def merged_cells(self):
		return self._merged_cells


	def set_value(self, value):
		if value is None:
			self._value = ""
		else:
			self._value = value
		return self


	def append_merge_cell(self, obj_cell):
		"""
		Method to append merged cell

		Args:
			obj_cell (Cell object): Cell object
		"""
		self._merged_cells.append(obj_cell)
		return self


	def set_merged_cell(self, list_obj_cell):
		"""
		Method to set list of Cell object

		Args:
			list_obj_cell (list): list of Cell object
		"""
		self._merged_cells = list_obj_cell
		return self



# =============== Function =============== #
def get_cells(input_file, sheetname, cell_area):
	# open workbook
	obj_wb = openpyxl.load_workbook(input_file, data_only=True)

	# open worksheet
	obj_ws = None
	if sheetname is not None:
		if sheetname in obj_wb.sheetnames:
			obj_ws = obj_ws[sheetname]
		else:
			sys.stderr.write("ERROR: sheetname `{}` is not found.\n".format(sheetname))
			sys.exit(1)

	else:
		obj_ws = obj_wb.active

	if cell_area == [None, None]:
		cell_area[0] = "{}{}".format(obj_ws.min_column, obj_ws.min_row)
		cell_area[1] = "{}{}".format(obj_ws.max_column, obj_ws.max_row)


	list_cells = {}
	layout_cells = []
	for row in obj_ws["{}:{}".format(*cell_area)]:
		layout_cells.append([])
		for cell in row:
			obj_cell = Cell(cell.coordinate, cell.value)
			layout_cells[-1].append(obj_cell)
			list_cells[cell.coordinate] = obj_cell

	merged_cells = [["{}{}".format(openpyxl.utils.get_column_letter(pos[1]), pos[0]) for pos in obj_cell.cells] for obj_cell in obj_ws.merged_cells.ranges]
	for list_coordinate in merged_cells:
		list_obj_cell = [list_cells[coordinate] for coordinate in list_coordinate]
		for obj_cell in list_obj_cell:
			partner_obj_cell = set(list_obj_cell) - set([obj_cell])
			obj_cell.set_merged_cell(list(partner_obj_cell))

	return layout_cells


def convert_markdown(layout_cells):
	# check width
	list_width = []
	max_row = len(layout_cells)
	for col_i in range(len(layout_cells[0])):
		value_length = [len(str(layout_cells[row_i][col_i].value)) for row_i in range(max_row)]
		list_width.append(max(value_length))
	list_format = ["{0:>"+str(v)+"}" for v in list_width]

	contents = []

	# prepare horizontal border
	list_border_horizontal = [[] for _ in range(len(layout_cells))]
	for col_i in range(len(layout_cells[0])):
		prev_state = False
		for row_i in range(max_row):
			obj_cell = layout_cells[row_i][col_i]
			if not (obj_cell.is_merged & prev_state):
				# not merged -> border
				list_border_horizontal[row_i].append(BORDER_HORIZONTAL*list_width[col_i])

			else:
				# merge -> no border
				list_border_horizontal[row_i].append(list_format[col_i].format(""))

			prev_state = obj_cell.is_merged

	for row_i, row in enumerate(layout_cells):
		# draw top border
		border_horizontal = [""] + list_border_horizontal[row_i] + [""]
		contents.append(BORDER_CROSS.join(border_horizontal))

		# draw values and vertical border
		new_row = []
		prev_state = False
		for col_i, obj_cell in enumerate(row):
			if not (obj_cell.is_merged & prev_state):
				# not merged -> border
				new_row.append(BORDER_VERTICAL)

			else:
				# merged -> no border
				new_row.append(" ")

			new_row.append(list_format[col_i].format(obj_cell.value))
			prev_state = obj_cell.is_merged
		new_row.append(BORDER_VERTICAL)
		contents.append("".join([str(v) for v in new_row]))

	# draw bottom border
	border_horizontal = [""] + list_border_horizontal[0] + [""]
	contents.append(BORDER_CROSS.join(border_horizontal))

	return "\n".join(contents)



# =============== main =============== #
if __name__ == '__main__':
	parser = argparse.ArgumentParser(description="Program to convert .xlsx to grid_tables for markdown", formatter_class=argparse.RawTextHelpFormatter)
	parser.add_argument("-i", dest="INPUT_FILE", metavar="INPUT.xlsx", required=True, help="input .xlsx file")
	parser.add_argument("-s", dest="SHEET_NAME", metavar="SHEET_NAME", help="sheet name for converting table")
	parser.add_argument("-r", dest="CELL_AREA", metavar="CELL_NAME", nargs=2, default=[None, None], help="Start and End position cells for target square area")
	parser.add_argument("-c", dest="TO_CLIPBOARD", action="store_true", default=False, help="send clipboard")
	args = parser.parse_args()

	check_exist(args.INPUT_FILE, 2)

	layout_cells = get_cells(args.INPUT_FILE, args.SHEET_NAME, args.CELL_AREA)
	str_grid_table = convert_markdown(layout_cells)

	if args.TO_CLIPBOARD:
		pyperclip.copy(str_grid_table)
	else:
		print(str_grid_table)
