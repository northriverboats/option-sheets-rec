import pickle
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
import datetime

# Put your data at the top
pickles = [
	[r"K:\Links\2020\Options\options.pickle", r"output\Boat Options Parts Listing.xlsx"],
	[r"K:\Links\2020\Yamaha Rigging\yamaha rigging.pickle", r"output\Yamaha Options Parts Listing.xlsx"],
	[r"K:\Links\2020\Mercury Rigging\mercury rigging.pickle", r"output\Mercury Options Parts Listing.xlsx"],
	[r"K:\Links\2020\Honda Rigging\honda rigging.pickle", r"output\Honda Options Parts Listing.xlsx"],
]

def load_pickle(file_name):
	with open(file_name, "rb") as file:
		return pickle.load(file)

def process_lengths(wb, options, file_name):
	for length in [k.split(' ')[0] for k in options[next(iter(options))] if " TOTAL COST" in k]:
		ws = wb["L" + length]
		bold = Font(bold=True, underline="single", color="FF0000")
		# red = Font(color="FF0000")
		blue = Font(color="0000FF")
		row = 1
		for option in sorted(options):
			# if len(options[option]["OUTFITTING PARTS"]) + len(options[option]["CANVAS PARTS"]) + len(options[option]["FABRICATION PARTS"]) + len(options[option]["PAINT PARTS"]) > 0:
			if len(options[option]["OPTION NOTES"]) > 0:
					row += 1
					ws.cell(row = row, column = 1).value =  options[option]["OPTION NOTES"]
					ws.cell(row = row, column = 1).font = bold
			if len(options[option]["EOS OUTFITTING NOTES"]) > 0:
					row += 1
					ws.cell(row = row, column = 1).value =  options[option]["EOS OUTFITTING NOTES"]
					ws.cell(row = row, column = 1).font = blue

			for section_name, section_options, count in [
				[
					'Paint', 
					options[option]["PAINT PARTS"], 
					len(options[option]["PAINT PARTS"]),
				],
				[
					"Outfutting", 
					options[option]["OUTFITTING PARTS"] + options[option]["CANVAS PARTS"], 
					len(options[option]["OUTFITTING PARTS"]) + len(options[option]["CANVAS PARTS"]),
				],
				[
					'Fabrication',
					options[option]["FABRICATION PARTS"], 
					len(options[option]["FABRICATION PARTS"])
				],
				[
					'Paint', 
					options[option]["PAINT PARTS"],
					len(options[option]["PAINT PARTS"]),
				],
			]:
				if count > 0:
					row += 1
					ws.cell(row = row, column = 1).value = option + " " + section_name
					ws.cell(row = row, column = 1).font = bold 
					ws.cell(row = row, column = 2).value = options[option]["OPTION NAME"]
					ws.cell(row = row, column = 2).font = bold

				if count > 0:
					for item in section_options:
						row += 1
						print(option, item["PART NUMBER"])
						ws.cell(row = row, column = 1).value = item["VENDOR"]
						ws.cell(row = row, column = 1).alignment = Alignment(horizontal='left')
						ws.cell(row = row, column = 2).value = item["VENDOR PART"]
						ws.cell(row = row, column = 2).alignment = Alignment(horizontal='left')
						ws.cell(row = row, column = 3).value = item["DESCRIPTION"]
						ws.cell(row = row, column = 4).value = float(item["PRICE"])
						ws.cell(row = row, column = 4).number_format = r'_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'
						ws.cell(row = row, column = 5).value = item["UOM"]
						ws.cell(row = row, column = 6).value = item[length + " QTY"]
						ws.cell(row = row, column = 7).value = "=SUM(D" + str(row) + "*F" + str(row) + ")"
						ws.cell(row = row, column = 7).number_format = r'_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'
						ws.cell(row = row, column = 8).value = 0
						ws.cell(row = row, column = 8).number_format = r'_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'
						ws.cell(row = row, column = 9).value = "=SUM(G" + str(row) + "+H" + str(row) + ")"
						ws.cell(row = row, column = 9).number_format = r'_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'
					
			
				if len(options[option]["PAINT PARTS"]) > 0:
						ws.cell(row = row, column = 1).value = option + " Paint"
				if len(options[option]["PAINT PARTS"]) > 0:
					for item in options[option]["PAINT PARTS"]:
						pass
	wb.save(file_name)



def process_options(file_name_in, file_name_out):
	wb = load_workbook(r"templates\CostingSheetTemplate.xlsx")
	options =  load_pickle(file_name_in)
	process_lengths(wb, options, file_name_out)

for pickle_file, output in pickles:
	process_options(pickle_file, output)