import pickle
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
import datetime

lenghts = [
	"18.5",
	"20", 
	"21",
	"22",
	"23", 
	"24", 
	"25", 
	"27",
	"29",
	"31",
	"33",
	"35",
]

def process_options(file_name_in, file_name_out):
	# load data from piclk
	with open(file_name_in, "rb") as file:
		options =  pickle.load(file)
	kill=[]
	for option in options:
		co = options.get(option+" - CO")
		cr = options.get(option+" - CR")
		if co:
			kill.append(option+" - CO")
			options[option]["OUTFITTING PARTS"] += co["OUTFITTING PARTS"]
			options[option]["CANVAS PARTS"] += co["CANVAS PARTS"]
			options[option]["PAINT PARTS"] += co["PAINT PARTS"]
			options[option]["FABRICATION PARTS"] += co["FABRICATION PARTS"]
		if cr:
			kill.append(option+" - CR")
			options[option]["OUTFITTING PARTS"] += cr["OUTFITTING PARTS"]
			options[option]["CANVAS PARTS"] += cr["CANVAS PARTS"]
			options[option]["PAINT PARTS"] += cr["PAINT PARTS"]
			options[option]["FABRICATION PARTS"] += cr["FABRICATION PARTS"]

	for option in kill:
		del options[option]

	wb = load_workbook(r"templates\CostingSheetTemplate.xlsx") # create new workbook

	for length in lenghts:
		ws = wb["L" + length]
		bold = Font(bold=True, underline="single", color="FF0000")
		red = Font(color="FF0000")
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
	wb.save(file_name_out)

process_options(r"K:\Links\2020\Options\options.pickle", r"output\Boat Options Parts Listing.xlsx")
process_options(r"K:\Links\2020\Yamaha Rigging\yamaha rigging.pickle", r"output\Yamaha Options Parts Listing.xlsx")
process_options(r"K:\Links\2020\Mercury Rigging\mercury rigging.pickle", r"output\Mercury Options Parts Listing.xlsx")
process_options(r"K:\Links\2020\Honda Rigging\honda rigging.pickle", r"output\Honda Options Parts Listing.xlsx")
