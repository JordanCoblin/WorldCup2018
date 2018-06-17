import openpyxl
from os import listdir, path

excel_sheets = {
	"Jordan": "World Cup 2018 Sheet Jordan.xlsx",
	"Jeff" : "2018 WC Pool Jeff.xlsx",
	"Millie" : "World Cup 2018 Sheet - Emile.xlsx",
	"Nish" : "World Cup 2018 Sheet Nish.xlsx",
	"Shonah" : "World Cup 2018 Sheet Shonah.xlsx",
	"Patty" : "World Cup 2018 Sheet Shonah.xlsx",
	"Vineet" : "World Cup 2018 Sheet Shonah.xlsx",
	"Darnel" : "MD World Cup 2018 Sheet.xlsx",
}

def read_part4(workbook):
	wb = openpyxl.load_workbook(path.join("picks", workbook))
	ws = wb.active

	results = []
	for row in ws.iter_rows(min_row=69, max_row=116, min_col=2, max_col=7):
		for i, cell in enumerate(row):
			if str(cell.value).lower() == "x":
	   			results.append(i)
	return results

def calculate_part4_score(picks, results):
	score = 0
	for i, result in enumerate(results):
		if picks[i] == result:
			score += 1
	return score

results_part4 = read_part4('Results.xlsx')

part4_scores = {}
for player, workbook in excel_sheets.items():
	player_part4 = read_part4(workbook)
	player_part4_score = calculate_part4_score(player_part4, results_part4)
	part4_scores[player] = player_part4_score

print (part4_scores)