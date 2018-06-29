import openpyxl
from os import listdir, path

excel_sheets = {
	"Jordan": "World Cup 2018 Sheet Jordan.xlsx",
	"Jeff" : "2018 WC Pool Jeff.xlsx",
	"Millie" : "World Cup 2018 Sheet - Emile.xlsx",
	"Nish" : "World Cup 2018 Sheet Nish.xlsx",
	"Shonah" : "World Cup 2018 Sheet Shonah.xlsx",
	"Patty" : "World Cup 2018 Sheet_Pfingler.xlsx",
	"Vineet" : "VINNY_WORLD_CUP.xlsx",
	"Darnel" : "MD World Cup 2018 Sheet.xlsx",
}

RESULTS_WB =  "Results.xlsx"
STATS_WB = "bro_stats.xlsx"

# Results are of the form []
def read_part4(workbook):
	wb = openpyxl.load_workbook(path.join("picks", workbook))
	ws = wb.active

	results = []
	for row in ws.iter_rows(min_row=69, max_row=116, min_col=3, max_col=5):
		for i, cell in enumerate(row):
			if str(cell.value).lower() == "x":
				results.append(i)
	return results

def read_part5(workbook):
	wb = openpyxl.load_workbook(path.join("picks", workbook))
	ws = wb.active

	results = []
	group_result_1 = []
	group_result_2 = []
	for i, row in enumerate(ws.iter_rows(min_row=18, max_row=37, min_col=3, max_col=6)):

		# Next group result
		if (i+1)%5 == 0:
			results.append(group_result_1)
			results.append(group_result_2)
			group_result_1 = []
			group_result_2 = []
			continue

		if str(row[0].value).lower() == "x":
			group_result_1.append((i)%5)

		if str(row[3].value).lower() == "x":
			group_result_2.append((i)%5)

	return results

def calculate_part4_score(picks, results):
	score = 0
	for i, result in enumerate(results):
		if picks[i] == result:
			score += 1
	return score

def calculate_part5_score(picks, results):
	score = 0
	for i, group_result in enumerate(results):
		for team in group_result:
			if team in picks[i]:
				score += 2
	return score

def write_stats(col, scores):
	wb = openpyxl.load_workbook(path.join("picks", STATS_WB))
	ws = wb.active

	print(scores)
	i = 2
	for player, score in scores.items():
		score_col = col + str(i)
		ws[score_col] = score
		i += 1

	wb.save(path.join("picks", STATS_WB))

def write_total(score_map):
	i = 2
	total_score_map = score_map
	write_stats('C', total_score_map)

def write_headers(bro_map):
	wb = openpyxl.load_workbook(path.join("picks", STATS_WB))
	ws = wb.active

	i = 2
	for player, foo in bro_map.items():
		player_col = "A" + str(i)
		ws[player_col] = player
		i += 1

	wb.save(path.join("picks", STATS_WB))

results_part4 = read_part4(RESULTS_WB)
results_part5 = read_part5(RESULTS_WB)
print("Part4 results: ", results_part4)
print("Part5 resutls: ", results_part5)

jordan_part5 = read_part5(excel_sheets["Jordan"])
print("jordan results: ", jordan_part5)

jordan_part5_score = calculate_part5_score(jordan_part5, results_part5)
print("jordan p5 score: ", jordan_part5_score)

part4_scores = {}
for player, workbook in excel_sheets.items():
	player_part4 = read_part4(workbook)
	player_part4_score = calculate_part4_score(player_part4, results_part4)
	part4_scores[player] = player_part4_score

print (part4_scores)
write_headers(part4_scores)
write_stats('B', part4_scores)
write_total(part4_scores)