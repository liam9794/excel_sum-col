import openpyxl

filepath = "false"

sheetname = "false"

max = -1

def pathTest():
	try:
		val = input("What is the path to the file that we are editing?: ")
		val = val.replace('"', '')
		workbook = openpyxl.load_workbook(val)
		return val
	except:
		print("Please enter a valid path to your file.")
		return "false"

def sheetTest():
	try:
		val = input("Which sheet are we editing?: ")
		sheet = workbook[val]
		return val
	except:
		print("Please enter a valid sheet name.")
		return "false"

def intTest(textToDisplay):
	try:
		val = int(input(textToDisplay))
		return val
	except:
		print("Please enter a valid number.")
		return -1

# load excel file
while filepath == "false":
	filepath = pathTest()
workbook = openpyxl.load_workbook(filepath)

# open workbook
while sheetname == "false":
	sheetname = sheetTest()
sheet = workbook[sheetname]

# modify the desired cell
# sheet["C53"] = 777

max = intTest("What is the max row that you want to initiate?: ")
while max < 5:
	print("Please enter a number greater than 4.")
	max = intTest("What is the max row that you want to initiate?: ")

for x in range(4, max + 1):
	cell = "B" + str(x)
	startcell = "C" + str(x)
	endcell = "Z" + str(x)
	cellfill = "=SUM(" + startcell + ":" + endcell + ")"
	sheet[cell] = cellfill

# save the file
workbook.save(filepath)
