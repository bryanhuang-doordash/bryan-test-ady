# import dependencies
from docx import Document
import csv

# create document "object"
doc = Document('file.docx')

# create table "object", assumes only one table
table = doc.tables[0]

# create empty lists of attendees/missing for later
list_of_attendees = []
list_of_missing = []

# iterate through each row in the table
# learn nested for loops in python
for i, row in enumerate(table.rows):
	# iterate through each cell in each row
	for j, cell in enumerate(row.cells):
		# get text from cell
		cell_text = cell.text
		# for each text in cell, create a list split by the return character
		list_of_lines = cell_text.split("\n")
		# iterate through each list of strings to find names
		for line in list_of_lines:
			# if the checkbox is clicked, add it to attendees
			if '☒' in line:
				parsed_name = line.replace('☒ ', '')
				list_of_attendees.append(parsed_name)
			# if the checkbox exists but is _not_ clicked, add it to the list of missing
			elif '☐' in line:
				parsed_name = line.replace('☐ ', '')
				list_of_missing.append(parsed_name)

# print everything to terminal in order to verify output
print(list_of_attendees)
print(list_of_missing)


# take both lists, create a csv file called 'attendees.csv', then add rows
with open('attendees.csv', 'w', newline='') as file:
	writer = csv.writer(file)

	for attendee in list_of_attendees:
		writer.writerow([attendee, 'attended'])

	for missing in list_of_missing:
		writer.writerow([missing, 'did not attend'])

