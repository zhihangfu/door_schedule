"""
read from the active autocad document "门窗编号“ block instances
and create a door schedule from their attributes
"""

import win32com.client
import csv
import os


acad = win32com.client.Dispatch("AutoCAD.Application")
blockName = "门窗编号"  # block instance name to search for
searchAtts = ["门窗主编号1", "门窗主编号2", "门窗辅编号1", "门窗辅编号2"]  # attributes to decide if a block is unique


schedule = {}


def doorEntry(entity):
	if entity.HasAttributes:
		attTable = {}
		for attrib in entity.GetAttributes():
			attTable[attrib.TagString] = attrib.TextString
		doorEntry = " ".join([attTable[a] for a in searchAtts])
		return doorEntry
	else:
		return None


def addToSchedule(schedule, entity):
	exist = False
	entry = doorEntry(entity)
	
	if entry in schedule.keys():
		schedule[entry] += 1
		exist = True
	if not exist:
		schedule[entry] = 1


# iterate through all objects (entities) in the currently opened drawing
for entity in acad.ActiveDocument.ModelSpace:
	name = entity.EntityName
	if name == "AcDbBlockReference":
		if entity.Name == blockName:
			addToSchedule(schedule, entity)


# writing schedule to a csv file
fileName = "门窗表.csv"
filePath = acad.ActiveDocument.path
path = os.path.join(filePath, fileName)
with open(path, 'w', newline='') as f:
	w = csv.writer(f)
	w.writerow(["门窗编号","数量"])
	w.writerows(schedule.items())