"""
read from the active autocad document "门窗编号“ block instances
and create a door schedule from their attributes
the schedule data should already exist as autocad block attributes
"""

__author__ = "Zhihang Fu"
__license__ = "MIT"
__version__ = 0.1
__email__ = "fuzhihang6829@gmail.com"


import win32com.client
import csv
import os


acad = win32com.client.Dispatch("AutoCAD.Application")

# block instance name to search for
blockName = "门窗编号"
# attributes to search for and decide a block's uniqueness
searchAtts = ["门窗主编号1", "门窗主编号2", "门窗辅编号1", "门窗辅编号2"]

schedule = {}


def doorEntry(entity):
	"""
	convert door acad entity into a string entry for the schedule
	"""
	if entity.HasAttributes:
		# extract attributes as a dictionary
		attTable = {}  
		for attrib in entity.GetAttributes():
			attTable[attrib.TagString] = attrib.TextString
		# create door entry as a string
		doorEntry = "|".join([attTable[a] for a in searchAtts])
		return doorEntry
	else:
		return None


def addToSchedule(schedule, entity):
	entry = doorEntry(entity)

	if entry in schedule.keys():
		schedule[entry] += 1
	else:
		schedule[entry] = 1


# iterate through all objects (entities) in the currently opened drawing
for entity in acad.ActiveDocument.ModelSpace:
	if entity.EntityName == "AcDbBlockReference":
		if entity.Name == blockName:
			addToSchedule(schedule, entity)


# writing schedule to a csv file
fileName = acad.ActiveDocument.name.replace(".dwg", "_") + "门窗表.csv"
filePath = acad.ActiveDocument.path
path = os.path.join(filePath, fileName)
with open(path, 'w', newline='') as f:
	w = csv.writer(f)
	w.writerow(searchAtts + ["数量"])
	for key in schedule.keys():
		w.writerow(key.split("|") + [schedule[key]])

# open the file just saved
os.startfile(path)
