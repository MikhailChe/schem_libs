import xlrd

import sqlite3
import subprocess

#dict init
tables = dict()
tables["ResArray"] = 0
tables["ResE24"] = 1
tables["ResE192"] = 2
tables["CapE24"] = 3
tables["CapE192"] = 4

#open db
db = sqlite3.connect("./lib.sqlite")
cursor = db.cursor()

#open excel
rb = xlrd.open_workbook('./library.xls',formatting_info=True)


for t in tables:
	sheet = rb.sheet_by_index(tables[t])
	cursor.execute('''CREATE TABLE "{tname}" (
				"PartNumber"	TEXT ( 255 ) NOT NULL UNIQUE,
				"Library Ref"	TEXT ( 255 ) NOT NULL,
				"Footprint Ref"	TEXT ( 255 ) NOT NULL,
				"Library Path"	TEXT ( 255 ) NOT NULL,
				"Footprint Path"	TEXT ( 255 ) NOT NULL,
				"Value"	TEXT ( 255 ) NOT NULL,
				"ComponentLink1URL"	TEXT ( 255 ),
				"ComponentLink1Description"	TEXT ( 255 ),
				"Package"	TEXT ( 255 ) NOT NULL,
				PRIMARY KEY("PartNumber"))'''.format(tname=t))
	for rownum in range(1,sheet.nrows):
		vals = [str(d) for d in sheet.row_values(rownum)]
		cursor.execute("INSERT INTO '{tname}' VALUES ('{pn}','{lr}','{fr}','{lp}','{fp}','{v}','{clu}','{cld}','{pac}')"
						.format(pn=vals[0],
								lr=vals[1],
								fr=vals[2],
								lp=vals[3],
								fp=vals[4],
								v=vals[5],
								clu=vals[6],
								cld=vals[7],
								pac=vals[8],
								tname = t))
db.commit()
db.close()
	