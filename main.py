# -*- coding: utf-8 -*-
import shutil, os
import arcpy
import multiprocessing as mp
from functions import readExcelFile
from functions import findFields
from functions import makeCopyFeatureClass
from functions import FeatureClassSeparateColumn
from functions import CreateMXD
from functions import Utf8ToUnicode
from functions import isDirectoryExists, findFileDirectoryByExtension

def mainFunction(provinceCodes):
	# **********************************************
	# Resources directories
	# -----------------------------------provinceCodes-----------
	# 1) Project directory
	# 2) Base template .mxd file directory
	# 3) Geo-database file directory
	# 4) Shape Files directory
	# 5) Excel workbook file directory
	# 6) Results files and folders directory
	# 7) ?
	# **********************************************
	# prjDir = rootDrive % "NK//rasad"
	rootDrive = r"F://%s"
	mxdDir = rootDrive % "NK//rasad//src//mxd//"
	gdbDir = rootDrive % r"NK//rasad//src//gdb//regions.gdb"
	shpDir = rootDrive % r"NK//rasad//src//shp//"
	xlsxDir= rootDrive % r"NK//rasad//src//xlsx//"
	imgDir = rootDrive % r"NK//rasad//src//img//"
	resDir = rootDrive % r"NK//rasad//res//%s"

	countiesFeatureClassDir = gdbDir + '//county'
	citiesFeatureClassDir = gdbDir + '//city'
	regionsFeatureClassDir = gdbDir + '//regions'

	regionLyr = shpDir + "regions%s.lyr"
	cityLyr = shpDir + "city.lyr"
	countyLyr = shpDir + "county.lyr"
	countiesLyr = shpDir + "counties.lyr"
	layerFiles = {
		"lyrRegion": {
			'edu': regionLyr % '_edu',
			'pop': regionLyr % '_pop',
			'fam': regionLyr % '_fam',
			'hou': regionLyr % '_hou',
			'mig': regionLyr % '_mig',
			'lab': regionLyr % '_lab',
			'mar': regionLyr % '_mar',
		},
		"lyrCity": cityLyr,
		"lyrCounty": countyLyr,
		"lyrCounties": countiesLyr
	}
	# **********************************************
	# Main constant contents
	# ----------------------------------------------
	# 1) Base template .mxd file name
	# 2) Pre-defined legend style name
	# 3) List of all node folder which are needed
	# 4) List of none-attribute field of shape file
	# 5) Dictionary of main subjects like Population, Education and Housing
	# 6) ?
	# **********************************************
	mxd0Dir = mxdDir + 'template0.mxd'
	logo0Dir = imgDir + 'logo.png'
	df_mainStyle = 'df0_legend_style'
	contents = ['SHP', 'MXD', 'JPG']
	mainFeatureClassFields = ['address', 'province', 'county', 'section', 'city', 'areaCode']
	subjects = [
		{'key': 'edu', 'lbl': 'آموزش'},
		{'key': 'pop', 'lbl': 'جمعیت'},
		{'key': 'fam', 'lbl': 'خانوار'},
		{'key': 'hou', 'lbl': 'مسکن'},
		{'key': 'mig', 'lbl': 'مهاجرت'},
		{'key': 'lab', 'lbl': 'اشتغال'},
		{'key': 'mar', 'lbl': 'زناشویی'},
	]
	subjectKeys = ['edu', 'pop', 'fam', 'hou', 'mig', 'lab', 'mar']

	# **********************************************
	# Reading/Opening resources
	# ----------------------------------------------
	# 1) Opening main .xlsx file >> contains titles and other contents which are needed for DataFrame "Main".
	# 2) Fetching all fields of main .shp File
	# 3) ?
	# **********************************************
	src_xlsx_wb = readExcelFile(xlsxDir + "regions.xlsx")
	countiesFields = arcpy.ListFields(countiesFeatureClassDir)
	countiesFieldsNames = [f.name for f in countiesFields]
	citiesFields = arcpy.ListFields(citiesFeatureClassDir)
	citiesFieldsNames = [f.name for f in citiesFields]
	regionsFields = arcpy.ListFields(regionsFeatureClassDir)
	regionsFieldsNames = [f.name for f in regionsFields]



	mainFeatureClassFieldsObj = FeatureClassSeparateColumn(regionsFeatureClassDir, mainFeatureClassFields)
	addresses = mainFeatureClassFieldsObj['address']
	provinces = mainFeatureClassFieldsObj['province']
	counties = mainFeatureClassFieldsObj['county']
	sections = mainFeatureClassFieldsObj['section']
	cities = mainFeatureClassFieldsObj['city']


	# **********************************************
	# Main part
	# **********************************************
	cityCodes = []
	shpFiles = {}
	logoDir = ''
	console = 'Province: %s | County: %s | Section: %s | City: %s | Subject: %s'
	for address in addresses:
		dirs = []
		cityCode = address[0:10]
		if cityCode not in cityCodes:
			cityCodes.append(cityCode)
		else:
			continue
		areaGroup = [jj for jj in addresses if jj.startswith(cityCode)]
		index = addresses.index(areaGroup[0])
		address0 = addresses[index]
		prCode = address0[0:2]
		if prCode not in provinceCodes:
			continue

		coCode = address0[2:4]
		seCode = address0[4:6]
		ciCode = address0[6:10]

		prName = provinces[index].rstrip()
		coName = counties[index].rstrip()
		seName = sections[index].rstrip()
		ciName = cities[index].rstrip()

		province = prCode + "_" + prName.replace(" ", "_")		# province level
		county = coCode + "_" + coName.replace(" ", "_")		# county level
		section = seCode + "_" + seName.replace(" ", "_")		# section level
		city = ciCode + "_" + ciName.replace(" ", "_")			# city level
		divisions = Utf8ToUnicode('تقسیمات سیاسی')				# تقسیمات سیاسی
		divisions0 = Utf8ToUnicode('شهرستان‌ها')						# تقسیمات سیاسی > شهرستان ها
		divisions1 = Utf8ToUnicode('نقاط شهری')					# تقسیمات سیاسی > نقاط شهری

		geoCodes = {'prCode': prCode, 'coCode':prCode + coCode, 'ciCode':ciCode}

		dirLevel1 = resDir % province
		dirLevel2 = resDir % province + "//" + county
		dirLevel3 = resDir % province + "//" + county + "//" + section
		dirLevel4 = resDir % province + "//" + county + "//" + section + "//" + city
		dirLevel2_tmp = resDir % province + "//" + divisions
		dirLevel2_tmp1 = resDir % province + "//" + divisions + "//" + divisions0
		dirLevel2_tmp2 = resDir % province + "//" + divisions + "//" + divisions1

		if dirLevel1 not in dirs: dirs.append(dirLevel1)
		if dirLevel2 not in dirs: dirs.append(dirLevel2)
		if dirLevel3 not in dirs: dirs.append(dirLevel3)
		if dirLevel4 not in dirs: dirs.append(dirLevel4)
		if dirLevel2_tmp not in dirs: dirs.append(dirLevel2_tmp)
		if dirLevel2_tmp1 not in dirs: dirs.append(dirLevel2_tmp1)
		if dirLevel2_tmp2 not in dirs: dirs.append(dirLevel2_tmp2)

		for subject in subjects:
			subjectName = Utf8ToUnicode(subject['lbl'])
			subjectKey = subject['key']
			dirLevel5 = dirLevel4 + "//" + subjectName
			dirs.append(dirLevel5)

			for content in contents:
				dirLevel6 = dirLevel5 + "//" + content
				dirs.append(dirLevel6)

			for dir in dirs:
				if not isDirectoryExists(dir):
					try:
						os.mkdir(dir)
					except OSError as e:
						print("Error in MKDir %s" % dir)
					else:
						targetFolder = os.path.basename(os.path.normpath(dir))
						if divisions in targetFolder:
							shutil.copy(logo0Dir, dir)
							_newShpName = Utf8ToUnicode('شهرستان‌های_استان_') + prName + ".shp"
							if not findFileDirectoryByExtension(dir, _newShpName):
								_newShpFields = findFields(countiesFieldsNames, '', ['address', 'province', 'county'])
								_where = """address LIKE '%s'""" % (str(prCode) + '%')
								makeCopyFeatureClass(countiesFeatureClassDir, dir, _newShpName, _newShpFields, _where)
							shpFiles['shpCounties'] = findFileDirectoryByExtension(dir, _newShpName)
							logoDir = findFileDirectoryByExtension(dir, "logo.png")

						elif divisions0 in targetFolder:
							_newShpName = Utf8ToUnicode('شهرستان_') + coName + ".shp"
							if not findFileDirectoryByExtension(dir, _newShpName):
								_newShpFields = findFields(countiesFieldsNames, '', ['address', 'province', 'county'])
								_where = """address LIKE '%s'""" % (str(prCode + coCode) + '%')
								makeCopyFeatureClass(countiesFeatureClassDir, dir, _newShpName, _newShpFields, _where)
							shpFiles['shpCounty'] = findFileDirectoryByExtension(dir, _newShpName)

						elif divisions1 in targetFolder:
							_newShpName = Utf8ToUnicode('شهر_') + ciName + ".shp"
							if not findFileDirectoryByExtension(dir, _newShpName):
								_newShpFields = findFields(citiesFieldsNames, '', mainFeatureClassFields)
								_where = """address LIKE '%s'""" % ('%' + str(ciCode))
								makeCopyFeatureClass(citiesFeatureClassDir, dir, _newShpName, _newShpFields, _where)
							shpFiles['shpCity']= findFileDirectoryByExtension(dir, _newShpName)

						elif "SHP" in dir:
							_newShpName = city + ".shp"
							if not findFileDirectoryByExtension(dir, _newShpName):
								_newShpFields = findFields(regionsFieldsNames, subjectKey, mainFeatureClassFields)
								_where = """address LIKE '%s'""" % (str(cityCode) + '%')
								makeCopyFeatureClass(regionsFeatureClassDir, dir, _newShpName, _newShpFields, _where)
							shpFiles['shpRegion'] = findFileDirectoryByExtension(dir, _newShpName)

						elif "MXD" in dir:
							jpgDir = dir[:-4] + "JPG//"
							try:
								os.mkdir(jpgDir)
							except OSError as e:
								print("Error in MKDir %s" % dir)
							else:
								tmp = {'mxd': dir, 'jpg': jpgDir}
								print (console % (prCode, coCode,seCode, ciCode, subjectName))
								CreateMXD(mxd0Dir, shpFiles, tmp, layerFiles, geoCodes, df_mainStyle, src_xlsx_wb, subjectKeys, logoDir)
				else:
					targetFolder = os.path.basename(os.path.normpath(dir))
					if divisions0 in targetFolder:
						_newShpName = Utf8ToUnicode('شهرستان_') + coName + ".shp"
						if not findFileDirectoryByExtension(dir, _newShpName):
							_newShpFields = findFields(countiesFieldsNames, '', ['address', 'province', 'county'])
							_where = """address LIKE '%s'""" % (str(prCode + coCode) + '%')
							makeCopyFeatureClass(countiesFeatureClassDir, dir, _newShpName, _newShpFields, _where)
						shpFiles['shpCounty'] = findFileDirectoryByExtension(dir, _newShpName)

					elif divisions1 in targetFolder:
						_newShpName = Utf8ToUnicode('شهر_') + ciName + ".shp"
						if not findFileDirectoryByExtension(dir, _newShpName):
							_newShpFields = findFields(citiesFieldsNames, '', mainFeatureClassFields)
							_where = """address LIKE '%s'""" % ('%' + str(ciCode))
							makeCopyFeatureClass(citiesFeatureClassDir, dir, _newShpName, _newShpFields, _where)
						shpFiles['shpCity'] = findFileDirectoryByExtension(dir, _newShpName)

if __name__ == '__main__':
	mp.freeze_support()
	p1 = mp.Process(target=mainFunction, args=(['00'], ))
	p2 = mp.Process(target=mainFunction, args=(['01'], ))
	p3 = mp.Process(target=mainFunction, args=(['02'], ))
	# p4 = mp.Process(target=mainFunction, args=(['03'], ))
	# p5 = mp.Process(target=mainFunction, args=(['04'], ))
	# p6 = mp.Process(target=mainFunction, args=(['05'], ))
	p1.start()
	p2.start()
	p3.start()
	# p4.start()
	# p5.start()
	# p6.start()
	p1.join()
	p2.join()
	p3.join()
	# p4.join()
	# p5.join()
	# p6.join()