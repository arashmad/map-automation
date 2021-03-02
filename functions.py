# -*- coding: utf-8 -*-

from os import path, walk, chdir
import arcpy
import arcpy.mapping as amp
import xlrd
import re

class my_dictionary(dict):
	def __init__(self):
		self = dict()
	def add(self, key, value):
		self[key] = value


def isFileExists(_fileDir):
	return path.exists(_fileDir)
def isDirectoryExists(_dir):
	return path.exists(_dir)
def findFileDirectoryByExtension(_root, _fileType):
	chdir(_root)
	for root, dirs, files in walk(_root):
		for file in files:
			if file.endswith(_fileType):
				return path.abspath(file)
def Utf8ToUnicode(_string):
	return unicode(_string, "utf-8")
def readExcelFile(_xlsxDir):
	try:
		xl_workbook = xlrd.open_workbook(_xlsxDir)
		return xl_workbook
	except Exception as e:
		print("Error: .xlsx file opening failed!")
		if hasattr(e, 'message'):
			print(e.message)
		else:
			print(e)
			return
def FeatureClassSeparateColumn(_featureClassDir, _listOfColumn):
	res = my_dictionary()
	for column in _listOfColumn:
		res.add(column, [])
	with arcpy.da.SearchCursor(_featureClassDir, _listOfColumn) as cursor:
		for row in cursor:
			for ii in range(0, len(_listOfColumn)):
				res[_listOfColumn[ii]].append(row[ii])
	return res
def findFields(fieldList, fieldKey, field0):
	res = []
	if not fieldKey:
		for field in fieldList:
			if field in field0:
				res.append(field)
	else:
		for field in fieldList:
			if field in field0 or field[0:3] == fieldKey:
				res.append(field)
	return res

def maskTitle(_text):
	_tmp = ''
	_text = re.sub('-', ' - ', _text)
	words = _text.split()
	for word in words:
		if word.isdigit():
			if not(word[0:2] == '13' and len(word) == 4):
				word = "\t" + " " + word + " " + "\t"
		else:
			if "-" in word:
				id = words.index("-")
				try:
					wBefore = word[id - 1]
					word = "\t" + " " + wBefore + " " + "\t"
				except:
					a = 'A'
				try:
					wAfter = word[id + 1]
					word = "\t" + " " + wAfter + " " + "\t"
				except:
					a = 'A'
		_tmp += word + " "
	return re.sub("  ", " ", _tmp)
# def maskTitle(_text):
# 	_tmp = ''
# 	_text = re.sub('-', ' - ', _text)
# 	words = _text.split()
# 	for word in words:
# 		if word.isdigit():
# 			if not(word[0:2] == '13' and len(word) == 4):
# 				_text = _text.replace(word, "\t" + word + "\t")
# 		else:
# 			if "-" in word:
# 				id = words.index("-")
# 				try:
# 					wBefore = word[id - 1]
# 					if wBefore.isdigit():
# 						_text = _text.replace('-', ' -')
# 						_text = _text.replace(wBefore, "\t" + wBefore + "\t")
# 				except:
# 					a = 'A'
# 				try:
# 					wAfter = word[id + 1]
# 					if wAfter.isdigit():
# 						_text = _text.replace('-', '- ')
# 						_text = _text.replace(wAfter, "\t" + wAfter + "\t")
# 				except:
# 					a = 'A'
# 	return re.sub("  ", " ", _text)
def createWhereClause(ids):
	sql = """'address' IN (%s)"""
	whereClause = ''
	for ii in range(0, len(ids)):
		whereClause += "'" + ids[ii] + "'"
		if ii <len(ids) - 1:
			whereClause += ','

	return sql%whereClause
def findInCellArray(_arr, _item):
	res = -1
	for ii in range(0, len(_arr)):
		_item0 = _arr[ii].value
		if isinstance(_item0, int) or isinstance(_item0, float):
			_item0 = str(int(_arr[ii].value))
		if _item0 == _item:
			res = ii
			return res
	return res
def makeStandard(_str):
	return _str
def isStrNumber(_string):
	try:
		int(_string)
		return True
	except:
		return False
def generateDirs(_sheet_name, _root):
	nRows = _sheet_name.nrows
	nCols = _sheet_name.ncols
	addresses = _sheet_name.col_values(0)
	provinces = _sheet_name.col_values(1)
	counties = _sheet_name.col_values(2)
	sections = _sheet_name.col_values(3)
	cities = _sheet_name.col_values(4)

	area = []
	dirs = []
	for ii in range(2, nRows):
		address = _sheet_name.cell_value(ii, 0)
		areaCode = address[0:10]
		if areaCode not in area:
			area.append(areaCode)
		else:
			continue
		areaGroup = [jj for jj in addresses if jj.startswith(areaCode)]
		index = addresses.index(areaGroup[0])
		address0 = addresses[index]
		province = address0[0:2] + "_" + provinces[index].rstrip().replace(" ", "_")  # level1 label
		county = address0[2:4] + "_" + counties[index].rstrip().replace(" ", "_")  # level2 label
		section = address0[4:6] + "_" + sections[index].rstrip().replace(" ", "_")  # level3 label
		city = address0[6:10] + "_" + cities[index].rstrip().replace(" ", "_")  # level4 label

		dirLevel1 = _root % ("test\\" + province)
		dirLevel2 = _root % ("test\\" + province) + "\\" + county
		dirLevel3 = _root % ("test\\" + province) + "\\" + county + "\\" + section
		dirLevel4 = _root % ("test\\" + province) + "\\" + county + "\\" + section + "\\" + city

		if dirLevel1 not in dirs: dirs.append(dirLevel1)
		if dirLevel2 not in dirs: dirs.append(dirLevel2)
		if dirLevel3 not in dirs: dirs.append(dirLevel3)
		if dirLevel4 not in dirs: dirs.append(dirLevel4)

	return dirs
def whatisthis(s):
	if isinstance(s, str):
		print "ordinary string"
	elif isinstance(s, unicode):
		print "unicode string"
	else:
		print "not a string"
def makeCopyFeatureClass(in_fc, out_dir, out_name, keep_fields, where=''):
	"""
	Required:
		 in_fc			input feature class
		 out_fc			output feature class
		 keep_fields	names of fields to keep in output
	Optional:
		 where			optional where clause to filter records
	"""
	fmap = arcpy.FieldMappings()
	fmap.addTable(in_fc)

	fields = {f.name: f for f in arcpy.ListFields(in_fc)}
	for fname, fld in fields.iteritems():
		if fld.type not in ('OID', 'Geometry') and 'shape' not in fname.lower():
			if fname not in keep_fields:
				fmap.removeFieldMap(fmap.findFieldMapIndex(fname))

	return arcpy.FeatureClassToFeatureClass_conversion(in_fc, out_dir, out_name, where, fmap)
def createRanges(_list):
	res = []
	for sym in _list:
		if "-" in sym:
			sym0 = sym.split(" - ")[0]
			sym0 = round(float(sym0), 2)
			# if isinstance(sym0, int):
			# 	sym0
			sym1 = sym.split(" - ")[1]
			sym1 = round(float(sym1), 2)
			res.append(str(sym0) + " - " + str(sym1))
		else:
			sym = round(float(sym), 2)
			res.append(sym)
	return res
def CreateMXD(srcMXD, srcSHPDirs, targetDir, srcLyrDirs, srcGeo, srcStyleName, srcXlsxWB, srcSubjects, srcLogo):
	mxd = amp.MapDocument(srcMXD)
	mxdDF0 = amp.ListDataFrames(mxd, "Main")[0]
	mxdDF1 = amp.ListDataFrames(mxd, "Index")[0]

	prCode = srcGeo['prCode']
	coCode = srcGeo['coCode']
	ciCode = srcGeo['ciCode']

	shpCity = srcSHPDirs['shpCity']
	shpCounty = srcSHPDirs['shpCounty']
	shpCounties = srcSHPDirs['shpCounties']
	shpRegion = srcSHPDirs['shpRegion']

	lyrCity = srcLyrDirs['lyrCity']
	lyrCounty = srcLyrDirs['lyrCounty']
	lyrCounties = srcLyrDirs['lyrCounties']

	srcLyrCity = amp.Layer(lyrCity)
	srcLyrCounty = amp.Layer(lyrCounty)
	srcLyrCounties = amp.Layer(lyrCounties)

	mxdLayer00 = amp.Layer(shpRegion)
	mxdLayer10 = amp.Layer(shpCity)
	mxdLayer11 = amp.Layer(shpCounty)
	mxdLayer12 = amp.Layer(shpCounties)

	amp.AddLayer(mxdDF0, mxdLayer00, "TOP")
	amp.AddLayer(mxdDF1, mxdLayer12, "TOP")
	amp.AddLayer(mxdDF1, mxdLayer11, "TOP")
	amp.AddLayer(mxdDF1, mxdLayer10, "TOP")

	addLayer = amp.ListLayers(mxd, "", mxdDF1)[0]
	amp.UpdateLayer(mxdDF0, addLayer, srcLyrCity, True)
	addLayer = amp.ListLayers(mxd, "", mxdDF1)[1]
	amp.UpdateLayer(mxdDF0, addLayer, srcLyrCounty, True)
	addLayer = amp.ListLayers(mxd, "", mxdDF1)[2]
	amp.UpdateLayer(mxdDF0, addLayer, srcLyrCounties, True)

	addLayer = amp.ListLayers(mxd, "", mxdDF0)[0]
	fields = arcpy.ListFields(shpRegion)
	for field in fields:
		fieldName = field.name
		fieldCategory = fieldName[0:3]
		if fieldCategory in srcSubjects:
			lyrRegion = srcLyrDirs['lyrRegion'][fieldCategory]
			srcLyrRegion = amp.Layer(lyrRegion)
			amp.UpdateLayer(mxdDF0, addLayer, srcLyrRegion, True)

			if addLayer.supports("LABELCLASSES"):
				for labelClass in addLayer.labelClasses:
					labelClass.showClassLabels = True
					labelClass.expression = "\"<CLR red = '0' green = '0' blue = '0'><FNT size = '10' name = 'B Yekan'>\" & [areaCode] & \"</FNT></CLR>\""
					addLayer.showLabels = True
					arcpy.RefreshActiveView()

			if addLayer.symbologyType == 'GRADUATED_COLORS':
				addLayer.symbology.valueField = fieldName
				labels = addLayer.symbology.classBreakLabels
				try:
					addLayer.symbology.classBreakLabels = createRanges(labels)
				except:
					print('Error in Symbology | %s' % fieldName)

			style0 = amp.ListStyleItems("USER_STYLE", "Legend Items", srcStyleName)[0]
			mxd_legend = amp.ListLayoutElements(mxd, "LEGEND_ELEMENT")[0]
			mxd_legend.title = ""
			mxd_legend.updateItem(addLayer, style0)


			for element in amp.ListLayoutElements(mxd, "PICTURE_ELEMENT"):
				elementName = element.name
				if elementName == 'Logo':
					element.sourceImage = srcLogo

			variableKeys = srcXlsxWB.sheet_by_index(0).row(0)
			colId = findInCellArray(variableKeys, fieldName)
			mapTitles = srcXlsxWB.sheet_by_index(0).cell_value(1, colId)

			for sheet in srcXlsxWB.sheets():
				sheetName = sheet.name
				if sheetName == 'total':
					countryValue = sheet.cell_value(2, colId)
				elif sheetName == 'province':
					featureKeys = sheet.col(0)
					rowId = findInCellArray(featureKeys, makeStandard(prCode))

					provinceName = sheet.cell_value(rowId, 1)
					provinceValue = sheet.cell_value(rowId, colId)
				elif sheetName == 'county':
					featureKeys = sheet.col(0)
					rowId = findInCellArray(featureKeys, makeStandard(coCode))
					countyName = sheet.cell_value(rowId, 1)
					countyValue = sheet.cell_value(rowId, colId)
				elif sheetName == 'city':
					featureKeys = sheet.col(0)
					rowId = findInCellArray(featureKeys, makeStandard(ciCode))

					cityName = sheet.cell_value(rowId, 1)
					cityName0 = cityName[0: len(cityName) - 1]
					cityName1 = cityName[len(cityName) - 1]
					if (isStrNumber(cityName1)):
						cityName = Utf8ToUnicode('منطقه ') + cityName1 + Utf8ToUnicode('شهر ') + cityName0
					else:
						cityName = Utf8ToUnicode('شهر ') + cityName
					cityValue = sheet.cell_value(rowId, colId)
				elif sheetName == 'unit':
					unitText = sheet.cell_value(2, colId)

			for element in amp.ListLayoutElements(mxd, "TEXT_ELEMENT"):
				elementName = element.name
				if elementName == 'elLegend':
					mapTitles = maskTitle(" ".join(mapTitles.split()))
					defWidth = 8
					element.fontSize = 16
					element.text = mapTitles

					if element.elementWidth >= defWidth:
						words = mapTitles.split(' ')
						lines = []
						line = []
						tmp = ''
						itr = 0
						while itr < len(words):
							word = words[itr]
							itr += 1
							tmp += word + ' '
							element.text = tmp
							line.append(word)
							if element.elementWidth >= defWidth:
								line.pop()
								lines.append(line)
								line = []
								tmp = ''
								itr = itr - 1
							if itr == len(words):
								count = 0
								for l in lines:
									count += len(l)
								if count < len(words):
									lines.append(line)

						mapTitlesNew = ''
						for jj in range(0, len(lines)):
							lineStr = " ".join(lines[jj])
							mapTitlesNew += lineStr
							if jj < len(lines) - 1:
								mapTitlesNew += "\n"

						element.text = mapTitlesNew

				elif elementName == 'elUnit' or elementName == 'elUnit2':
					element.text = unitText
				elif elementName == 'countryValue':
					element.text = round(countryValue,2)
				elif elementName == 'elProvinceTitle':
					element.text = Utf8ToUnicode('مناطق شهری استان ') + provinceName
				elif elementName == 'provinceValue':
					element.text = round(provinceValue,2)
				elif elementName == 'elCountyTitle':
					element.text = Utf8ToUnicode('مناطق شهری شهرستان ') + countyName
				elif elementName == 'countyValue':
					element.text = round(countyValue,2)
				elif elementName == 'elCityTitle':
					element.text = cityName
				elif elementName == 'cityValue':
					element.text = round(cityValue,2)

			try:
				mxd_name = targetDir['mxd'] + "//" + fieldName + ".mxd"
				mxd.saveACopy(mxd_name)
			except arcpy.ExecuteError:
				print(arcpy.GetMessages())

			try:
				mxd_jpg_name = targetDir['jpg'] + fieldName + ".jpg"
				amp.ExportToJPEG(mxd, mxd_jpg_name, resolution=300)
				# multiprocessing.freeze_support()
				# p = multiprocessing.Process(target=test, args=(mxd.__getattribute__('filePath'), mxd_jpg_name))
				# p.start()
				# p.join()
			except arcpy.ExecuteError:
				print(arcpy.GetMessages())