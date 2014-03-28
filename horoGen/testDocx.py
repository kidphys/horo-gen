from docx import Document
import xlrd
import unittest

class TestHoroInput(unittest.TestCase):

	def testSanity(self):
		input = HoroInput()
		input.load('input.xlsx')
		self.assertEqual(['Sheet1'], input.getReportNames())
		
	def testGetSectionNames(self):
		input = HoroInput()
		input.load('input.xlsx')
		self.assertEqual('Header', input.getSection()[0]['title'])

class HoroInput():

	def __init__(self):
		self._reportNames = []
		self._maxRow = 500
		self._maxCol = 500

	def load(self, filename):
		self._workbook = xlrd.open_workbook(filename)
		self._reportNames = self._workbook.sheet_names()
		for name in self._reportNames:
			self._loadSection(name)

	def getSection(self):
		return self._section

	def _loadSection(self, sheet_name):
		self._section = []
		sheet = self._workbook.sheet_by_name(sheet_name)
		for row in range(self._maxRow):
			self._section.append({'title': sheet.cell_value(row, 1)})


	def getReportNames(self):
		return self._reportNames

	def getStructure(self):
		return self._structure

class ExcelReader():

	def load(self, filename):
		workbook = xlrd.open_workbook(filename)
		sheet = workbook.sheet_by_name('1.1')
		return sheet.cell_value(1,2)


class TestDocx(unittest.TestCase):

	def testSanity(self):
		doc = Document()	
		excel = ExcelReader()
		content = excel.load('AstroReport_Content.xlsx')
		heading = doc.add_heading('Heading, level 1', level=1)
		doc.add_paragraph(content)
		doc.save('demo.docx')

if __name__ == '__main__':
	unittest.main()


