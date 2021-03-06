# coding=utf-8
import unittest
from horoscope import HoroInput
from horoscope import HoroReport
from horoscope import HoroDocument
from horoscope import ExcelReader


class TestDocxGen(unittest.TestCase):

	def testsSanity(self):
		doc = HoroDocument()
		content = []
		content.append({
			'section_title': u'Bạch Dương title',
			'paragraph': [
				{
					'title': u'Bạch Dương paragraph 1',
					'content': u'Bạch Dương content 1'
				},
				{
					'title': u'Bạch Dương paragraph 2',
					'content': u'Bạch Dương content 2'
				}
			]})
		doc.generate(content)
		doc.save('test.docx')

	def testIntegrationGen(self):
		input = HoroInput()
		input.load('sample.xlsx')
		report = HoroReport()
		report.loadDataSource('AstroReport_Content.xlsx')
		content = report.loadContentUsing(input.getReport('Sheet1'))
		doc = HoroDocument()
		doc.generate(content)
		doc.save('sample.docx')


class TestHoroReport(unittest.TestCase):

	def setUp(self):
		self.report = HoroReport()
		self.input = HoroInput()
		self.input.load('input.xlsx')
		content = [['Section']]
		self.report.loadDataSource('AstroReport_Content.xlsx')
		self.content = self.report.loadContentUsing(self.input.getReport('Sheet1'))

	def testSanity(self):
		self.assertEqual(2, len(self.content))

	def testLoadSectionCorrectly(self):
		self.assertEqual(u'Tình yêu', self.content[0]['section_title'])

	def testLoadParagraph(self):
		paragraphs = self.content[0]['paragraph']
		self.assertEqual(3, len(paragraphs))
		self.assertEqual(u'Cung Mọc ở Bạch Dương', paragraphs[0]['title'])

	def testLoadNonExistParagraph(self):
		paragraph = self.report.getParagraphFromSource('1.1', '1.3')
		self.assertEqual('Cannot found', paragraph['title'])

	def testLoadNonExistSheet(self):
		paragraph = self.report.getParagraphFromSource('1111', '1,3')
		self.assertEqual('Cannot found', paragraph['title'])


class TestHoroInput(unittest.TestCase):

	def setUp(self):
		self.input = HoroInput()
		self.input.load('input.xlsx')
		self.data = [['Section', u'Tình yêu'],
				[1.1, 1], [2.1, 3]]
		self.content = [{'Section': u'Tình yêu'},
					{'Paragraph': {'sheet': 1.1, 'value': 1}}]

	def testCleanInput(self):
		actual = self.input.clean(self.data)
		self.assertEqual(self.content[0], actual[0])
		self.assertDictEqual(self.content[1], actual[1])

	def testSanity(self):
		self.assertEqual(['Sheet1'], self.input.getReportNames())
		
	def testGetParagraphInfo(self):
		input = self.input.getReport('Sheet1')
		self.assertEqual('Section', input[0]['type'])
		self.assertEqual(u'Tình yêu', input[0]['value'])
		self.assertEqual('Paragraph', input[1]['type'])
		self.assertDictEqual({'sheet': 1.1, 'value': 1}, input[1]['value'])


class TestExcelReader(unittest.TestCase):

	def testSanity(self):
		input = ExcelReader('input.xlsx')
		self.assertEqual(['Sheet1'], input.getSheets())
		content = input.getSheetContent('Sheet1')
		self.assertEqual(8, len(content))
		self.assertEqual('Section', content[0][0])
		self.assertEqual(u'Tình yêu', content[0][1])


if __name__ == '__main__':
	unittest.main()