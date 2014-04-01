from horoscope import HoroInput
from horoscope import HoroReport
from horoscope import HoroDocument
import logging

def getLogger(name):
	logger = logging.getLogger(name)
	logger.setLevel(logging.INFO)
	formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
	streamHandler = logging.StreamHandler()
	streamHandler.setFormatter(formatter)
	logger.addHandler(streamHandler)
	return logger

mainLogger = getLogger('Main')

def testLogger():
	mainLogger.info('Printing logg')

def printDefaultReport():
	input = HoroInput()
	input.load('input.xlsx')
	report = HoroReport()
	report.loadDataSource('AstroReport_Content.xlsx')
	
	# print all sheets to docx	
	reportNames = input.getReportNames()
	for name in reportNames:
		mainLogger.info('Generate report for ' + name)
		content = report.loadContentUsing(input.getReport(name))
		doc = HoroDocument()
		doc.generate(content)
		doc.save(name + '.docx')

if __name__ == '__main__':
	printDefaultReport()
	# testLogger()