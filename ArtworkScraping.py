from selenium import webdriver
import time
import numpy
import pandas as pd
import xlsxwriter
import math

def chooseObjectType(objectNumber):
	#object Number Starts from zero
	global driver

	time.sleep(1)
	dropDownButton = driver.find_element_by_xpath('//*[@id="x-auto-53"]/img')
	dropDownButton.click()

	objectType = driver.find_element_by_xpath('//*[@id="field_10303"]/option[%s +1]' % (objectNumber + 1)) 
	objectType.click()

def openSearch():
	#Locate and open the search menu
	global driver

	startButton = driver.find_element_by_xpath('//*[@id="searchForm"]/p/span[2]/a')
	startButton.click()

def getItemsOnPage():
	#Finds how many pieces of art per page
	global driver

	pageInfo = driver.find_element_by_xpath('//*[@id="content"]/div[1]/ul/li[3]/ul/li[2]/span').text
	numberOnPage = int( pageInfo.split()[-1]) - int(pageInfo.split()[0]) + 1
	return numberOnPage

def getNumberOfPages():
	#Finds the total number of pages
	global driver

	numberOfitemOnPage = int(driver.find_element_by_xpath('//*[@id="content"]/div[1]/ul/li[3]/ul/li[4]/span').text)
	numberOfPages = math.ceil(numberOfitemOnPage / 25) 
	print('Number of pages: ', numberOfPages)
	
	return numberOfPages 

def openItem(itemNumber):
	#Opens individual item
	#item count starts at zero 

	global driver
	titleButtom = driver.find_element_by_xpath('//*[@id="detailListItem-%s"]/dd[1]/ul/li[2]/div[1]/h3/a/span' %itemNumber )
	titleButtom.click()

def writeInfo(heading, content):
	#Write information onto spreadsheet
	global itemNumber
	global worksheet	
	
	infoDict = { 'Artist:': 0 ,'Title:' : 1, 'Location:' : 2, 'Date:' : 3, 'Category/Object Type:' : 4,
				'Material/Technique:' : 5, 'Measure:': 6, 'Catalogue Raisonné:' : 7, 'Copyright:' : 8, 'EK-Title:' : 9,
				'NS Inventar EK-Nr.:' : 10 , 'Museum of Origin:' : 11 , 'Inventory of Origin:' : 12,  'Loss through:' : 13,
				'Date of Loss:': 14 }
	try:
		head = infoDict[heading]
		worksheet.write(itemNumber + 1, head, content)

	except:
		pass
	
def getProvenance():
	#Finds and writes providence info for item.
	global driver 
	global itemNumber
	global workprov

	provButton = driver.find_element_by_xpath('//*[@id="referenceTab-02"]/a')
	provButton.click()

	provText = driver.find_element_by_xpath('//*[@id="collectionReferences-captionBlock"]/div/span[1]').text
	workprov.write(itemNumber + 1, 2, provText)

def collectInformation():
	#item number is how far it needs to be down the sheet, starts from zero
	global driver
	global worksheet 
	global itemNumber
	
	heading = []
	content = []
	textString = ""

	#Rips all headings and content from item page
	for element in driver.find_elements_by_xpath('//*[@id="collectionDetailItem"]/div[2]/ul[1]'):
		textString += element.text
	textString += '\n'
	for element in driver.find_elements_by_xpath('//*[@id="collectionDetailItem"]/div[2]/ul[2]'):
		textString += element.text
	textString += '\n'
	for element in driver.find_elements_by_xpath('//*[@id="collectionDetailItem"]/div[2]/ul[3]'):
		textString += element.text

	#Splits text list
	textList = textString.split('\n')
	if len(textList) %2 != 0:
		textList.pop()

	#Turns text into content or heading
	for num in range(len(textList)):
		if num %2 == 0:
			heading.append(textList[num])
		else:
			content.append(textList[num])
	
	#adds information to spreadsheet
	for num in range(len(heading)):
		writeInfo(heading[num], content[num])
		#adds artist and title to Provenance file
		if heading[num] == 'Artist:':
			workprov.write(itemNumber+1, 0, content[num])
		elif heading[num] == 'Title:':
			workprov.write(itemNumber+1, 1, content[num])
	try:
		getProvenance()
		driver.back()
	except:
		pass
	driver.back()

def getToCorrectPage(num):
	global driver

	for i in range(num):
		nextPage = driver.find_element_by_xpath('//*[@id="pageSetEntries-nextSet"]/a/span') 
		nextPage.click()

def RunForPage():
	#collect data for one page
	global itemNumber
	print(getItemsOnPage())
	for item in range(getItemsOnPage()):
		openItem(item)
		try:
			collectInformation()
		except:
			pass
		itemNumber += 1

def loopThroughPages(startShift):
	#run through pages of a particular type of artwork
	global driver

	for page in range(numberOfPages() - startShift):
		if page > 0:
			nextPage = driver.find_element_by_xpath('//*[@id="pageSetEntries-nextSet"]/a/span') 
			nextPage.click()
			print('Changing to page:', page)
			time.sleep(2)
		RunForPage()
		
def Run(): 
	#run for all types of artwork
	numberOfObjectTypes = 5

	for objectNumber in range(4, numberOfObjectTypes):
		chooseObjectType(objectNumber)
		openSearch()
		getToCorrectPage(501)
		loopThroughPages(501)
		driver.get("http://emuseum.campus.fu-berlin.de/eMuseumPlus?service=ExternalInterface&moduleFunction=search")


#Initialize and format Excel files
excelFileData = xlsxwriter.Workbook('ArtworkData6.xlsx')
excelFileProv = xlsxwriter.Workbook('ArtworkProv6.xlsx')
worksheet = excelFileData.add_worksheet()
workprov = excelFileProv.add_worksheet()

bold = excelFileData.add_format({'bold': True})
bold = excelFileProv.add_format({'bold': True})

#Create headers for each category
worksheet.write('A1','Artist', bold)
worksheet.write('B1', 'Title', bold)
worksheet.write('C1', 'Location', bold)
worksheet.write('D1', 'Date', bold)
worksheet.write('E1', 'Type', bold)
worksheet.write('F1', 'Material/Technique', bold) 
worksheet.write('G1', 'Size', bold)
worksheet.write('H1', 'Catalogue Raisoinne', bold)
worksheet.write('I1', 'Copyright', bold)
worksheet.write('J1', 'Ek-Title', bold)
worksheet.write('K1', 'NS Inventar', bold)
worksheet.write('L1', 'Museum of Origin', bold)
worksheet.write('M1', 'Inventory of Origin', bold)
worksheet.write('N1', 'Loss Through', bold)
worksheet.write('O1', 'Date of Loss', bold)

workprov.write('A1','Artist', bold)
workprov.write('B1', 'Title', bold)
workprov.write('C1', 'Provenance', bold)

#Initialize driver
driver = webdriver.Chrome(r'C:\Users\domar\Documents\Python files\chromedriver')
driver.get("http://emuseum.campus.fu-berlin.de/eMuseumPlus?service=ExternalInterface&moduleFunction=search")
time.sleep(2)
t0 = time.time()

#Run scraping
itemNumber = 0
Run()
t1 = time.time()

#finish
excelFileData.close()
excelFileProv.close()
print('Finished')

T = t1 - t0
print('Time taken: ', T)
driver.quit()
