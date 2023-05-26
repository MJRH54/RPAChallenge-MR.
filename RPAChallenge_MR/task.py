"""Template robot with Python."""

from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.Robocorp.WorkItems import WorkItems
from datetime import datetime
from WebControl import WebControl
from os import getcwd
import time


        
def exportToExcel(wi:WorkItems,excelPathName: str, datesDict: dict, valuesDict: dict):


    """
    Export data dicts to Excel.
    """
    file = Files()
    file.create_workbook(excelPathName)

    titlesColumn = []
    descColumn = []
    pictureColumn = []
    datesColumn = []
    moneyColumn = []
    countColumn = []

    for key_, value in datesDict.items():
        datesColumn.append(value)
        value_ =  valuesDict[key_]
        titlesColumn.append(value_[0])
        descColumn.append(value_[1])
        pictureColumn.append(value_[2])
        countColumn.append(value_[3])
        moneyColumn.append(str(value_[4]))


    file.append_rows_to_worksheet(content=
        {
            'Date':datesColumn,
            'Title':titlesColumn,
            'Description':descColumn,
            'Picture Filename': pictureColumn,
            'Count of Phrases in title or description': countColumn,
            'Amount of money': moneyColumn,
        }, header=True)
    
    file.save_workbook(excelPathName)
    wi.create_output_work_item(files=excelPathName, save=True)

def calculateDates(months:int):
    """
        This function returns the dates from today to (today - months )
    """
    if months < 1:
        months = 1
    today = datetime.now().strftime('%m/%d/%Y')
    currentMonth = int(today.split('/')[0])
    currentYear = int(today.split('/')[2])
    for month in range(months-1):
        if currentMonth > 1:
            currentMonth-=1
        else:
            currentMonth = 12
            currentYear-=1

    toDate = '0' + str(currentMonth) + '/' + '01' + '/' + str(currentYear)
    return (today, toDate)

def orderByDates(dictDates: dict):
    """
        This function returns a order Dates Dict
    """
    newDictDates = {}
    auxList = []
    values = list(dictDates.values())

    for i_value in range(len(values)):
        try:
            date = values[i_value].split(' ')
            abr = date[0][0:3]
            auxList.append((datetime.strptime(str(abr) + ' ' + str(date[1]),'%b %d').strftime('%b %d'),values[i_value],i_value))
        except:
            continue
    
    auxList.sort(reverse=True)

    for value in auxList:
        newDictDates[value[2]] = value[1]

    return newDictDates

def main():
    #Initializing Work Items
    wi = WorkItems()
    wi.get_input_work_item()
    text = wi.get_work_item_variable('SEARCH_PHRASE')
    categories = wi.get_work_item_variable('CATEGORIES')
    months = wi.get_work_item_variable('MONTHS')

    #Initializing Driver
    driver = Selenium()
    webController = WebControl()
    #LocatorsDict (Is easier to control web navigation in this form) 
    locDict = {
        'SEARCH_ICON':"//button[@data-test-id = 'search-button']",
        'SEARCH_INPUT':"//input[@name = 'query']",
        'GO':"//button[@data-test-id = 'search-submit']",
        'SECTION_MENU': "//button[@data-testid = 'search-multiselect-button']",
        'SECTION_CBOX': "//ul[@data-testid = 'multi-select-dropdown-list']/",
        'DATES_MENU':"//button[@data-testid = 'search-date-dropdown-a']",
        'DATES_LINK':"//button[@value = 'Specific Dates']",
        'DATE1':'//*[@id="startDate"]',
        'DATE2':'//*[@id="endDate"]',
        'SHOW_MORE': "//div[@class = 'css-1t62hi8']",
        'IMAGE_BASE': "//ol[@data-testid = 'search-results']/",
    }

    try:
        driver.open_available_browser('https://www.nytimes.com/') #Open Browser
        webController.clickElement(driver,locDict['SEARCH_ICON']) #Make click in searchIcon

        #Send the PHRASE
        webController.send_text_to_webBrowser(driver,locDict['SEARCH_INPUT'],text)    #Send the PRHASE into input
        webController.clickElement(driver,locDict['GO'])                              #Click on GO button.

        #Applying Filters
        #Selecting months:
        webController.clickElement(driver,locDict['DATES_MENU'])                      #Click on DATES MENU
        webController.clickElement(driver,locDict['DATES_LINK'])                      #Custom DATES
        datesSender = calculateDates(months)                                          #Calculating Dates to filter
        webController.clickElement(driver,locDict['DATE1'])                           #Click on DATE START INPUT
        webController.send_text_to_webBrowser(driver,locDict['DATE1'],datesSender[1]) #Send start Date
        webController.send_key(driver, locDict['DATE1'])                              #Send Enter
        webController.clickElement(driver,locDict['DATE2'])                           #Click on DATE END INPUT
        webController.send_text_to_webBrowser(driver,locDict['DATE2'],datesSender[0]) #Send end Date
        webController.send_key(driver, locDict['DATE2'])                              #Send Enter
        time.sleep(1)

        #Selecting categories:
        categories =  categories + 1 if categories > 1 else 1  #Any Section is always selected
        webController.clickElement(driver,locDict['SECTION_MENU']) #Select Section Menu
        for i in range(categories):
            webController.clickElement(driver,locDict['SECTION_CBOX'] + 'li[{}]/label/input'.format(i+1))
        webController.send_key(driver, locDict['SECTION_MENU'], 'ESC') #Send ESC
        time.sleep(1)

        #Extract Information:
  
        #Charging all elements in the current filter page

        pageHeight = webController.execute_script(driver,'return document.body.scrollHeight')       #Calculate the height size in the current filter page
        webController.execute_script(driver, f'window.scrollTo(0,{pageHeight})') #Go to end of the page
        print('ANTES DE HACER CLICK')
      
        while driver.is_element_visible(locDict['SHOW_MORE']):                                      #Click in show more while show more button exists
            print('In While')
            webController.clickElement(driver,locDict['SHOW_MORE'])
            time.sleep(0.5)
            pageHeight = webController.execute_script(driver,'return document.body.scrollHeight')
            webController.execute_script(driver, f'window.scrollTo(0,{pageHeight})')
            time.sleep(0.5)
        pageHeight = webController.execute_script(driver,f'window.scrollTo(0,0)')                   #Go to start of the page

        #Extracting the info from notes:
        datesDict = {}
        valuesDict = {}
        picturesCont = 0
        imagesList = []
        #Get all nodes in the list of browser. 
        notesNodesLength = webController.execute_script(driver,"return document.getElementsByClassName('css-46b038')[0].children[0].children.length")
        for i in range(notesNodesLength):
            try:
                #Initializing counters:
                countPhrases = 0
                #Initializing flag Money
                flagMoney = False
                #Get Date
                dateIntoNote = webController.execute_script(driver,"return document.getElementsByClassName('css-46b038')[0].children[0].children[{}].children[0].children[0].innerText".format(i))
                if 'advertisement' in str(dateIntoNote).lower():
                    datesDict[i] = ''
                    valuesDict[i] = ''
                    continue
                #Save the date value into dictDates
                datesDict[i] = dateIntoNote
                #Get Title
                title = webController.execute_script(driver,"return document.getElementsByClassName('css-46b038')[0].children[0].children[{}].children[0].children[1].children[0].children[1].children[0].innerText".format(i))
                #Calculating the children length to evaluate if description exists
                lengthChildrenElements = int(webController.execute_script(driver,"return document.getElementsByClassName('css-46b038')[0].children[0].children[{}].children[0].children[1].children[0].children[1].children.length".format(i)))
                #Get Description
                if lengthChildrenElements > 1:
                    description = webController.execute_script(driver,"return document.getElementsByClassName('css-46b038')[0].children[0].children[{}].children[0].children[1].children[0].children[1].children[1].innerText".format(i))
                else:
                    description = ''
                #Get Image FileName (URL_SRC)
                picturesCont+=1
                imageName = 'pictureNote_{}.png'.format(picturesCont)
                #Download Picture:
                fileName = webController.download_image(driver,locDict['IMAGE_BASE']+'li[{}]/div/div/figure/div/img'.format(i+1),imageName)
                #count of frases in the tile
                countPhrases += title.lower().count(text.lower())
                #count of frases in the description
                countPhrases += description.lower().count(text.lower())
                #Changing the flagMoney value
                if ('$' in title.lower() or 'usd' in title.lower()) or ('$' in description.lower() or 'usd' in description.lower()):
                    flagMoney = True
                #Save the data values into dictValues
                valuesDict[i] = (title,description,imageName,countPhrases,flagMoney)
                #Appending image whole path into imageList
                imagesList.append(fileName)

            except BufferError as e:
                exit(e)

        newDatesDict = orderByDates(datesDict)      #Obtain the dates in order to export excel.

        #Export data to excel.
        exportToExcel(wi,getcwd()+'/OUTPUT.xlsx',newDatesDict,valuesDict)
        #Save the images as a WI output
        wi.create_output_work_item(files=imagesList, save=True)

    except BufferError as e:
        exit(e)
    pass




if __name__ == "__main__":
    main()
