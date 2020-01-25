#! python3
# inGroovesDownloads.py - downloads InGrooves files from the web
# finds and formats the files
# puts the files in their folder and unzips the files


import shutil, zipfile, os, time, datetime, re, pandas as pd, openpyxl, calendar, xlrd
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait

# gets the current username to pull from the downloads folder
path = os.path.abspath('C:\\Users')
pathFolders = os.listdir(path)
for i in pathFolders:
    if i == 'chehem':
        userName = i
    elif i == 'Kevqua':
        userName = i
    elif i == 'keleng':
        userName = i


print('Downloads Folder file path created.')


# get the dates for the files
today = datetime.date.today()
year = datetime.datetime.today().year
month = today.month

if month == 1:
    year = year -1
    month = 12
else:
    month = month -1

fileDate = str(year) + '-' + str(month) + '-15'


print('Gathered dates.')

# open chrome browser and go to site
browser = webdriver.Chrome()
browser.get('https://statements.InGrooves.com')
browser.maximize_window()
##signInElem = browser.find_element_by_class_name('sign-in-btn').click()
##
### switch tabs
##for handle in browser.window_handles:
##    browser.switch_to.window(handle)
# logging in
time.sleep(2)
emailElem = browser.find_element_by_id('email')
time.sleep(2)
emailElem.send_keys('username')
time.sleep(2)
passwordElem = browser.find_element_by_id('password')
passwordElem.send_keys('password')
time.sleep(2)
signInButtonElem = browser.find_element_by_id('login').click()
time.sleep(20)

### go to statements
##statementsElem = browser.find_element_by_xpath("//*[@id='main-nav']/ul[1]/li[5]/a").click()
##time.sleep(10)
# choose the date range
dateRangeElem = browser.find_element_by_xpath("//*[@id='daterange-dropdown']").click()
time.sleep(1)
rangeElem = browser.find_element_by_xpath('/html/body/div/div[2]/div/div/div[2]/div[1]/div[1]/div[1]/div/div/div/div[1]/div/select').click()
time.sleep(1)
lastMonthElem = browser.find_element_by_xpath(".//*[contains(text(), 'Last Month')]").click()
time.sleep(1)
applyElem = browser.find_element_by_xpath(".//*[contains(text(), 'APPLY')]").click()
time.sleep(5)
# pick the reports
concordJazzElem = browser.find_element_by_xpath(".//*[contains(text(), 'Concord Jazz')]").click()
time.sleep(1)
concordPicanteElem = browser.find_element_by_xpath(".//*[contains(text(), 'Concord Picante')]").click()
time.sleep(1)
concordRecordsElem = browser.find_element_by_xpath(".//*[contains(text(), 'Concord Records.')]").click()
time.sleep(1)
headsUpElem = browser.find_element_by_xpath(".//*[contains(text(), 'Heads Up')]").click()
time.sleep(1)
mcgJazzElem = browser.find_element_by_xpath(".//*[contains(text(), 'MCG Jazz')]").click()
time.sleep(1)
peakRecordsElem = browser.find_element_by_xpath(".//*[contains(text(), 'Peak Records / Telarc')]").click()
time.sleep(1)
razorAndTieElem = browser.find_element_by_xpath(".//*[contains(text(), 'Razor & Tie Direct LLC')]").click()
time.sleep(1)
rounderElem = browser.find_element_by_xpath(".//*[contains(text(), 'Rounder')]").click()
time.sleep(1)
#telarcElem = browser.find_element_by_xpath('/html/body/div/div[2]/div/div/div[2]/div[3]/div/div[1]/div[2]/div/div[8]/div[1]/div/span').click()
telarcElem = browser.find_element_by_xpath("./html/body/div/div[2]/div/div/div[2]/div[3]/div/div[1]/div[2]/div/div[9]/div[2]/div/span").click()
time.sleep(1)
telarcHUElem = browser.find_element_by_xpath(".//*[contains(text(), 'Telarc/Heads Up International a division of Concord Music Group Inc.')]").click()
time.sleep(1)
tigressRecordsElem = browser.find_element_by_xpath(".//*[contains(text(), 'Tigress Records')]").click()
time.sleep(1)
windUpRecordsElem = browser.find_element_by_xpath(".//*[contains(text(), 'Wind-up Records')]").click()
time.sleep(1)
downloadAllElem = browser.find_element_by_class_name('download-btn.btn.btn-default.ng-scope').click()
time.sleep(20)



print('Files Downloaded')



# move files, create destination folder and unzip the files
sourcePath = os.path.abspath('C:\\Users\\' + userName + '\Downloads')
sourceFiles = os.listdir(sourcePath)
os.makedirs(r'\\INgrooves\\'+ str(year) +'\\OriginalFiles\\' + str(year) + '-' + '{:02d}'.format(month)) 
destinationPath = os.path.abspath(r'\\INgrooves\\'+ str(year) +'\\OriginalFiles\\' + str(year) + '-' + '{:02d}'.format(month))

for file in sourceFiles:
    if file.endswith('.zip'):
         shutil.move(os.path.join(sourcePath, file), os.path.join(destinationPath, file))

print('Files Moved To Folder')

for file in sourceFiles:
    if file.endswith('.zip'):
         filename = os.path.join(destinationPath, file)
         zip = zipfile.ZipFile(filename)
         zip.extractall(destinationPath)

print('Are the files unzipped?')
answer = input('Y or N:')

while answer != 'y' or if answer != 'Y':
    print('is it ready now?')
    answer = input('Y or N:')


         
destinationPath = r'S:\INgrooves\\'+ str(year) +'\\OriginalFiles\\' + str(year) + '-' + '{:02d}'.format(month)+ '\\Statements'
destinationFiles = os.listdir(destinationPath)


for file in destinationFiles:
    if file.endswith('.zip'):
         filename = os.path.join(destinationFiles, file)
         zip = zipfile.ZipFile(filename)
         zip.extractall(destinationPath)
         
         
print('Files Unzipped')
print('Download And File Movement Complete')


#########################################Formatting###########################################

#Directory Variables
destPath = r'S:\INgrooves\\'+ str(year) +'\\Formatted-Completed'


#Lists
 
neededFiles=['Concord_Jazz_'+str(year)+str(month)+'_DSR.xlsx'
            ,'Concord_Picante_'+str(year)+str(month)+'_DSR.xlsx'
            ,'Concord_Records__'+str(year)+str(month)+'_DSR.xlsx'
            ,'Heads_Up_'+str(year)+str(month)+'_DSR.csv'
            ,'MCG_Jazz_'+str(year)+str(month)+'_DSR.xlsx'
            ,'Peak_Records___Telarc_'+str(year)+str(month)+'_DSR.xlsx'
            ,'Razor___Tie_Direct_'+str(year)+str(month)+'_DSR.csv'
            ,'Rounder_'+str(year)+str(month)+'_DSR.xlsx'
            ,'Telarc_'+str(year)+str(month)+'_DSR.csv'
            ,'Wind-up_Records_'+str(year)+str(month)+'_DSR.xlsx']

#Functions
def formatExcel(xlFile):
    data = pd.read_excel(os.path.join(sourcePath, xlFile), 'Digital Sales Details')
    data.dropna(axis = 1, how = 'all', inplace = True )
    data.dropna(axis = 0, how = 'all', inplace = True )
    data.drop(data[data['Sales Classification'] == 'Total'].index, axis = 0, inplace=True)
    data['UPC/EAN'] = data['UPC/EAN'].astype(str)
    data['UPC/EAN'] = data['UPC/EAN'].str.rstrip('.0')
    data['FileDate'] = fileDate
    data['LabelID'] = data['Label']
    return data

def formatRazor(csvFile):
    data = pd.read_csv(os.path.join(sourcePath, csvFile))
    data.dropna(axis = 1, how = 'all', inplace = True )
    data.dropna(axis = 0, how = 'all', inplace = True )
    data.drop(data[data['Period'] == 'Totals'].index, axis = 0, inplace=True)
    data['UPC/EAN'] = data['UPC/EAN'].astype(str)
    data['UPC/EAN'] = data['UPC/EAN'].str.rstrip('.0')
    data.rename(columns = {'US$ After Fees':'Net Dollars after Fees'} , inplace = True)
    data['FileDate'] = fileDate
    data['LabelID'] = 'Razor'
    return data

def formatCSV(csvFile):
    data = pd.read_csv(os.path.join(sourcePath, csvFile))
    data.dropna(axis = 1, how = 'all', inplace = True )
    data.dropna(axis = 0, how = 'all', inplace = True )
    data.drop(data[data['Sales Classification'] == 'Total'].index, axis = 0, inplace=True)
    data.drop(data[data['Period'] == 'Totals'].index, axis = 0, inplace=True)
    data['UPC/EAN'] = data['UPC/EAN'].astype(str)
    data['UPC/EAN'] = data['UPC/EAN'].str.rstrip('.0')
    data.rename(columns = {'US$ After Fees':'Net Dollars after Fees'} , inplace = True)
    data['FileDate'] = fileDate
    data['LabelID'] = data['Label']
    return data

#Data frames    
formattedDF = pd.DataFrame()
razorDF = pd.DataFrame()

for file in neededFiles:
    if file.endswith('.csv') and file[:5] == 'Razor':
        data = formatRazor(file)
        razorDF = pd.DataFrame(data) 
        print(file + ' has been formatted!')
        print('$' + data['Net Dollars after Fees'].sum())
    elif file.endswith('.csv'):
        data = formatCSV(file)
        formattedDF = formattedDF.append(data, sort = False)
        print(file + ' has been formatted!')
        print('$' + data['Net Dollars after Fees'].sum())
    elif file.endswith('.xlsx'):
        data = formatExcel(file)
        formattedDF = formattedDF.append(data, sort = False)
        print(file + ' has been formatted!')
        print('$' + data['Net Dollars after Fees'].sum())

formattedDF.to_excel(os.path.join(destPath, fileDate[:7] + '_IngroovesImport.xlsx'),index = False)
razorDF.to_excel(os.path.join(destPath, fileDate[:7] + '_RazorImport.xlsx'),index = False)

print('---------------------------------------')            
print(formattedDF.LabelID.unique())
print(formattedDF['FileDate'].count())
print('---------------------------------------')            
print(razorDF.LabelID.unique())
print(razorDF['FileDate'].count())
print('---------------------------------------')



#Old Code
#########################################RAZOR AND TIE#######################################
### open chrome browser and go to site
##browser = webdriver.Chrome()
##browser.get('https://statements.InGrooves.com')
##browser.maximize_window()
####signInElem = browser.find_element_by_class_name('sign-in-btn').click()
####
##### switch tabs
####for handle in browser.window_handles:
####	browser.switch_to.window(handle)
##
### logging in	
##emailElem = browser.find_element_by_id('email')
##time.sleep(2)
##emailElem.send_keys('username')
##time.sleep(2)
##passwordElem = browser.find_element_by_id('password')
##passwordElem.send_keys('password')
##time.sleep(2)
##signInButtonElem = browser.find_element_by_id('login').click()
##time.sleep(15)
##
##### go to statements
####statementsElem = browser.find_element_by_xpath("//*[@id='main-nav']/ul[1]/li[3]/a").click()
####time.sleep(10)
### choose the date range
##dateRangeElem = browser.find_element_by_xpath("//*[@id='daterange-dropdown']").click()
##time.sleep(1)
##rangeElem = browser.find_element_by_xpath('/html/body/div/div[2]/div/div/div[2]/div[1]/div[1]/div[1]/div/div/div/div[1]/div/select').click()
##time.sleep(1)
##lastMonthElem = browser.find_element_by_xpath(".//*[contains(text(), 'Last Month')]").click()
##time.sleep(1)
##applyElem = browser.find_element_by_xpath(".//*[contains(text(), 'APPLY')]").click()
##time.sleep(5)
### pick the reports
##razorAndTieElem = browser.find_element_by_xpath(".//*[contains(text(), 'Razor & Tie Direct LLC')]").click()
##time.sleep(1)
##downloadAllElem = browser.find_element_by_class_name('download-btn.btn.btn-default.ng-scope').click()
##time.sleep(10)
##
##print('Razor And Tie Downloaded')
         

