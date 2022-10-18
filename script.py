import json
from datetime import date
from yattag import Doc
import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import urllib.request
import zipfile
import os
import shutil
import csv

settings = None
currentDate = None

def readJson(fileName):
    with open(fileName, encoding='utf-8', mode='r') as file:
        return json.load(file)
    
def writeJson(fileName, data):
    with open(fileName, encoding='utf-8', mode='w') as file:
        json.dump(data, file, ensure_ascii=False, indent=4)
        
def getPartOfMonth(d):
    firstSalaryDay = date.fromisoformat(settings['salary']['firstSalaryDay'])
    diff = d - firstSalaryDay
    if diff.days < 14: return 1
    if (((diff.days - (diff.days % 14))/14) % 2) == 0: return 1
    return 2

def todayIsSalaryDay():
    firstSalaryDay = date.fromisoformat(settings['salary']['firstSalaryDay'])
    diff = currentDate - firstSalaryDay
    if diff.days < 14: return False
    if diff.days % 14 == 0: return True
    return False    

def getExpensesForCurrentPartOfMonth(part):
    currentPartOfMonth = getPartOfMonth(currentDate)
    expenses = []
    if part == "weekly": return settings['expenses'][part]
    for expense in settings['expenses'][part]:
        if expense['salaryPart'] == currentPartOfMonth: expenses.append(expense)
    return expenses

def generateReportForCurrentPartOfMonth():
    exchangeRate = getCurrencyExchangeRate()
    currentSalary = 0
    if getPartOfMonth(currentDate) == 1: currentSalary = settings['salary']['firstSalaryPartDollarAmount']
    else: currentSalary = settings['salary']['secondSalaryPartDollarAmount']
    
    doc, tag, text, line = Doc().ttl()
    
    total = 0

    with tag('html'):
        with tag('body'):
            line('p', f'Привет! Сегодня {currentDate} и это день зарплаты!')
            line('p', f'Зарплата USD: {currentSalary}')
            line('p', f'Средний курс USD: {exchangeRate}')
            line('p', f'Зарплата RUB: {currentSalary * exchangeRate}')
            with tag('table'):
                mainTotal = 0
                for exp in getExpensesForCurrentPartOfMonth('main'):
                    with tag('tr'):       
                        with tag('td'):
                            text(f"{(exp['caption'])}")
                        with tag('td'):
                            text(f"{(exp['amountRub'])}")
                    mainTotal += exp['amountRub']
            line('p', f'итого: {mainTotal}')
            total += mainTotal
            with tag('table'):
                montlyTotal = 0
                for exp in getExpensesForCurrentPartOfMonth('monthly'):
                    with tag('tr'):       
                        with tag('td'):
                            text(f"{(exp['caption'])}")
                        with tag('td'):
                            text(f"{(exp['amountRub'])}")
                    montlyTotal += exp['amountRub']
            line('p', f'итого: {montlyTotal}')
            total += montlyTotal
            with tag('table'):
                weeklyTotal = 0
                for exp in getExpensesForCurrentPartOfMonth('weekly'):
                    with tag('tr'):       
                        with tag('td'):
                            text(f"{(exp['caption'])}")
                        with tag('td'):
                            text(f"{(exp['amountRub'])}")
                    weeklyTotal += exp['amountRub']
            line('p', f'итого: {weeklyTotal*2}')
            total += weeklyTotal*2
        buffer = (currentSalary * exchangeRate - total) * (settings['salary']['bufferPercentageFromBalance'])/100
        line('p', f"Tак же оставить на не предвиденные расходы : {buffer}")
        line('p', f'Остаток на фин обязательства/накопление/подушку : {(currentSalary * exchangeRate - total) - buffer}')
    return doc.getvalue()

def generateReportForWeek():
    doc, tag, text, line = Doc().ttl()
    with tag('html'):
        with tag('body'):
            line('p', f'Привет! Сегодня {currentDate} и плановые траты на неделю!')
            with tag('table'):
                total = 0
                for exp in getExpensesForCurrentPartOfMonth('weekly'):
                    with tag('tr'):       
                        with tag('td'):
                            text(f"{(exp['caption'])}")
                        with tag('td'):
                            text(f"{(exp['amountRub'])}")
                    total += exp['amountRub']
            line('p', f'Итого: {total}')
    return doc.getvalue()

def sendReport(content, subject):
    print(content)
    message = MIMEMultipart("alternative")
    message["Subject"] = subject
    message["From"] = settings['email']['senderEmail']
    message["To"] = settings['email']['receiverEmail']
    
    messageContent = MIMEText(content, "html")
    
    message.attach(messageContent)
    
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(settings['email']['smtpAddress'], settings['email']['smtpPort'], context=context) as server:
        server.login(settings['email']['senderEmail'], settings['email']['senderPwd'])
        server.sendmail(
            settings['email']['senderEmail'], settings['email']['receiverEmail'], message.as_string()
        )
        
def todayIsDayForMontlyReport():
    previousDate = None
    if settings['reports']['previousMonthlyReportDate'] == "": return True
    else: previousDate = date.fromisoformat(settings['reports']['previousMonthlyReportDate'])
    
    if currentDate.year == previousDate.year and currentDate.month == previousDate.month and currentDate.day == previousDate.day: return False

    if getPartOfMonth(currentDate) == getPartOfMonth(previousDate): return False
    return todayIsSalaryDay()

def todayIsDayForWeeklyReport():
    previousDate = None
    if settings['reports']['previousWeeklyReportDate'] == "": return True
    else: previousDate = date.fromisoformat(settings['reports']['previousWeeklyReportDate'])
    
    if currentDate.year == previousDate.year and currentDate.month == previousDate.month and currentDate.day == previousDate.day: return False
    
    return currentDate.weekday() == 0

def processMonthlyReport():
    if not todayIsDayForMontlyReport(): return
    report = generateReportForCurrentPartOfMonth()
    sendReport(report, "monthly report")
    
    global settings
    settings['reports']['previousMonthlyReportDate'] = currentDate.strftime("%Y-%m-%d")
    writeJson('settings.json', settings)
    settings = readJson('settings.json')
    
def processWeeklyReport():
    if not todayIsDayForWeeklyReport(): return
    report = generateReportForWeek()
    sendReport(report, "weekly report")

    global settings
    settings['reports']['previousWeeklyReportDate'] = currentDate.strftime("%Y-%m-%d")
    writeJson('settings.json', settings)
    settings = readJson('settings.json')
    
def getCurrencyExchangeRate():
    with urllib.request.urlopen(settings['exchangeRates']['url']) as response:
        data = response.read()
        with open(f"{(settings['exchangeRates']['fileMask'])}.zip", "wb") as file:
            file.write(data)
            with zipfile.ZipFile(f"{(settings['exchangeRates']['fileMask'])}.zip", 'r') as zip:
                zip.extractall(settings['exchangeRates']['fileMask'])
    
    #lets find ids for our currencies
    exchangeFromId = 0
    exchangeToId = 0
    currencyIdentifiersFile = open(f"{(settings['exchangeRates']['fileMask'])}/bm_cy.dat")
    currencyIdentifiersReader = csv.reader(currencyIdentifiersFile, delimiter = ';')
    for dataLine in currencyIdentifiersReader:
        if dataLine[2] == settings['exchangeRates']['fromKey']: exchangeFromId = dataLine[0]
        if dataLine[2] == settings['exchangeRates']['toKey']: exchangeToId = dataLine[0]
    currencyIdentifiersFile.close()
    
    #lets find exhange rate four our currencies by id
    exchangers = []
    currencyRatesFile = open(f"{(settings['exchangeRates']['fileMask'])}/bm_rates.dat")
    currencyRatesReader = csv.reader(currencyRatesFile, delimiter = ';')
    for dataLine in currencyRatesReader:
        if dataLine[0] == exchangeFromId and dataLine[1] == exchangeToId: exchangers.append(dataLine)
    currencyRatesFile.close()
    
    exchangers.sort(reverse=True, key=reviews)
    temp = exchangers[:5]
    
    totl = 0.0
    for el in temp: totl += float(el[4])   
    
    #delete zip file and related directory
    os.remove(f"{(settings['exchangeRates']['fileMask'])}.zip")
    shutil.rmtree(settings['exchangeRates']['fileMask'])
    return totl/len(temp)   

def reviews(el):
    temp = el[6].split('.')
    return int(temp[1])

if __name__ == "__main__":
    settings = readJson("settings.json")
    currentDate = date.today()    
    processMonthlyReport()
    processWeeklyReport()