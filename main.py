import httplib2
from googleapiclient import discovery
from oauth2client.service_account import ServiceAccountCredentials
from pprint import pprint
import re
from bs4 import BeautifulSoup
import requests

import wrapper as wr
from styles import Colors as c
from styles import Borders as b

class GoogleTabs:

    def __init__(self):
        # Service-объект, для работы с Google-таблицами
        CREDENTIALS_FILE = 'creds.json'  # имя файла с закрытым ключом
        credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE,
                                                                       ['https://www.googleapis.com/auth/spreadsheets'])
        httpAuth = credentials.authorize(httplib2.Http())
        self.service = discovery.build('sheets', 'v4', http=httpAuth)

        self.spreadsheet_id = '1cbcfJaNbM_OByFE7RkTQHG8pLUgfMdUtnFnF7xQzc0Q'

        self.sheetTitle = "Figures"

    def prepareTable(self, spId, length):

        request = []

        # стили ячеек
        request.append(wr.repeatCells(spId, "A1:G1", c.light_purple))
        request.append(wr.repeatCells(spId,"A2:G{0}".format(length+1), c.white, textFormat= False))

        request.append(wr.updateDimensionProperties(spId, [0, 1], 30, dimension='ROWS'))
        request.append(wr.updateDimensionProperties(spId,[1, length+1], 70, dimension='ROWS'))
        request.append(wr.updateDimensionProperties(spId, [0, 1], 70))
        request.append(wr.updateDimensionProperties(spId, [3, 7], 100))
        request.append(wr.updateDimensionProperties(spId, [1, 2], 200))

        request.append(wr.setCellBorder(spId, "A1:G{0}".format(length+1),bstyleList=b.plain_black))

        return request

    def insertValues(self, spId, listInfo):

        body = {}
        body["valueInputOption"] = "USER_ENTERED"

        data = []

        headers = ['иконка', 'название', 'тип', 'релиз', 'бренд', 'черновик', 'блок']

        for i in range(ord('A'), ord('G') + 1):
            data.append(wr.insertValue(spId, self.sheetTitle+"!{0}1".format(chr(i)), headers[i - ord('A')]))

        for i in range(len(listInfo)):
          formula = '=IMAGE("{0}")'.format(listInfo[i]['icon'])
          data.append(wr.insertValue(spId, self.sheetTitle+"!A{0}".format(i+2), formula))
          formula = '=HYPERLINK("{1}"; "{0}")'.format(listInfo[i]['name'].replace('"', '＂'), listInfo[i]['url'])
          data.append(wr.insertValue(spId, self.sheetTitle+"!B{0}".format(i+2),formula))

          data.append(wr.insertValue(spId, self.sheetTitle+"!C{0}".format(i+2), listInfo[i]['type']))
          data.append(wr.insertValue(spId, self.sheetTitle + "!D{0}".format(i + 2), listInfo[i]['releaseDate']))
          data.append(wr.insertValue(spId, self.sheetTitle + "!E{0}".format(i + 2), listInfo[i]['brand']))

          answer = 'нет' if listInfo[i]['noDraft'] == True else 'да'
          data.append(wr.insertValue(spId, self.sheetTitle + "!F{0}".format(i + 2), answer))
          answer = 'нет' if listInfo[i]['noBlock'] == True else 'да'
          data.append(wr.insertValue(spId, self.sheetTitle + "!G{0}".format(i + 2), answer))

        body["data"] = data

        return body

    def createUpdateChart(self, spId, length):

        body = {}
        body["valueInputOption"] = "USER_ENTERED"
        data = []

        data.append(wr.insertValue(spId, self.sheetTitle + "!I1", "=COUNTUNIQUE(E2:E{0})".format(length+1)))
        body["data"] = data

        self.service.spreadsheets().values().batchUpdate(spreadsheetId=self.spreadsheet_id,
                                                         body=body).execute()

        get = self.service.spreadsheets().values().get(spreadsheetId=self.spreadsheet_id,
                                                       range = self.sheetTitle+"!I1"
                                                      ).execute()
        uniqLength = int(get['values'][0][0])
        data = []

        data.append(wr.insertValue(spId, self.sheetTitle+"!J1", "=UNIQUE(E2:E{0})".format(length+1)))

        for i in range(uniqLength):
            data.append(wr.insertValue(spId, self.sheetTitle+"!K{0}".format(i+1), "=COUNTIF(E2:E{0};J{1})".format(length+1, i+1)))

        body["data"] = data

        self.service.spreadsheets().values().batchUpdate(spreadsheetId=self.spreadsheet_id,
                                                         body=body).execute()

        requestSp = []
        requestSp.append(wr.repeatCells(spId, "J1:K{0}".format(uniqLength),c.white))

        try:
            requestSp.append(wr.addPieChart(spId, spId+5, "J1:J{0}".format(uniqLength), "K1:K{0}".format(uniqLength)))
            self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheet_id,
                                                    body={"requests": requestSp}).execute()
        except:
            requestSp.pop(-1)
            requestSp.append(wr.updateChart(spId, spId+5, "J1:J{0}".format(uniqLength), "K1:K{0}".format(uniqLength)))
            self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheet_id,
                                                    body={"requests": requestSp}).execute()

    def createTable(self, spId, listInfo):

        self.clearTable(spId)

        self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheet_id,
                                                body={"requests": self.prepareTable(spId, len(listInfo))}).execute()

        self.service.spreadsheets().values().batchUpdate(spreadsheetId=self.spreadsheet_id,
                                                           body=self.insertValues(spId, listInfo)).execute()

        self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheet_id,
                                                body={"requests": wr.deleteDuplicates(spId, "A1:G{0}".format(len(listInfo)+1), [0, 6])}).execute()

        self.createUpdateChart(spId, len(listInfo))

    def clearTable(self, spId):

        get = self.service.spreadsheets().values().get(spreadsheetId=self.spreadsheet_id,
                                                       range = self.sheetTitle+"!I1",
                                                       valueRenderOption = "FORMULA"
                                                      ).execute()

        formula, lastItem = get['values'][0][0].replace(')', '').split(":")

        request = []
        request.append(wr.repeatCells(spId, "A1:G1".format(lastItem[1:]), c.white, textFormat=False))
        request.append(wr.setCellBorder(spId, "A1:G{0}".format(lastItem[1:])))
        request.append(wr.updateDimensionProperties(spId, [0, 7], 100))
        request.append(wr.updateDimensionProperties(spId, [0, lastItem[1:]], 21, dimension='ROWS'))

        self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheet_id,
                                                body={"requests": request}).execute()

        self.service.spreadsheets().values().clear(spreadsheetId=self.spreadsheet_id,
                                                         range = self.sheetTitle+"!A1:G{0}".format(lastItem[0:])).execute()


def createFiguresList(page):

    soup = BeautifulSoup(page.text, "html.parser")

    allFigures = soup.findAll('div', class_='apiReturn')

    figuresByMonth = []

    t = allFigures[0].find_next('div')
    current_month = ''

    while t.find_next('div') is not None:
        try:
            class_ = t.get('class')


            if class_[0] == 'fs18':
                current_month = t.text
            elif class_[0] == 'figure':
                figuresByMonth.append([current_month, t])

            t = t.find_next('div')
        except Exception as e:
            break


    figuresDict = {}
    figuresInfo = []

    for fig in figuresByMonth:
        class_ = fig[1].get('class')

        if class_[0] == 'figure':


            figuresDict['noDraft'] = class_[5] == 'draft_'
            figuresDict['noBlock'] = class_[7] == 'block_0'

            src = fig[1].find('a')

            figuresDict['url'] = 'https://ru.myanimeshelf.com'+src.get('href')
            figuresDict['icon'] = 'https://ru.myanimeshelf.com' + src.find('img').get('src')
            figuresDict['name'] = (fig[1].find('div', class_='nameWrap ofh')).find('a').text
            figuresDict['brand'] = fig[1].find('span', class_ = 'fs11 pright20').text.replace('\t', '').replace('\n', '')
            figuresDict['type'] = fig[1].find('span', class_ = 'softColor fs11 pright20').text.replace('\t', '').replace('\n', '')

            figuresDict['releaseDate'] = fig[0]

            figuresInfo.append(figuresDict.copy())

    return figuresInfo

def parcePage(url):
       headers = {'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/51.0.2704.103 Safari/537.36'}

       page = requests.get(url, headers)

       return createFiguresList(page)

def getInfo(url):
    figuresInfo = []

    for i in range(3, 5):
        figuresInfo.extend(parcePage(url.format(i)))

    return figuresInfo

gt = GoogleTabs()

url = 'https://ru.myanimeshelf.com/figures/?page={0}'

figuresInfo = getInfo(url)

gt.createTable(0, figuresInfo)
