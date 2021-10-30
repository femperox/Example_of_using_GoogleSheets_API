from styles import Borders

def toRangeType(spId, range):
    '''
    Конвертирует диапозон ячеек в словарь

    :param spId: айди листа в таблице
    :param range: диапозон ячеек в формате "В1:С45" - пример
    :return: возвращает словарь для json запроса
    '''

    rangeType = {}
    rangeType["sheetId"] = spId

    try: # Формат *#:*#
      startCell, endCell = range.split(":")[0:2]

      rangeType["startRowIndex"] = int(startCell[1:]) -1
      rangeType["startColumnIndex"] = ord(startCell[0]) - ord('A')

      rangeType["endRowIndex"] = endCell[1:]
      rangeType["endColumnIndex"] = ord(endCell[0]) - ord('A') + 1
    except: # Формат *#
      rangeType["rowIndex"] = int(range[1:]) -1
      rangeType["columnIndex"] = ord(range[0]) - ord('A')

    return rangeType

def setCellBorder(spId, range, all_same = True, only_outer = False, bstyleList = Borders.no_border):
    '''
    подготовка json запроса для обрамления диапозона ячеек границами с определёнными стилями

    :param spId: айди листа в таблице
    :param range: диапозон ячеек в формате "В1:С45" - пример
    :param all_same: Применять ли ко всем сторонам ячейки одинаковый стиль границ. По умолчанию - да
    :param bstyleList: Список стилей границ. По умолчанию - "без граниу"
    :return: возвращает json запрос
    '''

    if not isinstance(bstyleList, list): bstyleList = [bstyleList]

    if len(bstyleList) < 6:
        bstyleList = bstyleList + [Borders.no_border]*(4-len(bstyleList))

    if all_same:
        bstyleList = [bstyleList[0]]*6

    if only_outer:
        bstyleList[4] = Borders.no_border
        bstyleList[5] = Borders.no_border

    request = {"updateBorders":
                { "range" : toRangeType(spId, range),
                  "top": bstyleList[0],
                  "bottom": bstyleList[1],
                  "left": bstyleList[2],
                  "right": bstyleList[3],
                  "innerHorizontal": bstyleList[4],
                  "innerVertical": bstyleList[5]
                }
              }
    return request

def insertValue(spId, range, text ="", majorDime = "ROWS"):
    '''
    подготовка json запроса для заполнения ячеек заданным текстом

    :param spId: айди листа в таблице
    :param range: диапозон ячеек в формате "НазваниеЛиста!В1:С45" - пример
    :param text: текст ячейки
    :param majorDime: формат заполнения ячеек
    :return: возвращает json запрос
    '''

    data = {}

    data["range"] = range
    data["majorDimension"] = majorDime
    data["values"] = [[text]]

    return data

def repeatCells(spId, range, color, hali = "CENTER", textFormat = True):
    '''
    подготовка json запроса для оформления диапозона ячеек определённым стилем

    :param spId: айди листа в таблице
    :param range: диапозон ячеек в формате "В1:С45" - пример
    :param color: цвет ячейки
    :return: возвращает json запрос
    '''

    request = { "repeatCell":
                { "range": toRangeType(spId, range),
                  "cell": { "userEnteredFormat": toUserEnteredFormat(color, hali = hali, textFormat=textFormat) },
                  "fields": "userEnteredFormat"
                }

    }

    return request

def toUserEnteredFormat(color, hali = 'CENTER', vali = 'MIDDLE', textFormat = 'True'):
    '''
    составляет словарь стиля ячейки

    :param color: цвет
    :param hali: вертикальное выравнивание
    :param vali: горизонтальное выравнивание
    :param textFormat: формат текста
    :return: возвращает словарь для json запроса
    '''
    userEnteredFormat = {}

    userEnteredFormat['horizontalAlignment'] = hali
    userEnteredFormat['verticalAlignment'] = vali
    userEnteredFormat['textFormat'] = {'bold': textFormat}
    userEnteredFormat["backgroundColor"] = color
    userEnteredFormat["wrapStrategy"] = "WRAP"

    return userEnteredFormat

def toDimensionRange(spId, range, dimension = 'COLUMNS'):

    dimensionRange = {}
    dimensionRange["sheetId"] = spId
    dimensionRange["dimension"] = dimension
    dimensionRange["startIndex"] =  range[0]
    dimensionRange["endIndex"] = range[1]

    return dimensionRange

def updateDimensionProperties(spId, range, size, dimension = 'COLUMNS'):

    request = { "updateDimensionProperties":
                { "properties": { "pixelSize": size},
                   "range" : toDimensionRange(spId, range, dimension= dimension),
                   "fields": "pixelSize"
                }
              }

    return request

def deleteDuplicates(spId, gridRange, dimeRange):

    request = { "deleteDuplicates":
                { "range": toRangeType(spId, gridRange),
                  "comparisonColumns": [toDimensionRange(spId, dimeRange)]
                }
              }

    return request

def createDomainsForChart(spId, range):

    domain = { "sourceRange": {  "sources": [ toRangeType(spId, range) ]} ,
               "aggregateType": "COUNT"
             }

    return domain

def createSeriesForChart(spId, rangeList, targetAxis = "BOTTOM_AXIS"):

    series = []

    for range in rangeList:

        ser = { "sourceRange": {"sources": [toRangeType(spId, range)]},
                "aggregateType": "COUNT"
              }

        series.append(ser.copy())

    return series

def chartSpec(spId, domainRange, seriesRange, legendPosition = "LABELED_LEGEND"):
    spec = {"title": "Статистика брендов",
                              "pieChart":
                                  {"legendPosition": legendPosition,
                                   "domain": createDomainsForChart(spId, domainRange),
                                   "series": (createSeriesForChart(spId, [seriesRange]))[0]
                                   }

            }
    return spec

def addPieChart(spId, chartId, domainRange, seriesRange, legendPosition = "LABELED_LEGEND"):
    request = {"addChart":
                   {"chart":
                        {"chartId": chartId,
                         "spec": chartSpec(spId, domainRange, seriesRange),
                         "position": {"newSheet": True}
                         }
                    }
               }

    return request

def updateChart(spId, chartId, domainRange, seriesRange):

    request = {  "updateChartSpec":
                  { "chartId" : chartId,
                    "spec": chartSpec(spId, domainRange, seriesRange)
                  }
              }

    return request

def deleteRange(spId, range):
    request = { "deleteRange":
                {   "range": {  toRangeType(spId, range)  },
                    "shiftDimension": "ROWS"
                }
              }

    return request
