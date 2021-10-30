class Colors:

    light_purple = {'blue': 0.8392157, 'green': 0.654902, 'red': 0.7058824} #Оформление
    light_green = {'blue': 0.65882355, 'green': 0.84313726, 'red': 0.7137255} #Уплачено
    light_red = {'blue': 0.4, 'green': 0.4, 'red': 0.8784314} #Неуплачено
    light_yellow = {'blue': 0.6, 'green': 0.8980392, 'red': 1} #Отсрочка/Цвет окаймляющей рамки
    light_blue = {'blue': 0.9098039, 'green': 0.77254903, 'red': 0.62352943} #Результрующее число при расчётах

    white = {'blue': 1, 'green': 1, 'red': 1}
    black = {'blue': 0, 'green': 0, 'red': 0}
    green = {'green': 1}

class Borders:

    '''
     Задаёт стили границ ячеек
    '''

    plain_black =  { "style" : "SOLID",
                     "color" : Colors.black
                    }

    thick_yellow = { "style" : "SOLID_THICK",
                     "color" : Colors.light_yellow
                    }

    thick_green = { "style" : "SOLID_THICK",
                     "color" : Colors.green
                    }

    no_border = {"style": "NONE"}