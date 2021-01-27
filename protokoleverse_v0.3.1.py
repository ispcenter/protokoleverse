'''protokoleverse v0.31'''

import os
import pandas as pd
from shutil import copy2 as copy

def stendNumberDefiner(razmer):
	return {
	razmer < 50: 'razmerError',
	razmer == 50: 1,
	60 <= razmer < 90: 2,
	90 <= razmer <= 100: 3,
	110 <= razmer <= 130: 4,
	130 < razmer: 5
	}[True]

def find_all(name, path):
    result = []
    for root, dirs, files in os.walk(path):
        if name in files:
            result.append(os.path.join(root, name))
    return result

def protFinder(stendNumber, year, month, serialNumber):
	'''in path
	conditions'''
	pass
	
def parametersChecking():
	'''Checking parameters in protokol on stend'''
	pass

#0 Setup
taskFound = False
task_models = []
protList = []
container = r'D:\coded\protokoleverse\Container'
monthWorder = {
	'01': 'января',
	'02': 'февраля',
	'03': 'марта',
	'04': 'апреля',
	'05': 'мая',
	'06': 'июня',
	'07': 'июля',
	'08': 'августа',
	'09': 'сентября',
	'10': 'октября',
	'11': 'ноября',
	'12': 'декабря'
}

#1 Получаем задание

# 1.1 Находим задание
listdir = os.listdir()

for fileName in listdir:
	if 'task' in fileName:
		taskFilePath = os.path.abspath(fileName)
		taskFound = True

if taskFound == False: raise('Задание не найдено! Это должен быть файл с фрагментом "task" в имени.')

# 1.2 Читаем xlsx задание
task_0 = pd.read_excel(taskFilePath, header=None)

if task_0.shape[1] != 8: raise('Неправильный формат задания. Откройте его в Excel и проверьте: где-то за его пределами лишняя информация.')

# 1.3 Сформируем ПП-часть задания в удобном формате
task = pd.DataFrame(columns=["Тип Протокола", "№ протокола", "Модель", "Серийный номер", "Путь для поиска", "Дата"])
j = 1 #счётчик сдвига курсора

for i in range(len(task_0.index)):
	task.loc[i] = [
	task_0.loc[i][0],	#"Тип Протокола"
	task_0.loc[i][1],	#"№ протокола"
	task_0.loc[i][0] if task_0.loc[i][0] != 'пп' else task_0.loc[i-j][0],	#"Модель"
	task_0.loc[i][2], #"Серийный номер"
	'path', #"Путь для поиска"
	task_0.loc[i][3] #"Дата"
	]

	if task_0.loc[i][0] == 'пп':
		j += 1
	else:
		j = 1

# 1.4 Сформируем ПC-часть задания в удобном формате
j = 1
for i in range(len(task_0.index)):
	task.loc[i + len(task_0.index)] = [
	task_0.loc[i][4],	#"Тип Протокола"
	task_0.loc[i][5],	#"№ протокола"
	task_0.loc[i][0] if task_0.loc[i][4] != 'пс' else task_0.loc[i-j][0],	#"Модель"
	task_0.loc[i][6], #"Серийный номер"
	'path', #"Путь для поиска"
	task_0.loc[i][7] #"Дата"
	]

	if task_0.loc[i][4] == 'пс':
		j += 1
	else:
		j = 1

# 1.5 Почистим задание от мусора
mask = task.isnull().any(axis=1)
task = task.loc[~mask]
task.astype('str')
print(task)
print(task.loc[[26],["№ протокола"]])
print(type(task.loc[26]["№ протокола"]))

psDate = task_0.loc[1][7]

psYear = psDate.year
psMonth = str(psDate.month)
if len(psMonth) == 1: psMonth = '0' + psMonth
psDay = str(psDate.day)
if len(psDay) == 1: psDay = '0' + psDay


#2 подготавливаем папки для протоколов новой приёмки

# containerYearPriemka = container + '\\' + f'{psYear}' + '\\' + f'{psYear}.{psMonth}.{psDay}'
# os.makedirs(containerYearPriemka)

#3 Находим нужные файлы на сервере и копируем в нашу папку
# 3.1 Копируются все протоколы нужных ТКР
# 3.2 Удаляются плохие аналоги, оставляются по одному самому лучшему
# 3.3 Проверяем полученные протоколы, исправляем

#4 Проверяем все ли файлы из задания найдены
# 4.1 Все ли найдены, сообщаем
# 4.2 Если не все, генерируем их?

#5 Спрашиваем нужно ли печатать говые протоколы, печатаем

# печатать в порядке номеров протоколов