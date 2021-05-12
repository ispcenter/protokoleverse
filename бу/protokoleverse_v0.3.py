'''protokoleverse v0.3'''

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

class ProtTask:
	"""Задание по одному конкретному протоколу"""
	def __init__(self, stendNumber, modelName, row):
		self.modelName = modelName
		self. = 
		self. = 
		self. = 
		self. = 
		self. = 

class Task_model_pp:
	'Содержит задание ПП по одному конкретному типу ТКР'
	def __init__(self, startRow, finishRow, modelName):
		self.startRow = startRow
		self.finishRow = finishRow
		self.modelName = modelName
		self.stendNumber = stendNumberDefiner(int(self.modelName.split('.')[0].split('-')[0].split(' ')[1]))
		







		self.serverPath = fr'\\Saukip\arm_{self.stendNumber}\Аттестация\{psYear}\{monthWorder[psMonth]}'
		
		if self.finishRow - self.startRow > 10: raise('В данной программе предусмотрено только 10 мест для различных ПП протоколов одной модели. Обратитесь к разработчику либо разбейте задание на меньшие части')
		if self.finishRow - self.startRow >= 1: pp1 = ProtTask(self.stendNumber, self.modelName, self.startRow + 1)
		if self.finishRow - self.startRow >= 2: pp2 = ProtTask(self.stendNumber, self.modelName, self.startRow + 2)
		if self.finishRow - self.startRow >= 3: pp3 = ProtTask(self.stendNumber, self.modelName, self.startRow + 3)
		if self.finishRow - self.startRow >= 4: pp4 = ProtTask(self.stendNumber, self.modelName, self.startRow + 4)
		if self.finishRow - self.startRow >= 5: pp5 = ProtTask(self.stendNumber, self.modelName, self.startRow + 5)
		if self.finishRow - self.startRow >= 6: pp6 = ProtTask(self.stendNumber, self.modelName, self.startRow + 6)
		if self.finishRow - self.startRow >= 7: pp7 = ProtTask(self.stendNumber, self.modelName, self.startRow + 7)
		if self.finishRow - self.startRow >= 8: pp8 = ProtTask(self.stendNumber, self.modelName, self.startRow + 8)
		if self.finishRow - self.startRow >= 9: pp9 = ProtTask(self.stendNumber, self.modelName, self.startRow + 9)
		if self.finishRow - self.startRow == 10: pp10 = ProtTask(self.stendNumber, self.modelName, self.startRow + 10)

#0 Setup
taskFound = False
task_models = {}
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

if taskFound == False: raise('Задание не найдено!')

# 1.2 Читаем задание
df = pd.read_excel(taskFilePath, header=None)

mask = df.isnull().all(axis=1)

cursor = 0


while cursor < len(mask):

	startRow = cursor

	try:
		while mask[cursor] == False:
			cursor += 1
	except:
		pass
	cursor -= 1
	finishRow = cursor
	modelName = df.loc[startRow][0]

	cursor += 2
	task_models.append((startRow, finishRow, modelName))

print(task_models)
print('len(task_models) = ', len(task_models))

if len(task_models) > 6: raise('В данной программе предусмотрено только 6 мест для различных моделей. Обратитесь к разработчику либо разбейте задание на меньшие части')
if len(task_models) >= 1: task_model_1 = Task_model(task_models[0]) 
if len(task_models) >= 2: task_model_2 = Task_model(task_models[1]) 
if len(task_models) >= 3: task_model_3 = Task_model(task_models[2]) 
if len(task_models) >= 4: task_model_4 = Task_model(task_models[3]) 
if len(task_models) >= 5: task_model_5 = Task_model(task_models[4]) 
if len(task_models) == 6: task_model_6 = Task_model(task_models[5])

print(task_model_1.modelName)
print(task_model_1.stendNumber)

psDate = df.loc[1][7]

psYear = psDate.year
psMonth = str(psDate.month)
if len(psMonth) == 1: psMonth = '0' + psMonth
psDay = str(psDate.day)
if len(psDay) == 1: psDay = '0' + psDay


#2 подготавливаем папки для протоколов новой приёмки

containerYearPriemka = container + '\\' + f'{psYear}' + '\\' + f'{psYear}.{psMonth}.{psDay}'
os.makedirs(containerYearPriemka)

#3 Находим нужные файлы на сервере и копируем в нашу папку
# 3.1 Копируются все протоколы нужных ТКР
# 3.2 Удаляются плохие аналоги, оставляются по одному самому лучшему
# 3.3 Проверяем полученные протоколы, исправляем

#4 Проверяем все ли файлы из задания найдены
# 4.1 Все ли найдены
# 4.2 Если не все, генерируем их

#5 Спрашиваем нужно ли печатать говые протоколы, печатаем

# печатать в порядке номеров протоколов