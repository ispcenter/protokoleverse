'''protokoleverse v0.1'''

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
    }

def protFinder():
	'''in path
	conditions'''
	
def parametersChecking():
	'''Checking parameters in protokol on stend'''
	pass

class Task_model:
	'Содержит задание по одному конкретному типу ТКР'
	def __init__(self, adress):
		self.adress = adress
		self.modelName = df.loc[self.adress[0]][0]
		self.stendNumber = stendNumberDefiner(int(self.modelName.split('.')[0].split('-')[0].split(' ')[1]))[True]
		# self.protNumbers = 
		# self. =

#0 Setup
taskFound = False
serverPath = 'Saukip'

#1 Получаем задание

# 1.1 Находим задание
# print(os.getcwd())
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
task_models = []

while cursor < len(mask):

	startRow = cursor

	try:
		while mask[cursor] == False:
			cursor += 1
	except:
		pass
	cursor -= 1
	finishRow = cursor

	cursor += 2
	task_models.append((startRow, finishRow))


print(len(task_models))

if len(task_models) > 6: raise('В интерпретаторе данной программы предусмотрено только 6 мест для различных моделей. Обратитесь к разработчику либо разбейте задание на меньшие части')

if len(task_models) >= 1: task_model_1 = Task_model(task_models[0]) 
if len(task_models) >= 2: task_model_2 = Task_model(task_models[1]) 
if len(task_models) >= 3: task_model_3 = Task_model(task_models[2]) 
if len(task_models) >= 4: task_model_4 = Task_model(task_models[3]) 
if len(task_models) >= 5: task_model_5 = Task_model(task_models[4]) 
if len(task_models) == 6: task_model_6 = Task_model(task_models[5])

secondModel = df.iloc[task_models[1][0]][0]

print(task_model_1.modelName)
print(task_model_2.modelName)
print(task_model_3.modelName)
print(task_model_4.modelName)
print(task_model_5.modelName)

print(task_model_1.stendNumber)
print(task_model_2.stendNumber)
print(task_model_3.stendNumber)
print(task_model_4.stendNumber)
print(task_model_5.stendNumber)

#2 подготавливаем папки для протоколов новой приёмки
# 2.1 если это первая приёмка в году, то надо создать папку нового года
# 2.2 Создать папку для новой приёмки

#3 Находим нужные файлы на сервере и копируем в нашу папку
# 3.1 Проверяем все ли файлы из задания найдены

#4 проверяем скопированные файлы на наличие ошибок, исправляем

#5 Спрашиваем нужно ли печатать говые протоколы, печатаем