import pandas as pd
import copy

connection_dict = {
    'Начало': ['a', 'a', 'b', 'c', 'd'],
    'Конец': ['b', 'e', 'c', 'd', 'e']}
connection_frame = pd.DataFrame(connection_dict)


# connection_frame.to_excel(r'C:\123.xlsx')


# Создание списка из уникальных названий вершин графа
def createNamesList():
    nameslist = []
    for c in connection_frame.columns:
        for i in range(connection_frame[c].size):
            if connection_frame[c][i] not in nameslist:
                nameslist.append(connection_frame[c][i])
    return nameslist


# Создание словаря со списками связностей для каждой вершины графа
def createNamesDict(NamesList):
    NamesDict = {}
    for i in range(len(NamesList)):
        connection_list = [0] * len(NamesList)
        # first column 'Начало'
        for j in range(connection_frame['Начало'].size):
            if NamesList[i] == connection_frame.loc[j]['Начало']:
                connection_list[NamesList.index(connection_frame.loc[j]['Конец'])] = 1
        # second column 'Конец'
        for j in range(connection_frame['Конец'].size):
            if NamesList[i] == connection_frame.loc[j]['Конец']:
                connection_list[NamesList.index(connection_frame.loc[j]['Начало'])] = 1
        # same element (self crossing) - matrix diagonal
        connection_list[i] = 1
        # making element in dictionary
        NamesDict[NamesList[i]] = connection_list
    return NamesDict


# Меняем 2 строчки и столбца местами
def swapAandB(NamesDict, NamesList, index_a, index_b):
    # Сначала идет перемещение строчек по ключам словаря
    key_a = NamesList[index_a]
    key_b = NamesList[index_b]
    value_a = NamesDict[key_a]
    NamesDict[key_a] = NamesDict[key_b]
    NamesDict[key_b] = value_a
    NamesDict = dict(zip(NamesList, list(NamesDict.values())))
    # Затем идет замена столбцов непосредственно в строках (в списках, лежащих в словаре)
    for i in NamesList:
        value_c = NamesDict[i][index_a]
        NamesDict[i][index_a] = NamesDict[i][index_b]
        NamesDict[i][index_b] = value_c
    # Возвращаем словарь (новую матрицу связности)
    return NamesDict


# Определение ширины ленты
def countZeros(List, index):
    counter = 0
    pos = -1
    while List[pos] != 1:
        counter += 1
        pos -= 1
    return (len(List) - counter - (index + 1))  # все это *2 и +1, если нужно подсчитать ширину ленты как в книге
    # я считаю ширину над диагональю, без включения самой диагонали


# Определение двух вершин, приводящих к максимальной ширине ленты
def findMaxBandLength(NamesDict, NamesList):
    maxindex = 0
    for i in range(len(NamesList)):
        if countZeros(NamesDict[NamesList[i]], i) > countZeros(NamesDict[NamesList[maxindex]], maxindex):
            maxindex = i
    return (maxindex, maxindex + countZeros(NamesDict[NamesList[maxindex]], maxindex))


# Возвращает список с индаксами вершин, смежных с рассматриваемой (только выше диагонали)
def findIndexes(row, NamesDict, NamesList):
    indexList = []
    for i in range(row + 1, len(NamesDict[NamesList[row]])):
        if NamesDict[NamesList[row]][i] == 1:
            indexList.append(i)
    return indexList


# Сам метод
def firstMethod(namesdict, nameslist):
    swaplist = []
    end = False
    while end != True:
        for row in range(len(nameslist) - 1):
            indexes = findIndexes(row, namesdict, nameslist)
            print('Row = ', row, ', Connection indexes = ', indexes)
            for index in range(-1, -len(indexes) - 1, -1):
                if_nameslist = nameslist.copy()
                if_nameslist[row] = nameslist[indexes[index]]
                if_nameslist[indexes[index]] = nameslist[row]
                if_namesdict = copy.deepcopy(namesdict)
                if_namesdict = swapAandB(if_namesdict, if_nameslist, row, indexes[index])
                old_length = countZeros(namesdict[nameslist[findMaxBandLength(namesdict, nameslist)[0]]],
                                        findMaxBandLength(namesdict, nameslist)[0])
                new_length = countZeros(if_namesdict[if_nameslist[findMaxBandLength(if_namesdict, if_nameslist)[0]]],
                                        findMaxBandLength(if_namesdict, if_nameslist)[0])
                print('Old length = ', old_length, ', New length = ', new_length, '\n')
                if new_length < old_length:
                    nameslist = if_nameslist
                    namesdict = copy.deepcopy(if_namesdict)
                    print('Swap success!\n')
        for row in range(len(nameslist) - 1):
            exit = False
            indexes = findIndexes(row, namesdict, nameslist)
            for index in range(-1, -len(indexes) - 1, -1):
                if_nameslist = nameslist.copy()
                if_nameslist[row] = nameslist[indexes[index]]
                if_nameslist[indexes[index]] = nameslist[row]
                if_namesdict = copy.deepcopy(namesdict)
                if_namesdict = swapAandB(if_namesdict, if_nameslist, row, indexes[index])
                old_length = countZeros(namesdict[nameslist[findMaxBandLength(namesdict, nameslist)[0]]],
                                        findMaxBandLength(namesdict, nameslist)[0])
                new_length = countZeros(if_namesdict[if_nameslist[findMaxBandLength(if_namesdict, if_nameslist)[0]]],
                                        findMaxBandLength(if_namesdict, if_nameslist)[0])
                if new_length == old_length:
                    if (row, indexes[index]) not in swaplist:
                        nameslist = if_nameslist
                        namesdict = copy.deepcopy(if_namesdict)
                        print('Save swap success! Starting loop from the beginning... \n')
                        swaplist.append((row, indexes[index]))
                        exit = True
                        break
            if exit == True:
                break
        else:
            end = True

    return (nameslist, namesdict)


# Запись исходных списка вершин и словаря
nameslist = createNamesList()
namesdict = createNamesDict(nameslist)
print('Исходный список вершин: ', nameslist)
print('\nИсходная матрица: ', namesdict)

# Запись конечных списка вершин и словаря
(nameslist, namesdict) = firstMethod(namesdict, nameslist)
print('\nКонечный список вершин: ', nameslist)
print('\nКонечная матрица: ', namesdict)

# Создание датафрейма
# connection_matrix = pd.DataFrame(namesdict)
# for i in range(len(nameslist)):
#   connection_matrix.rename(index={i: nameslist[i]}, inplace=True)

