import pandas as pd
from fuzzywuzzy import fuzz  # в этом модуле сравниватель строк ratio
import re  # для удаления русских букв

#fpath = 'E:/Temp/Выгрузки/'
fpath = 'E:/Temp/_bsu_all.xlsx'

# !сопоставление по названиям для EID и UT по данным НОРА
try:
    dfs = pd.read_excel(fpath, sheet_name='Scopus')
except:
    dfs = pd.DataFrame(
        columns=['Авторы', 'Название', 'Год', 'Название источника', 'Том', 'Выпуск ', 'Статья №', 'Страница начала',
                 'Страница окончания', 'Количество страниц', 'Source', 'EID', 'ISSN', 'Document Type'])

try:
    dfw = pd.read_excel(fpath, sheet_name='WoS')
except:
    dfw = pd.DataFrame(
        columns=['Author Full Names', 'Article Title', 'Publication Year', 'Source Title', 'Volume', 'Issue',
                 'Article Number', 'Start Page', 'End Page', 'Количество страниц', 'Source', 'UT', 'ISSN', 'Document Type'])

# удаление из названия русских букв
dfs['Название'] = dfs['Название'].apply(lambda x: re.sub('\s+', ' ', re.sub('[А-Яа-я]', '', x)).strip())

# это блок поиска и выборки одинаковых статей для сопоставления EID и UT, результат - в список где Scopus
dfs1 = dfs[['Название', 'EID']]
dfw1 = dfw[['Article Title', 'UT']]
dfs1['Название'] = dfs1['Название'].str.upper()
dfw1['Article Title'] = dfw1['Article Title'].str.upper()

lst = []
for i in dfs1.index:
    for j in dfw1.index:
        koef = fuzz.ratio(dfs1['Название'][i], dfw1['Article Title'][j])
        if koef > 90:
            lst.append([dfs1['EID'][i], dfw1['UT'][j], 'да'])
# конец блока сопоставления одинаковых, на выходе список lst

# Этот блок формирует экспорт Scopus по нужному шаблону, а также вписывает туда сопоставления с UT по более чем 90% совпадению титулов

# создание DF из сопоставленных
dfsw = pd.DataFrame(lst, columns=['EID', 'UT', 'WoS'])
# объединение сопоставленных с экспортом Scopus по EID
dfs = dfs.merge(dfsw, on=['EID'], how='left')

writer = pd.ExcelWriter('_Export.xlsx', engine='xlsxwriter')
dfs.to_excel(writer, 'ScopusWoS')


writer.save()
# !конец блока сопоставления по названиям для EID и UT по данным НОРА

# !сопоставление по названиям+журналам и именам файлов из присланных НОРА
dfs = pd.read_excel(fpath, sheet_name='ScopusWoS')
dfw = pd.read_excel(fpath, sheet_name='Файлы')

# это блок поиска и выборки одинаковых статей для сопоставления EID и UT, результат - в список где Scopus
dfs1 = dfs[['Title', 'Наз+Жур']]
dfw1 = dfw[['Файлы-', 'Файлы']]
dfs1['Наз+Жур'] = dfs1['Наз+Жур'].str.upper()
dfw1['Файлы'] = dfw1['Файлы'].str.upper()

lst = []
for i in dfs1.index:
    for j in dfw1.index:
        koef = fuzz.ratio(dfs1['Наз+Жур'][i], dfw1['Файлы'][j])
        if koef > 80:
            lst.append([dfs1['Title'][i], dfw1['Файлы-'][j], koef])
# конец блока сопоставления одинаковых, на выходе список lst

# создание DF из сопоставленных
dfsw = pd.DataFrame(lst, columns=['Название', 'Файл', 'Коэф'])
writer = pd.ExcelWriter('_Export.xlsx', engine='xlsxwriter')
dfsw.to_excel(writer, 'Сопост')

writer.save()
# !конец блока сопоставления по названиям+журналам и именам файлов из присланных НОРА

# !сопоставление по названиям для EID и UT по моим выгрузкам Scopus и WoS Open Access за все года
dfs = pd.read_csv(fpath+'scopus.csv')
dfw = pd.read_excel(fpath+'savedrecs.xls')


# удаление из названия русских букв из списка Scopus
dfs['Title'] = dfs['Title'].apply(lambda x: re.sub('\s+', ' ', re.sub('[А-Яа-я]', '', x)).strip())

# это блок поиска и выборки одинаковых статей для сопоставления EID и UT, результат - в список где Scopus
dfs1 = dfs[['Title', 'EID']]
dfw1 = dfw[['Article Title','UT (Unique WOS ID)']]
dfs1['Title'] = dfs1['Title'].str.upper()
dfw1['Article Title'] = dfw1['Article Title'].str.upper()

lst = []
for i in dfs1.index:
    for j in dfw1.index:
        koef = fuzz.ratio(dfs1['Title'][i], dfw1['Article Title'][j])
        if koef>90:
            lst.append([dfs1['Title'][i], dfw1['Article Title'][j], dfs1['EID'][i], dfw1['UT (Unique WOS ID)'][j], 'да', koef])
# конец блока сопоставления одинаковых, на выходе список lst

# Этот блок формирует экспорт Scopus по нужному шаблону, а также вписывает туда сопоставления с UT по более чем 90% совпадению титулов

# создание DF из сопоставленных
dfsw = pd.DataFrame(lst, columns =['ScopusT', 'WoST', 'EID', 'UT', 'WoS', 'Ratio'])
# объединение сопоставленных с экспортом Scopus по EID
dfs = dfs.merge(dfsw, on=['EID'], how='left')

dfs = dfs[['Authors', 'Title', 'Year', 'Source title', 'Volume', 'Issue', 'Page start', 'Page end', 'Abstract', 'ISSN',
           'Language of Original Document', 'Document Type', 'Source', 'WoS', 'EID', 'UT', 'ScopusT', 'WoST', 'Ratio']]

writer = pd.ExcelWriter('_Export.xlsx', engine='xlsxwriter')
#dfs.to_excel(writer, 'Scopus')

# Этот блок формирует экспорт WoS по нужному шаблону

# исключаем те, которые есть в списке Scopus
#lst = []
lst = dfs['UT'].values
dfw = dfw.loc[~dfw['UT (Unique WOS ID)'].isin(lst)]


dfw = dfw[['Author Full Names', 'Article Title', 'Source Title', 'Language', 'Document Type', 'Abstract',
           'ISSN', 'Publication Year', 'Volume', 'Issue', 'Start Page', 'End Page', 'UT (Unique WOS ID)']]
dfw.insert(13, 'WoS', 'да')

#dfw.to_excel(writer, 'WoS')
dfw.rename(columns={'Author Full Names': 'Authors', 'Article Title': 'Title', 'UT (Unique WOS ID)': 'UT',
                    'Source Title': 'Source title', 'Language': 'Language of Original Document',
                    'Start Page': 'Page start', 'End Page': 'Page end', 'Publication Year': 'Year'}, inplace=True)

dfsw = dfs.append(dfw, sort=False)
dfsw.to_excel(writer, 'Export')
writer.save()
# !конец блока сопоставления по названиям для EID и UT по моим выгрузкам Scopus и WoS Open Access за все года

# !поиск дубликатов в подготовленном по данным НОРА списке с уже имеющимися экспортами коллекций
dfw = pd.read_csv('123456789-90.csv')
dfs = pd.read_excel(fpath, sheet_name='dc. авторыбез{}')
# это блок поиска и выборки одинаковых статей для сопоставления EID и UT, результат - в список где Scopus
dfw1 = dfw[['dc.title[ru]', 'dc.identifier.uri']]
dfs1 = dfs[['dc.title']]
dfw1[['dc.title[ru]']] = dfw1[['dc.title[ru]']].astype(str)
SUB = str.maketrans("₀₁₂₃₄₅₆₇₈₉", "0123456789", )
dfw1['dc.title[ru]'] = dfw1['dc.title[ru]'].str.translate(SUB)
dfw1['dc.title[ru]'] = dfw1['dc.title[ru]'].str.upper()
dfs1['dc.title'] = dfs1['dc.title'].str.upper()

lst = []
for i in dfs1.index:
    print(int(i / len(dfs1.index)*100))
    for j in dfw1.index:
        koef = fuzz.ratio(dfw1['dc.title[ru]'][j], dfs1['dc.title'][i])
        if koef > 80:
            lst.append([dfw1['dc.title[ru]'][j], dfs1['dc.title'][i], dfw1['dc.identifier.uri'][j], koef])
# конец блока сопоставления одинаковых, на выходе список lst

# создание DF из сопоставленных
dfsw = pd.DataFrame(lst, columns=['DSpace', 'НОРА', 'URL', 'Коэф'])
writer = pd.ExcelWriter('_Result.xlsx', engine='xlsxwriter')
dfsw.to_excel(writer, 'Дубли')

writer.save()
# !конец блока поиска дубликатов в подготовленном по данным НОРА списке с уже имеющимися экспортами коллекций

print('Finish')