# импортируем модули
import requests # для запросов cо страницы сайта
from bs4 import BeautifulSoup # парсер для синтаксического разбора страницы HTML/XML
import shutil, os # для работы с ОС


# определяем базовые параметры
link1 = 'https://www.atsenergo.ru/nreport' # первая часть ссылки на отчеты о равновесных ценах в наиболее крупных узлах
link2 = '?rname=big_nodes_prices_pub&rdate=' # вторая часть ссылки на отчеты о равновесных ценах в наиболее крупных узлах
link3 = [] # третья часть ссылки, состоит из года, месяца, дня (дней по умолчанию 31)

# задаем параметры для данного конкретного расчета (год, месяц, дни)
destination_folder = 'D:\\УЗЛцены\\test\\февраль 2018'
filename = '_eur_big_nodes_prices_pub.xls'
year = '2018'
month = '02'
days = ['01','02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12', '13', '14', '15', '16', '17', '18', '19', '20', '21', '22', '23', '24', '25', '26', '27', '28', '29', '30', '31']

# формируем третью часть ссылки
for i in days:
    link3.append(year + month + i)
    
# формируем ссылки на страницу отчета за весь месяц (из расчета 31 день)
linkpage = []
for j in link3:
    linkpage.append(link1 + link2 + j)
   
# формируем ссылки на сам отчет 
for link in linkpage:
    page = requests.get(link) # обращаемся к странице по сформированной сслыке
    soup = BeautifulSoup(page.text, 'html.parser') # читаем эту страницу
    l = soup.find_all('a') # находим теги со ссылками и добавляем в перечень
    for i in range(len(l)): # проверяем, содержит ли тэг со ссылкой наименование необходимого файла, по нахождению прерываем цикл
        if str(l[i]).rfind(filename) != -1:
            print('найден тэг', l[i])
            filenameRSV = soup.find_all('a')[i].get_text() # определяем имя скачиваемого файла в найденно тэге
            linkRSV = link1 + soup.find_all('a')[i].get('href') # определяем ссылку на файл для скачивания
            r = requests.get(linkRSV)  # запрашиваем файл для скачивания
            output = open(filenameRSV, 'wb')
            output.write(r.content)
            output.close()
            source_files = os.getcwd()
            shutil.move(source_files + '\\' + filenameRSV, destination_folder + '\\' + filenameRSV)
            print('файл', filenameRSV, 'скачен и сохранен в папке', destination_folder)
            break
        if i == len(l)-1 and str(l[i]).rfind(filename) == -1:
            print('искомое имя файла по ссылке', link, 'в отобранных тэгах не найдено')




# чтобы посмотреть разметку страницы print(soup.prettify())

# проводим анализ

destination_folder = 'D:\\УЗЛцены\\январь 2018'

# составляем список файлов в папке назначения
import os
files = os.listdir(destination_folder)

# определяем узлы расчетной модели для анализа
node1 = 330916
node2 = 330095
node3 = 330015


# определяем таблицу для анализа (многомерный список)

price = [[0 for i in range(5)] for j in range(744)]

# читаем файлы
index = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23] # индекс соответствует часу
i = 0
import xlrd
for f in files:
    file = xlrd.open_workbook(destination_folder + '\\' + f)
    for j in index:
        sheet = file.sheet_by_index(j)
        for k in range(sheet.nrows):
            price[i][0] = f[6:8] 
            price[i][1] = j
            if sheet.cell_value(k, 0) == node1:
                price[i][2] = sheet.cell_value(k, 4)
            if sheet.cell_value(k, 0) == node2:
                price[i][3] = sheet.cell_value(k, 4) 
            if sheet.cell_value(k, 0) == node3:
                price[i][4] = sheet.cell_value(k, 4)            
        i += 1    
    
# разносим цены по выходным и дням недели 
priceweek = {'выходной' : [], 'понедельник' : [], 'вторник' : [], 'среда' : [], 'четверг' : [], 'пятница' : [], 'суббота' : [], 'воскресенье' : []}
jan = {'выходной' : ['01', '02', '03', '04', '05', '06', '07', '08', '13', '14', '20', '21', '27', '28'], 'понедельник' : ['01', '08', '15', '22', '29'], 'вторник' : ['02', '09', '16', '23', '30'], 'среда' : ['03', '10', '17', '24', '31'], 'четверг' : ['04', '11', '18', '25'], 'пятница' : ['05', '12', '19', '26'], 'суббота' : ['06', '13', '20', '27'], 'воскресенье' : ['07', '14', '21', '28']}

def dayweek(**month):
    print("введите день недели из списка: выходной, понедельник, вторник, среда, четверг, пятница, суббота, воскресенье")
    day = input()
    for i in range(len(price)):
        for key, value in month.items():
            if key == day:
                for j in value:
                    if j == price[i][0]:
                        priceweek[day].append([price[i][0], price[i][1], price[i][2], price[i][3],price[i][4]])
            i += 1
i = 0
dayweek(**jan)
# провекра priceweek['выходной']

# считаем средню цену в выходной и каждый день недели
def avg(**kw):
    print("введите день недели из списка: выходной, понедельник, вторник, среда, четверг, пятница, суббота, воскресенье")
    day = input()
    for key, value in kw.items():
        if key == day:
            for j in range(len(value)):
                global s
                s += value[j][2]

def createfileavg(**month):
    import xlwt
    wb =  xlwt.Workbook()
    ws = wb.add_sheet('1')
    for key in month:
        for j in len(jan):
            ws.write(0, j, key)
        wb.save(destination_folder + '\\' + 'avg.xls')

s = 0           
avg(priceweek)
       
createfileavg(jan)            
