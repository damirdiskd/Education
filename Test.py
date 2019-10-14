import os #Для возможности указать директорию с файлом
import re #Для работы с регулярными выражениями
import xlwt #Для выгрузки в Excel

frequency = {}
stop_list = ['не', 'а', 'в', 'у', 'и', 'с', ''] #список стоп-слов

# Ввести путь к файлу:
os.chdir("E:\\Python\\scripts")
# Ввести название файла:
file = open('text.txt')
text = file.read()
print (text)
file.close()

# Вызываем предустановленную библиотеку pymystem3 (Установлен Mystem) для приведения слов к нормальной форме
from pymystem3 import Mystem
mystem = Mystem()
lemmas = mystem.lemmatize(text)
lemmas_string=(''.join(lemmas))
lemmas_list = re.split(r'\W+', lemmas_string)

#Подсчет частоты повторения слова в тексте с исключением в подсчете стоп-слов
for word in lemmas_list:
    if word in stop_list:
        continue
    else:
        count = frequency.get(word,0)
        frequency[word] = count + 1

#сортирую список слов по убыванию количества повторений
sorted_list = sorted(frequency.items(), key=lambda x: x[1], reverse=True)

#Запись в Excel
book = xlwt.Workbook()
sheet = book.add_sheet('Frequency list')

sheet.write(0, 0, 'Слово')
sheet.write(0, 1, 'Частота')

for i in range(len(sorted_list)):
    print (sorted_list[i][0], sorted_list[i][1])
    sheet.write(i+1, 0, sorted_list[i][0])
    sheet.write(i+1, 1, sorted_list[i][1])

book.save('Test.xls')


