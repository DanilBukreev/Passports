from win32com.client import Dispatch as comDispatch
import re


word = comDispatch('Word.Application')
doc = word.Documents.Open("C:/Users/dbukreev/PycharmProjects/Passports/002_Паспорт_АСУ_котлов.docx")
word.Application.Run("C:/Users/dbukreev/PycharmProjects/Passports/002_Паспорт_АСУ_котлов.docx!MakeList")  # возможно здесь тоже нужен полный путь
doc.Save()
doc.Close()

def multiple_replace(target_str, replace_values):
    # получаем заменяемое: подставляемое из словаря в цикле
    for i, j in replace_values.items():
        # меняем все target_str на подставляемое
        target_str = target_str.replace(i, j)
    return target_str

my_str=['df "df" f', 'dfdf f f','/fdf']
# создаем словарь со значениями и строку, которую будет изменять
replace_values = {' ': "_", '"': "_", '/': '_'}

my_st=[]
# изменяем и печатаем строку
for i in range(0,len(my_str)):
    my_st.append(multiple_replace(my_str[i],replace_values))
    print(my_st)


#kek=[]
#for i in range(0,len(lol)):
 #   kek.append(lol[i].replace('"','_'))
#
#print(kek)