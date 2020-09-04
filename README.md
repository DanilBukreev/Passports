# Passports
Краткая инструкция по использованию вспомогательной программы для переноса данных из Excele в WORD
1.	Перед запуском программы необходимо удостовериться в правильности заполненного шаблона.
a.	Не должно быть работающих процессов Word, шаблон должен быть закрытым во время работы программы.
b.	В названии пути до паспортов не должно быть никаких пробелов.
c.	В столбце с кратким наименование (в шаблоне Эксель он называется Краткое_наим) не должно быть таких знаков как /,"" и иных знаков, которые не предусмотрены системой Windows для создания имени файла, так как из этого поля подтягиваются данные на генерацию имен для паспортов. Если в этом столбце будет написано -Исключено, то программа не будет обрабатывать это паспорт. 
d.	Первый столбец должен быть пронумерован трехзначной нумерацией. Пример - 001 , и ко всем ячейкам этого столбца должны быть применены все границы (последняя ячейка с примененными границами должна быть пустая, чтобы программа смогла идентифицировать конец списка АСУ как на рис.1).
Формат всех ячеек должен быть “Текстовым”, что бы удостовериться в этом в верхнем углу ячейки будет зеленый уголок.(см рис. 1) 

![Image alt](https://github.com/DanilBukreev/Passports/raw/Master/DanilBukreev/Passports/image.png)
 
                              рис. 1
e.	
2.	Запуск программы осуществляется использованием файла с названием First_menu.exe (он находится в папке dist)
a.	Если это первая итерация работы с генерацией паспортов, то необходимо сначала нажать на первый чекбокс "Сгенерировать паспорта", далее следовать указаниям ИИ
b.	В указанной вами папке создастся директория “Шаблоны_паспортов” в ней будут храниться шаблоны под каждый паспорт. 
3.	Если вы уже сгенерировали паспорта и шаблоны к паспортам, но хотите обновить информацию, то информация, которая заносится вручную в Word изменяется через шаблон паспорта в папке “Шаблоны_паспортов”. Для обновления данных запускаем программу заново, нажимаем на второй чекбокс и указываем путь до папки “Шаблоны_паспортов”, к эксельке и до папки, куда необходимо сохранить паспорта.
4.	Программу каждый раз необходимо запускать заново, повторное нажатие кнопки Start, после завершения процесса переноса может вызвать неправильную работу программы. Так что при завершении работы программы, необходимо нажать на Cancel или крестик, чтобы безопасно завершить ее работу.
Если Вы уже нажали кнопку Cancel после нажатия кнопки Start и до завершения работы программы, то закроется только меню, к сожалению, сам процесс работы 
Если необходимо внести изменения в таблице Excel, вышли из программы, внесли изменения, сохранили эксель файл, запустили программу заново с чекбоксом обновить данные в паспортах.
