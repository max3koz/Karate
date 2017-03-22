import xlrd
import xlwt
import array

# Ввод данных для проведения сортировки данных

print("Введите название файла со списком участников соревнований без расширения:")
Name_Workbook_Competitor = input() + ".xls"

Workbook_Competitor = xlrd.open_workbook(Name_Workbook_Competitor)
Worksheet_Competitor = Workbook_Competitor.sheet_by_name('Участники')

print("Введите название соревнований:")
Name_Competition = input()

print("Ведеите число дня соревнований, например, 23:")
Day_Competition = input()
print("Ведеите число месяца соревнований, например, 02:")
Month_Competition = input()
print("Ведеите число года соревнований, например, 2017:")
Year_Competition = input()

print("Введите виды соревнований, например, ката, кумитэ, котен или Джунро:")
List_Type_Competition = []
Type_Competition = "+"
Qty_Type_Competition = 0
while Type_Competition != "":
    Type_Competition = input()
    if Type_Competition == "ката" or Type_Competition == "Ката" or Type_Competition == "kata" or Type_Competition == "Kata":
        Type_Competition = Worksheet_Competitor.cell(0, 10).value
        List_Type_Competition.append(Type_Competition)
    elif Type_Competition == "кумитэ" or Type_Competition == "Кумитэ" or Type_Competition == "kumite" or Type_Competition == "Kumite":
        Type_Competition = Worksheet_Competitor.cell(0, 11).value
        List_Type_Competition.append(Type_Competition)
    elif Type_Competition == "котен" or Type_Competition == "Котен" or Type_Competition == "koten" or Type_Competition == "Koten" or Type_Competition == "Джунро" or Type_Competition == "джунро" or Type_Competition == "dzhunro" or Type_Competition == "Dzhunro":
        Type_Competition = Worksheet_Competitor.cell(0, 8).value
        List_Type_Competition.append(Type_Competition)
    elif Type_Competition == "":
        break
    else:
        print("Нет такого вида соревнований.")
    Qty_Type_Competition += 1
print (List_Type_Competition)

print("Введите категории спортменов:")
List_Type_Category = []
Type_Category = "+"
Qty_Type_Category = 0
while Type_Category != "":
    Type_Category = input()
    if Type_Category == "а" or Type_Category == "A" or Type_Category == "а" or Type_Category == "А":
        List_Type_Category.append("А")
    elif Type_Category == "б" or Type_Category == "Б" or Type_Category == "b" or Type_Category == "B":
        List_Type_Category.append("Б")
    elif Type_Category == "":
        break
    else:
        print("Нет такой категории.")
    Qty_Type_Category += 1
print (List_Type_Category)

print("Введите возрастные категории для ката:")
List_Type_Age_Kata = []
Type_Age_Kata = "+"
Qty_Type_Age_Kata = 0
while Type_Age_Kata != "":
    Type_Age_Kata = input()
    if Type_Age_Kata != "":
        List_Type_Age_Kata.append(Type_Age_Kata)
    else:
        break
    Qty_Type_Age_Kata += 1
print (List_Type_Age_Kata)

print("Введите возрастные категории для котен ката:")
List_Type_Age_Koten = []
Type_Age_Koten = "+"
Qty_Type_Age_Koten = 0
while Type_Age_Koten != "":
    Type_Age_Koten = input()
    if Type_Age_Koten != "":
        List_Type_Age_Koten.append(Type_Age_Koten)
    else:
        break
    Qty_Type_Age_Koten += 1
print (List_Type_Age_Koten)

print("Введите возрастные категории для кумитэ:")
List_Type_Age_Kumite = []
Type_Age_Kumite = "+"
Qty_Type_Age_Kumite = 0
while Type_Age_Kumite != "":
    Type_Age_Kumite = input()
    if Type_Age_Kumite != "":
        List_Type_Age_Kumite.append(Type_Age_Kumite)
    else:
        break
    Qty_Type_Age_Kumite += 1
print (List_Type_Age_Kumite)

List_Type_Weight_Category = []
for i in range(Qty_Type_Age_Kumite):
    print("Введите граничный вес для категории ",List_Type_Age_Kumite[i])
    Border_Weight = int(input())
    List_Type_Weight_Category.append([])
    List_Type_Weight_Category[0].append(int(List_Type_Age_Kumite[i]))
    List_Type_Weight_Category[1].append(int(Border_Weight))
    print ("Rjytw")
#for i in range(len(List_Type_Weight_Category)):
#    for j in range(len(List_Type_Weight_Category[i])):
#    print(List_Type_Weight_Category[i][j], end=' ')