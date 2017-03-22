import xlrd
import xlwt


print ("Введите название файла со список участников, без расширения файла:", )
# Name_Workbook_Competitor = input() + ".xls"
Name_Workbook_Competitor = "Test.xls"

print (Name_Workbook_Competitor)

Workbook_Competitor = xlrd.open_workbook(Name_Workbook_Competitor)
Worksheet_Competitor = Workbook_Competitor.sheet_by_name('Участники')

#Выбор участников соревнований по видам соревнований
Person = {}

Qty_String = 1
while Worksheet_Competitor.cell(Qty_String, 1).value != xlrd.empty_cell.value:

    NamePerson = Worksheet_Competitor.cell(Qty_String,1).value
    SexPerson = Worksheet_Competitor.cell(Qty_String, 6).value
    GroupPerson = Worksheet_Competitor.cell(Qty_String, 10).value
    AgePerson = int(Worksheet_Competitor.cell(Qty_String, 7).value)
    WeightPerson = Worksheet_Competitor.cell(Qty_String, 9).value
    DzunroPerson = Worksheet_Competitor.cell(Qty_String, 8).value
    KataPerson = Worksheet_Competitor.cell(Qty_String, 10).value
    KumitePerson = Worksheet_Competitor.cell(Qty_String, 11).value

    Person[Qty_String] = {'id': Qty_String, 'DataPerson' : {'name':NamePerson, 'sex':SexPerson, 'group':GroupPerson, 'age':AgePerson, 'weight':WeightPerson, 'dzunro':DzunroPerson, 'kata':KataPerson, 'kumite':KumitePerson}}

    Qty_String += 1

print(Person.items())
print(Person[2]['DataPerson']['name'])
print(Person[1]['DataPerson']['name'])
print(Qty_String-1)

Femail_A_6_Kata = {}
Femail_B_6_Kata = {}
Femail_A_6_Kumite = {}
Femail_B_6_Kumite = {}

Femail_A_7_Kata = {}
Femail_B_7_Kata = {}
Femail_A_7_Kumite = {}
Femail_B_7_Kumite = {}

Femail_A_8_Kata = {}
Femail_B_8_Kata = {}
Femail_A_8_Kumite = {}
Femail_B_8_Kumite = {}

Femail_A_9_Kata = {}
Femail_B_9_Kata = {}
Femail_A_9_Kumite = {}
Femail_B_9_Kumite = {}

Femail_A_1011_Kata = {}
Femail_B_1011_Kata = {}
Femail_A_1011_Kumite = {}
Femail_B_1011_Kumite = {}

Femail_A_1213_Kata = {}
Femail_B_1213_Kata = {}
Femail_A_1213_Kumite = {}
Femail_B_1213_Kumite = {}

Femail_A_1415_Kata = {}
Femail_B_1415_Kata = {}
Femail_A_1415_Kumite = {}
Femail_B_1415_Kumite = {}

Femail_A_1617_Kata = {}
Femail_B_1617_Kata = {}
Femail_A_1617_Kumite = {}
Femail_B_1617_Kumite = {}

Femail_A_18_Kata = {}
Femail_B_18_Kata = {}
Femail_A_18_Kumite = {}
Femail_B_18_Kumite = {}

Mail_A_5_Kata = {}
Mail_B_5_Kata = {}
Mail_A_5_Kumite = {}
Mail_B_5_Kumite = {}

Mail_A_6_Kata = {}
Mail_B_6_Kata = {}
Mail_A_6_Kumite = {}
Mail_B_6_Kumite = {}
Mail_A_6_Kumite_22 = {}
Mail_B_6_Kumite_22 = {}

Mail_A_7_Kata = {}
Mail_B_7_Kata = {}
Mail_A_7_Kumite = {}
Mail_B_7_Kumite = {}
Mail_A_7_Kumite_25 = {}
Mail_B_7_Kumite_25 = {}

Mail_A_8_Kata = {}
Mail_B_8_Kata = {}
Mail_A_8_Kumite = {}
Mail_B_8_Kumite = {}
Mail_A_8_Kumite_28 = {}
Mail_B_8_Kumite_28 = {}

Mail_A_9_Kata = {}
Mail_B_9_Kata = {}
Mail_A_9_Kumite = {}
Mail_B_9_Kumite = {}
Mail_A_9_Kumite_30 = {}
Mail_B_9_Kumite_30 = {}

Mail_A_1011_Kata = {}
Mail_B_1011_Kata = {}
Mail_A_1011_Kumite = {}
Mail_B_1011_Kumite = {}
Mail_A_1011_Kumite_37 = {}
Mail_B_1011_Kumite_37 = {}

Mail_A_1213_Kata = {}
Mail_B_1213_Kata = {}
Mail_A_1213_Kumite = {}
Mail_B_1213_Kumite = {}
Mail_A_1213_Kumite_46 = {}
Mail_B_1213_Kumite_46 = {}

Mail_A_1415_Kata = {}
Mail_B_1415_Kata = {}
Mail_A_1415_Kumite = {}
Mail_B_1415_Kumite = {}
Mail_A_1415_Kumite_60 = {}
Mail_B_1415_Kumite_60 = {}

Mail_A_1617_Kata = {}
Mail_B_1617_Kata = {}
Mail_A_1617_Kumite = {}
Mail_B_1617_Kumite = {}
Mail_A_1617_Kumite_68 = {}
Mail_B_1617_Kumite_68 = {}

Mail_A_18_Kata = {}
Mail_B_18_Kata = {}
Mail_A_18_Kumite = {}
Mail_B_18_Kumite = {}

Mail_A_35_Kata = {}
Mail_B_35_Kata = {}
Mail_A_35_Kumite = {}
Mail_B_35_Kumite = {}

for i in range(1, Qty_String-1):

    Name = Person[i]['DataPerson']['name']
    Sex = Person[i]['DataPerson']['sex']
    Group = Person[i]['DataPerson']['group']
    Age = Person[i]['DataPerson']['age']
    Weight = Person[i]['DataPerson']['weight']
    Dzunro = Person[i]['DataPerson']['dzunro']
    Kata = Person[i]['DataPerson']['kata']
    Kumite = Person[i]['DataPerson']['kumite']

# Девочки 6 лет
    if (Sex=="д" or Sex=="Д" or Sex=="ж" or Sex=="Ж") and (Kata=="a" or Kata=="A" or Kata=="а" or Kata=="А") and Age==6:
        Femail_A_6_Kata[i] = Person[i]
    if (Sex=="д" or Sex=="Д" or Sex=="ж" or Sex=="Ж") and (Kata=="б" or Kata=="Б" or Kata=="b" or Kata=="B") and Age==6:
        Femail_B_6_Kata[i] = Person[i]
    if (Sex=="д" or Sex=="Д" or Sex=="ж" or Sex=="Ж") and (Kumite=="a" or Kumite=="A" or Kumite=="а" or Kumite=="А") and Age==6:
        Femail_A_6_Kumite[i] = Person[i]
    if (Sex=="д" or Sex=="Д" or Sex=="ж" or Sex=="Ж") and (Kumite=="б" or Kumite=="Б" or Kumite=="b" or Kumite=="B") and Age==6:
        Femail_B_6_Kumite[i] = Person[i]

# Девочки 7 лет
    if (Sex=="д" or Sex=="Д" or Sex=="ж" or Sex=="Ж") and (Kata=="a" or Kata=="A" or Kata=="а" or Kata=="А") and Age==7:
        Femail_A_7_Kata[i] = Person[i]
    if (Sex=="д" or Sex=="Д" or Sex=="ж" or Sex=="Ж") and (Kata=="б" or Kata=="Б" or Kata=="b" or Kata=="B") and Age==7:
        Femail_B_7_Kata[i] = Person[i]
    if (Sex=="д" or Sex=="Д" or Sex=="ж" or Sex=="Ж") and (Kumite=="a" or Kumite=="A" or Kumite=="а" or Kumite=="А") and Age==7:
        Femail_A_7_Kumite[i] = Person[i]
    if (Sex=="д" or Sex=="Д" or Sex=="ж" or Sex=="Ж") and (Kumite=="б" or Kumite=="Б" or Kumite=="b" or Kumite=="B") and Age==7:
        Femail_B_7_Kumite[i] = Person[i]

# Девочки 8 лет
    if (Sex=="д" or Sex=="Д" or Sex=="ж" or Sex=="Ж") and (Kata=="a" or Kata=="A" or Kata=="а" or Kata=="А") and Age==8:
        Femail_A_8_Kata[i] = Person[i]
    if (Sex=="д" or Sex=="Д" or Sex=="ж" or Sex=="Ж") and (Kata=="б" or Kata=="Б" or Kata=="b" or Kata=="B") and Age==8:
        Femail_B_8_Kata[i] = Person[i]
    if (Sex=="д" or Sex=="Д" or Sex=="ж" or Sex=="Ж") and (Kumite=="a" or Kumite=="A" or Kumite=="а" or Kumite=="А") and Age==8:
        Femail_A_8_Kumite[i] = Person[i]
    if (Sex=="д" or Sex=="Д" or Sex=="ж" or Sex=="Ж") and (Kumite=="б" or Kumite=="Б" or Kumite=="b" or Kumite=="B") and Age==8:
        Femail_B_8_Kumite[i] = Person[i]

# Девочки 9 лет
    if (Sex == "д" or Sex == "Д" or Sex == "ж" or Sex == "Ж") and (Kata == "a" or Kata == "A" or Kata == "а" or Kata == "А") and Age == 9:
        Femail_A_9_Kata[i] = Person[i]
    if (Sex == "д" or Sex == "Д" or Sex == "ж" or Sex == "Ж") and (Kata == "б" or Kata == "Б" or Kata == "b" or Kata == "B") and Age == 9:
        Femail_B_9_Kata[i] = Person[i]
    if (Sex == "д" or Sex == "Д" or Sex == "ж" or Sex == "Ж") and (Kumite == "a" or Kumite == "A" or Kumite == "а" or Kumite == "А") and Age == 9:
        Femail_A_9_Kumite[i] = Person[i]
    if (Sex == "д" or Sex == "Д" or Sex == "ж" or Sex == "Ж") and (Kumite == "б" or Kumite == "Б" or Kumite == "b" or Kumite == "B") and Age == 9:
        Femail_B_9_Kumite[i] = Person[i]
    '''
# Девочки 10-11 лет
    if (Sex == "д" or Sex == "Д" or Sex == "ж" or Sex == "Ж") and (Group == "a" or Group == "A" or Group == "а" or Group == "А") and (Age == 10 or Age ==11) and Kata != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Ката девочки 10-11 лет группа А")
    if (Sex == "д" or Sex == "Д" or Sex == "ж" or Sex == "Ж") and (Group == "б" or Group == "Б" or Group == "b" or Group == "B") and (Age == 10 or Age ==11) and Kata != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Ката девочки 10-11 лет группа Б")
    if (Sex == "д" or Sex == "Д" or Sex == "ж" or Sex == "Ж") and (Group == "a" or Group == "A" or Group == "а" or Group == "А") and (Age == 10 or Age ==11) and Kumite != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ девочки 10-11 лет группа А")
    if (Sex == "д" or Sex == "Д" or Sex == "ж" or Sex == "Ж") and (Group == "б" or Group == "Б" or Group == "b" or Group == "B") and (Age == 10 or Age ==11) and Kumite != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ девочки 10-11 лет группа Б")

# Девочки 12-13 лет
    if (Sex == "д" or Sex == "Д" or Sex == "ж" or Sex == "Ж") and (Group == "a" or Group == "A" or Group == "а" or Group == "А") and (Age == 12 or Age == 13) and Kata != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Ката девочки 12-13 лет группа А")
    if (Sex == "д" or Sex == "Д" or Sex == "ж" or Sex == "Ж") and (Group == "б" or Group == "Б" or Group == "b" or Group == "B") and (Age == 12 or Age == 13) and Kata != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Ката девочки 12-13 лет группа Б")
    if (Sex == "д" or Sex == "Д" or Sex == "ж" or Sex == "Ж") and (Group == "a" or Group == "A" or Group == "а" or Group == "А") and (Age == 12 or Age == 13) and Kumite != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ девочки 12-13 лет группа А")
    if (Sex == "д" or Sex == "Д" or Sex == "ж" or Sex == "Ж") and (Group == "б" or Group == "Б" or Group == "b" or Group == "B") and (Age == 12 or Age == 13) and Kumite != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ девочки 12-13 лет группа Б")

# Девушки 14-15 лет
    if (Sex == "д" or Sex == "Д" or Sex == "ж" or Sex == "Ж") and (Group == "a" or Group == "A" or Group == "а" or Group == "А") and (Age == 14 or Age == 15) and Kata != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Ката девушки 14-15 лет группа А")
    if (Sex == "д" or Sex == "Д" or Sex == "ж" or Sex == "Ж") and (Group == "б" or Group == "Б" or Group == "b" or Group == "B") and (Age == 14 or Age == 15) and Kata != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Ката девушки 14-15 лет группа Б")
    if (Sex == "д" or Sex == "Д" or Sex == "ж" or Sex == "Ж") and (Group == "a" or Group == "A" or Group == "а" or Group == "А") and (Age == 14 or Age == 15) and Kumite != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ девушки 14-15 лет группа А")
    if (Sex == "д" or Sex == "Д" or Sex == "ж" or Sex == "Ж") and (Group == "б" or Group == "Б" or Group == "b" or Group == "B") and (Age == 14 or Age == 15) and Kumite != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ девушки 14-15 лет группа Б")

# Девушки 16-17 лет
    if (Sex == "д" or Sex == "Д" or Sex == "ж" or Sex == "Ж") and (Group == "a" or Group == "A" or Group == "а" or Group == "А") and (Age == 16 or Age == 17) and Kata != "":
            print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Ката девушки 16-17 лет группа А")
    if (Sex == "д" or Sex == "Д" or Sex == "ж" or Sex == "Ж") and (Group == "б" or Group == "Б" or Group == "b" or Group == "B") and (Age == 16 or Age == 17) and Kata != "":
            print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Ката девушки 16-17 лет группа Б")
    if (Sex == "д" or Sex == "Д" or Sex == "ж" or Sex == "Ж") and (Group == "a" or Group == "A" or Group == "а" or Group == "А") and (Age == 16 or Age == 17) and Kumite != "":
            print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ девушки 16-17 лет группа А")
    if (Sex == "д" or Sex == "Д" or Sex == "ж" or Sex == "Ж") and (Group == "б" or Group == "Б" or Group == "b" or Group == "B") and (Age == 16 or Age == 17) and Kumite != "":
            print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ девушки 16-17 лет группа Б")

#Женщины 18+ лет
    if (Sex == "д" or Sex == "Д" or Sex == "ж" or Sex == "Ж") and (Group == "a" or Group == "A" or Group == "а" or Group == "А") and (Age >= 18 and Age < 35) and Kata != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Ката женщины 18+ лет группа А")
    if (Sex == "д" or Sex == "Д" or Sex == "ж" or Sex == "Ж") and (Group == "б" or Group == "Б" or Group == "b" or Group == "B") and (Age >= 18 and Age < 35) and Kata != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Ката женщины 18+ лет группа Б")
    if (Sex == "д" or Sex == "Д" or Sex == "ж" or Sex == "Ж") and (Group == "a" or Group == "A" or Group == "а" or Group == "А") and (Age >= 18 and Age < 35) and Kumite != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ женщины 18+ лет группа А")
    if (Sex == "д" or Sex == "Д" or Sex == "ж" or Sex == "Ж") and (Group == "б" or Group == "Б" or Group == "b" or Group == "B") and (Age >= 18 and Age < 35) and Kumite != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ женщины 18+ лет группа Б")

# Мальчики до 6 лет
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group=="a" or Group=="A" or Group=="а" or Group=="А") and Age<=6 and Kata!="":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Ката мальчики до 6 лет группа А")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group=="б" or Group=="Б" or Group=="b" or Group=="B") and Age<=6 and Kata!="":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Ката мальчики до 6 лет группа Б")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group=="a" or Group=="A" or Group=="а" or Group=="А") and Age<=6 and Kumite != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ мальчики до 6 лет группа А")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group=="б" or Group=="Б" or Group=="b" or Group=="B") and Age<=6 and Kumite != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ мальчики до 6 лет группа Б")

# Мальчики 6 лет
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group=="a" or Group=="A" or Group=="а" or Group=="А") and Age==6 and Kata!="":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Ката мальчики 6 лет группа А")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group=="б" or Group=="Б" or Group=="b" or Group=="B") and Age==6 and Kata!="":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Ката мальчики 6 лет группа Б")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group=="a" or Group=="A" or Group=="а" or Group=="А") and Age==6 and Weight < 22 and Kumite != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ мальчики 6 лет группа А")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group=="б" or Group=="Б" or Group=="b" or Group=="B") and Age==6 and Weight < 22 and Kumite != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ мальчики 6 лет группа Б")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group=="a" or Group=="A" or Group=="а" or Group=="А") and Age==6 and Weight >= 22 and Kumite != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ мальчики 6 лет группа А")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group=="б" or Group=="Б" or Group=="b" or Group=="B") and Age==6 and Weight >= 22 and Kumite != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ мальчики 6 лет группа Б")

# Мальчики 7 лет
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group=="a" or Group=="A" or Group=="а" or Group=="А") and Age==7 and Kata!="":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Ката мальчики 7 лет группа А")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group=="б" or Group=="Б" or Group=="b" or Group=="B") and Age==7 and Kata!="":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Ката мальчики 7 лет группа Б")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group=="a" or Group=="A" or Group=="а" or Group=="А") and Age==7 and Weight < 25 and Kumite != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ мальчики 7 лет группа А")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group=="б" or Group=="Б" or Group=="b" or Group=="B") and Age==7 and Weight < 25 and Kumite != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ мальчики 7 лет группа Б")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group=="a" or Group=="A" or Group=="а" or Group=="А") and Age==7 and Weight >= 25 and Kumite != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ мальчики 7 лет группа А")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group=="б" or Group=="Б" or Group=="b" or Group=="B") and Age==7 and Weight >= 25 and Kumite != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ мальчики 7 лет группа Б")

# Мальчики 8 лет
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group=="a" or Group=="A" or Group=="а" or Group=="А") and Age==8 and Kata!="":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Ката мальчики 8 лет группа А")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group=="б" or Group=="Б" or Group=="b" or Group=="B") and Age==8 and Kata!="":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Ката мальчики 8 лет группа Б")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group=="a" or Group=="A" or Group=="а" or Group=="А") and Age==8 and Weight < 28 and Kumite!="":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ мальчики 8 лет группа А")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group=="б" or Group=="Б" or Group=="b" or Group=="B") and Age==8 and Weight < 28 and Kumite!= "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ мальчики 8 лет группа Б")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group=="a" or Group=="A" or Group=="а" or Group=="А") and Age==8 and Weight >= 28 and Kumite!="":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ мальчики 8 лет группа А")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group=="б" or Group=="Б" or Group=="b" or Group=="B") and Age==8 and Weight >= 28 and Kumite!= "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ мальчики 8 лет группа Б")

# Мальчики 9 лет
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group == "a" or Group == "A" or Group == "а" or Group == "А") and Age == 9 and Kata != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Ката мальчики 9 лет группа А")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group == "б" or Group == "Б" or Group == "b" or Group == "B") and Age == 9 and Kata != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Ката мальчики 9 лет группа Б")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group == "a" or Group == "A" or Group == "а" or Group == "А") and Age == 9 and Weight < 30 and Kumite != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ мальчики 9 лет группа А")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group == "б" or Group == "Б" or Group == "b" or Group == "B") and Age == 9 and Weight < 30 and Kumite != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ мальчики 9 лет группа Б")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group == "a" or Group == "A" or Group == "а" or Group == "А") and Age == 9 and Weight >= 30 and Kumite != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ мальчики 9 лет группа А")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group == "б" or Group == "Б" or Group == "b" or Group == "B") and Age == 9 and Weight >= 30 and Kumite != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ мальчики 9 лет группа Б")

# Мальчики 10-11 лет
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group == "a" or Group == "A" or Group == "а" or Group == "А") and (Age == 10 or Age ==11) and Kata != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Ката мальчики 10-11 лет группа А")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group == "б" or Group == "Б" or Group == "b" or Group == "B") and (Age == 10 or Age ==11) and Kata != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Ката мальчики 10-11 лет группа Б")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group == "a" or Group == "A" or Group == "а" or Group == "А") and (Age == 10 or Age ==11) and Weight < 37 and Kumite != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ мальчики 10-11 лет группа А")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group == "б" or Group == "Б" or Group == "b" or Group == "B") and (Age == 10 or Age ==11) and Weight < 37 and Kumite != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ мальчики 10-11 лет группа Б")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group == "a" or Group == "A" or Group == "а" or Group == "А") and (Age == 10 or Age ==11) and Weight >= 37 and Kumite != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ мальчики 10-11 лет группа А")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group == "б" or Group == "Б" or Group == "b" or Group == "B") and (Age == 10 or Age ==11) and Weight >= 37 and Kumite != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ мальчики 10-11 лет группа Б")

# Мальчики 12-13 лет
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group == "a" or Group == "A" or Group == "а" or Group == "А") and (Age == 12 or Age == 13) and Kata != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Ката мальчики 12-13 лет группа А")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group == "б" or Group == "Б" or Group == "b" or Group == "B") and (Age == 12 or Age == 13) and Kata != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Ката мальчики 12-13 лет группа Б")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group == "a" or Group == "A" or Group == "а" or Group == "А") and (Age == 12 or Age == 13) and Weight < 46 and Kumite != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ мальчики 12-13 лет группа А")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group == "б" or Group == "Б" or Group == "b" or Group == "B") and (Age == 12 or Age == 13) and Weight < 46 and Kumite != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ мальчики 12-13 лет группа Б")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group == "a" or Group == "A" or Group == "а" or Group == "А") and (Age == 12 or Age == 13) and Weight >= 46 and Kumite != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ мальчики 12-13 лет группа А")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group == "б" or Group == "Б" or Group == "b" or Group == "B") and (Age == 12 or Age == 13) and Weight >= 46 and Kumite != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ мальчики 12-13 лет группа Б")

# Юноши 14-15 лет
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group == "a" or Group == "A" or Group == "а" or Group == "А") and (Age == 14 or Age == 15) and Kata != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Ката юноши 14-15 лет группа А")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group == "б" or Group == "Б" or Group == "b" or Group == "B") and (Age == 14 or Age == 15) and Kata != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Ката юноши 14-15 лет группа Б")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group == "a" or Group == "A" or Group == "а" or Group == "А") and (Age == 14 or Age == 15) and Weight < 60 and Kumite != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ юноши 14-15 лет группа А")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group == "б" or Group == "Б" or Group == "b" or Group == "B") and (Age == 14 or Age == 15) and Weight < 60 and Kumite != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ юноши 14-15 лет группа Б")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group == "a" or Group == "A" or Group == "а" or Group == "А") and (Age == 14 or Age == 15) and Weight >= 60 and Kumite != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ юноши 14-15 лет группа А")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group == "б" or Group == "Б" or Group == "b" or Group == "B") and (Age == 14 or Age == 15) and Weight >= 60 and Kumite != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ юноши 14-15 лет группа Б")

# Юноши 16-17 лет
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group == "a" or Group == "A" or Group == "а" or Group == "А") and (Age == 16 or Age == 17) and Kata != "":
            print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Ката юноши 16-17 лет группа А")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group == "б" or Group == "Б" or Group == "b" or Group == "B") and (Age == 16 or Age == 17) and Kata != "":
            print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Ката юноши 16-17 лет группа Б")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group == "a" or Group == "A" or Group == "а" or Group == "А") and (Age == 16 or Age == 17) and Weight < 68 and Kumite != "":
            print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ юноши 16-17 лет группа А")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group == "б" or Group == "Б" or Group == "b" or Group == "B") and (Age == 16 or Age == 17) and Weight < 68 and Kumite != "":
            print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ юноши 16-17 лет группа Б")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group == "a" or Group == "A" or Group == "а" or Group == "А") and (Age == 16 or Age == 17) and Weight >= 68 and Kumite != "":
            print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ юноши 16-17 лет группа А")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group == "б" or Group == "Б" or Group == "b" or Group == "B") and (Age == 16 or Age == 17) and Weight >= 68 and Kumite != "":
            print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ юноши 16-17 лет группа Б")

#Мужчины 18+ лет
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group == "a" or Group == "A" or Group == "а" or Group == "А") and (Age >= 18 and Age < 35) and Kata != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Ката мужчины 18+ лет группа А")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group == "б" or Group == "Б" or Group == "b" or Group == "B") and (Age >= 18 and Age < 35) and Kata != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Ката мужчины 18+ лет группа Б")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group == "a" or Group == "A" or Group == "а" or Group == "А") and (Age >= 18 and Age < 35) and Kumite != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ мужчины 18+ лет группа А")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group == "б" or Group == "Б" or Group == "b" or Group == "B") and (Age >= 18 and Age < 35) and Kumite != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ мужчины 18+ лет группа Б")

#Мужчины 35+ лет
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group == "a" or Group == "A" or Group == "а" or Group == "А") and (Age >= 18 and Age < 35) and Kata != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Ката мужчины 35+ лет группа А")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group == "б" or Group == "Б" or Group == "b" or Group == "B") and (Age >= 18 and Age < 35) and Kata != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Ката мужчины 35+ лет группа Б")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group == "a" or Group == "A" or Group == "а" or Group == "А") and (Age >= 18 and Age < 35) and Kumite != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ мужчины 35+ лет группа А")
    if (Sex=="м" or Sex=="М" or Sex=="ч" or Sex=="Ч") and (Group == "б" or Group == "Б" or Group == "b" or Group == "B") and (Age >= 18 and Age < 35) and Kumite != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, Qty_String, "Кумитэ мужчины 35+ лет группа Б")

#Джунро микс
    if Age >= 9 and Age <= 11 and Dzunro != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, "Джунро 9-11 лет")
    if Age >= 12 and Age <= 15 and Dzunro != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, "Джунро 12-15 лет")
    if Age >= 16 and Dzunro != "":
        print(Worksheet_Competitor.cell(Qty_String, 1).value, "Джунро 16+ лет")

'''

print("Девочки Група А 6 лет ката")
print(Femail_A_6_Kata)
print()
print("Девочки Група B 6 лет ката")
print(Femail_B_6_Kata)
print()
print("Девочки Група A 6 лет кумитэ")
print(Femail_A_6_Kumite)
print()
print("Девочки Група B 6 лет кумитэ")
print(Femail_B_6_Kumite)
print()

print("Девочки Група А 7 лет ката")
print(Femail_A_7_Kata)
print()
print("Девочки Група B 7 лет ката")
print(Femail_B_7_Kata)
print()
print("Девочки Група A 7 лет кумитэ")
print(Femail_A_7_Kumite)
print()
print("Девочки Група B 7 лет кумитэ")
print(Femail_B_7_Kumite)
print()

print("Девочки Група А 8 лет ката")
print(Femail_A_8_Kata)
print()
print("Девочки Група B 8 лет ката")
print(Femail_B_8_Kata)
print()
print("Девочки Група A 8 лет кумитэ")
print(Femail_A_8_Kumite)
print()
print("Девочки Група B 8 лет кумитэ")
print(Femail_B_8_Kumite)
print()

print("Девочки Група А 9 лет ката")
print(Femail_A_9_Kata)
print()
print("Девочки Група B 9 лет ката")
print(Femail_B_9_Kata)
print()
print("Девочки Група A 9 лет кумитэ")
print(Femail_A_9_Kumite)
print()
print("Девочки Група B 9 лет кумитэ")
print(Femail_B_9_Kumite)
print()