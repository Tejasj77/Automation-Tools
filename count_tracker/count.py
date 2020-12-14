import pandas as pd
import xlwings as xw
import datetime


book = xw.Book('count.xlsx')
shit = book.sheets('Sheet1')

now = datetime.datetime.now()
time = now.strftime("%d %y")

main_frame = pd.read_excel('count.xlsx',name='Sheet1')

date = list(main_frame['Date'])
project = list(main_frame['Project'])
section = list(main_frame['Section'])
type = list(main_frame['Type'])
count = list(main_frame['Count'])
total = list(main_frame['Total'])

proj_inp = []
sec_inp = []
type_inp = []
count_inp = []

list1 = []
list2 = []
list3 = []
list4 = []
sum2 = 0
totality = 0
summation = 0

number_of_programs = int(input("Enter the number of programs"))
for i in range(number_of_programs):
    proj_inp.append(str(input("Enter the project")))
    sec_inp.append(str(input("Enter the section")))
    type_inp.append(str(input("Enter the type")))
    count_inp.append(int(input("Enter the count")))

def printing():
    shit.range('A' + str(len(date) + 2)).value = time
    shit.range('B' + str(len(project) + 2)).options(transpose=True).value = proj_inp
    shit.range('C' + str(len(section) + 2)).options(transpose=True).value = sec_inp
    shit.range('D' + str(len(type) + 2)).options(transpose=True).value = type_inp
    shit.range('E' + str(len(count) + 2)).options(transpose=True).value = count_inp

def list_cut(list,i):
    fin_list = list[i:]
    return fin_list

def list_sum(list):
    y = 0
    for i in range(len(list)):
        y = y + list[i]
    return y

def list_remove(list,iter):
    list.remove(list[iter])


def del_check():
    for iter in range(len(list1)):
        proj_inp.remove(list1[iter])
    for iter in range(len(list2)):
        sec_inp.remove(list2[iter])
    for iter in range(len(list3)):
        type_inp.remove(list3[iter])
    for iter in range(len(list4)):
        count_inp.remove(list4[iter])




def same_check():
    #print("And we should be here")
    for iter in range(len(proj_inp)):
        for iter1 in range(len(project1)):
            if project1[iter1] == proj_inp[iter]:
                if section1[iter1] == sec_inp[iter]:
                    if type1[iter1] == type_inp[iter]:
                        #print("Increment the count")
                        x = 0
                        x = count1[iter1] + count_inp[iter]
                        list1.append(proj_inp[iter])
                        list2.append(sec_inp[iter])
                        list3.append(type_inp[iter])
                        list4.append(count_inp[iter])
                        shit.range('E' + str(reference + iter1 + 2)).value = x
                    else:
                        continue
                else:
                    continue
            else:
                continue


new_sum = 0
existing_sum = 0
new_sum = list_sum(count_inp)
reference = 0

if len(date):
        if time in date:
            for iter in range(len(date)):
                if date[iter] == time:
                    project1 = list_cut(project,iter)
                    section1 = list_cut(section,iter)
                    type1 = list_cut(type,iter)
                    count1 = list_cut(count,iter)
                    existing_sum = list_sum(count1)
                    reference = iter
                    same_check()
                    del_check()
                    #print(proj_inp,sec_inp,type_inp,count_inp)
                    for iter in range(len(proj_inp)):
                        shit.range('B' + str(len(project1) + reference + 2)).value = proj_inp[iter]
                        shit.range('C' + str(len(section1) + reference + 2)).value = sec_inp[iter]
                        shit.range('D' + str(len(type1) + reference + 2)).value = type_inp[iter]
                        shit.range('E' + str(len(count1) + reference + 2)).value = count_inp[iter]
        else:
            #print("But we are here")
            printing()
            sum1 = list_sum(count_inp)
            print(sum1)
            shit.range('F' + str(len(type) + 2)).value = sum1

else:
    print("But we are here")
    printing()
    sum2 = list_sum(count_inp)
    print(sum2)
    shit.range('F' + str(len(type) + 2)).value = sum2

"""
else:
    printing()
    sum2 = list_sum(count_inp)
    print(sum2)
    shit.range('F' + str(2)).value = sum2
"""

"""
number_of_programs = int(input("Enter the number of programs"))
for x in range(number_of_programs):
    different.append(input('Which program did you work on Today'))
    number.append(int(input("How many requirements did you work on?")))
    type1.append(str(input("What type was it ? Press A for analyzed, Press D for drafted , Press R for rework")))

print(number_of_programs)
print(different)
print(number)
print(type1)



print(type(time))
print(type(date))
for i in range(len(date)):
    print(date[i])
    print(type(date[i]))
    if date[i] == time:
        print(date[i])
        print("Here")
        print(time)


if len(date):
    for iter in range(len(date)):
        if date[iter] == time:
            break
        else:
            shit.range('C' + str(len(date) + 3)).value = time
else:
    shit.range('C' + str(len(date) + 3)).value = time

#Program
program = main_frame['Program']
if len(program):
    for iter2 in range(len(program)):
        shit.range('D' + str(len(program) + 3)).options(transpose=True).value = different
else:
    shit.range('D' + str(len(program) + 3)).options(transpose=True).value = different

#No.of requirements
requirements = main_frame['No.of requirements']
if len(requirements):
    for iter3 in range(len(requirements)):
        shit.range('E' + str(len(requirements) + 3)).options(transpose=True).value = number
else:
    shit.range('E' + str(len(requirements) + 3)).options(transpose=True).value = number

#Type_of_requirement
types = main_frame['Type']
if len(types):
    for iter4 in range(len(types)):
        for iter5 in range(len(type1)):
            if type1[iter5] =='A':
                shit.range('F' + str(len(types) + 3)).value = 'Analyzed'
            elif type1[iter5] =='D':
                shit.range('F' + str(len(types) + 3)).value = 'Drafted'
            elif type1[iter5] =='R':
                shit.range('F' + str(len(types) + 3)).value = "Reworked"
else:
    for iter4 in range(len(types)):
        for iter5 in range(len(type1)):
            if type1[iter5] == "A":
                shit.range('F' + str(len(types) + 3)).value = 'Analyzed'
            elif type1[iter5] == "D":
                shit.range('F' + str(len(types) + 3)).value = 'Drafted'
            elif type1[iter5] == "R":
                shit.range('F' + str(len(types) + 3)).value = "Reworked"

#Total
y = 0
for iter6 in range(len(number)):
    y = y + number[iter6]
print(y)
total = main_frame['Total']
if len(total):
    for iter7 in range(len(total)):
        shit.range('G' + str(len(total) + 3)).value = y
else:
    shit.range('G' + str(len(total) + 3)).value = y

book.save()

"""
"""
def same_check(list1,list2,list3,list4,number):
    for iter in range(len(proj_inp)):
        for iter1 in range(len(list1)):
            if proj_inp[iter] in list1:
                print("The project exists")
                if proj_inp[iter] == list1[iter1]:
                    if sec_inp[iter] == list2[iter1]:
                        if type_inp[iter] == list3[iter1]:
                            print("Increase the count")
                            x = 0
                            x = list4[iter1] + count_inp[iter]
                            print(x)
                            shit.range('E' + str(number + iter1 + 2)).value = x
            else:
                print("The project does not exist")
                shit.range('B' + str(number + iter + 2)).value = proj_inp[iter]
                shit.range('C' + str(number + iter + 2)).value = sec_inp[iter]
                shit.range('D' + str(number + iter + 2)).value = type_inp[iter]
                shit.range('E' + str(number + iter + 2)).value = count_inp[iter]
"""

final_sum = existing_sum + new_sum
shit.range('F' + str(reference + 2)).value = final_sum
book.save()
book.close()
