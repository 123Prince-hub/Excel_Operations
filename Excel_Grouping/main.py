import xlwings as xw
import random
from time import sleep

wb = xw.Book("data.xlsx").sheets['data']
ws = wb.range("A2").expand().options(transpose=True, numbers=int).value
rowlen = str(len(ws))
amount = ws[2]
amount.sort() 


total = 0
count = 0
while((count < 10) or (total < 10000)):
    total += amount[count]
    count += 1
    print(amount[count])
print(total)



# using random ************************************
# total = 0
# count = 0
# while(count != 10):
#     guess = random.choices(amount)
#     for gu in guess:
#         guess = int(gu)
#     print(f":- {guess}")
#     count += 1
#     total += guess
#     if (total >= 9000) & (total <= 10000):
#         print(f"Total_1 : {total}")
#         break
#     elif (total >= 8000) & (total <= 9000):
#         total += guess
#         if total > 10000:
#             total -= guess
#         else:
#             total += guess
#         print(f"Total_2 : {total}")
#         break
#     elif (total >= 7000) & (total <= 8000):
#         total += guess
#         if total > 10000:
#             total -= guess
#         else:
#             total += guess
#         print(f"Total_3 : {total}")
#         break
#     elif (total >= 6000) & (total <= 7000):
#         total += guess
#         if total > 10000:
#             total -= guess
#         else:
#             total += guess
#         print(f"Total_4 : {total}")
#         break
#     else:
#         pass


# without random ************************************
# while (amount != False):
#     total = 0
#     count = 0
#     while(count != 10):
#         guess = amount[count]
#         total += guess
#         count += 1
#         # print(guess)
#         # print(count)
#         if total <= 10000:
#             print(f"Total_1st : ", total)
#             # break
#         elif total >= 10000:
#             total -= guess 
#             print(f"Total_2nd : ", total)
#             break
#         else:
#             print(f"Total_3rd : ", total)
#             break

#         if count == 10:
#             del amount[0:10]
#             print(amount)
        

    # if (total >= 9000) & (total <= 10000):
    #     print(f"Total_1 : {total}")
    #     break
    # elif (total >= 8000) & (total <= 9000):
    #     # total += guess
    #     print(f"Total_2 : {total}")
    #     break
    # elif (total >= 7000) & (total <= 8000):
    #     # total += guess
    #     print(f"Total_3 : {total}")
    #     break
    # elif (total >= 6000) & (total <= 7000):
    #     # total += guess
    #     print(f"Total_4 : {total}")
    #     break
    # else:
    #     # print(f"Total_5 : {total}")
    #     # break
    #     pass
        


# total = 0
# count = 0
# for amnt in amount:
#     total = total + int(amnt)
#     count += 1
#     if count == 10:
#         break
# del amount[0:count]
# Group_1st = total
# print(f"First 10 Total :- {Group_1st}")
    

# total = 0
# count = 0
# for amnt in amount:
#     total = total + int(amnt)
#     count += 1
#     if count == 10:
#         break
# del amount[0:count]
# Group_2nd = total
# print(f"Second 10 Total :- {Group_2nd}")


# while True:
#     if  9000 > total :
#         pass



    