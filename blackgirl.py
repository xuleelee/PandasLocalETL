import random

# Score = int(input("请输入分数: "))
# Score = random.randint(0,100)
# print(Score)
# if Score>=90:
#     print('A')
# elif Score >= 80:
#     print('B')
# elif Score>=60:
#     print('C')
# elif Score>=40:
#     print('D')
# else:
#     print('E')


# print(random.randint(0,10))

# count = 0
# while count<100:
#     print(count)
#     count+=1

#
# for i in range(1,10):
#     print()
#     for b in range (1,10):
#         if b<=i:
#             print(i,"*",b,"=",i*b,"  ",end=" ")

# for i in range (1,10):
#     print(i)

# for i in range(11):
#     if i<=5:
#         print("*"*i)
#     else:
#         print("*"*(10-i))

saving = 10000
count = 1
while saving <=20000:
    saving = ((saving *0.0325) + saving)
    count+=1
    print(count)
    print(saving)