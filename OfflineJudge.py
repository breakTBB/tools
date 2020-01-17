from os import system

exec = 'a.exe'
prob = input('问题名: ')
cnt = int(input('数据组数: '))
for i in range(1, cnt + 1):
    cmd = "a.exe < %s > %s" % (prob + str(i) + '.in', "out")
    print(cmd)
    print('i = %d' % i)
    system(cmd)
    if system("fc out " + prob + str(i) + ".ans"):
        print("Passed #" + str(i))
    else:
        print("GG")
        break
