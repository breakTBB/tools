# @Author Prism17
# @Date 1/17/2020
# 学习委员用，教务处成绩处理。

import openpyxl


class Course:

    def __init__(self, n, w, s):
        self.name = n
        self.weight = w
        self.score = s

    def __lt__(self, other):
        if self.weight != other.weight:
            return self.weight < other.weight
        return self.name < other.name

    def __hash__(self):
        return hash(self.name)

    def __eq__(self, other):
        return self.name == other.name


class Student:

    def __init__(self, i, n):
        self.sid = i
        self.name = n
        self.courses = []
        self.gpa = 0

    def __lt__(self, other):
        if self.gpa != other.gpa:
            return self.gpa < other.gpa
        return self.name < other.name

    def setGPA(self):
        p = 0
        q = 0
        for co in self.courses:
            we = float(co.weight)
            sc = float(co.score)
            p += we * sc
            q += we
        self.gpa = p / q


wb = openpyxl.load_workbook('score.xlsx')
score = wb.get_active_sheet()

lines = list(score.rows)

datas = {}
courseList = set()
footer = True

# process the excel file

for line in lines:
    sid = line[1].value
    name = line[2].value
    courseName = line[3].value
    attr = line[4].value
    weight = line[5].value
    score = line[8].value
    ava = line[10].value

    if footer:
        footer = False
        continue
    if attr != '必修' or ava == '否':
        continue

    courseList.add(Course(courseName, weight, score))
    try:
        datas[sid].courses.append(Course(courseName, weight, score))
    except:
        datas.update({sid: Student(sid, name)})
        datas[sid].courses.append(Course(courseName, weight, score))

outs = []
for data in datas.values():
    data.setGPA()
    outs.append(data)

# save the result

courseList = list(courseList)
courseList.sort(reverse=True)
header = ['学号', '姓名'] + [i.name + '(' + i.weight + ')' for i in courseList] + ['必修加权成绩']

wb = openpyxl.Workbook()
fileName = '计教一班必修加权成绩.xlsx'
ws = wb.active

ws.append(header)

outs.sort(reverse=True)

WIDTH = 30
line = ['*' for i in range(4)]
for o in outs:
    line[0] = '-' * WIDTH
    line[1] = o.name
    line[2] = '加权平均分: ' + str(o.gpa)
    line[3] = '-' * WIDTH + '\n\n'
    for i in range(4):
        print(line[i])
    d = [o.sid, o.name]
    for s in courseList:
        fd = '0'
        for c in o.courses:
            if c.name == s.name:
                fd = c.score
        d.append(fd)
    d.append(o.gpa)
    ws.append(d)
print('共计' + str(len(datas)) + '人')

wb.save(fileName)
