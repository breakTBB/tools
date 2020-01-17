# @Author Prism17
# @Date 1/17/2020
# 学习委员用，教务处成绩处理。

import openpyxl

class Course:
    name = ''
    weight = 0
    score = 0

    def __init__(self, name, weight, score):
        self.name = name
        self.weight = weight
        self.score = score

    def __lt__(self, other):
        if self.weight != other.weight:
            return self.weight < other.weight
        return self.name < other.name

    def __hash__(self):
        return hash(self.name)

    def __eq__(self, other):
        return self.name == other.name


class Student:
    sid = 201803523
    name = '张李政'
    courses = []
    gpa = 5.0

    def __init__(self, sid, sname):
        self.sid = sid
        self.name = sname
        self.courses = []

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
cs = set()
footer = True

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
    if attr != '必修' or ava != '是':
        continue

    cs.add(Course(courseName, weight, score))
    try:
        datas[sid].courses.append(Course(courseName, weight, score))
    except:
        datas.update({sid: Student(sid, name)})
        datas[sid].courses.append(Course(courseName, weight, score))

outs = []
for data in datas.values():
    data.setGPA()
    data.courses.sort()
    outs.append(data)

# save the result

cs = list(cs)
cs.sort(reverse=True)
header = ['学号', '姓名'] + [i.name + '(' + i.weight + ')' for i in cs] + ['GPA']

wb = openpyxl.Workbook()
fname = '计教一班必修加权成绩.xlsx'
ws = wb.active

ws.append(header)

outs.sort(reverse=True)

for o in outs:
    print(o.name)
    print('加权平均分: ', end='')
    print(o.gpa)
    print('-' * 20)
    d = [o.sid, o.name]
    for s in cs:
        fd = '0'
        for c in o.courses:
            if c.name == s.name:
                fd = c.score
        d.append(fd)
    d.append(o.gpa)
    ws.append(d)
print('本班共计' + str(len(datas)) + '人')

wb.save(fname)
