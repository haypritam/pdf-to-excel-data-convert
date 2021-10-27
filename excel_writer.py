import pandas as pd
import numpy as np
import os

def column(c):
    q = c // 26
    r = c % 26
    if q > 0:
        col = chr(65 + q - 1)+chr(65 + r)
    else:
        col = chr(65 + r)
    return col

def getColRange(start, end):
    col1 = column(start)
    col2 = column(end)
    return col1 + ':' + col2

def getCellRange(start_row, start_col, end_row, end_col):
    return column(start_col) + str(start_row + 1) + ':' + column(end_col) + str(end_row + 1)

def getPermittedValues(marks):
    return [x for x in np.arange(0, float(marks)+0.5, 0.5)]

def getSubQuestion(pm):
    info = pm.split('×')
    parts = int(info[1].split('=')[0])
    total = int(info[1].split('=')[1])
    if total / parts != info[0]:
        marks = total / parts
    else:
        marks = info[0]
    return (parts, marks)

def addDataValidation(ws, start_row, start_col, end_row, end_col, fm):
    cell_range = getCellRange(start_row, start_col, end_row, end_col)
    validation = {'validate': 'list', 'source': getPermittedValues(fm)}
    ws.data_validation(cell_range, validation)

#=IFERROR(SUM(LARGE(F8:L8,{1,2,3,4})),SUM(F8:L8))
def write_questions(ws, r, c, questions, nr_of_students, col_fmt):
    qCols = []
    r1, c1 = r, c
    for q in questions.keys():
        s = c
        if len(questions[q]['parts']) > len(questions[q]['partmarks']):
            (ans, marks) = getSubQuestion(questions[q]['partmarks'][0])
            for p in questions[q]['parts']:
                ws.write(r, c, q+'.'+p)
                addDataValidation(ws, r + 1, c, r + nr_of_students, c, marks)
                c += 1
            ws.set_column(getColRange(s, c - 1), 5, None, {'level' : 1, 'hidden': True})
            for row in range(1, nr_of_students + 1):
                cell_range = getCellRange(r + row, s, r + row, c -  1)
                formula = "=iferror(sum(large({},{})), sum({}))".format(
                          cell_range, '{'+','.join([str(x) for x in range(1, ans+1)])+'}', cell_range)
                ws.write_formula(r+row, c, formula, col_fmt)
        else:
            for i in range(len(questions[q]['partmarks'])):
                pm = questions[q]['partmarks'][i].split('+')
                p = ''
                if i < len(questions[q]['parts']):
                    p = questions[q]['parts'][i]
                if len(pm) > 1:
                    for n in range(len(pm)):
                        ws.write(r, c, q+'.'+p+'.'+str(n+1))
                        addDataValidation(ws, r + 1, c, r + nr_of_students, c, pm[n])
                        c += 1
                else:
                    if '×' in questions[q]['partmarks'][i]:
                        (subparts, submarks) = getSubQuestion(questions[q]['partmarks'][i])
                        for k in range(subparts):
                            ws.write(r, c, q+'.'+p+'.'+str(k+1))
                            addDataValidation(ws, r + 1, c, r + nr_of_students, c, submarks)
                            c += 1
                    else:
                        ws.write(r, c, q+'.'+p)
                        addDataValidation(ws, r + 1, c, r + nr_of_students, c, pm[0])
                        c += 1
            ws.set_column(getColRange(s, c - 1), 5, None, {'level' : 1, 'hidden': True})
            for row in range(1, nr_of_students + 1):
                cell_range = getCellRange(r + row, s, r + row, c -  1)
                ws.write_formula(r+row, c, "=sum({})".format(cell_range), col_fmt)
        ws.write(r, c, 'Q'+q, col_fmt)
        qCols.append(column(c))
        c += 1
    return qCols

def write_head(ws, rows):
    lines = 1
    for r in range(len(rows)):
        for c in range(len(rows[r])):
            if pd.notnull(rows[r][c]):
                ws.write(r, c, rows[r][c])
            else:
                ws.write(r, c, '')
            lines += 1
    return lines

def write_body(ws, rows, nr_of_students, start, nameCol, nam_format, col_format, qCols):
    for r in range(len(rows)):
        for c in range(len(rows[r])):
            (width, fmt) = (15, col_format) if c != nameCol else (45, nam_format)
            if pd.notnull(rows[r][c]):
                ws.write(start+r, c, rows[r][c], fmt)
            else:
               ws.write_formula(start + r, c, 
                                "=ceiling(sum({}), 1)".format(",".join([x + str(start+r+1) for x in qCols])), 
                                fmt) 
            ws.set_column(c, c, width)

def write_title(ws, cols, nStudents, row, col_fmt, questions):
    for col in range(len(cols)):
        ws.write(row, col, cols[col], col_fmt)
    return write_questions(ws, row, len(cols), questions, nStudents, col_fmt)

def write_xl(template, questions):
    #questions = review(questions)
    print(questions)
    filename = os.path.basename(template).split('.')[0] + "_paperchecking.xlsx"

    infile = pd.ExcelFile(template)
    head = pd.read_excel(infile, header=None, nrows=5, usecols = [0])
    title = pd.read_excel(infile, skiprows=6, nrows=1)
    students = pd.read_excel(infile, header=None, dtype=np.str, skiprows=7)
    nr_of_students = len(students.index)

    writer = pd.ExcelWriter(filename, engine = "xlsxwriter")
    wb = writer.book
    col_fmt = wb.add_format({'num_format' : '0.0;;', 'align' : 'center', 'border' : 1})
    nam_fmt = wb.add_format({'align' : 'left', 'border' : 1})
    ws = wb.add_worksheet()
    lines = write_head(ws, head.values.tolist())
    nameCol = title.columns.get_loc('NAME')
    qCols = write_title(ws, title.columns.values.tolist(), nr_of_students, lines, col_fmt, questions)
    write_body(ws, students.values.tolist(), nr_of_students, 
               lines+1, nameCol, nam_fmt, col_fmt, qCols)

    writer.save()
