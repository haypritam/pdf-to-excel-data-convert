import PyPDF2
import sys, re, os, io

def parseQuestion(qfile):
    with open(qfile, 'rb') as pdf:
        pdfReader = PyPDF2.PdfFileReader(pdf)
        q = 0
        questions = {}
        for pageNo in range(pdfReader.numPages):
            txt = pdfReader.getPage(pageNo).extractText()
            txt = re.split('\n\s\n\s\n', txt)
            for field in txt:
                field = field.replace('\n', '')
                field.strip()
                #print(field)
                match = re.search('(\d+)\.\s*Answer any.+(\d)+×(\d+)\s*=\s*(\d+)\s*\(', field)
                if match:
                    q = match.group(1)
                    marks = float(match.group(2))
                    ans = int(match.group(3))
                    total = int(match.group(4))
                    if marks * ans != total:
                        marks = total / ans
                    #print("question", q, "has", total, "marks:", marks, "*", ans)
                else:
                    match = re.search('\s*(\d+)\.', field)
                    if match:
                        q = match.group(1)

                cbcs = re.search('^\s*(CBCS.*)\s*\d+\.', field)
                if cbcs:
                    field = re.sub(cbcs.group(1), '', field)
                if (q != 0 and 'CBCS' not in field):
                    #print(field)
                    items = re.findall('\(([a-h])\)', field)
                    question = questions.get(q, {})
                    parts = question.get('parts', []) 
                    for item in items:
                        parts.append(item)
                    question['parts'] = parts
                    questions[q] = question
                    partmarks = questions[q].get('partmarks', [])
                    for match in re.finditer('\s*(\S+)\s*(\(|$)', field):
                        if bool(re.search('\d(?!\.)', match.group(1))):
                            if not match.group(1).isnumeric() or int(match.group(1)) <= 10:
                                partmarks.append(match.group(1))
                    questions[q]['partmarks'] = partmarks

        print(questions)

#--- main ---
parseQuestion(sys.argv[1])
