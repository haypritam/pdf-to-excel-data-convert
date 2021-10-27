import PyPDF2
import sys, re, os, io
import excel_writer

def check_parts(questions):
    for q in questions:
        parts = questions[q]['parts']
        partmarks = questions[q]['partmarks']
        if len(parts) != len(partmarks):
            partmarks = "".join(partmarks)
            if partmarks.count('+') + 1 == len(parts):
                partmarks = partmarks.split('+')
                questions[q]['partmarks'] = partmarks

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
                match = re.search('(\d+)\.\s*Answer any.+(\d+×\d+\s*=\s*\d+)\s*\(', field)
                if match:
                    q = match.group(1)
                    pms = match.group(2)
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
                    if (len(pms) == 0):
                        for match in re.finditer('\s*(\S+)\s*(\(|$)', field):
                            if bool(re.search('\d(?!\.)', match.group(1))):
                                if not match.group(1).isnumeric() or int(match.group(1)) <= 10:
                                    partmarks.append(match.group(1))
                    else:
                        partmarks.append(pms)
                        pms = []
                    questions[q]['partmarks'] = partmarks

        check_parts(questions)
        return questions

#--- main ---
questions = parseQuestion(sys.argv[1])
excel_writer.write_xl(sys.argv[2], questions)
