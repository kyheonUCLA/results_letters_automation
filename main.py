import os, sys
from unicodedata import name
from docxtpl import DocxTemplate, RichText
from parser import DataParser
from datetime import date

# os.chdir(sys.path[0])

test_times = ['8:30 AM PST', '8:30 AM PST', '8:30 AM PST', '8:30 AM PST', '8:30 AM PST', '8:30 AM PST', '8:30 AM PST']


def main():
    data = DataParser('templates/job_sheet.xlsx')
    for index, row in data.data.iterrows():
        if not(row['TESTED']):
            continue

        print(index)

        letter_date = date.today()
        time = '8:30 AM PST'
        doc = DocxTemplate('templates/letter_template.docx')
        context = create_context(row, letter_date, time)
        doc.render(context)
        document_title = context['NAME'] + 'results letter.pdf'
        doc.save('results_letters/' + document_title)
        

def create_context(row, date, time):
    name = row['FIRST NAME'].strip() + " " + row['LAST NAME'].strip()
    dob = row['DOB']
    test_result = RichText()
    if row['NEG']:
        test_result.add(text='NEGATIVE', bold=True, underline=True)
    elif row['POS']:
        test_result.add(text='POSITIVE', bold=True, underline=True)
    else:
        test_result.add(text='INCONCLUSIVE', bold=True, underline=True)

    print(test_result)

    context =  { 'NAME': name.title(), 'DOB': dob.strftime('%m/%d/%Y'), 
    'TIME': time, 'DATE': date.strftime('%m/%d/%Y'), 
    'DATEPRETTY': date.strftime("%B %d, %Y"), 'RESULT': test_result }

    return context


main()
# python-docx, docxtpl

