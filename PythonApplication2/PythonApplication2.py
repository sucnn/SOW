from docx import *
document = Document('abc.docx')

#read paragraph 
for par in document.paragraphs: 
    pa=par.text
    print(pa)

#read table 
for table in document.tables:
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                print(paragraph.text)


#Author 
print(document.core_properties.author)

#Title
print(document.core_properties.title)

#Category
print(document.core_properties.category)

#Language
print(document.core_properties.language)

#Subject
print(document.core_properties.subject)

"""
Code to extract data in xml format
body_element = document._body._body
print(body_element.xml) 

"""