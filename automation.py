import os, sys  
from docxtpl import DocxTemplate  
import pandas as pd  


#READING VALUES FROM THE EXCEL FILE USING PANDAS LIBRARY

df = pd.read_excel('data.xlsx')

#STORING THE KEYS AND VALUES IN FORM OF PYTHON LISTS

keys=df['keys'].values
values = df['values'].values

#CONVERTING THE ABOVE TWO LISTS INTO PYTHON DICTIONARY TO RENDER IT IN THE WORD FILE

context = {keys[i]: values[i] for i in range(len(keys))}

#CREATING INSTANCES OF THE GIVEN 5 WORD DOCUMENTS 

doc1 = DocxTemplate("temp_file_1.docx")
doc2 = DocxTemplate("temp_file_2.docx")
doc3 = DocxTemplate("temp_file_3.docx")
doc4 = DocxTemplate("temp_file_4.docx")
doc5 = DocxTemplate("temp_file_5.docx")

#RENDERING THE CONTENT INTO THE WORD FILES

doc1.render(context)
doc2.render(context)
doc3.render(context)
doc4.render(context)
doc5.render(context)

#SAVING THE DESIRED FINAL OUTPUT INTO THE NESTED FOLDERS GIVEN IN THE QUESTION

doc1.save('Master Folder/Master File.docx')
doc2.save('Master Folder/Annex A/Sub A1.docx')
doc3.save('Master Folder/Annex B/Sub B2.docx')
doc4.save('Master Folder/Annex C/Sub C1.docx')
doc5.save('Master Folder/Annex C/Sub C2.docx')