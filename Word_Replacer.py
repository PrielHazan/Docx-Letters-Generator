import docx
import os
import sys
import ntpath
import pathlib
from time import sleep
import comtypes.client

# attention: if bold or any style it wont replace, or not?
# cant use space because its seperate args
# commands:
# pyinstaller --additional-hooks-dir=hooks -F -i "Wicon.ico" Word_Replacer.py
# python docx_to_pdf.py C:\Users\priel\Downloads\old_projects\Word_replacer\Hello.docx C:\Users\priel\OneDrive\Desktop Hello HelloReplace This ThisReplace
# python Word_Replacer.py C:\Users\priel\Downloads\old_projects\Word_replacer\file.docx C:\Users\priel\Downloads\old_projects\Word_replacer name1 פריאל_חזן כסף 1000_שקלים_חדשים כתובת הברוש_107_בני_ראם
# C:\Users\priel\Downloads\old_projects\Word_replacer\Word_Replacer.exe C:\Users\priel\Downloads\old_projects\Word_replacer\file.docx C:\Users\priel\Downloads\old_projects\Word_replacer name1 פריאל_חזן כסף 1000_שקלים_חדשים כתובת הברוש_107_בני_ראם


# CMD Input:
if len(sys.argv)%2 != 1:
    print('The software was expecting an even number of arguments')
    sys.exit()
file_path = sys.argv[1]
output_dir = sys.argv[2]
old_words = []
new_words = []
for idx, arg in enumerate(sys.argv):
    if idx <= 2:
        continue
    if idx%2 == 0:
        new_words.append(arg)
    else:
        old_words.append(arg)
# for idx, val in enumerate(sys.argv):
    # print(f'index: {idx}, val: {val}')
# print(old_words)
# print(new_words)

# Debug input
# file_path = r'C:\Users\priel\OneDrive\Desktop\Hello.docx'
# output_dir = r'C:\Users\priel\OneDrive\Desktop'
# old_words = ['Hello', 'This']
# new_words = ['HelloReplace', 'thisReplace']



# check for valid input:
if not(os.path.isfile(file_path)):
    print('First argument should be a file path!')
    sys.exit()

if not(os.path.isdir(output_dir)):
    print('Second argument should be a directory path!')
    sys.exit()

def _toSpaces(names):
    for idx, name in enumerate(names):
        names[idx] = name.replace("_", " ")

_toSpaces(old_words)
_toSpaces(new_words)

filename = ntpath.basename(file_path)
name, ext = filename.split('.')

new_filename = name + '_.' + ext
pdf_name = name + '.pdf'

save_as = os.path.join(output_dir, new_filename)

doc = docx.Document(file_path)
for p in doc.paragraphs:
    for idx, old_text in enumerate(old_words):
        if old_text in p.text:
            inline = p.runs
            # Loop added to work with runs (strings with same style)
            for i in range(len(inline)):
                # print(i)
                # print(idx)
                if old_text in inline[i].text:
                    text = inline[i].text.replace(old_text, new_words[idx])
                    inline[i].text = text
            # print(p.text)

doc.save(save_as)
sleep(2)
r_save_as = rf'{save_as}'
temp = os.path.join(output_dir, pdf_name)
pdf_saved = rf'{temp}'

wdFormatPDF = 17

in_file = os.path.abspath(r_save_as)
out_file = os.path.abspath(pdf_saved)

word = comtypes.client.CreateObject('Word.Application')
doc = word.Documents.Open(in_file)
doc.SaveAs(out_file, FileFormat=wdFormatPDF)
doc.Close()
word.Quit()


os.remove(save_as)
