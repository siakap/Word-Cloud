#packcakge need: install PyPDF2, pdf2docx, docx, wordcloud

import os
#to merge
from PyPDF2 import PdfFileMerger
#to convert to docx
from pdf2docx import Converter
#to read docx
import docx
#stuffs for wordcloud
import matplotlib.pyplot as plt
from wordcloud import WordCloud, STOPWORDS

cwd = os.getcwd()
print("Current working directory:", cwd)
#place merger in same loc as pdf folder

merger = PdfFileMerger(strict=False)

for item in os.listdir(cwd):
	if item.endswith('pdf'):
		merger.append(item)

merger.write('Combined.pdf')
merger.close()

#merged pdf created


word_file = 'combined.docx'

cv = Converter('Combined.pdf')
cv.convert(word_file, start=0, end=None)
cv.close()
#docx version of merged pdf created

def getText(filename):
    doc = docx.Document(filename)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
    return '\n'.join(fullText)

text = getText(word_file)

stopwords = STOPWORDS
#need to check quality of stopwords
#something suitable for investment/market talk

wc = WordCloud(
	background_color = 'white',
	stopwords= stopwords,
	height = 600,
	width = 400
	)

wc.generate(text)	
wc.to_file('Wordcloud.png')

print("Done")	
	
	
	
