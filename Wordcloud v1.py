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
end_name = input("Pleasee name your output file:")
user_color = input("Background color for the wordcloud:")
combined_pdf = 'combined.pdf'
output_png = end_name + '.png'

merger = PdfFileMerger(strict=False)

for item in os.listdir(cwd):
	if item.endswith(combined_pdf):
		os.remove(combined_pdf)

for item in os.listdir(cwd):
	if item.endswith('pdf'):
		merger.append(item)

merger.write(combined_pdf)
merger.close()

#merged pdf created


word_file = 'combined.docx'

cv = Converter(combined_pdf)
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
stopwordori = STOPWORDS

with open('stoplist.txt') as f:
	new_words = f.read().splitlines()
#add more words to stop word list

stopwords = new_words + list(STOPWORDS)
#need to check quality of stopwords
#something suitable for investment/market talk

wc = WordCloud(
	background_color = user_color,
	stopwords= stopwords,
	height = 600,
	width = 400
	)

wc.generate(text)
wc.to_file(output_png)
os.remove(word_file)

print("Done")
