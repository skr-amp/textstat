from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
import string
import pymorphy2
import xlwt
import sys
import argparse

	
argparser = argparse.ArgumentParser(description = "Программа для частотного анализа текста", epilog = "Программа написанна в качестве вступительного испытания для поступления на курс 'Анализ данных средствами Python'  2019")
argparser.add_argument ('-i', '--input', type=argparse.FileType(mode='r', encoding="UTF8"), required=True, help="Имя входного файла. Обязательный параметр.", metavar = "ИМЯ ФАЙЛА")
argparser.add_argument ('-o', '--output', default='output.xls', help="Имя выходного файла. По умолчанию output.xls", metavar = "ИМЯ ФАЙЛА")
args = argparser.parse_args()

text = args.input.read()
text = text.translate(str.maketrans(' ', ' ', string.punctuation))
words = word_tokenize(text)
stop_words = stopwords.words('russian')
words = [i for i in words if ( i not in stop_words )]

morph = pymorphy2.MorphAnalyzer()

stat = []

for i in range(len(words)):
	word_in_stat = False
	p = morph.parse(words[i])
	
	for j in range(len(stat)):
		if p[0].normal_form == stat[j][0]:
			stat[j][1] += 1
			word_in_stat = True
	
	if not word_in_stat:
		stat.append([])
		stat[len(stat)-1].append(p[0].normal_form)
		stat[len(stat)-1].append(1)

stat.sort(key = lambda row: row[1], reverse=True)

book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet("Частота слов")
sheet1.write(0, 0, "Слово в нормальной форме")
sheet1.write(0, 1, "Количество слов в тексте")
for i in range(len(stat)):
	sheet1.write(i+1, 0, stat[i][0])
	sheet1.write(i+1, 1, stat[i][1])
	
book.save(args.output)
print("Результаты частотного анализа сохранены в файле", args.output)

