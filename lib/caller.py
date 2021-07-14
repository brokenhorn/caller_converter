import shutil
import pandas as pd
import win32com.client as win32
from  docx2csv import extract, extract_tables

#Путь подается в ковычках  'path'
#создает файлы xls, csv  по одному на таблицу
#pip install pywin32
#pip install pandas

class XlsResultSaver:

	def ft_docx_to_xls(self,result, new_directory, path):
		len_t = len(extract_tables(path))
		i = len_t
		extract(path, format="xls", singlefile=False)
		while i > 0:
			path_cp = path[:-5]
			path_cp = path_cp + '_' + str(i) + '.xls'
			new_path = path_cp
			try:
				new_path = shutil.move(path_cp, new_directory)
				print(new_path)
				xls_to_csv(new_path)
			except:
				print()
			i = i - 1
		return new_path
# def ft_docx_to_csv_xls_converteer(path_to_docx):
# 		#extract(filename=path_to_docx, format="csv",singlefile=False)
# 		#extract(filename=path_to_docx, format="xls",singlefile=True)
# 		#ft_docx_to_csv(path_to_docx)

def ft_doc_to_docx(path):
		word = win32.Dispatch("Word.Application")
		wdFormatDocumentDefault = 16
		wdHeaderFooterPrimary = 1
		doc = word.Documents.Open(path)
		doc.SaveAs(path[:-3] + 'docx', FileFormat=wdFormatDocumentDefault)
		doc.Close()
		word.Quit()

def xls_to_csv(path):
	read_file = pd.read_excel(path)
	path_cp = path[:-4]
	path_cp = path_cp + '.csv'
	read_file.to_csv(path_cp, index=None, header=True)
	print(path_cp)

def ft_rtf_to_docx(path):
	word = win32.Dispatch("Word.Application")
	wdFormatDocumentDefault = 16
	wdHeaderFooterPrimary = 1
	doc = word.Documents.Open(path)
	doc.SaveAs(path[:-3] + 'docx' , FileFormat=wdFormatDocumentDefault)
	doc.Close()
	word.Quit()

