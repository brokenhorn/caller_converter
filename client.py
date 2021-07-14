import os
import sys
from lib.caller import *
from lib.FileManager import *


file_path = sys.argv[1]
save_directory = sys.argv[2]
file_manager = FileManager(file_path)

result = file_manager.manage()
if file_path.endswith(".rtf"):
    ft_rtf_to_docx(file_path)
    file_path = file_path[:-3] + 'docx'
    saved_path_xls = XlsResultSaver().ft_docx_to_xls(result,save_directory,file_path)
elif file_path.endswith('.doc'):
    ft_doc_to_docx(file_path)
    file_path = file_path[:-3] + 'docx'
    saved_path_xls = XlsResultSaver().ft_docx_to_xls(result, save_directory, file_path)
elif file_path.endswith('.docx'):
    saved_path_xls = XlsResultSaver().ft_docx_to_xls(result, save_directory, file_path)


#

