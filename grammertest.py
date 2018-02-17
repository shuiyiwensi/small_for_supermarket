from win32com.client import Dispatch

word = Dispatch('Word.Application')
word.Visible = False
word = word.Documents.Open("demo.docx")

#get number of sheets
word.Repaginate()
num_of_sheets = word.ComputeStatistics(2)


