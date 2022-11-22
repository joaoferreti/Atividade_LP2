import win32com.client

e = win32com.client.Dispatch("Excel.Application") # Passando o Excel para a variável.
e.Visible = 1 # Parâmetros
e.Workbooks.Add() # Propriedade do Excel (workbook)
e.Cells(1, 1).Value = "Hello World"
