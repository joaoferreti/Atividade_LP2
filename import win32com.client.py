import win32com.client

e = win32com.client.Distpatch("Excel.Application") # Passando o Excel para à variável.
e.Visible = 1 # Parâmetros
e.Workbooks.Add() # Propriedade do Excel (workbook)
e.Cells(1, 1).Value = "Hello World"