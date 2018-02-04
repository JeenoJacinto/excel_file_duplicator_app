Dim args, objExcel

Set args = WScript.Arguments
Set objExcel = CreateObject("Excel.Application")

objExcel.Workbooks.Open args(0)
objExcel.Visible = False


objExcel.Run "delete_row"

objExcel.ActiveWorkbook.Save
objExcel.ActiveWorkbook.Close(0)
objExcel.Quit
