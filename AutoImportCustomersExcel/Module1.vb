'Imports Microsoft.Office.Interop

Module Module1

    'Sub Main()
    '    'Dim table As New DataTable("CurrencyRate")

    '    Dim xlApp As Excel.Application
    '    Dim xlWorkbook As Excel.Workbook
    '    Dim xlWorkSheet As Excel.Worksheet
    '    Dim xlRange As Excel.Range

    '    'Dim xlCol As Integer

    '    Dim xlRow As Integer


    '    xlApp = New Excel.Application
    '    xlWorkbook = xlApp.Workbooks.Open("D:\Marky\knowlsemployee.xls")
    '    xlWorkSheet = xlWorkbook.ActiveSheet()
    '    xlRange = xlWorkSheet.UsedRange

    '    If xlRange.Columns.Count > 0 Then
    '        If xlRange.Rows.Count > 0 Then
    '            For xlRow = 2 To xlRange.Rows.Count 'here the xlRow is start from 2 coz in exvel sheet mostly 1st row is the header row
    '                MsgBox(xlRange.Cells(xlRow, 1).Text & " | " & xlRange.Cells(xlRow, 2).Text & " | " & _
    '                       xlRange.Cells(xlRow, 3).Text & " | " & xlRange.Cells(xlRow, 4).Text & " | " & _
    '                       xlRange.Cells(xlRow, 5).Text)
    '                'For xlCol = 1 To xlRange.Columns.Count
    '                '    Data(xlCol - 1) = xlRange.Cells(xlRow, xlCol).text
    '                'Next
    '                '.LoadDataRow(Data, True)
    '            Next
    '            xlWorkbook.Close()
    '            xlApp.Quit()
    '            'KillExcelProcess()
    '        End If
    '    End If

    'End Sub

End Module
