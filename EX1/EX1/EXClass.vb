Imports Excel = Microsoft.Office.Interop.Excel
Public Class EXClass
    Function exTranAs(ByVal inputFilePath As String, ByVal inPWD As String, ByVal outputFilePath As String, ByVal outPWD As String, ByRef errDes As String) As Integer
        On Error GoTo Error_Handler
        Dim extFileName, fileSaveName As String
        Dim wsCount As Integer
        Dim firstActSheetNameStr, tempStr As String

        'Excel define
        Dim EXAPP As New Excel.Application
        Dim wb As Excel.Workbook
        Dim wSheet As Excel.Worksheet
        Dim wSheet2 As Excel.Worksheet
        Dim range As Excel.Range
        Dim cell As Excel.Range

        On Error Resume Next

        ' Check File Path
        If outputFilePath = inputFilePath Then
            exTranAs = 4
            Exit Function
        End If
        extFileName = GetFileExtFromPath(outputFilePath)
        If extFileName = "" Then
            exTranAs = 3
            Exit Function
        End If

        extFileName = GetFileExtFromPath(inputFilePath)
        If extFileName = "" Then
            exTranAs = 2
            Exit Function
        End If
        fileSaveName = inputFilePath


        'exl = CreateObject("Excel.Application")
        ' open file

        'wb = EXAPP.Workbooks.Open(inputFilePath)
        wb = EXAPP.Workbooks.Open(Filename:=inputFilePath, ReadOnly:=True, Password:=inPWD)

        If Err.Number <> 0 Then
            'Could not open Excel
            'Err.Clear()
            'On Error GoTo Error_Handler
            exTranAs = 5
            GoTo Error_Handler
        Else
            On Error GoTo Error_Handler
        End If

        ' Save to other path
        EXAPP.DisplayAlerts = False
        wb.SaveAs(Filename:=outputFilePath, Password:=outPWD, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges)

        If Err.Number <> 0 Then
            'Could not save Excel
            'Err.Clear()
            'On Error GoTo Error_Handler
            exTranAs = 6
            GoTo Error_Handler
        Else
            On Error GoTo Error_Handler
        End If

        wb.Close()
        'If (wb.AutoUpdateSaveChanges.Equals(wb)) Then

        'exTran = 3

        'Exit Function
        'End If
        wb = EXAPP.Workbooks.Open(Filename:=outputFilePath, ReadOnly:=False, Password:=outPWD)

        If Err.Number <> 0 Then
            'Could not open Excel
            'Err.Clear()
            'On Error GoTo Error_Handler
            exTranAs = 7
            GoTo Error_Handler
        Else
            On Error GoTo Error_Handler
        End If

        On Error GoTo 0

        ' get sheet count
        wsCount = wb.Sheets.Count

        wSheet2 = wb.ActiveSheet
        firstActSheetNameStr = wb.ActiveSheet.Name   'get activate sheet name
        'range = wSheet2.Cells(1, 1)
        'MsgBox(range.Value)

        For Each wSheet In wb.Worksheets
            wSheet.Activate()  ' set activate

            range = wSheet.UsedRange
            For Each cell In range
                tempStr = cell.Value           ' get value
                If (Right$(tempStr, 1) <> " " And tempStr <> "") Then   ' check last char not space and not null
                    cell.NumberFormatLocal = "@"   ' set cell format is text
                    cell.Value = tempStr & " "     ' add space in last char
                End If
            Next cell
        Next wSheet

        If firstActSheetNameStr <> "" Then
            wb.Activate()
            wSheet2.Activate()
            'EXAPP.Worksheets(firstActSheetNameStr).Activate()
        End If

        wb.Save()

        If Err.Number <> 0 Then
            'Could not save Excel
            'Err.Clear()
            'On Error GoTo Error_Handler
            exTranAs = 8
            GoTo Error_Handler
        Else
            On Error GoTo Error_Handler
        End If

        wb.Close()
        EXAPP.Quit()

        Exit Function
Error_Handler_Exit:
        wb.Close()
        Err.Clear()
        On Error Resume Next
        Exit Function

Error_Handler:
        errDes = Err.Description
        'MsgBox("The following error has occured" & vbCrLf & vbCrLf & _
        '       "Error Number: " & Err.Number & vbCrLf & _
        '       "Error Source: OpenPwdXLS" & vbCrLf & _
        '       "Error Description: " & Err.Description)
        Resume Error_Handler_Exit
    End Function
    Private Function GetFileExtFromPath(ByVal strPath As String) As String
        ' Returns the rightmost characters of a string upto but not including the rightmost '\'
        ' e.g. 'c:\winnt\win.ini' returns 'ini'
        GetFileExtFromPath = ""
        If Right$(strPath, 1) <> "." And Len(strPath) > 0 Then
            GetFileExtFromPath = GetFileExtFromPath(Left$(strPath, Len(strPath) - 1)) + Right$(strPath, 1)
        End If
    End Function
End Class
