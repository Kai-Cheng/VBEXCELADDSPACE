Imports EX1
Module Module1

    Sub Main()
        Dim class1 As New EXClass
        Dim errStr As String
        Dim Result As Integer

        errStr = ""
        ' Change the Input/Output Path of Excel File, and those password
        Result = class1.exTranAs("D:\Book4.xls", "123abc", "D:\Book4_as.xls", "", errStr)

        If Result <> 0 Then
            If errStr <> "" Then
                Debug.Print(errStr)
            Else
                Debug.Print("ERROR!")
            End If
        Else
            Debug.Print("OK!")
        End If

    End Sub

End Module
