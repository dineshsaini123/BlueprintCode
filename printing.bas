Attribute VB_Name = "Printing"
Public Function printreport(SQL As String, header As String) As Boolean
    Dim rs1 As ADODB.Recordset
    Set rs1 = New ADODB.Recordset
    printreport = True
    rs1.Open SQL, con, adOpenKeyset, adLockReadOnly, adCmdText
    If Not rs1.BOF Then
        rs1.MoveLast
        Y = rs1.RecordCount
        rs1.MoveFirst
    End If
    '////////////////*********************
        
        
    Dim Line As Integer
    Line = 0
    If Not rs1.BOF Then
        Open "" + App.Path + "\vipin.txt" For Output As #1
        If Not rs1.EOF Then
            Print #1, Space(2); Chr(27) + Chr(15); lsets("CATEGORY", rs1(0).DefinedSize); lsets("GEN. LEDGER NAME", rs1(1).DefinedSize - 5); lsets("PLC", 6); lsets("BSC", 6); lsets("SLF", 6); lsets("YEAROPENING", 13)
            Print #1, "----------------------------------------------------------------------------"
            Line = 2
        End If
            Waitwindow.pb1.Min = 0
            Waitwindow.pb1.value = 0
            Waitwindow.pb1.Max = Y
        
            Do While Not rs1.EOF
                Print #1, Space(2); lsets(rs1(0), rs1(0).DefinedSize); lsets(rs1(1), rs1(1).DefinedSize); bsets(rs1(2)); bsets(rs1(3)); bsets(rs1(4)); rsets(rs1(5), rs1(5).DefinedSize)
                Line = Line + 1
                Waitwindow.pb1.value = Waitwindow.pb1.value + 1
                If Not rs1.EOF Then
                    rs1.MoveNext
                End If
            Loop
            Waitwindow.Hide
            If Line < 80 Then
                Do While Not Line = 80
                    Print #1, " "
                    Line = Line + 1
                Loop
            End If
            Close #1
            GRIDpreview.SQL (SQL)
        Else
            printreport = False
        End If
End Function
Public Function lsets(ST As String, length As Integer) As String
    Dim kk As String
            kk = Trim(ST)
            If Len(kk) < length Then
                Do While Not Len(kk) = length
                    kk = kk + " "
                Loop
            End If
            If Len(kk) > length Then
                Do While Not Len(kk) = length
                    kk = Mid$(kk, 0, Len(kk) - 1)
                Loop
            End If
        lsets = kk
End Function
Public Function bsets(ST As Boolean) As String
    Dim kk As String
        If ST = True Then
            kk = "True  "
        Else
            kk = "False "
        End If
            
        bsets = kk
End Function
Public Function rsets(ST As String, length As Integer) As String
Dim kk As String
kk = Trim(ST)
If Len(kk) < length Then
Do While Not Len(kk) = length
kk = " " + kk
Loop
End If
If Len(kk) > length Then
Do While Not Len(kk) = length
kk = Mid$(kk, 0, Len(kk) - 1)
Loop
End If
rsets = kk
End Function
Function repli(char As String, t As Integer) As String
Dim I As Integer
repli = ""
If t = 143 Then
repli = "-----------------------------------------------------------------------------------------------------------------------------------------------"
Exit Function
ElseIf t = 136 Then
repli = "----------------------------------------------------------------------------------------------------------------------------------------"
Exit Function
ElseIf t = 96 Then
repli = "------------------------------------------------------------------------------------------------"
Exit Function
End If


For I = 1 To t
repli = Trim(repli) + Trim(char)
Next
End Function
