VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form viewgenledger 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   1410
   ClientWidth     =   13245
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   13245
   ShowInTaskbar   =   0   'False
   Begin RichTextLib.RichTextBox r1 
      Height          =   7320
      Left            =   30
      TabIndex        =   4
      Top             =   90
      Width           =   13185
      _ExtentX        =   23257
      _ExtentY        =   12912
      _Version        =   393217
      BackColor       =   12632319
      ScrollBars      =   3
      RightMargin     =   20000
      TextRTF         =   $"viewGLed.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton printCOMMAND 
      Caption         =   "Print"
      Height          =   405
      Left            =   4785
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7560
      Width           =   705
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5805
      TabIndex        =   2
      Text            =   "50 %"
      Top             =   7560
      Width           =   1095
   End
   Begin VB.CommandButton export 
      Caption         =   "Export"
      Height          =   405
      Left            =   4065
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7560
      Width           =   705
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   435
      Left            =   7005
      TabIndex        =   0
      Top             =   7560
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog C1 
      Left            =   3420
      Top             =   8100
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FromPage        =   1
      PrinterDefault  =   0   'False
      ToPage          =   1
   End
End
Attribute VB_Name = "viewgenledger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQLSTRING As String
'Dim CON As ADODB.Connection
'Dim CON As ADODB.Connection
Dim RS As ADODB.Recordset
Private Sub Combo1_Change()
If Trim(Combo1.Text) = "50 %" Then
    r1.Font.Size = 5
End If
If Trim(Combo1.Text) = "75 %" Then
    r1.Font.Size = 8
End If
If Trim(Combo1.Text) = "100 %" Then
    r1.Font.Size = 10
End If
If Trim(Combo1.Text) = "125 %" Then
    r1.Font.Size = 12
End If
If Trim(Combo1.Text) = "150 %" Then
    r1.Font.Size = 14
End If
If Trim(Combo1.Text) = "200 %" Then
    r1.Font.Size = 18
End If
End Sub

Private Sub Combo1_Click()
'r1.row = 1
If Trim(Combo1.Text) = "50 %" Then
    r1.Font.Size = 5
End If
If Trim(Combo1.Text) = "75 %" Then
    r1.Font.Size = 8
End If
If Trim(Combo1.Text) = "100 %" Then
    r1.Font.Size = 10
End If
If Trim(Combo1.Text) = "200 %" Then
    r1.Font.Size = 18
End If
If Trim(Combo1.Text) = "125 %" Then
    r1.Font.Size = 12
End If
If Trim(Combo1.Text) = "150 %" Then
    r1.Font.Size = 14
End If


End Sub

Private Sub Command1_Click()
    Unload Me
    
End Sub
Private Sub export_Click()
    d1.ShowPrinter
   ' MsgBox "copies =" + Str(d1.copies)
    'D1.copies
    'Printer.PaperSize
End Sub
Public Function printnow()
    Dim X As Long
    Dim p As Printer
    For I = 0 To Printers.Count - 1
        Set p = Printers(I)
        If Trim(p.DeviceName) = Trim(Printdlg.Comboprinters.Text) Then
            Exit For
        End If
    Next
    For I = 1 To (Printdlg.UpDown1.value)
        X = Shell("" + App.Path + "\ppp.bat " + App.Path + "\vipin.txt " & Trim(p.Port))
    Next
    Printdlg.UpDown1.value = 1
    Printdlg.Text1.Text = "1"
End Function
Private Sub Form_Load()
    
    
    Me.Width = MainMenu.Width - 3000
    Me.Height = (MainMenu.Height - 1500)
    
    
    Command1.Top = r1.Top + r1.Height + 30
    Combo1.Top = r1.Top + r1.Height + 30
    
    
    Set RS = New ADODB.Recordset
    genreport
    r1.FileName = "" + App.Path + "\vipin.txt"
    r1.LoadFile (r1.FileName)
    
    BackColorFrom Me

End Sub
Private Sub Form_Resize()


If Me.Width > 350 And Me.Height > 1500 Then
    r1.Width = Me.Width - 250
    r1.Height = Me.Height - 1000
    Command1.Top = r1.Top + r1.Height + 30
    Combo1.Top = r1.Top + r1.Height + 30
    export.Top = r1.Top + r1.Height + 30
    printCOMMAND.Top = r1.Top + r1.Height + 30

End If
End Sub

Private Sub return1_Click()
    Unload Me
End Sub


Private Sub print_Click()
    On Error GoTo errorcancel
    Dim frompage, topage, copies As Integer
    c1.Flags = 0
    c1.Max = 11
    c1.Flags = &H40000 Or &H4&
    c1.ShowPrinter
    frompage = c1.frompage
    topage = c1.topage
    copies = c1.copies
    If c1.Flags And &H20 Then
        c2.ShowSave
    End If
    X = Shell("" + App.Path + "\ppp.bat " + App.Path + "\vipin.txt " & Trim(Printer.Port))
errorcancel:
End Sub
Function genreport()
    
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        DoEvents

    
    'con.Execute "delete from Winrpt where uid=" & UId & ""
    con.Execute "delete from Winrpt"
    
    Dim called1, called2 As Boolean
    Dim MaxLine As Integer
    Dim T1, T2, T3, T4, T5, T6, T7, T8 As Integer
    Dim paperWidth As Integer
    Dim xtemp As String
    Dim trs As ADODB.Recordset
    Dim Pno As Integer
    Set trs = New ADODB.Recordset
    
    Dim cr, dr
    
    Dim cname, add1 As String
    
 
    dr = 0
    cr = 0
    
    paperWidth = 150
        T1 = 8  '10
        T2 = 16 '16
        T3 = 40
        T4 = 55
        T5 = 65  '70
        T6 = 80  '85
        T7 = 100
        T8 = 115
        
        called1 = False
        called2 = False
        Dim Line As Integer
        Dim rs1 As ADODB.Recordset
        Dim kkk As ADODB.Recordset
        Dim FooterYes As Integer
        
        
        Set kkk = New ADODB.Recordset
        Set rs1 = New ADODB.Recordset
        
        If kkk.State = 1 Then kkk.close
        kkk.Open "Select * from setup1 where " & stringyear & "", con, adOpenKeyset, adLockReadOnly
        If kkk.EOF = False Then
        cname = dspace(kkk!cname)
        add1 = dspace(kkk!add1)
        End If
        
        
        FooterYes = False
        main.reportdata
        main.repors.Find "reportname='" + Trim(main.reportname) + "'"
        MaxLine = main.repors!totalline
        If main.repors!comp = True Then
            paperWidth = Int(main.repors!totalcolumn * 1.75)
        Else
            paperWidth = main.repors!totalcolumn
        End If
        paperWidth = 136
        MaxLine = 72
        Open "" + App.Path + "\vipin.txt" For Output As #1
        Line = 0
        Pno = 1
header:
        
        Dim I As Integer
        For I = 1 To main.repors!TopMargin
            Print #1, ""
            Line = Line + 1
        Next
                            
       If FooterYes = True Then
            Print #1, Tab(LEFTM); repli("-", paperWidth)
            Line = Line + 1
            Do While Line < 72
                    Print #1, " "
                    Line = Line + 1
            Loop
            Line = 0
            FooterYes = False
       End If
       If kkk.State = 1 Then kkk.close
       CNSetup
       kkk.Open "Select * from setup1 where " & stringyear & "", con, adOpenKeyset, adLockReadOnly
       If Not kkk.BOF Then
            Print #1, ""
            Print #1, ""
            Print #1, Chr(27) + Chr(15) + Chr(14)
            Print #1, Tab(125); "Page No:  " & Pno
            Print #1, Tab(((paperWidth - (Len(Trim(kkk!cname)) * 2)) / 2) + LEFTM - 15); Chr(27) + Chr(15) + Chr(14); cname    'dspace(Trim(kkk!cname))
            Print #1, Tab(((paperWidth - (Len(Trim(kkk!add1)) * 2)) / 2) + LEFTM); Chr(27) + Chr(15); add1   'dspace(Trim(kkk!add1))
            Line = Line + 6
        End If
        If trs.State = 1 Then
            trs.close
        End If
        'trs.Open "treportGen", CON, adOpenKeyset, adLockReadOnly, adcmdtext
        trs.Open "select top 1 * from treportGen where " & stringyear & " and userid=" & main.UId & " order by sno", con, adOpenKeyset, adLockReadOnly, adCmdText
        
        xstr = trs!Period
        Print #1, Chr(27) + Chr(14); Tab((69 - Len(Trim(trs!header))) / 2); Trim(trs!header); Chr(27) + Chr(15)
        Print #1, Tab(LEFTM + ((paperWidth - Len(Trim("Period : " + Trim(xstr)))) / 2)); Trim("Period : " + Trim(xstr))
        Print #1, Tab(LEFTM); repli("-", paperWidth)
        Print #1, Tab(0); Chr(27) + Chr(71); "V.DATE"; Tab(T1 + 1 + LEFTM); "V.T."; Tab(T1 + 6 + LEFTM); "V.N."; Tab(T2 + 7 + LEFTM); "N A R R A T I O N"; Tab(T4 + 10 + LEFTM); "AMOUNT(Dr.)"; Tab(T5 + 15 + LEFTM); "AMOUNT(Cr.)"; Tab(T6 + 13 + LEFTM); "CHEQ./BILL NO. & DATE"; Tab(T6 + 20 + LEFTM + Len("CHEQ./BILL NO. & DATE")); "BALANCE"; Chr(27) + Chr(72)
        Print #1, Tab(LEFTM); repli("-", paperWidth)
        Line = Line + 5
        trs.close
        If called1 = True Then GoTo printagain1
        If rs1.State = 1 Then rs1.close
        
        
        'rs1.Open "select * from treportGen order by vdate,vtype,vno", CON, adOpenKeyset, adLockReadOnly, adCmdText
        rs1.Open "select * from treportGen order by vdate,vtype,vno", con, adOpenKeyset, adLockReadOnly, adCmdText
        If Not rs1.BOF Then
          
          '--------------------------------
           ' Balance = Balance + IIf(IsNull(rs1!aD), 0, rs1!aD) - IIf(IsNull(rs1!aC), 0, rs1!aC)
           ' con.Execute "update treportGen set Balance = " & Round(Balance, 2) & " where sno=" & rs1!sno
           '---------------------------------
        
            Print #1, ""
            Print #1, Tab(LEFTM); rs1!SUBLEDGER
            Print #1, Tab(LEFTM); "** Opening Balance as on " & GLEDGERPRINT.date1.Text; Tab(T4 + 9 + LEFTM); IIf(Val(rs1!aD & "") <> 0, rsets(Trim(Format(Str(Val(rs1!aD & "")), "0.00")), 12), ""); Tab(T5 + 13 + LEFTM); IIf(Val(rs1!aC & "") <> 0, rsets(Trim(Format(Str(Val(rs1!aC & "")), "0.00")), 12), "")
            Print #1, ""
            Line = Line + 4
            con.Execute "insert into winrpt(Party,Narration,Receipt,Payment,FromDate,ToDate,Description,uid) values('" & rs1.Fields("subledger").value & "', '" & "Opening Balance as on " & GLEDGERPRINT.date1.Text & "' ," & IIf(IsNull(rs1!aD), 0, rs1!aD) & "," & IIf(IsNull(rs1!aC), 0, rs1!aC) & ",'" & Format(GLEDGERPRINT.date1.Text, "MM/dd/yyyy") & "','" & Format(GLEDGERPRINT.date2.Text, "MM/dd/yyyy") & "','" & "GEN. LEDGER ACCOUNT" & "'," & UId & ")"
            
            rs1.MoveNext
        End If
        If Not rs1.BOF Then
            Do While Not rs1.EOF
                Print #1, Tab(LEFTM); Trim(rs1!vdate); Tab(12 + LEFTM); (Trim(rs1!vtype)); Tab(15 + LEFTM); (Trim(rs1!vno)); Tab(T2 + 8 + LEFTM);
                If IsNull(Trim(rs1!narration)) Then
                    xtemp = " "
                Else
                    xtemp = Mid(Trim(rs1!narration), 1, 39)
                End If
                Print #1, Trim(xtemp);
                If rs1!aD > 0 Then
                   Print #1, Tab(T4 + 9 + LEFTM); rsets(Trim(Format(Str(rs1!aD), "0.00")), 12);
                   dr = dr + rs1!aD
                End If
                If rs1!aC > 0 Then
                   Print #1, Tab(T5 + 13 + LEFTM); rsets(Trim(Format(Str(rs1!aC), "0.00")), 12);  'Tab(T7 + 2 + LEFTM);
                   cr = cr + rs1!aC
                End If
                
                If Trim(rs1!cbno) <> "" Then
                   xtemp = Trim(rs1!cbno)
                   Print #1, Tab(T6 + 15 + LEFTM); Trim(xtemp);
                Else
                   xtemp = " "
                End If
                
                
                
                'Print #1, Tab(T6 + 15 + LEFTM); lsets(cnullstr(xtemp), 18);
                
                Print #1, Tab(T7 + 17 + LEFTM); rsets(Trim(Format(Str(rs1!Balance), "0.00")), 14)
                Line = Line + 1
                
                con.Execute "insert into winrpt(date1,op,Narration,Receipt,Payment,FromDate,ToDate,Description,cbno,uid) values('" & Format(rs1!vdate, "MM/dd/yyyy") & "'," & rs1!Balance & ",'" & rs1!vtype & " " & rs1!vno & " " & rs1!narration & "'," & IIf(IsNull(rs1!aD), 0, rs1!aD) & "," & IIf(IsNull(rs1!aC), 0, rs1!aC) & ",'" & Format(GLEDGERPRINT.date1.Text, "MM/dd/yyyy") & "','" & Format(GLEDGERPRINT.date2.Text, "MM/dd/yyyy") & "','" & "GEN. LEDGER ACCOUNT" & "','" & IIf(IsNull(rs1!cbno), "", rs1!cbno) & "'," & UId & ")"
                
                If Line > MaxLine - 8 Then
                   called1 = True
printnext:
                   Pno = Pno + 1
                   FooterYes = True
                   GoTo header
                   
                   
printagain1:
                    
                    called1 = False
                End If
                If Not rs1.EOF Then
                    rs1.MoveNext
                End If
            Loop
printfooter:
              Print #1, Tab(LEFTM); repli("-", paperWidth)
              Line = Line + 1
              
              
              
                If dr > 0 Then
                   Print #1, Tab(T4 + 9 + LEFTM); rsets(Trim(Format(Str(dr), "0.00")), 12);
                End If
                If cr > 0 Then
                   Print #1, Tab(T5 + 13 + LEFTM); rsets(Trim(Format(Str(cr), "0.00")), 12);  'Tab(T7 + 2 + LEFTM);
                End If

              Print #1, Tab(LEFTM); repli("-", paperWidth)
              Line = Line + 1

              
              
              Do While Line <= 72
                     Print #1, " "
                     Line = Line + 1
              Loop
        End If
        Close #1
End Function



Function genreport1()
    Dim called1, called2 As Boolean
    Dim MaxLine As Integer
    Dim T1, T2, T3, T4, T5, T6, T7, T8 As Integer
    Dim paperWidth As Integer
    Dim xtemp As String
    Dim trs As ADODB.Recordset
    Dim Pno As Integer
    Set trs = New ADODB.Recordset
       paperWidth = 150
        T1 = 10
        T2 = 25
        T3 = 40
        T4 = 55
        T5 = 70
        T6 = 85
        T7 = 100
        T8 = 115
        MaxLine = 72
        called1 = False
        called2 = False
        Dim Line As Integer
        Dim Balance As Double
        Dim FooterYes As Boolean
        Dim rs1 As ADODB.Recordset
        Dim kkk As ADODB.Recordset
        Set kkk = New ADODB.Recordset
        Set rs1 = New ADODB.Recordset
        
        main.reportdata
        main.repors.Find "reportname='" + Trim(main.reportname) + "'"
        MaxLine = main.repors!totalline
        If main.repors!comp = True Then
            paperWidth = Int(main.repors!totalcolumn * 1.75)
        Else
            paperWidth = main.repors!totalcolumn
        End If
        Open "" + App.Path + "\vipin.txt" For Output As #1
        Line = 0
        Pno = 1
        FooterYes = False
header:
        Dim I As Integer
        For I = 1 To main.repors!TopMargin
            Print #1, ""
            Line = Line + 1
        Next
        If FooterYes = True Then
           Do While Line < 72
                    Print #1, " "
                    Line = Line + 1
            Loop
            Line = 0
            FooterYes = False
        End If
        If kkk.State = 1 Then kkk.close
        CNSetup
        kkk.Open "Select * from setup where " & stringyear & "", con, adOpenKeyset, adLockReadOnly, adCmdText
        If Not kkk.BOF Then
            Print #1, ""
            Print #1, ""
            Print #1, Chr(27) + Chr(15) + Chr(14)
            Print #1, Tab(135); "Page No:  " & Pno
            Print #1, Tab((((paperWidth - (Len(Trim(kkk!cname)) * 2)) / 2) - 15) + LEFTM); Chr(27) + Chr(15) + Chr(14); dspace(Trim(kkk!cname))
            Print #1, Tab(((paperWidth - (Len(Trim(kkk!add1)) * 2)) / 2) + LEFTM); Chr(27) + Chr(15); dspace(Trim(kkk!add1)); Chr(27) + Chr(14)
            Line = Line + 6
        End If
        If trs.State = 1 Then trs.close
        trs.Open "treportGen", con, adOpenKeyset, adLockReadOnly, adCmdText
        If trs.RecordCount <= 0 Then
          GoTo EndFun
        End If
        'xstr = Trs!period
        Print #1, Tab(((paperWidth - (Len(Trim(trs!header)))) / 2) + LEFTM); Chr(27) + Chr(15) + Chr(14); Trim(trs!header)
        Print #1, Tab(LEFTM + ((paperWidth - Len(Trim("Period : " + Trim(xstr)))) / 2)); Trim("Period : " + Trim(xstr))
        Print #1, Tab(LEFTM); repli("-", paperWidth)
        Print #1, Tab(LEFTM); "V.DATE"; Tab(T1 + 1 + LEFTM); "V.T."; Tab(T1 + 6 + LEFTM); "V.NUMBER"; Tab(T2 + 7 + LEFTM); "N A R R A T I O N"; Tab(T4 + 7 + LEFTM); "AMOUNT(Dr.)"; Tab(T5 + 12 + LEFTM); "AMOUNT(Cr.)"; Tab(T6 + 12 + LEFTM); "CHEQ./BILL NO. & DATE"; Tab(T6 + 20 + LEFTM + Len("CHEQ./BILL NO. & DATE")); "BALANCE"
        Print #1, Tab(LEFTM); repli("-", paperWidth)
        Line = Line + 5
        trs.close
        If called1 Then
           called1 = False
           GoTo printagain1
        End If
        If rs1.State = 1 Then
           rs1.close
        End If
        Dim rss As ADODB.Recordset
        Set rss = New ADODB.Recordset
        If rss.State = 1 Then rss.close
        rss.Open "Select DISTINCT genledger, Subledger from treportGen order by genledger,Subledger", con, adOpenKeyset
        If rss.RecordCount > 0 Then
            While Not rss.EOF
                    Balance = 0
                    If rs1.State = 1 Then rs1.close
                    rs1.Open "select * from treportGen  WHERE " & stringyear & " and GENLEDGER= '" & rss!Genledger & "'  AND SUBLEDGER = '" & rss!SUBLEDGER & "' order by vdate,vtype,vno", con, adOpenKeyset, adLockReadOnly, adCmdText
                    Print #1, Tab(LEFTM); rs1!SUBLEDGER
                    Line = Line + 1
                    Dim Ts As Integer
                    Ts = 0
                    If rs1!OpeningBalance > 0 Then
                           Ts = T4 + 6 + LEFTM
                    ElseIf rs1!OpeningBalance < 0 Then
                           Ts = T5 + 10 + LEFTM
                    End If
                    Print #1, Tab(LEFTM); "** Opening Balance as on " & SLEDGERPRINT.date1.Text; Tab(Ts); IIf(rs1!OpeningBalance <> 0, rsets(Trim(Format(Str(Abs(rs1!OpeningBalance)), "0.00")), 12), "")
                    Line = Line + 1
                    Balance = rs1!OpeningBalance
                    Do While Not rs1.EOF
                            Line = Line + 1
                            Print #1, Tab(LEFTM); Trim(rs1!vdate); Tab(13 + LEFTM); (Trim(rs1!vtype)); Tab(17 + LEFTM); (Trim(rs1!vno)); Tab(T2 + 5 + LEFTM);
                            If IsNull(Trim(rs1!narration)) Then
                                xtemp = " "
                            Else
                               xtemp = Trim(rs1!narration)
                            End If
                            Print #1, Trim(xtemp);
                            If rs1!aD > 0 Then
                                    Print #1, Tab(T4 + 6 + LEFTM); rsets(Trim(Format(Str(rs1!aD), "0.00")), 12);
                            End If
                            
                            If rs1!aC > 0 Then
                                 Print #1, Tab(T5 + 10 + LEFTM); rsets(Trim(Format(Str(rs1!aC), "0.00")), 12);
                            End If
                            Balance = Balance + rs1!aD - rs1!aC
                            If Trim(rs1!cbno) <> "" Then
                                 xtemp = Trim(rs1!cbno)
                            Else
                                 xtemp = " "
                            End If
                            Print #1, Tab(T7 - 4); lsets(cnullstr(xtemp), 18);
                            Print #1, Tab(T7 + 20 + LEFTM); rsets(Trim(Format(Str(Balance), "0.00")), 12)
                            Line = Line + 1
                            If Line > MaxLine - 4 Then
                                    called1 = True
                                    FooterYes = True
                                    Pno = Pno + 1
                                    GoTo header
printagain1:
                                    
                                    called1 = False
                            End If
                            If Not rs1.EOF Then
                                   rs1.MoveNext
                            End If
                    Loop
                    Print #1, Tab(LEFTM); repli("-", paperWidth)
                    Print #1, " "
                    Line = Line + 2
                    rss.MoveNext
               Wend
        End If
printfooter:
        Do While Line < 72
            Print #1, " "
            Line = Line + 1
        Loop
  
EndFun:
            Close #1
End Function


Private Sub printCOMMAND_Click()
'''  '  Printdlg.Show
'''   ' Printdlg.ZOrder 0
'''   ' Unload Me
'''   ' c1.Flags = &H8&
'''    c1.PrinterDefault = True
'''    c1.ShowPrinter
'''    printnow
'''    Dim X As Long
'''    Dim p As Printer
'''    For I = 0 To Printers.Count - 1
'''        Set p = Printers(I)
'''        If Trim(p.DeviceName) = Trim(Printdlg.Comboprinters.Text) Then
'''            Exit For
'''        End If
'''    Next
'''    For I = 1 To (Printdlg.UpDown1.Value)
'''        X = Shell("" + App.Path + "\ppp.bat " + App.Path + "\vipin.txt " & Trim(p.Port))
'''    Next
'''    Printdlg.UpDown1.Value = 1
'''    Printdlg.Text1.Text = "1"
'''
'''
'''
'''
'''
'''
    On Error GoTo errorcancel
    Dim frompage, topage, copies As Integer
    c1.Flags = 0
    c1.Max = 11
    c1.Flags = &H40000 Or &H4&
    c1.ShowPrinter
    frompage = c1.frompage
    topage = c1.topage
    copies = c1.copies
    If c1.Flags And &H20 Then
        c1.ShowSave
    End If
    'X = Shell("" + App.Path + "\ppp.bat " + App.Path + "\vipin.txt " & Trim(Printer.Port))
errorcancel:
    
End Sub

