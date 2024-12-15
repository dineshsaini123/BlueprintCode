VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form viewledger 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   1410
   ClientWidth     =   11460
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "viewledg.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   11460
   ShowInTaskbar   =   0   'False
   Begin RichTextLib.RichTextBox r1 
      Height          =   5985
      Left            =   30
      TabIndex        =   4
      Top             =   90
      Width           =   11340
      _ExtentX        =   20003
      _ExtentY        =   10557
      _Version        =   393217
      ScrollBars      =   3
      RightMargin     =   20000
      TextRTF         =   $"viewledg.frx":000C
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
      Height          =   345
      Left            =   4200
      Picture         =   "viewledg.frx":008C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6300
      Width           =   345
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "viewledg.frx":01FE
      Left            =   4650
      List            =   "viewledg.frx":0217
      TabIndex        =   2
      Text            =   "50 %"
      Top             =   6300
      Width           =   1095
   End
   Begin VB.CommandButton export 
      Height          =   345
      Left            =   3780
      Picture         =   "viewledg.frx":024A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6300
      Width           =   345
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   435
      Left            =   5820
      TabIndex        =   0
      Top             =   6300
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog C1 
      Left            =   3240
      Top             =   6300
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FromPage        =   1
      PrinterDefault  =   0   'False
      ToPage          =   1
   End
End
Attribute VB_Name = "viewledger"
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
    Command1.Top = r1.Top + r1.Height + 30
    Combo1.Top = r1.Top + r1.Height + 30

 

'    Set CON = New ADODB.Connection
'    Set CON = New ADODB.Connection
    Set RS = New ADODB.Recordset
''    With CON
'        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + App.Path + "\" + Trim(main.directory) + "\data.mdb"
'        .Open
'    End With
'    With CON
'        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + App.Path + "\" + Trim(main.directory) + "\tchitra.mdb"
'        .Open
'    End With
    
    If SLEDGERPRINT.Alpha.Visible = True Then
      genreport1
    Else
      genreport1
    End If
    r1.FileName = "" + App.Path + "\vipin.txt"
    r1.LoadFile (r1.FileName)
End Sub

Private Sub Form_Resize()
If Me.Width > 350 And Me.Height > 1500 Then
    r1.Width = Me.Width - 250
    r1.Height = Me.Height - 1000
    Command1.Top = r1.Top + r1.Height + 30
    Combo1.Top = r1.Top + r1.Height + 30
    export.Top = Command1.Top
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
        MaxLine = 50
        called1 = False
        called2 = False
        Dim Line As Integer
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
header:
        Dim I As Integer
        For I = 1 To main.repors!TopMargin
            Print #1, ""
            Line = Line + 1
        Next
        If kkk.State = 1 Then
            kkk.close
        End If
        CNSetup
        kkk.Open "Select * from setup1 where " & stringyear & "", con, adOpenKeyset, adLockReadOnly, adCmdText
        If Not kkk.BOF Then
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, Chr(27) + Chr(15) + Chr(14)
            Print #1, Tab(135); "Page No:  " & Pno
            Print #1, Tab((((paperWidth - (Len(Trim(kkk!cname)) * 2)) / 2) - 15) + LEFTM); Chr(27) + Chr(15) + Chr(14); dspace(Trim(kkk!cname))
            Print #1, Tab(((paperWidth - (Len(Trim(kkk!add1)) * 2)) / 2) + LEFTM); Chr(27) + Chr(15); dspace(Trim(kkk!add1))
            
            Line = Line + 8
        End If
        If trs.State = 1 Then
            trs.close
        End If
        Print #1, Chr(27) + Chr(14)
        Line = Line + 1
        trs.Open "treport", con, adOpenKeyset, adLockReadOnly, adCmdText
        xstr = trs!Period
        Print #1, Tab(((paperWidth - (Len(Trim(trs!header)))) / 2) + LEFTM); Chr(27) + Chr(15) + Chr(14); Trim(trs!header)
        Line = Line + 1
        Print #1, Tab(LEFTM + ((paperWidth - Len(Trim("Period : " + Trim(xstr)))) / 2)); Trim("Period : " + Trim(xstr))
        Line = Line + 1
        Print #1, Tab(LEFTM); repli("-", paperWidth)
        Line = Line + 1
        Print #1, Tab(LEFTM); "V.DATE"; Tab(T1 + 1 + LEFTM); "V.T."; Tab(T1 + 6 + LEFTM); "V.NUMBER"; Tab(T2 + 7 + LEFTM); "N A R R A T I O N"; Tab(T4 + 7 + LEFTM); "AMOUNT(Dr.)"; Tab(T5 + 12 + LEFTM); "AMOUNT(Cr.)"; Tab(T6 + 12 + LEFTM); "CHEQ./BILL NO. & DATE"; Tab(T6 + 20 + LEFTM + Len("CHEQ./BILL NO. & DATE")); "BALANCE"
        Print #1, Tab(LEFTM); repli("-", paperWidth)
        Line = Line + 1
        trs.close
        If called1 Then
            GoTo printagain1
        End If
        If rs1.State = 1 Then
            rs1.close
        End If
      rs1.Open "select * from treport order by vdate,vtype,vno", con, adOpenKeyset, adLockReadOnly, adCmdText
       'rs1.Open "select * from treport", CON, adOpenKeyset, adLockReadOnly, adCmdText
        If Not rs1.BOF Then
            Print #1, Tab(LEFTM); rs1!SUBLEDGER
            Line = Line + 1
            Print #1, Tab(LEFTM); cnullstr(Trim(rs1!Text)); Tab(T4 + 6 + LEFTM); rsets(Trim(Format(Str(rs1!aD), "0.00")), 12)
            Line = Line + 1
            rs1.MoveNext
        End If
        If Not rs1.BOF Then
            Do While Not rs1.EOF
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
                   Print #1, Tab(T5 + 10 + LEFTM); rsets(Trim(Format(Str(rs1!aC), "0.00")), 12); Tab(T7 + 2 + LEFTM);
               End If
               
               If Trim(rs1!cbno) <> "" Then
                    xtemp = Trim(rs1!cbno)
               Else
                    xtemp = " "
               End If
               Print #1, lsets(cnullstr(xtemp), 18);
               Print #1, Tab(T7 + 20 + LEFTM); rsets(Trim(Format(Str(rs1!Balance), "0.00")), 12)
               Line = Line + 1
               If Line > MaxLine Then
                    called1 = True
                    GoTo printfooter
printnext:
                    Line = 0
                    Print #1, Chr(12)
                    Pno = Pno + 1
                    GoTo header
printagain1:
                    Line = 0
                    called1 = False
                End If
                If Not rs1.EOF Then
                    rs1.MoveNext
                End If
            Loop
printfooter:

            Print #1, Tab(LEFTM); repli("-", paperWidth)
            If Line < MaxLine Then
               Do While Line < MaxLine
                    Print #1, " "
                    Line = Line + 1
                Loop
            End If
            'Print #1, Tab(LEFTM); repli("-", paperwidth)
            Dim tempdata As ADODB.Recordset
            Set tempdata = New ADODB.Recordset
            CNSetup
            tempdata.Open "Select * from setup1 where " & stringyear & "", con, adOpenKeyset, adLockReadOnly, adCmdText
            If called1 Then
                GoTo printnext
            End If
           End If
           Print #1, ""
           Print #1, ""
           Print #1, ""
       If trs.State = 1 Then
       trs.close
       End If
       
        Close #1
End Function


'******************* For  Sub ledger printing
Function SelectedParty()
    Dim rsqty As New ADODB.Recordset
    Dim called1, called2 As Boolean
    Dim MaxLine As Integer
    Dim T1, T2, T3, T4, T5, T6, T7, T8 As Integer
    Dim paperWidth As Integer
    Dim xtemp As String
    Dim qtyop
    Dim trs As ADODB.Recordset
    Dim Pno As Integer
    Dim debit, credit As Double
    Dim debitT, creditT As Double
    Dim sub_party As String
    Dim TotalQty
    TotalQty = 0
    qtyop = 0
    debit = 0
    credit = 0
    debitT = 0
    creditT = 0
    din = 0
    
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
        Dim DR_SUM As Double
        Dim CR_SUM As Double
        Dim qty_sum As Double
        Dim Balance As Double
        Dim rs1 As ADODB.Recordset
        Dim kkk As ADODB.Recordset
        Set kkk = New ADODB.Recordset
        Set rs1 = New ADODB.Recordset
        con.Execute "delete from Winrpt where uid=" & UId & ""
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
        Dim FooterYes As Boolean
        FooterYes = False
        called1 = False
        MaxLine = 72
header:
header1:
        
        Dim I As Integer
        Dim NARR As String
        For I = 1 To main.repors!TopMargin
            Line = Line + 1
            Print #1, ""
            
        Next
        If FooterYes = True Then
             Do While Line < 72
               Line = Line + 1
               Print #1, ""
               
            Loop
            Line = 0
            FooterYes = False
        End If
        If kkk.State = 1 Then kkk.close
        CNSetup
        kkk.Open "Select * from setup where " & stringyear & "", con, adOpenDynamic, adLockReadOnly, adCmdText
        If Not kkk.BOF Then
            Print #1, ""
            'Print #1, ""
            'Print #1, ""
            
            Print #1, Chr(27) + Chr(15) + Chr(14)
            Print #1, Tab(122); "Page No:  " & Pno
            Line = Line + 3
            Print #1, Tab((((paperWidth - (Len(Trim(kkk!cname)) * 2)) / 2)) + LEFTM); Chr(27) + Chr(15) + Chr(14); Trim(kkk!cname)
            Print #1, Tab(((paperWidth - (Len(Trim(kkk!add1)) * 2)) / 2) + LEFTM); Chr(27) + Chr(15); dspace(Trim(kkk!add1)); Chr(27) + Chr(14)
            Line = Line + 2
        End If
        
        
        xstr = SLEDGERPRINT.date1 & " To " & SLEDGERPRINT.date2
        If SLEDGERPRINT.Alpha.Visible = True Then
           Print #1, Tab(((paperWidth - (Len(Trim(SLEDGERPRINT.COMBOGENLEDGER.Text)) * 2)) / 2) + LEFTM); Chr(27) + Chr(15) + Chr(14); Trim(SLEDGERPRINT.COMBOGENLEDGER.Text)
           Line = Line + 1
        Else
           Print #1, Tab(((paperWidth - (Len(Trim("SUB LEDGER ACCOUNT")) * 2)) / 2) + LEFTM); Chr(27) + Chr(15) + Chr(14); Trim("SUB LEDGER ACCOUNT")
           Line = Line + 1
        End If
        Print #1, Tab(LEFTM + ((paperWidth - Len(Trim("Period : " + Trim(xstr)))) / 2)); Trim("Period : " + Trim(xstr))
        Print #1, Tab(LEFTM); repli("-", paperWidth)
        Print #1, Tab(LEFTM); "V.NUMBER"; Tab(T1 + 1 + LEFTM); ""; Tab(T1 + 2 + LEFTM); "V.DATE"; Tab(T2 + 2 + LEFTM); "N A R R A T I O N"; Tab(T4 + 10 + LEFTM); "AMOUNT(Dr.)"; Tab(T5 + 12 + LEFTM); "AMOUNT(Cr.)"; Tab(T6 + 12 + LEFTM); "Bill No."; Tab(T6 + 28 + LEFTM + Len("Bill No.")); "BALANCE"
        Print #1, Tab(LEFTM); repli("-", paperWidth)
        Line = Line + 4
        Line = Line + 1
        Print #1, ""
        
        If called1 = True Then
            called1 = False
            GoTo printagain1
            GoTo printagain2
        End If
     
      
        
        If trs.State = 1 Then trs.close
        ''' line- 6
        '=************
        trs.Open "select * from  treport where " & stringyear & " and userid = " & UId & "", con, adOpenStatic, adLockReadOnly
        '*************
        If trs.RecordCount <= 0 Then
          GoTo EndFun
        End If
         
        
        If rs1.State = 1 Then
            rs1.close
        End If
      Dim rss As ADODB.Recordset
      Set rss = New ADODB.Recordset
      If rss.State = 1 Then rss.close
      If SLEDGERPRINT.Alpha.Visible = True Then
        rss.Open "Select DISTINCT genledger, Subledger from treport  where  " & stringyear & " and Openingbalance <>0 or ad<>0 or ac<>0   order by genledger,Subledger", con, adOpenStatic
      Else
        rss.Open "Select DISTINCT genledger, Subledger from treport where " & stringyear & " and Openingbalance <>0 or ad<>0 or ac<>0 order by genledger,Subledger", con, adOpenStatic
      End If
      
      ''line -6
      
      If rss.RecordCount > 0 Then
            While Not rss.EOF
                   Dim X1, X2 As Integer
                    Balance = 0
                    If rs1.State = 1 Then rs1.close
                    If IsNull(rss!SUBLEDGER) Then
                    rs1.Open "select * from treport  WHERE  " & stringyear & " and GENLEDGER= '" & rss!Genledger & "'  AND isnull(SUBLEDGER) order by vdate,vtype,vno,sno", con, adOpenStatic, adLockReadOnly, adCmdText
                    Else
                    rs1.Open "select * from treport  WHERE  " & stringyear & " and GENLEDGER= '" & rss!Genledger & "'  AND SUBLEDGER = '" & rss!SUBLEDGER & "' order by vdate,vtype,vno,sno", con, adOpenStatic, adLockReadOnly, adCmdText
                    End If
                    Line = Line + 1
                    Print #1, Tab(LEFTM); Line & " " & rs1!SUBLEDGER & "  (" & SLEDGERPRINT.COMBOGENLEDGER.Text & ")"
                    
                    If rs1.RecordCount > 0 Then
                    If rs1!OpeningBalance >= 0 Then
                         
                         X1 = T4 + 6 + LEFTM
                         Line = Line + 1
                         Print #1, Tab(LEFTM + 10); "** Opening Balance as on " & SLEDGERPRINT.date1.Text & " **"; Tab(X1); rsets(Trim(Format(Abs(rs1!OpeningBalance), "0.00")), 12); Tab(114); rsets(Trim(Format(Abs(rs1!OpeningBalance), "0.00")), 12) & " Dr"
                         
                         con.Execute "insert into winrpt(Narration,Receipt,Payment,Description,FromDate,ToDate,party,uid,fyear,setupid) values( '" & "Opening Balance as on " & SLEDGERPRINT.date1.Text & "' ," & rs1!OpeningBalance & "," & 0 & ",'" & "SUB LEDGER ACCOUNT" & "',convert(smalldatetime,'" & SLEDGERPRINT.date1.Text & "',103),convert(smalldatetime,'" & SLEDGERPRINT.date2.Text & "',103),'" & rs1!SUBLEDGER & "'," & UId & ",'" & main.session & "'," & main.setupid & ")"
                         
                    ElseIf rs1!OpeningBalance < 0 Then
                         
                         X1 = T5 + 8 + LEFTM
                         Line = Line + 1
                         Print #1, Tab(LEFTM + 10); "** Opening Balance as on " & SLEDGERPRINT.date1.Text & " **"; Tab(X1); rsets(Trim(Format(Str((rs1!OpeningBalance)), "0.00")), 12); Tab(114); rsets(Trim(Format(Abs(rs1!OpeningBalance), "0.00")), 12) & " Cr"
                         con.Execute "insert into winrpt(Narration,Receipt,Payment,Description,FromDate,ToDate,party,uid,fyear,setupid) values( '" & "Opening Balance as on " & SLEDGERPRINT.date1.Text & "' ," & 0 & "," & (-1 * rs1!OpeningBalance) & ",'" & "SUB LEDGER ACCOUNT" & "',convert(smalldatetime,'" & SLEDGERPRINT.date1.Text & "',103),convert(smalldatetime,'" & SLEDGERPRINT.date2.Text & "',103),'" & rs1!SUBLEDGER & "'," & UId & ",'" & main.session & "'," & main.setupid & ")"
                         
                         
                    End If
                    End If
                    
                    'Print #1, ""
                    
                    Balance = Val(rs1!OpeningBalance & "")
                    ''''''''''''''
                    
                    If Line > MaxLine - 15 Then
                    called1 = True
printnext1:
                    FooterYes = True
                    Pno = Pno + 1
                    GoTo header1
printagain2:
                    
                    called1 = False
                    End If
                    Do While Not rs1.EOF
                       If Trim(rs1!vno) <> 0 Then
                                    If IsNull(Trim(rs1!narration)) Or Trim(rs1!narration) = "" Then
                                       xtemp = ""
                                    Else
                                       xtemp = Left(Trim(rs1!narration), 40)
                                    End If
                                    'Line = Line + 1
                                    If Not IsNull(rs1!vtype) Then
                                    Line = Line + 1
                                    Print #1, Tab(LEFTM); Trim(rs1!vtype); Tab(5 + LEFTM); Trim(rs1!vno); Tab(13 + LEFTM); IIf(IsNull(rs1!vdate) = True, "", Trim(rs1!vdate)) & "-"; Tab(T2 + LEFTM);
                                    End If
                                    credit = 0
                                    debit = 0
                                    'Line = Line + 1
                                    Print #1, Left(Trim(xtemp), 38);
                                    If rs1!aD <> 0 And rs1!dorc = "D" Then
                                        'Line = Line + 1
                                        Print #1, Tab(T4 + 6 + LEFTM); IIf(rs1!aD <> 0, rsets(Trim(Format(Str(rs1!aD), "0.00")), 12), "");
                                        Balance = Balance + rs1!aD
                                        debit = rs1!aD
                                    ElseIf rs1!aD <> 0 And rs1!dorc = "C" Then
                                        'Line = Line + 1
                                        Print #1, Tab(T5 + 8 + LEFTM); IIf(rs1!aD <> 0, rsets(Trim(Format(Str(rs1!aD), "0.00")), 12), "");
                                        Balance = Balance - rs1!aD
                                        credit = rs1!aD
                                    End If
                                    'Line = Line + 1
                                    If Val(Trim(rs1!cbno)) <> 0 Then
                                       
                                       Print #1, Tab(T7 - 2); Trim(rs1!cbno);
                                    End If
                                    If Balance < 0 Then
                                    
                                    Print #1, Tab(T7 + 14 + LEFTM); rsets(Trim(Format(Str((-1 * Balance)), "0.00")), 12) & " Cr"
                                    Else
                                    
                                    Print #1, Tab(T7 + 14 + LEFTM); rsets(Trim(Format(Str(Balance), "0.00")), 12) & " Dr"
                                    End If
                                    
                                    
                                    qty_sum = qty_sum + Val(rs1!cbno)
                                    DR_SUM = DR_SUM + debit
                                    CR_SUM = CR_SUM + credit
                             
                            
                            If IsNull(rs1!vdate) Then
                            If rs1!header <> "" Then
                              NARR = rs1!header & vbCrLf & rs1.Fields("vtype").value & " " & rs1!vno & " " & rs1!narration
                            Else
                              NARR = rs1.Fields("vtype").value & " " & rs1!vno & " " & rs1!narration
                            End If
   
                               con.Execute "insert into winrpt(Narration,Receipt,Payment,Qty,op,party,uid,fyear,setupid) values('" & NARR & "'," & debit & "," & credit & "," & Val(rs1!cbno) & "," & Balance & ",'" & rs1!SUBLEDGER & "'," & UId & ",'" & main.session & "'," & main.setupid & ")"
                            Else
                               
                            If rs1!header <> "" Then
                              NARR = rs1!header & vbCrLf & rs1.Fields("vtype").value & " " & rs1!vno & " " & rs1!narration
                            Else
                              NARR = rs1.Fields("vtype").value & " " & rs1!vno & " " & rs1!narration
                            End If
                               
                               con.Execute "insert into winrpt(date1,Narration,Receipt,Payment,Qty,op,party,uid,fyear,setupid) values(convert(smalldatetime,'" & rs1!vdate & "',103),'" & NARR & "'," & debit & "," & credit & "," & Val(rs1!cbno) & "," & Balance & ",'" & rs1!SUBLEDGER & "'," & UId & ",'" & main.session & "'," & main.setupid & ")"
                            End If
                       End If
                            
                            If Line > MaxLine - 15 Then
                                    called1 = True
printnext:
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
                    
                    If DR_SUM = 0 And CR_SUM = 0 And qty_sum = 0 Then
                    Else
                    'Print #1, Tab(60); "------------------------------------------------"
                    Print #1, Tab(60); "---------------------------------"
                    Print #1, Tab(T4 + 6 + LEFTM); IIf(DR_SUM <> 0, rsets(Trim(Format(Str(DR_SUM), "0.00")), 12), "");
                    Print #1, Tab(T5 + 8 + LEFTM); IIf(CR_SUM <> 0, rsets(Trim(Format(Str(CR_SUM), "0.00")), 12), "");
                    'Print #1, Tab(T7 - 2); Trim(IIf(qty_sum = 0, "", qty_sum));
                    Print #1, Tab(T7 - 2); "";
                    Print #1, Tab(60); "---------------------------------"
                    Print #1, " "
                    Line = Line + 4
                    End If
                    TotalQty = TotalQty + qty_sum
                    debitT = debitT + DR_SUM
                    creditT = creditT + CR_SUM
                    qty_sum = 0
                    DR_SUM = 0
                    CR_SUM = 0
                    rss.MoveNext
           Wend
     End If
printfooter:
            
            Print #1, Tab(60); "---------------------------------"
            Print #1, Tab(T4 + 6 + LEFTM); IIf(debitT <> 0, rsets(Trim(Format(Str(debitT), "0.00")), 12), "");
            Print #1, Tab(T5 + 8 + LEFTM); IIf(creditT <> 0, rsets(Trim(Format(Str(creditT), "0.00")), 12), "");
            'Print #1, Tab(T7 - 8); rsets(Trim(Format(Str(TotalQty), "0.000")), 10);
            Print #1, Tab(T7 - 8); "";
            Print #1, Tab(60); "---------------------------------"
            Line = Line + 3
            Do While Line < 72
                Print #1, ""
                Line = Line + 1
            Loop
            If trs.State = 1 Then
                trs.close
            End If
EndFun:
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
    
    Dim dr, cr
    dr = 0
    cr = 0
        
        DoEvents
        DoEvents
        SLEDGERPRINT.pb.Visible = True
        
        paperWidth = 150
        T1 = 10
        T2 = 23
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
        Dim rs1 As ADODB.Recordset
        Dim kkk As ADODB.Recordset
        Set kkk = New ADODB.Recordset
        Set rs1 = New ADODB.Recordset
        
        con.Execute "delete from Winrpt where uid=" & UId & ""
        
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
        Dim FooterYes As Boolean
        FooterYes = False
        called1 = False
        MaxLine = 72
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
        kkk.Open "Select * from setup1 where " & stringyear & "", con, adOpenKeyset, adLockReadOnly, adCmdText
        If Not kkk.BOF Then
            Print #1, ""
            Print #1, ""
            Print #1, Chr(27) + Chr(15) + Chr(14)
            Print #1, Tab(122); "Page No:  " & Pno
            Print #1, Tab((((paperWidth - (Len(Trim(kkk!cname)) * 2)) / 2) - 15) + LEFTM); Chr(27) + Chr(15) + Chr(14); dspace(Trim(kkk!cname))
            Print #1, Tab(((paperWidth - (Len(Trim(kkk!add1)) * 2)) / 2) + LEFTM); Chr(27) + Chr(15); dspace(Trim(kkk!add1)); Chr(27) + Chr(14)
            Line = Line + 6
        End If
        If trs.State = 1 Then trs.close
        trs.Open "select * from  treport where userid = " & UId & "", con, adOpenStatic, adLockReadOnly
        If trs.RecordCount <= 0 Then
          GoTo EndFun
        End If
        xstr = SLEDGERPRINT.date1 & " To " & SLEDGERPRINT.date2
        If SLEDGERPRINT.Alpha.Visible = True Then
           Print #1, Tab(((paperWidth - (Len(Trim(SLEDGERPRINT.COMBOGENLEDGER.Text)) * 2)) / 2) + LEFTM); Chr(27) + Chr(15) + Chr(14); Trim(SLEDGERPRINT.COMBOGENLEDGER.Text)
        Else
           Print #1, Tab(((paperWidth - (Len(Trim("SUB LEDGER ACCOUNT")) * 2)) / 2) + LEFTM); Chr(27) + Chr(15) + Chr(14); Trim("SUB LEDGER ACCOUNT")
        End If
        Line = Line + 1
        
        Print #1, Tab(LEFTM + ((paperWidth - Len(Trim("Period : " + Trim(xstr)))) / 2)); Trim("Period : " + Trim(xstr))
        Line = Line + 1
        Print #1, Tab(LEFTM); repli("-", paperWidth)
        Print #1, Tab(LEFTM); "V.DATE"; Tab(T1 + 1 + LEFTM); "V.T."; Tab(T1 + 6 + LEFTM); "V.N."; Tab(T2 + 4 + LEFTM); "N A R R A T I O N"; Tab(T4 + 7 + LEFTM); "AMOUNT(Dr.)"; Tab(T5 + 12 + LEFTM); "AMOUNT(Cr.)"; Tab(T6 + 12 + LEFTM); "CHEQ./BILL NO. & DATE"; Tab(T6 + 20 + LEFTM + Len("CHEQ./BILL NO. & DATE")); "BALANCE"
        Print #1, Tab(LEFTM); repli("-", paperWidth)
        Print #1, ""
        Line = Line + 4
        trs.close
        If called1 = True Then
            called1 = False
            GoTo printagain1
        End If
        
          
    
      
      If rs1.State = 1 Then
        rs1.close
      End If
      SLEDGERPRINT.pb.value = 0
      Dim rss As ADODB.Recordset
      Set rss = New ADODB.Recordset
      If rss.State = 1 Then rss.close
      rss.Open "Select DISTINCT genledger, Subledger from treport  where userid = " & UId & " and " & stringyear & " and (Openingbalance <>0 or ad<>0 or ac<>0 )  order by genledger,Subledger", con, adOpenStatic
      If rss.RecordCount > 0 Then
         
            SLEDGERPRINT.pb.Max = rss.RecordCount + 1
        
            While Not rss.EOF
                 
                   Dim X1, X2 As Integer
                   dr = 0
                   cr = 0
                   
                   SLEDGERPRINT.pb.value = SLEDGERPRINT.pb.value + 1
                   
                    Balance = 0
                    If rs1.State = 1 Then rs1.close
                    rs1.Open "select * from treport  WHERE userid = " & UId & " and " & stringyear & " and GENLEDGER= '" & rss!Genledger & "'  AND SUBLEDGER = '" & rss!SUBLEDGER & "' order by vdate,vtype,vno,sno", con, adOpenStatic, adLockReadOnly, adCmdText
                    Print #1, Tab(LEFTM); rs1!SUBLEDGER
                    Line = Line + 1
                    If rs1.RecordCount > 0 Then
                    If rs1!OpeningBalance >= 0 Then
                         X1 = T4 + 6 + LEFTM
                         Print #1, Tab(LEFTM + 10); "** Opening Balance as on " & SLEDGERPRINT.date1.Text & " **"; Tab(X1); rsets(Trim(Format(Abs(rs1!OpeningBalance), "0.00")), 12)
                         Line = Line + 1
                         con.Execute "insert into winrpt(Narration,op,Receipt,Payment,Description,FromDate,ToDate,party,uid) values( '" & "Opening Balance as on " & SLEDGERPRINT.date1.Text & "'," & rs1!OpeningBalance & " ," & rs1!OpeningBalance & "," & 0 & ",'" & "SUB LEDGER ACCOUNT" & "','" & Format(SLEDGERPRINT.date1.Text, "MM/dd/yyyy") & "','" & Format(SLEDGERPRINT.date2.Text, "MM/dd/yyyy") & "','" & rs1!SUBLEDGER & "'," & UId & ")"
                         
                    ElseIf rs1!OpeningBalance < 0 Then
                         X1 = T5 + 10 + LEFTM
                         Print #1, Tab(LEFTM + 10); "** Opening Balance as on " & SLEDGERPRINT.date1.Text & " **"; Tab(X1); rsets(Trim(Format(Str(rs1!OpeningBalance), "0.00")), 12)
                         Line = Line + 1
                         con.Execute "insert into winrpt(Narration,op,Receipt,Payment,Description,FromDate,ToDate,party,uid) values( '" & "Opening Balance as on " & SLEDGERPRINT.date1.Text & "'," & (rs1!OpeningBalance) & "," & 0 & "," & (-1 * rs1!OpeningBalance) & ",'" & "SUB LEDGER ACCOUNT" & "','" & Format(SLEDGERPRINT.date1.Text, "MM/dd/yyyy") & "','" & Format(SLEDGERPRINT.date2.Text, "MM/dd/yyyy") & "','" & rs1!SUBLEDGER & "'," & UId & ")"
                    End If
                    End If
                    Print #1, ""
                    Line = Line + 1
                    Balance = rs1!OpeningBalance
                    Do While Not rs1.EOF
                       If Trim(rs1!vno) <> 0 Then
                                    Print #1, Tab(LEFTM); IIf(IsNull(rs1!vdate) = True, "", Trim(rs1!vdate)); Tab(13 + LEFTM); Trim(rs1!vtype); Tab(16 + LEFTM); Trim(rs1!vno); Tab(T2 + 2 + LEFTM);
                                    If IsNull(Trim(rs1!narration)) Or Trim(rs1!narration) = "" Then
                                        xtemp = ""
                                    Else
                                       xtemp = Trim(Mid(rs1!narration, 1, 36))    '38
                                    End If
                                    Print #1, Trim(xtemp);
                                    If rs1!aD <> 0 And rs1!dorc = "D" Then
                                        Print #1, Tab(T4 + 6 + LEFTM); IIf(rs1!aD <> 0, rsets(Trim(Format(Str(rs1!aD), "0.00")), 12), "");
                                        Balance = Balance + rs1!aD
                                        dr = dr + rs1!aD
                                    ElseIf rs1!aD <> 0 And rs1!dorc = "C" Then
                                        Print #1, Tab(T5 + 10 + LEFTM); IIf(rs1!aD <> 0, rsets(Trim(Format(Str(rs1!aD), "0.00")), 12), "");
                                        Balance = Balance - rs1!aD
                                        cr = cr + rs1!aD
                                    End If
                                    If Trim(rs1!cbno) <> "" Then
                                         xtemp = Trim(rs1!cbno)
                                    Else
                                         xtemp = " "
                                    End If
                                    Print #1, Tab(T7 - 4); lsets(cnullstr(xtemp), 18);
                                    Print #1, Tab(T7 + 20 + LEFTM); rsets(Trim(Format(Str(Balance), "0.00")), 12)
                                    Line = Line + 1
                                    
                                    If IsNull(rs1!vdate) Then
                                        con.Execute "insert into winrpt(Narration,Receipt,Payment,Qty,op,party,uid) values('" & rs1.Fields("vtype").value & " " & rs1!vno & " " & rs1!narration & "'," & IIf(rs1!dorc = "D", rs1!aD, 0) & "," & IIf(rs1!dorc = "C", rs1!aD, 0) & ",'" & rs1!cbno & "'," & Balance & ",'" & rs1!SUBLEDGER & "'," & UId & ")"
                                    Else
                                        con.Execute "insert into winrpt(date1,Narration,Receipt,Payment,Qty,op,party,uid) values('" & Format(rs1!vdate, "MM/dd/yyyy") & "','" & rs1.Fields("vtype").value & " " & rs1!vno & " " & rs1!narration & "'," & IIf(rs1!dorc = "D", rs1!aD, 0) & "," & IIf(rs1!dorc = "C", rs1!aD, 0) & ",'" & rs1!cbno & "'," & Balance & ",'" & rs1!SUBLEDGER & "'," & UId & ")"
                                    End If
                                    
                           End If
                            If Line > MaxLine - 10 Then
                                    called1 = True
printnext:
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
                    
                    ''Print #1, Tab(T4 + 6 + LEFTM); IIf(dr_ <> 0, rsets(Trim(Format(str(dr_), "0.00")), 12), ""); Tab(T5 + 8 + LEFTM); IIf(cr_ <> 0, rsets(Trim(Format(str(cr_), "0.00")), 12), "")
                   Print #1, Tab(T4 + 6 + LEFTM); IIf(dr <> 0, rsets(Trim(Format(Str(dr), "0.00")), 12), ""); Tab(T5 + 10 + LEFTM); IIf(cr <> 0, rsets(Trim(Format(Str(cr), "0.00")), 12), "")
                    Line = Line + 2

                    '------------------------------------------
''                    If dr > 0 Then
''                        Print #1, Tab(T4 + 6 + LEFTM); IIf(dr <> 0, rsets(Trim(Format(Str(dr), "0.00")), 12), "");
''                    End If
''
''                    If cr > 0 Then
''                    Print #1, Tab(T5 + 10 + LEFTM); IIf(cr <> 0, rsets(Trim(Format(Str(cr), "0.00")), 12), "")
''                    End If
''
''                    If (dr > 0 Or cr > 0) Then
''                       Line = Line + 1
''                    End If
                    
                    
                    '------------------------------------------

                     
                    
                    rss.MoveNext
           Wend
     End If
printfooter:
            Print #1, Tab(LEFTM); repli("-", paperWidth)
            Line = Line + 1
            

            
            Do While Line <= 72
                Print #1, " "
                Line = Line + 1
            Loop
            
            
            If trs.State = 1 Then
                trs.close
            End If
EndFun:
            Close #1
            
            DoEvents
            DoEvents
            SLEDGERPRINT.pb.Visible = False
End Function















Private Sub printCOMMAND_Click()
  '  Printdlg.Show
   ' Printdlg.ZOrder 0
   ' Unload Me
   ' c1.Flags = &H8&
    c1.PrinterDefault = True
    c1.ShowPrinter
'    printnow
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
End Sub







'''''''Function genreport1()
'''''''    Dim called1, called2 As Boolean
'''''''    Dim MaxLine As Integer
'''''''    Dim T1, T2, T3, T4, T5, T6, T7, T8 As Integer
'''''''    Dim paperWidth As Integer
'''''''    Dim xtemp As String
'''''''    Dim trs As ADODB.Recordset
'''''''    Dim Pno As Integer
'''''''    Set trs = New ADODB.Recordset
'''''''       paperWidth = 150
'''''''        T1 = 10
'''''''        T2 = 25
'''''''        T3 = 40
'''''''        T4 = 55
'''''''        T5 = 70
'''''''        T6 = 85
'''''''        T7 = 100
'''''''        T8 = 115
'''''''        MaxLine = 72
'''''''        called1 = False
'''''''        called2 = False
'''''''        Dim Line As Integer
'''''''        Dim BALANCE As Double
'''''''        Dim rs1 As ADODB.Recordset
'''''''        Dim kkk As ADODB.Recordset
'''''''        Set kkk = New ADODB.Recordset
'''''''        Set rs1 = New ADODB.Recordset
'''''''
'''''''        main.reportdata
'''''''        main.repors.Find "reportname='" + Trim(main.reportname) + "'"
'''''''        MaxLine = main.repors!totalline
'''''''        If main.repors!comp = True Then
'''''''            paperWidth = Int(main.repors!totalcolumn * 1.75)
'''''''        Else
'''''''            paperWidth = main.repors!totalcolumn
'''''''        End If
'''''''        Open "" + App.Path + "\vipin.txt" For Output As #1
'''''''        Line = 0
'''''''        Pno = 1
'''''''        Dim FooterYes As Boolean
'''''''        FooterYes = False
'''''''
'''''''header:
'''''''        Dim I As Integer
'''''''        For I = 1 To main.repors!topmargin
'''''''            Print #1, ""
'''''''            Line = Line + 1
'''''''        Next
'''''''        If FooterYes = True Then
'''''''            Print #1, Tab(LEFTM); repli("-", paperWidth)
'''''''            Line = Line + 1
'''''''            Do While Line < 72
'''''''                    Print #1, " "
'''''''                    Line = Line + 1
'''''''            Loop
'''''''            Line = 0
'''''''            FooterYes = True
'''''''        End If
'''''''
'''''''        If kkk.State = 1 Then
'''''''            kkk.Close
'''''''        End If
'''''''        CNSetup
'''''''        kkk.Open "select * from setup where  " & stringyear & "  ", CON, adOpenKeyset, adLockReadOnly, adCmdText
'''''''        If Not kkk.BOF Then
'''''''            Print #1, ""
'''''''            Print #1, ""
'''''''            Print #1, ""
'''''''            Print #1, ""
'''''''            Print #1, Chr(27) + Chr(15) + Chr(14)
'''''''            Print #1, Tab(115); "Page No:  " & Pno
'''''''            Print #1, Tab((((paperWidth - (Len(Trim(kkk!Cname)) * 2)) / 2) - 15) + LEFTM); Chr(27) + Chr(15) + Chr(14); dspace(Trim(kkk!Cname))
'''''''            Print #1, Tab(((paperWidth - (Len(Trim(kkk!add1)) * 2)) / 2) + LEFTM); Chr(27) + Chr(15); dspace(Trim(kkk!add1))
'''''''            Line = Line + 8
'''''''        End If
'''''''        If trs.State = 1 Then
'''''''            trs.Close
'''''''        End If
'''''''        Print #1, Chr(27) + Chr(14)
'''''''        Line = Line + 1
'''''''        trs.Open "treport", CON, adOpenKeyset, adLockReadOnly, adcmdtext
'''''''        If trs.RecordCount <= 0 Then
'''''''          GoTo EndFun
'''''''        End If
'''''''        xstr = SLEDGERPRINT.date1 & " To " & SLEDGERPRINT.date2
'''''''        'xstr = trs!period
'''''''        If SLEDGERPRINT.Alpha.Visible = True Then
'''''''           Print #1, Tab(((paperWidth - (Len(Trim(SLEDGERPRINT.COMBOGENLEDGER.Text)))) / 2) + LEFTM); Chr(27) + Chr(15) + Chr(14); Trim(SLEDGERPRINT.COMBOGENLEDGER.Text)
'''''''        Else
'''''''          Print #1, Tab(((paperWidth - (Len(Trim("SUB LEDGER ACCOUNT")))) / 2) + LEFTM); Chr(27) + Chr(15) + Chr(14); Trim("SUB LEDGER ACCOUNT")
'''''''        End If
'''''''        Line = Line + 1
'''''''        Print #1, Tab(LEFTM + ((paperWidth - Len(Trim("Period : " + Trim(xstr)))) / 2)); Trim("Period : " + Trim(xstr))
'''''''        Line = Line + 1
'''''''        Print #1, Tab(LEFTM); repli("-", paperWidth)
'''''''        Print #1, Tab(LEFTM); "V.DATE"; Tab(T1 + 1 + LEFTM); "V.T."; Tab(T1 + 6 + LEFTM); "V.NUMBER"; Tab(T2 + 7 + LEFTM); "N A R R A T I O N"; Tab(T4 + 7 + LEFTM); "AMOUNT(Dr.)"; Tab(T5 + 12 + LEFTM); "AMOUNT(Cr.)"; Tab(T6 + 12 + LEFTM); "CHEQ./BILL NO. & DATE"; Tab(T6 + 20 + LEFTM + Len("CHEQ./BILL NO. & DATE")); "BALANCE"
'''''''        Print #1, Tab(LEFTM); repli("-", paperWidth)
'''''''        Line = Line + 3
'''''''        trs.Close
'''''''        If called1 = True Then
'''''''            GoTo printagain1
'''''''        End If
'''''''        If rs1.State = 1 Then
'''''''            rs1.Close
'''''''        End If
'''''''      Dim rsS As ADODB.Recordset
'''''''      Set rsS = New ADODB.Recordset
'''''''      If rsS.State = 1 Then rsS.Close
'''''''      rsS.Open "Select DISTINCT genledger, Subledger from treport order by genledger,Subledger", CON, adOpenKeyset
'''''''      If rsS.RecordCount > 0 Then
'''''''          While Not rsS.EOF
'''''''                   Dim X1, X2 As Integer
'''''''                    BALANCE = 0
'''''''                    If rs1.State = 1 Then rs1.Close
'''''''                    rs1.Open "select * from treport  WHERE GENLEDGER= '" & rsS!Genledger & "'  AND SUBLEDGER = '" & rsS!SUBLEDGER & "' order by vdate,vtype,vno", CON, adOpenKeyset, adLockReadOnly, adCmdText
'''''''                    Print #1, Tab(LEFTM); rs1!SUBLEDGER
'''''''                    Line = Line + 1
'''''''                    If rs1!openingbalance > 0 Then
'''''''                         X1 = T4 + 6 + LEFTM
'''''''                    ElseIf rs1!openingbalance < 0 Then
'''''''                         X1 = T5 + 10 + LEFTM
'''''''                    End If
'''''''                    Print #1, Tab(LEFTM); "** Opening Balance as on " & SLEDGERPRINT.date1.Text; Tab(X1); IIf(rs1!openingbalance <> 0, rsets(Trim(Format(Str(Abs(rs1!openingbalance)), "0.00")), 12), "")
'''''''                    Line = Line + 1
'''''''                    BALANCE = rs1!openingbalance
'''''''                    Do While Not rs1.EOF
'''''''                            Print #1, Tab(LEFTM); IIf(IsNull(rs1!vdate), "", Trim(rs1!vdate)); Tab(13 + LEFTM); (Trim(rs1!vtype)); Tab(17 + LEFTM); (Trim(rs1!vno)); Tab(T2 + 5 + LEFTM);
'''''''                            If IsNull(Trim(rs1!narration)) Or Trim(rs1!narration) = "" Then
'''''''                                xtemp = ""
'''''''                            Else
'''''''                               xtemp = Trim(rs1!narration)
'''''''                            End If
'''''''                            Print #1, Trim(xtemp);
'''''''                            If rs1!aD <> 0 And rs1!dorc = "D" Then
'''''''                                Print #1, Tab(T4 + 6 + LEFTM); IIf(rs1!aD <> 0, rsets(Trim(Format(Str(rs1!aD), "0.00")), 12), "");
'''''''                                BALANCE = BALANCE + rs1!aD
'''''''                            ElseIf rs1!aD <> 0 And rs1!dorc = "C" Then
'''''''                                 Print #1, Tab(T5 + 10 + LEFTM); IIf(rs1!aD <> 0, rsets(Trim(Format(Str(rs1!aD), "0.00")), 12), "");
'''''''                                 BALANCE = BALANCE - rs1!aD
'''''''                            End If
'''''''
'''''''                            If Trim(rs1!cbno) <> "" Then
'''''''                                 xtemp = Trim(rs1!cbno)
'''''''                            Else
'''''''                                 xtemp = " "
'''''''                            End If
'''''''                            Print #1, Tab(T7 - 4); lsets(cnullstr(xtemp), 18);
'''''''                            Print #1, Tab(T7 + 20 + LEFTM); rsets(Trim(Format(Str(BALANCE), "0.00")), 12)
'''''''                            Line = Line + 1
'''''''                            If Line > MaxLine - 9 Then
'''''''                                    called1 = True
'''''''printnext:
'''''''                                    FooterYes = True
'''''''                                    Pno = Pno + 1
'''''''                                    GoTo header
'''''''printagain1:
'''''''                                    Line = 0
'''''''                                    called1 = False
'''''''                            End If
'''''''                            If Not rs1.EOF Then
'''''''                                    rs1.MoveNext
'''''''                            End If
'''''''                    Loop
'''''''                    Print #1, Tab(LEFTM); repli("-", paperWidth)
'''''''                    Print #1, " "
'''''''                    Line = Line + 2
'''''''                    rsS.MoveNext
'''''''               Wend
'''''''        End If
'''''''printfooter:
'''''''            Print #1, Tab(LEFTM); repli("-", paperWidth)
'''''''            Line = Line + 1
'''''''            Do While Line < 72
'''''''                Print #1, " "
'''''''                Line = Line + 1
'''''''            Loop
'''''''
'''''''            If trs.State = 1 Then
'''''''                trs.Close
'''''''            End If
'''''''EndFun:
'''''''            Close #1
'''''''End Function
'''''''

