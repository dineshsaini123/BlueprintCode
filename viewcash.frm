VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form viewcash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   1410
   ClientWidth     =   12555
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "viewcash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   12555
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton print 
      Height          =   465
      Left            =   3180
      Picture         =   "viewcash.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8280
      Width           =   825
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "viewcash.frx":017E
      Left            =   4260
      List            =   "viewcash.frx":0191
      TabIndex        =   3
      Text            =   "100 %"
      Top             =   8340
      Width           =   3255
   End
   Begin VB.CommandButton export 
      Caption         =   "Export"
      Height          =   465
      Left            =   2220
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8280
      Width           =   885
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   7560
      TabIndex        =   1
      Top             =   8340
      Width           =   1035
   End
   Begin RichTextLib.RichTextBox r1 
      Height          =   8175
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   12525
      _ExtentX        =   22093
      _ExtentY        =   14420
      _Version        =   393217
      ScrollBars      =   3
      RightMargin     =   20000
      TextRTF         =   $"viewcash.frx":01B7
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
   Begin MSComDlg.CommonDialog c1 
      Left            =   900
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "viewcash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQLSTRING As String
'Dim CON As ADODB.Connection
'Dim CON As ADODB.Connection
Dim RS As ADODB.Recordset
Private Sub Combo1_Change()

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

Private Sub Combo1_Click()
'r1.row = 1
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
    MsgBox "copies =" + Str(d1.copies)
    'd1.Copies
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
    'Set CON = New ADODB.Connection
'    Set CON = New ADODB.Connection
    Set RS = New ADODB.Recordset
    'With CON
    '    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + App.Path + "\" + Trim(main.directory) + "\data.mdb"
    '    .Open
    'End With
'    With CON
'        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + App.Path + "\" + Trim(main.directory) + "\tchitra.mdb"
'        .Open
'    End With
    GenrepNoFooter
    r1.filename = "" + App.Path + "\vipin.txt"
    r1.LoadFile (r1.filename)
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
   c1.PrinterDefault = True
   c1.ShowPrinter
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
    'Printdlg.Show
End Sub
Function genreport()
Dim called1, called2 As Boolean
    Dim Pno As Integer
    Dim MaxLine As Integer
    Dim OPENB As Double
    Dim T1, T2, T3, T4, T5, T6, T7, T8 As Integer
    Dim paperWidth As Integer
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
        Open "" + App.Path + "\vipin.txt" For Output As #1
        Line = 0
        Pno = 1
header:
        If kkk.State = 1 Then
            kkk.close
        End If
        CNSetup
        kkk.Open "Select * from setup1 where " & stringyear & "", con, adOpenKeyset, adLockReadOnly, adCmdText
        If Not kkk.BOF Then
            Print #1, Chr(27) + Chr(15) + Chr(14)
            Print #1, Tab(135); "Page No:  " & Pno
            Print #1, Tab(((paperWidth - (Len(Trim(kkk!cname)) * 2)) / 2) - 15); Chr(27) + Chr(15) + Chr(14); dspace(Trim(kkk!cname))
            Print #1, Tab((paperWidth - (Len(Trim(kkk!add1)) * 2)) / 2); Chr(27) + Chr(15); dspace(Trim(kkk!add1))
            Print #1, Tab((paperWidth - Len(Trim(kkk!phone1))) / 2); Trim(kkk!phone1)
            Line = Line + 3
        End If
        If rs1.State = 1 Then
            rs1.close
        End If
        Print #1, Chr(27) + Chr(14)
       
        Line = Line + 1
        rs1.Open "SELECT * FROM treport order by vdate,vtype,vno  ", con, adOpenKeyset, adLockOptimistic, adCmdText
        xstr = rs1!Period
        Print #1, Chr(27) + Chr(14); Tab((paperWidth - Len(Trim(rs1!header)))); Trim(rs1!header)
        Line = Line + 1
        Print #1, Tab((paperWidth - Len(Trim("Period : " + Trim(xstr)))) / 2); Trim("Period : " + Trim(xstr))
        Line = Line + 1
        Print #1, dspace("RECEIPTS"); Tab(T8 + 7); dspace("PAYMENTS")
        Print #1, repli("-", 146)
        Line = Line + 1
        Print #1, "V.TYPE"; Tab(8); "E.NO."; Tab(T2 + 5); "GEN./SUB. LEDGER ACCOUNT"; Tab(T5 + 5); "V.TYPE"; Tab(T1 + 1 + 75); "E.NO."; Tab(T2 + 5 + 75); "GEN./SUB. LEDGER ACCOUNT";
        Print #1, Tab(T2 + 8); "[  Narration  ]"; Tab(T4 + 8); "Amount"; Tab(T2 + 8 + 75); "[  Narration  ]"; Tab(T4 + 8 + 75); "Amount"
        Print #1, repli("-", 146)
        Line = Line + 1
        rs1.close
        If called1 Then
            GoTo printagain1
        End If
        rs1.Open "select * from treport order by vdate,sno", con, adOpenKeyset, adLockReadOnly, adCmdText
        If Not rs1.BOF Then
            Do While Not rs1.EOF
                If Left$(Trim(rs1!Text), 10) = Trim("** Opening") Then
                    OPENB = rs1!ad
                    Exit Do
                End If
                rs1.MoveNext
            Loop
        End If
        rs1.close
        rs1.Open "select distinct vdate from treport where vdate is not null order by vdate", con, adOpenKeyset, adLockReadOnly, adCmdText
        If Not rs1.BOF Then
            Dim prevdate As String
            Dim prevbal As Double
            Dim prevpayment As Double
            Dim rsr As ADODB.Recordset
            Dim rsp As ADODB.Recordset
            Set rsr = New ADODB.Recordset
            Set rsp = New ADODB.Recordset
            prevdate = ""
            prevpayment = 0
            prevbal = 0
            prevbal = OPENB
            Do While Not rs1.EOF
               If Trim(rs1!vdate) <> "" Then
                    If Trim(prevdate) <> Trim(rs1!vdate) Then
                        Print #1, Chr(27 + 71); "DATE:  "; rs1!vdate; "      ***** Balance B/F *****"; Tab(T4 + 10); rsets(Trim(Format(Str(prevbal), "0.00")), 12); Chr(27 + 72)
                        prevdate = Trim(rs1!vdate)
                        prevbal = prevbal - prevpayment
                        prevpayment = 0
                    End If
               End If
               Set rsr = con.Execute("select * from treport where convert(smalldatetime,vdate,103)=convert(smalldatetime,'" + Trim(prevdate) + "',103) and  dorc='D' order by vdate,sno")
               Set rsp = con.Execute("select * from treport where convert(smalldatetime,vdate,103)=convert(smalldatetime,'" + Trim(prevdate) + "',103) and  dorc='C' order by vdate,sno")
               Do While Not rsr.EOF And Not rsp.EOF
                        Print #1, Tab(2); rsr!vtype; Tab(7); rsr!vno; Tab(T2); rsr!SUBLEDGER; Tab(2 + 75); rsp!vtype; Tab(7 + 75); rsp!vno; Tab(T2 + 75); rsp!SUBLEDGER
                        Print #1, Tab(T2); rsr!narration; Tab(T5); rsets(Trim(Format(Str(rsr!ad), "0.00")), 12); Tab(T2 + 75); rsp!narration; Tab(T5 + 75); rsets(Trim(Format(Str(rsp!ac), "0.00")), 12)
                        prevbal = prevbal + rsr!ad
                        prevpayment = prevpayment + rsp!ac
                        Line = Line + 2
                        If Not rsr.EOF Then
                            rsr.MoveNext
                        End If
                        If Not rsp.EOF Then
                            rsp.MoveNext
                        End If
                Loop
                Do While Not rsr.EOF
                        Print #1, Tab(2); rsr!vtype; Tab(7); rsr!vno; Tab(T2); rsr!SUBLEDGER
                        Print #1, Tab(T2); rsr!narration; Tab(T5); rsets(Trim(Format(Str(rsr!ad), "0.00")), 12)
                        prevbal = prevbal + rsr!ad
                        'prevpayment = prevpayment + rsp!ac
                        
                        Line = Line + 1
                        If Not rsr.EOF Then
                            rsr.MoveNext
                        End If
                    Loop
                    Do While Not rsp.EOF
                        Print #1, Tab(2 + 71); rsp!vtype; Tab(7 + 71); rsp!vno; Tab(T2 + 71); rsp!SUBLEDGER
                        Print #1, Tab(T2 + 75); rsp!narration; Tab(T5 + 75); rsets(Trim(Format(Str(rsp!ac), "0.00")), 12)
                        'prevbal = prevbal + rsr!ad
                        prevpayment = prevpayment + rsp!ac
                        Line = Line + 1
                        If Not rsp.EOF Then
                            rsp.MoveNext
                        End If
                    Loop
                    
                    If CASHBOOK.CheckCash = True Then
                        If prevbal < 0 Then
                            MsgBox "Amount Going In Credit Amount.." & Chr(13) & "Please Check the Amount for Date & rs1!vdate"
                            Close #1
                            CASHBOOK.CheckCash = False
                            Exit Function
                         End If
                         
                    End If
                    
                    Print #1, Tab(T2 + 71); "***** Balance C/F *****"; Tab(T4 + 71); rsets(Trim(Format(Str(prevbal - prevpayment), "0.00")), 12)
                    Print #1, Tab(T2 + 71); '******************'
printnext:
printagain1:
                If Not rs1.EOF Then
                    rs1.MoveNext
                End If
            Loop
printfooter:
            If Line < MaxLine Then
                 Print #1, repli("-", 146)
                Do While Line < MaxLine
                    Print #1, " "
                    Line = Line + 1
                Loop
            End If
            Dim tempdata As ADODB.Recordset
            Set tempdata = New ADODB.Recordset
            CNSetup
            tempdata.Open "Select * from setup1 where " & stringyear & "", con, adOpenKeyset, adLockReadOnly, adCmdText
            Print #1, "E.& O.E"
            Print #1, tempdata!COURT; Tab(paperWidth - ((Len(tempdata!COURT) + Len(tempdata!cname)) * 0.75)); "FOR " + Trim(tempdata!cname)
            Pno = Pno + 1
            If called1 Then
                GoTo printnext
            End If
        End If
        Close #1
End Function
Function GenrepNoFooterNew_Ledger()
    con.Execute "delete from tempCash"
    Dim called1, called2 As Boolean
    Dim id As Integer
    Dim mydate As String
    
    id = 1
    Dim MaxLine As Integer
    Dim OPENB As Double
    Dim T1, T2, T3, T4, T5, T6, T7, T8 As Integer
    Dim paperWidth As Integer
    Dim Pno As Integer
    Dim VARLABEL As Integer
    VARLABEL = -40
    paperWidth = 150
        T1 = 10
        T2 = 25
        T3 = 40
        T4 = 55
        T5 = 70
        T6 = 85
        T7 = 100
        T8 = 115
        called1 = False
        called2 = False
        Dim Line As Integer
        Dim rs1 As ADODB.Recordset
        Dim rs2 As ADODB.Recordset
        Dim kkk As ADODB.Recordset
        Set kkk = New ADODB.Recordset
        Set rs1 = New ADODB.Recordset
        
        Set rs2 = New ADODB.Recordset
        Open "" + VB.App.Path + "\vipin.txt" For Output As #1
        MaxLine = 72
        Pno = 0
header:
        If VARLABEL >= 0 Then
            Do While Line < 72
               Print #1, " "
               Line = Line + 1
            Loop
        Else
           If called1 = True Then
               Do While Line < 72
                    Print #1, " "
                    Line = Line + 1
               Loop
            End If
        End If
        Line = 0
        Pno = Pno + 1
        If kkk.State = 1 Then kkk.close
        CNSetup
        kkk.Open "select * from setup1", con, adOpenStatic, adLockReadOnly, adCmdText
        If Not kkk.BOF Then
            Print #1, ""
            Print #1, ""
            Print #1, Chr(27) + Chr(15) + Chr(14)
            Print #1, Tab(127); "Page No:  " & Pno
            Print #1, Tab(((paperWidth - (Len(Trim(kkk!cname)) * 2)) / 2)); Chr(27) + Chr(15) + Chr(14); Trim(kkk!cname)
            Print #1, Tab((paperWidth - (Len(Trim(kkk!add1)) * 2)) / 2); Chr(27) + Chr(15); dspace(Trim(kkk!add1))
            Line = Line + 6
        End If
        If rs2.State = 1 Then
            rs2.close
        End If
        rs2.Open "treport", con, adOpenStatic, adLockOptimistic, adCmdTable
        xstr = rs2!Period
        Print #1, Chr(27) + Chr(14); Tab((73 - Len(Trim(rs2!header))) / 2); Trim(rs2!header); Chr(27) + Chr(15)
        Line = Line + 1
        'Print #1, Tab((paperWidth - Len(Trim("Period : " + Trim(xstr)))) / 2); Trim("Period : " + Trim(xstr))
        'Line = Line + 1
        Print #1, dspace("RECEIPTS"); Tab(T8 + 7); dspace("PAYMENTS")
        Print #1, repli("-", 143)
        Line = Line + 2
        Print #1, "V.TYPE"; Tab(8); "E.NO."; Tab(T2 - 3); "GEN./SUB. LEDGER ACCOUNT"; Tab(T5 + 7); "V.TYPE"; Tab(T1 + 74); "E.NO."; Tab(T2 + 72); "GEN./SUB. LEDGER ACCOUNT"
        Print #1, Tab(T2 + 5); "[  Narration  ]"; Tab(T4 + 12); "Amount"; Tab(T2 + 80); "[  Narration  ]"; Tab(T4 + 82); "Amount"; Chr(27) + Chr(72)
        Print #1, repli("-", 143)
        Line = Line + 3
        rs2.close
        If called1 Then
            called1 = False
            GoTo printagain0
        End If
        If VARLABEL = 1 Then
            GoTo printagain1
        End If
        If VARLABEL = 2 Then
            GoTo printagain2
        End If
        If VARLABEL = 3 Then
            GoTo printagain3
        End If
        If VARLABEL = 4 Then
            GoTo printagain4
        End If
        If VARLABEL = 5 Then
            GoTo printagain5
        End If
        If VARLABEL = 6 Then
            GoTo printagain6
        End If
        If VARLABEL = 7 Then
            GoTo printagain7
        End If
        If VARLABEL = 8 Then
            GoTo printagain8
        End If
        If VARLABEL = 11 Then
            GoTo printagain11
        End If
        rs1.Open "select * from treport order by vdate,sno", con, adOpenStatic, adLockReadOnly, adCmdText
        If Not rs1.BOF Then
         If IsDate(rs1!vdate) Then
            mydate = rs1!vdate
         End If
        
            Do While Not rs1.EOF
            
            If Left$(Trim(rs1!Text), 10) = Trim("** Opening") Then
                  If rs1!ad > 0 Then
                    OPENB = rs1!ad
                  Else
                     OPENB = -rs1!ac
                  End If
                    Exit Do
                End If
                rs1.MoveNext


              
            Loop
        End If
        rs1.close
        rs1.Open "select Distinct vdate from treport order by vdate", con, adOpenStatic, adLockReadOnly, adCmdText
        If rs1.RecordCount > 0 Then
           rs1.MoveNext
        End If
        If Not rs1.BOF Then
            Dim prevdate As String
            Dim prevbal As Double
            Dim prevpayment As Double
            Dim rsr As ADODB.Recordset
            Dim rsp As ADODB.Recordset
            Dim rsj As ADODB.Recordset
            Dim rsjC As ADODB.Recordset
            Set rsr = New ADODB.Recordset
            Set rsp = New ADODB.Recordset
            Set rsj = New ADODB.Recordset
            Set rsjC = New ADODB.Recordset
            prevdate = ""
            prevpayment = 0
            prevbal = 0
            prevbal = OPENB
            Do While Not rs1.EOF
               If Trim(rs1!vdate) <> "" Then
                    If Trim(prevdate) <> Trim(rs1!vdate) Then
printagain1:
                    Line = Line + 1
                        If Line > MaxLine - 5 Then
                           Line = Line - 1
                           VARLABEL = 1
                           GoTo header
                        End If
                        prevdate = Trim(rs1!vdate)
                        prevbal = prevbal - prevpayment
                        
                        Print #1, Chr(27) + Chr(71); "DATE:  "; rs1!vdate; "      ***** Balance B/F *****"; Tab(T4 + 8); rsets(Trim(Format(Str(prevbal), "0.00")), 12); Chr(27) + Chr(72)
                        
                           
                         id = MaxSNo("tempCash", "id")
                         con.Execute "insert into tempCash(dates,R1,R2,R3,R4,id) values('" & rs1!vdate & "','-','" & "-" & "','***** Balance B/F *****'," & prevbal & "," & id & ")"
                        '-----------------------------------
                        
                        If CASHBOOK.CheckCash = True Then
                             If prevbal < 0 Then
                                      MsgBox "Amount Going In Credit..." & Chr(13) & "Please Check  Amount for Date :" & rs1!vdate
                                      Close #1
                                      CASHBOOK.CheckCash = False
                                      Exit Function
                             End If
                         
                       End If
                       prevpayment = 0
                    End If
               End If
               'Set rsr = con.Execute("select * from treport where vdate=cdate('" + Trim(prevdate) + "') and vtype = 'R' and dorc='C' order by vdate,sno")
               'Set rsp = con.Execute("select * from treport where vdate=cdate('" + Trim(prevdate) + "') and vtype = 'P' and dorc='D' order by vdate,sno")
               ''Set rsj = con.Execute("select * from treport where vdate=cdate('" + Trim(prevdate) + "') and vtype = 'J' and dorc='D' order by vdate,sno")
               'Set rsjC = con.Execute("select * from treport where vdate=cdate('" + Trim(prevdate) + "')and vtype = 'J' and dorc='C' order by vdate,sno")
               'Set rsi = con.Execute("select * from treport where vdate=cdate('" + Trim(prevdate) + "') and vtype = 'S' and dorc='C' order by vdate,sno")
               
               Set rsr = con.Execute("select * from treport where " & stringyear & " and convert(smalldatetime,vdate,103)=convert(smalldatetime,'" + Trim(prevdate) + "',103) and  dorc='C' and vtype ='R' order by vdate,sno")
               Set rsp = con.Execute("select * from treport where " & stringyear & " and convert(smalldatetime,vdate,103)=convert(smalldatetime,'" + Trim(prevdate) + "',103) and  dorc='D' and vtype='P' order by vdate,sno")
               Set rsj = con.Execute("select * from treport where " & stringyear & " and convert(smalldatetime,vdate,103)=convert(smalldatetime,'" + Trim(prevdate) + "',103) and vtype='J' and  dorc='D' order by vdate,sno")
               Set rsjC = con.Execute("select * from treport where " & stringyear & " and convert(smalldatetime,vdate,103)=convert(smalldatetime,'" + Trim(prevdate) + "',103) and vtype='J' and dorc='C' order by vdate,sno")
               Set rsi = con.Execute("select * from treport where " & stringyear & " and convert(smalldatetime,vdate,103)=convert(smalldatetime,'" + Trim(prevdate) + "',103) and vtype='S' and dorc='C' order by vdate,sno")

               
               Do While Not rsr.EOF And Not rsp.EOF
printagain6:
                   Line = Line + 2
                   If Line > MaxLine - 5 Then
                      Line = Line - 2
                      VARLABEL = 6
                      GoTo header
                   End If
                   Print #1, Tab(2); rsr!vtype; Tab(7); rsr!vno; Tab(T2 - 3); IIf(IsNull(rsr!SUBLEDGER) Or rsr!SUBLEDGER = "", rsr!Genledger, rsr!SUBLEDGER); Tab(5 + 72); rsp!vtype; Tab(11 + 72); rsp!vno; Tab(T2 + 72); IIf(IsNull(rsp!SUBLEDGER) Or rsp!SUBLEDGER = "", rsp!Genledger, rsp!SUBLEDGER)
                   Print #1, Tab(T2 - 3); Left(rsr!narration, 35); Tab(T5 - 9); rsets(Trim(Format(Str(rsr!ac), "0.00")), 12); Tab(T2 + 72); Left(rsp!narration, 35); Tab(T5 + 62); rsets(Trim(Format(Str(rsp!ad), "0.00")), 12) '      Format(rsp!aD, "0.00")
                   
                   id = id + 1
                   con.Execute "insert into tempCash(dates,R1,R2,R4,P1,P2,P4,id) values('" & rsr!vdate & "','" & rsr!vtype & " " & rsr!vno & "','" & IIf(IsNull(rsr!SUBLEDGER) Or rsr!SUBLEDGER = "", rsr!Genledger, rsr!SUBLEDGER) & Chr(13) & rsr!narration & "','" & rsr!ac & "','" & rsp!vtype & " " & rsp!vno & "','" & IIf(IsNull(rsp!SUBLEDGER) Or rsp!SUBLEDGER = "", rsp!Genledger, rsp!SUBLEDGER) & Chr(13) & rsp!narration & "','" & rsp!ad & "'," & id & ")"
                  
                   prevbal = prevbal + rsr!ac
                   prevpayment = prevpayment + rsp!ad
                   If Not rsr.EOF Then
                        rsr.MoveNext
                   End If
                   If Not rsp.EOF Then
                        rsp.MoveNext
                   End If
               Loop
               
               Do While Not rsi.EOF And Not rsp.EOF
printagain11:
                        Line = Line + 2
                        If Line > MaxLine - 4 Then
                            Line = Line - 2
                           VARLABEL = 11
                           GoTo header
                        End If
                        Print #1, Tab(2); rsi!vtype; Tab(7); rsi!vno; Tab(T2 - 3); IIf(IsNull(rsi!SUBLEDGER) Or rsi!SUBLEDGER = "", rsi!Genledger, rsi!SUBLEDGER); Tab(5 + 72); rsp!vtype; Tab(11 + 72); rsp!vno; Tab(T2 + 72); IIf(IsNull(rsp!SUBLEDGER) Or rsp!SUBLEDGER = "", rsp!Genledger, rsp!SUBLEDGER)
                        Print #1, Tab(T2 - 3); Left(rsi!narration, 35); Tab(T5 - 9); rsets(Trim(Format(Str(rsi!ac), "0.00")), 12); Tab(T2 + 72); Left(rsp!narration, 35); Tab(T5 + 62); rsets(Trim(Format(Str(rsp!ad), "0.00")), 12)
                        
                        id = id + 1
                        con.Execute "insert into tempCash(dates,R1,R2,R4,P1,P2,P4,id) values('" & rsi!vdate & "','" & rsi!vtype & " " & rsi!vno & "','" & IIf(IsNull(rsi!SUBLEDGER) Or rsi!SUBLEDGER = "", rsi!Genledger, rsi!SUBLEDGER) & Chr(13) & rsi!narration & "','" & rsi!ac & "','" & rsp!vtype & " " & rsp!vno & "','" & IIf(IsNull(rsp!SUBLEDGER) Or rsp!SUBLEDGER = "", rsp!Genledger, rsp!SUBLEDGER) & Chr(13) & rsp!narration & "','" & rsp!ad & "'," & id & ")"

                        
                        prevbal = prevbal + rsi!ac
                        prevpayment = prevpayment + rsp!ad
                        If Not rsp.EOF Then
                            rsp.MoveNext
                        End If
                        If Not rsi.EOF Then
                            rsi.MoveNext
                        End If
               Loop
               
               Do While Not rsj.EOF And Not rsjC.EOF
printagain3:
                       
                   Line = Line + 2
                   If Line > MaxLine - 5 Then
                           Line = Line - 2
                           VARLABEL = 3
                           GoTo header
                   End If
                   Print #1, Tab(2); rsj!vtype; Tab(7); rsj!vno; Tab(T2 - 3); IIf(IsNull(rsj!SUBLEDGER) Or rsj!SUBLEDGER = "", rsj!Genledger, rsj!SUBLEDGER); Tab(5 + 72); rsjC!vtype; Tab(11 + 72); rsjC!vno; Tab(T2 + 72); IIf(IsNull(rsjC!SUBLEDGER) Or rsjC!SUBLEDGER = "", rsjC!Genledger, rsjC!SUBLEDGER)
                   Print #1, Tab(T2 - 3); Left(rsj!narration, 35); Tab(T5 - 9); rsets(Trim(Format(Str(rsj!ad), "0.00")), 12); Tab(T2 + 72); Left(rsjC!narration, 30); Tab(T5 + 62); rsets(Trim(Format(Str(rsjC!ac), "0.00")), 12) '      Format(rsp!aD, "0.00")
                   
                   id = id + 1
                   con.Execute "insert into tempCash(dates,R1,R2,R4,P1,P2,P4,id) values('" & rsj!vdate & "','" & rsj!vtype & " " & rsj!vno & "','" & IIf(IsNull(rsj!SUBLEDGER) Or rsj!SUBLEDGER = "", rsj!Genledger, rsj!SUBLEDGER) & Chr(13) & rsj!narration & "','" & rsjC!ac & "','" & rsjC!vtype & " " & rsjC!vno & "','" & IIf(IsNull(rsjC!SUBLEDGER) Or rsjC!SUBLEDGER = "", rsjC!Genledger, rsjC!SUBLEDGER) & Chr(13) & rsjC!narration & "','" & rsjC!ad & "'," & id & ")"
 
                   prevbal = prevbal + rsj!ad
                   prevpayment = prevpayment + rsjC!ac
                   If Not rsj.EOF Then
                        rsj.MoveNext
                   End If
                   If Not rsjC.EOF Then
                        rsjC.MoveNext
                   End If
               Loop
               
    
               
               Do While Not rsj.EOF

printagain4:
                      Line = Line + 2
                      If Line > MaxLine - 5 Then
                           Line = Line - 2
                           VARLABEL = 4
                           GoTo header
                        End If
                        Print #1, Tab(2); rsj!vtype; Tab(7); rsj!vno; Tab(T2 - 3); IIf(IsNull(rsj!SUBLEDGER) Or rsj!SUBLEDGER = "", rsj!Genledger, rsj!SUBLEDGER)
                        Print #1, Tab(T2 - 3); Left(rsj!narration, 35); Tab(T5 - 9); rsets(Trim(Format(Str(rsj!ad), "0.00")), 12)
                        
                        id = id + 1
                        con.Execute "insert into tempCash(dates,R1,R2,R4,id) values('" & rsj!vdate & "','" & rsj!vtype & " " & rsj!vno & "','" & IIf(IsNull(rsj!SUBLEDGER) Or rsj!SUBLEDGER = "", rsj!Genledger, rsj!SUBLEDGER) & Chr(13) & rsj!narration & "','" & rsj!ad & "'," & id & ")"
  
                        prevbal = prevbal + rsj!ad
                        If Not rsj.EOF Then
                            rsj.MoveNext
                        End If
               Loop
                           
               
               Do While Not rsi.EOF
printagain2:
                        Line = Line + 2
                        If Line > MaxLine - 4 Then
                            Line = Line - 2
                           VARLABEL = 2
                           GoTo header
                        End If
                        Print #1, Tab(2); rsi!vtype; Tab(7); rsi!vno; Tab(T2 - 3); IIf(IsNull(rsi!SUBLEDGER) Or rsi!SUBLEDGER = "", rsi!Genledger, rsi!SUBLEDGER)
                        Print #1, Tab(T2 - 3); Left(rsi!narration, 35); Tab(T5 - 9); rsets(Trim(Format(Str(rsi!ac), "0.00")), 12)
                        
                        id = id + 1
                        con.Execute "insert into tempCash(dates,R1,R2,R4,id) values('" & rsi!vdate & "','" & rsi!vtype & " " & rsi!vno & "','" & IIf(IsNull(rsi!SUBLEDGER) Or rsi!SUBLEDGER = "", rsi!Genledger, rsi!SUBLEDGER) & Chr(13) & rsi!narration & "','" & rsi!ac & "'," & id & ")"
                        
                        prevbal = prevbal + rsi!ac
                        If Not rsi.EOF Then
                            rsi.MoveNext
                        End If
               Loop
               
               
   

               
               Do While Not rsjC.EOF
printagain5:
                       Line = Line + 2
                       If Line > MaxLine - 5 Then
                           Line = Line - 2
                           VARLABEL = 5
                           GoTo header
                        End If
                        Print #1, Tab(5 + 72); rsjC!vtype; Tab(11 + 72); rsjC!vno; Tab(T2 + 72); IIf(IsNull(rsjC!SUBLEDGER) Or rsjC!SUBLEDGER = "", rsjC!Genledger, rsjC!SUBLEDGER)
                        Print #1, Tab(T2 + 72); Left(rsjC!narration, 35); Tab(T5 + 62); rsets(Trim(Format(Str(rsjC!ac), "0.00")), 12)
                        
                        prevpayment = prevpayment + rsjC!ac
                        If Not rsjC.EOF Then
                            rsjC.MoveNext
                        End If
               Loop
    
    
               Do While Not rsr.EOF
printagain7:
                      Line = Line + 2
                      If Line > MaxLine - 5 Then
                           Line = Line - 2
                            VARLABEL = 7
                            GoTo header
                       End If
                       Print #1, Tab(2); rsr!vtype; Tab(7); rsr!vno; Tab(T2 - 3); IIf(IsNull(rsr!SUBLEDGER) Or rsr!SUBLEDGER = "", rsr!Genledger, rsr!SUBLEDGER)
                       Print #1, Tab(T2 - 3); Left(rsr!narration, 35); Tab(T5 - 9); rsets(Trim(Format(Str(rsr!ac), "0.00")), 12)
                       
                       prevbal = prevbal + rsr!ac
                       If Not rsr.EOF Then
                             rsr.MoveNext
                       End If
               Loop
               Do While Not rsp.EOF
printagain8:
                      Line = Line + 2
                      If Line > MaxLine - 5 Then
                           Line = Line - 2
                           VARLABEL = 8
                           GoTo header
                        End If
                        Print #1, Tab(5 + 72); rsp!vtype; Tab(11 + 72); rsp!vno; Tab(T2 + 72); IIf(IsNull(rsp!SUBLEDGER) Or rsp!SUBLEDGER = "", rsp!Genledger, rsp!SUBLEDGER)
                        Print #1, Tab(T2 + 72); Left(rsp!narration, 35); Tab(T5 + 62); rsets(Trim(Format(Str(rsp!ad), "0.00")), 12)
                       
                        id = id + 1
                        con.Execute "insert into tempCash(dates,P1,P2,P4,id) values('" & rsp!vdate & "','" & rsp!vtype & " " & rsp!vno & "','" & IIf(IsNull(rsp!SUBLEDGER) Or rsp!SUBLEDGER = "", rsp!Genledger, rsp!SUBLEDGER) & Chr(13) & rsp!narration & "','" & rsp!ad & "'," & id & ")"
                       
                        prevpayment = prevpayment + rsp!ad
                        If Not rsp.EOF Then
                            rsp.MoveNext
                        End If
               Loop
                    
printagain0:
               Line = Line + 6
               If Line > MaxLine - 10 Then
                  Line = Line - 6
                  called1 = True
                  GoTo header
               End If
               Print #1, ""
               Print #1, Chr(27) + Chr(71); Tab(T2 + 69); "***** Balance C/F *****"; Tab(T5 + 63); rsets(Trim(Format(Str(prevbal - prevpayment), "0.00")), 12); Chr(27) + Chr(72)    'Format(Str(prevbal - prevpayment), "0.00")
               
               id = id + 1
               con.Execute "insert into tempCash(P2,P4,id) values('***** Balance C/F *****'," & (prevbal - prevpayment) & "," & id & ")"
               id = id + 1
               con.Execute "insert into tempCash(R3,R4,P4,id) values('Total'," & prevbal & "," & prevbal & "," & id & ")"
               
               
               Print #1, Tab(T5 - 9); "---------------"; Tab(T5 + 61); "-------------"
               Print #1, Tab(T5 - 9); rsets(Trim(Format(Str(prevbal), "0.00")), 12); Tab(T5 + 62); rsets(Trim(Format(Str(prevbal), "0.00")), 12)
               Print #1, Tab(T5 - 9); "---------------"; Tab(T5 + 61); "-------------"
               Print #1, Tab(0); repli("-", 143)
printnext:     If Not rs1.EOF Then
                  rs1.MoveNext
               End If
            Loop
printfooter:
            Do While Line < 72
               Print #1, " "
               Line = Line + 1
            Loop
            
            Dim tempdata As ADODB.Recordset
            Set tempdata = New ADODB.Recordset
            CNSetup
            tempdata.Open "setup1", con, adOpenStatic, adLockReadOnly, adCmdTable
            Pno = Pno + 1
            If called1 Then
               GoTo printnext
            End If
        End If
        Close #1
End Function

Function GenrepNoFooterNew()
    
    con.Execute "delete from tempCash"
    
    Dim id As Integer
   
    id = 1
    
    Dim called1, called2 As Boolean
    Dim MaxLine As Integer
    Dim OPENB As Double
    Dim T1, T2, T3, T4, T5, T6, T7, T8 As Integer
    Dim paperWidth As Integer
    Dim Pno As Integer
    Dim VARLABEL As Integer
    VARLABEL = -40
    paperWidth = 150
        T1 = 10
        T2 = 25
        T3 = 40
        T4 = 55
        T5 = 70
        T6 = 85
        T7 = 100
        T8 = 115
        called1 = False
        called2 = False
        Dim Line As Integer
        Dim rs1 As ADODB.Recordset
        Dim rs2 As ADODB.Recordset
        Dim kkk As ADODB.Recordset
        Set kkk = New ADODB.Recordset
        Set rs1 = New ADODB.Recordset
        Set rs2 = New ADODB.Recordset
        Open "" + VB.App.Path + "\vipin.txt" For Output As #1
        MaxLine = 72
        Pno = 0
header:
        If VARLABEL >= 0 Then
            Do While Line < 72
               Print #1, " "
               Line = Line + 1
            Loop
        Else
           If called1 = True Then
               Do While Line < 72
                    Print #1, " "
                    Line = Line + 1
               Loop
            End If
        End If
        Line = 0
        Pno = Pno + 1
        If kkk.State = 1 Then kkk.close
        CNSetup
        kkk.Open "select * from setup1", con, adOpenStatic, adLockReadOnly, adCmdText
        If Not kkk.BOF Then
            Print #1, ""
            Print #1, ""
            Print #1, Chr(27) + Chr(15) + Chr(14)
            Print #1, Tab(127); "Page No:  " & Pno
            Print #1, Tab(((paperWidth - (Len(Trim(kkk!cname)) * 2)) / 2)); Chr(27) + Chr(15) + Chr(14); Trim(kkk!cname)
            Print #1, Tab((paperWidth - (Len(Trim(kkk!add1)) * 2)) / 2); Chr(27) + Chr(15); dspace(Trim(kkk!add1))
            Line = Line + 6
        End If
        If rs2.State = 1 Then
            rs2.close
        End If
        'rs2.Open "treport", con, adOpenStatic, adLockOptimistic, adCmdTable
        rs2.Open "select top 1 * from treport where " & stringyear & " and len(header)>0 order by sno ", con, adOpenKeyset, adLockOptimistic, adCmdText
        
        xstr = rs2!Period
        Print #1, Chr(27) + Chr(14); Tab((73 - Len(Trim(rs2!header))) / 2); Trim(rs2!header); Chr(27) + Chr(15)
        Line = Line + 1
        Print #1, Tab((paperWidth - Len(Trim("Period : " + Trim(xstr)))) / 2); Trim("Period : " + Trim(xstr))
        Line = Line + 1
        Print #1, dspace("RECEIPTS"); Tab(T8 + 7); dspace("PAYMENTS")
        Print #1, repli("-", 143)
        Line = Line + 2
        Print #1, "V.TYPE"; Tab(8); "E.NO."; Tab(T2 - 3); "GEN./SUB. LEDGER ACCOUNT"; Tab(T5 + 7); "V.TYPE"; Tab(T1 + 74); "E.NO."; Tab(T2 + 72); "GEN./SUB. LEDGER ACCOUNT"
        Print #1, Tab(T2 + 5); "[  Narration  ]"; Tab(T4 + 12); "Amount"; Tab(T2 + 80); "[  Narration  ]"; Tab(T4 + 82); "Amount"; Chr(27) + Chr(72)
        Print #1, repli("-", 143)
        Line = Line + 3
        rs2.close
        If called1 Then
            called1 = False
            GoTo printagain0
        End If
        If VARLABEL = 1 Then
            GoTo printagain1
        End If
        If VARLABEL = 2 Then
            GoTo printagain2
        End If
        If VARLABEL = 3 Then
            GoTo printagain3
        End If
        If VARLABEL = 4 Then
            GoTo printagain4
        End If
        If VARLABEL = 5 Then
            GoTo printagain5
        End If
        If VARLABEL = 6 Then
            GoTo printagain6
        End If
        If VARLABEL = 7 Then
            GoTo printagain7
        End If
        If VARLABEL = 8 Then
            GoTo printagain8
        End If
        If VARLABEL = 11 Then
            GoTo printagain11
        End If
        rs1.Open "select top 1 * from treport order by vdate,sno", con, adOpenStatic, adLockReadOnly, adCmdText
        If Not rs1.BOF Then
        
        
            Do While Not rs1.EOF
            
            If Left$(Trim(rs1!Text), 10) = Trim("** Opening") Then
                  If rs1!ad > 0 Then
                    OPENB = rs1!ad
                  Else
                    
                    If Not IsNull(rs1!ac) Then
                     OPENB = -rs1!ac
                     Else
                     OPENB = 0
                    End If
                     
                  End If
                    Exit Do
                End If
                rs1.MoveNext

                
            Loop
        End If
        rs1.close
        rs1.Open "select Distinct vdate from treport order by vdate", con, adOpenStatic, adLockReadOnly, adCmdText
        If rs1.RecordCount > 0 Then
           rs1.MoveNext
        End If
        If Not rs1.BOF Then
            Dim prevdate As String
            Dim prevbal As Double
            Dim prevpayment As Double
            Dim rsr As ADODB.Recordset
            Dim rsp As ADODB.Recordset
            Dim rsj As ADODB.Recordset
            Dim rsjC As ADODB.Recordset
            Set rsr = New ADODB.Recordset
            Set rsp = New ADODB.Recordset
            Set rsj = New ADODB.Recordset
            Set rsjC = New ADODB.Recordset
            prevdate = ""
            prevpayment = 0
            prevbal = 0
            prevbal = OPENB
            Do While Not rs1.EOF
               If Trim(rs1!vdate) <> "" Then
                    If Trim(prevdate) <> Trim(rs1!vdate) Then
printagain1:
                    Line = Line + 1
                        If Line > MaxLine - 5 Then
                           Line = Line - 1
                           VARLABEL = 1
                           GoTo header
                        End If
                        
                        
                         id = MaxSNo("tempCash", "id")
                         con.Execute "insert into tempCash(dates,R1,R2,R3,R4,id) values('" & rs1!vdate & "','-','" & "-" & "','***** Balance B/F *****'," & prevbal & "," & id & ")"

                         
                        '-----------------------------------

                        
                        prevdate = Trim(rs1!vdate)
                        prevbal = prevbal - prevpayment
                        Print #1, Chr(27) + Chr(71); "DATE:  "; rs1!vdate; "      ***** Balance B/F *****"; Tab(T4 + 8); rsets(Trim(Format(Str(prevbal), "0.00")), 12); Chr(27) + Chr(72)
                        If CASHBOOK.CheckCash = True Then
                             If prevbal < 0 Then
                                      MsgBox "Amount Going In Credit..." & Chr(13) & "Please Check  Amount for Date :" & rs1!vdate
                                      Close #1
                                      CASHBOOK.CheckCash = False
                                      Exit Function
                             End If
                         
                       End If
                       prevpayment = 0
                    End If
               End If
               
               'Set rsr = con.Execute("select * from treport where vdate=cdate('" + Trim(prevdate) + "') and vtype = 'R' and dorc='C' order by vdate,sno")
              ' Set rsp = con.Execute("select * from treport where vdate=cdate('" + Trim(prevdate) + "') and vtype = 'P' and dorc='D' order by vdate,sno")
              ' Set rsj = con.Execute("select * from treport where vdate=cdate('" + Trim(prevdate) + "') and vtype = 'J' and dorc='D' order by vdate,sno")
             '  Set rsjC = con.Execute("select * from treport where vdate=cdate('" + Trim(prevdate) + "')and vtype = 'J' and dorc='C' order by vdate,sno")
             '  Set rsi = con.Execute("select * from treport where vdate=cdate('" + Trim(prevdate) + "') and vtype = 'S' and dorc='C' order by vdate,sno")
               
                Set rsr = con.Execute("select * from treport where " & stringyear & " and convert(smalldatetime,vdate,103)=convert(smalldatetime,'" + Trim(prevdate) + "',103) and  dorc='C' and vtype ='R' order by vdate,sno")
               Set rsp = con.Execute("select * from treport where " & stringyear & " and convert(smalldatetime,vdate,103)=convert(smalldatetime,'" + Trim(prevdate) + "',103) and  dorc='D' and vtype='P' order by vdate,sno")
               Set rsj = con.Execute("select * from treport where " & stringyear & " and convert(smalldatetime,vdate,103)=convert(smalldatetime,'" + Trim(prevdate) + "',103) and vtype='J' and  dorc='D' order by vdate,sno")
               Set rsjC = con.Execute("select * from treport where " & stringyear & " and convert(smalldatetime,vdate,103)=convert(smalldatetime,'" + Trim(prevdate) + "',103) and vtype='J' and dorc='C' order by vdate,sno")
               Set rsi = con.Execute("select * from treport where " & stringyear & " and convert(smalldatetime,vdate,103)=convert(smalldatetime,'" + Trim(prevdate) + "',103) and vtype='S' and dorc='C' order by vdate,sno")

               
               
               Do While Not rsr.EOF And Not rsp.EOF
printagain6:
                   Line = Line + 2
                   If Line > MaxLine - 5 Then
                      Line = Line - 2
                      VARLABEL = 6
                      GoTo header
                   End If
                   
                   

                   
                   Print #1, Tab(2); rsr!vtype; Tab(7); rsr!vno; Tab(T2 - 3); IIf(IsNull(rsr!SUBLEDGER) Or rsr!SUBLEDGER = "", rsr!Genledger, rsr!SUBLEDGER); Tab(5 + 72); rsp!vtype; Tab(11 + 72); rsp!vno; Tab(T2 + 72); IIf(IsNull(rsp!SUBLEDGER) Or rsp!SUBLEDGER = "", rsp!Genledger, rsp!SUBLEDGER)
                   Print #1, Tab(T2 - 3); Left(rsr!narration, 35); Tab(T5 - 9); rsets(Trim(Format(Str(rsr!ac), "0.00")), 12); Tab(T2 + 72); Left(rsp!narration, 35); Tab(T5 + 62); rsets(Trim(Format(Str(rsp!ad), "0.00")), 12) '      Format(rsp!aD, "0.00")
                   
                   id = id + 1
                   con.Execute "insert into tempCash(dates,R1,R2,R4,P1,P2,P4,id) values('" & rsr!vdate & "','" & rsr!vtype & " " & rsr!vno & "','" & IIf(IsNull(rsr!SUBLEDGER) Or rsr!SUBLEDGER = "", rsr!Genledger, rsr!SUBLEDGER) & Chr(13) & rsr!narration & "','" & rsr!ac & "','" & rsp!vtype & " " & rsp!vno & "','" & IIf(IsNull(rsp!SUBLEDGER) Or rsp!SUBLEDGER = "", rsp!Genledger, rsp!SUBLEDGER) & Chr(13) & rsp!narration & "','" & rsp!ad & "'," & id & ")"

                   
                   prevbal = prevbal + rsr!ac
                   prevpayment = prevpayment + rsp!ad
                   If Not rsr.EOF Then
                        rsr.MoveNext
                   End If
                   If Not rsp.EOF Then
                        rsp.MoveNext
                   End If
               Loop
               
               Do While Not rsi.EOF And Not rsp.EOF
printagain11:
                        Line = Line + 2
                        If Line > MaxLine - 4 Then
                            Line = Line - 2
                           VARLABEL = 11
                           GoTo header
                        End If
                        
                        Print #1, Tab(2); rsi!vtype; Tab(7); rsi!vno; Tab(T2 - 3); IIf(IsNull(rsi!SUBLEDGER) Or rsi!SUBLEDGER = "", rsi!Genledger, rsi!SUBLEDGER); Tab(5 + 72); rsp!vtype; Tab(11 + 72); rsp!vno; Tab(T2 + 72); IIf(IsNull(rsp!SUBLEDGER) Or rsp!SUBLEDGER = "", rsp!Genledger, rsp!SUBLEDGER)
                        Print #1, Tab(T2 - 3); Left(rsi!narration, 35); Tab(T5 - 9); rsets(Trim(Format(Str(rsi!ac), "0.00")), 12); Tab(T2 + 72); Left(rsp!narration, 35); Tab(T5 + 62); rsets(Trim(Format(Str(rsp!ad), "0.00")), 12)
                        
                        id = id + 1
                        con.Execute "insert into tempCash(dates,R1,R2,R4,P1,P2,P4,id) values('" & rsi!vdate & "','" & rsi!vtype & " " & rsi!vno & "','" & IIf(IsNull(rsi!SUBLEDGER) Or rsi!SUBLEDGER = "", rsi!Genledger, rsi!SUBLEDGER) & Chr(13) & rsi!narration & "','" & rsi!ac & "','" & rsp!vtype & " " & rsp!vno & "','" & IIf(IsNull(rsp!SUBLEDGER) Or rsp!SUBLEDGER = "", rsp!Genledger, rsp!SUBLEDGER) & Chr(13) & rsp!narration & "','" & rsp!ad & "'," & id & ")"


                        
                        prevbal = prevbal + rsi!ac
                        prevpayment = prevpayment + rsp!ad
                        If Not rsp.EOF Then
                            rsp.MoveNext
                        End If
                        If Not rsi.EOF Then
                            rsi.MoveNext
                        End If
               Loop
               
               Do While Not rsj.EOF And Not rsjC.EOF
printagain3:
                       
                   Line = Line + 2
                   If Line > MaxLine - 5 Then
                           Line = Line - 2
                           VARLABEL = 3
                           GoTo header
                   End If
                   Print #1, Tab(2); rsj!vtype; Tab(7); rsj!vno; Tab(T2 - 3); IIf(IsNull(rsj!SUBLEDGER) Or rsj!SUBLEDGER = "", rsj!Genledger, rsj!SUBLEDGER); Tab(5 + 72); rsjC!vtype; Tab(11 + 72); rsjC!vno; Tab(T2 + 72); IIf(IsNull(rsjC!SUBLEDGER) Or rsjC!SUBLEDGER = "", rsjC!Genledger, rsjC!SUBLEDGER)
                   Print #1, Tab(T2 - 3); Left(rsj!narration, 35); Tab(T5 - 9); rsets(Trim(Format(Str(rsj!ad), "0.00")), 12); Tab(T2 + 72); Left(rsjC!narration, 30); Tab(T5 + 62); rsets(Trim(Format(Str(rsjC!ac), "0.00")), 12) '      Format(rsp!aD, "0.00")
                   
                    id = id + 1
                   con.Execute "insert into tempCash(dates,R1,R2,R4,P1,P2,P4,id) values('" & rsj!vdate & "','" & rsj!vtype & " " & rsj!vno & "','" & IIf(IsNull(rsj!SUBLEDGER) Or rsj!SUBLEDGER = "", rsj!Genledger, rsj!SUBLEDGER) & Chr(13) & rsj!narration & "','" & rsj!ad & "','" & rsjC!vtype & " " & rsjC!vno & "','" & IIf(IsNull(rsjC!SUBLEDGER) Or rsjC!SUBLEDGER = "", rsjC!Genledger, rsjC!SUBLEDGER) & Chr(13) & rsjC!narration & "','" & rsjC!ac & "'," & id & ")"

                   
                   prevbal = prevbal + rsj!ad
                   prevpayment = prevpayment + rsjC!ac
                   If Not rsj.EOF Then
                        rsj.MoveNext
                   End If
                   If Not rsjC.EOF Then
                        rsjC.MoveNext
                   End If
               Loop
               
    
               
               Do While Not rsj.EOF

printagain4:
                      Line = Line + 2
                      If Line > MaxLine - 5 Then
                           Line = Line - 2
                           VARLABEL = 4
                           GoTo header
                        End If
                        Print #1, Tab(2); rsj!vtype; Tab(7); rsj!vno; Tab(T2 - 3); IIf(IsNull(rsj!SUBLEDGER) Or rsj!SUBLEDGER = "", rsj!Genledger, rsj!SUBLEDGER)
                        Print #1, Tab(T2 - 3); Left(rsj!narration, 35); Tab(T5 - 9); rsets(Trim(Format(Str(rsj!ad), "0.00")), 12)
                        
                        id = id + 1
                        con.Execute "insert into tempCash(dates,R1,R2,R4,id) values('" & rsj!vdate & "','" & rsj!vtype & " " & rsj!vno & "','" & IIf(IsNull(rsj!SUBLEDGER) Or rsj!SUBLEDGER = "", rsj!Genledger, rsj!SUBLEDGER) & Chr(13) & rsj!narration & "','" & rsj!ad & "'," & id & ")"

                        
                        prevbal = prevbal + rsj!ad
                        If Not rsj.EOF Then
                            rsj.MoveNext
                        End If
               Loop
                           
               
               Do While Not rsi.EOF
printagain2:
                        Line = Line + 2
                        If Line > MaxLine - 4 Then
                            Line = Line - 2
                           VARLABEL = 2
                           GoTo header
                        End If
                        
 
                        
                        Print #1, Tab(2); rsi!vtype; Tab(7); rsi!vno; Tab(T2 - 3); IIf(IsNull(rsi!SUBLEDGER) Or rsi!SUBLEDGER = "", rsi!Genledger, rsi!SUBLEDGER)
                        Print #1, Tab(T2 - 3); Left(rsi!narration, 35); Tab(T5 - 9); rsets(Trim(Format(Str(rsi!ac), "0.00")), 12)
                        
                        id = id + 1
                        con.Execute "insert into tempCash(dates,R1,R2,R4,id) values('" & rsi!vdate & "','" & rsi!vtype & " " & rsi!vno & "','" & IIf(IsNull(rsi!SUBLEDGER) Or rsi!SUBLEDGER = "", rsi!Genledger, rsi!SUBLEDGER) & Chr(13) & rsi!narration & "','" & rsi!ac & "'," & id & ")"
 
                        
                        prevbal = prevbal + rsi!ac
                        If Not rsi.EOF Then
                            rsi.MoveNext
                        End If
               Loop
               
               
   

               
               Do While Not rsjC.EOF
printagain5:
                       Line = Line + 2
                       If Line > MaxLine - 5 Then
                           Line = Line - 2
                           VARLABEL = 5
                           GoTo header
                        End If
                        Print #1, Tab(5 + 72); rsjC!vtype; Tab(11 + 72); rsjC!vno; Tab(T2 + 72); IIf(IsNull(rsjC!SUBLEDGER) Or rsjC!SUBLEDGER = "", rsjC!Genledger, rsjC!SUBLEDGER)
                        Print #1, Tab(T2 + 72); Left(rsjC!narration, 35); Tab(T5 + 62); rsets(Trim(Format(Str(rsjC!ac), "0.00")), 12)
                        
                        id = id + 1
                        con.Execute "insert into tempCash(dates,P1,P2,P4,id) values('" & rsjC!vdate & "','" & rsjC!vtype & " " & rsjC!vno & "','" & IIf(IsNull(rsjC!SUBLEDGER) Or rsjC!SUBLEDGER = "", rsjC!Genledger, rsjC!SUBLEDGER) & Chr(13) & rsjC!narration & "','" & rsjC!ac & "'," & id & ")"

                        
                        prevpayment = prevpayment + rsjC!ac
                        If Not rsjC.EOF Then
                            rsjC.MoveNext
                        End If
               Loop
    
    
               Do While Not rsr.EOF
printagain7:
                      Line = Line + 2
                      If Line > MaxLine - 5 Then
                           Line = Line - 2
                            VARLABEL = 7
                            GoTo header
                       End If
                       Print #1, Tab(2); rsr!vtype; Tab(7); rsr!vno; Tab(T2 - 3); IIf(IsNull(rsr!SUBLEDGER) Or rsr!SUBLEDGER = "", rsr!Genledger, rsr!SUBLEDGER)
                       Print #1, Tab(T2 - 3); Left(rsr!narration, 35); Tab(T5 - 9); rsets(Trim(Format(Str(rsr!ac), "0.00")), 12)
                       
                       id = id + 1
                       con.Execute "insert into tempCash(dates,r1,r2,r4,id) values('" & rsr!vdate & "','" & rsr!vtype & " " & rsr!vno & "','" & IIf(IsNull(rsr!SUBLEDGER) Or rsr!SUBLEDGER = "", rsr!Genledger, rsr!SUBLEDGER) & Chr(13) & rsr!narration & "','" & rsr!ac & "'," & id & ")"

                       
                       
                       prevbal = prevbal + rsr!ac
                       If Not rsr.EOF Then
                             rsr.MoveNext
                       End If
               Loop
               Do While Not rsp.EOF
printagain8:
                      Line = Line + 2
                      If Line > MaxLine - 5 Then
                           Line = Line - 2
                           VARLABEL = 8
                           GoTo header
                        End If
                        Print #1, Tab(5 + 72); rsp!vtype; Tab(11 + 72); rsp!vno; Tab(T2 + 72); IIf(IsNull(rsp!SUBLEDGER) Or rsp!SUBLEDGER = "", rsp!Genledger, rsp!SUBLEDGER)
                        Print #1, Tab(T2 + 72); Left(rsp!narration, 35); Tab(T5 + 62); rsets(Trim(Format(Str(rsp!ad), "0.00")), 12)
                        
                       id = id + 1
                       con.Execute "insert into tempCash(dates,P1,P2,P4,id) values('" & rsp!vdate & "','" & rsp!vtype & " " & rsp!vno & "','" & IIf(IsNull(rsp!SUBLEDGER) Or rsp!SUBLEDGER = "", rsp!Genledger, rsp!SUBLEDGER) & Chr(13) & rsp!narration & "','" & rsp!ad & "'," & id & ")"

                        
                        prevpayment = prevpayment + rsp!ad
                        If Not rsp.EOF Then
                            rsp.MoveNext
                        End If
               Loop
                    
printagain0:
               Line = Line + 6
               If Line > MaxLine - 10 Then
                  Line = Line - 6
                  called1 = True
                  GoTo header
               End If
               Print #1, ""
               Print #1, Chr(27) + Chr(71); Tab(T2 + 69); "***** Balance C/F *****"; Tab(T5 + 63); rsets(Trim(Format(Str(prevbal - prevpayment), "0.00")), 12); Chr(27) + Chr(72)    'Format(Str(prevbal - prevpayment), "0.00")
               
               id = id + 1
               con.Execute "insert into tempCash(P2,P4,id) values('***** Balance C/F *****'," & (prevbal - prevpayment) & "," & id & ")"
               id = id + 1
               con.Execute "insert into tempCash(R3,R4,P4,id) values('Total'," & prevbal & "," & prevbal & "," & id & ")"
 
               
               Print #1, Tab(T5 - 9); "---------------"; Tab(T5 + 61); "-------------"
               Print #1, Tab(T5 - 9); rsets(Trim(Format(Str(prevbal), "0.00")), 12); Tab(T5 + 62); rsets(Trim(Format(Str(prevbal), "0.00")), 12)
               Print #1, Tab(T5 - 9); "---------------"; Tab(T5 + 61); "-------------"
               Print #1, Tab(0); repli("-", 143)
printnext:     If Not rs1.EOF Then
                  rs1.MoveNext
               End If
            Loop
printfooter:
            Do While Line < 72
               Print #1, " "
               Line = Line + 1
            Loop
            
            Dim tempdata As ADODB.Recordset
            Set tempdata = New ADODB.Recordset
            CNSetup
            tempdata.Open "setup1", con, adOpenStatic, adLockReadOnly, adCmdTable
            Pno = Pno + 1
            If called1 Then
               GoTo printnext
            End If
        End If
        Close #1
End Function

Function GenrepNoFooter()
    Dim called1, called2 As Boolean
    Dim MaxLine As Integer
    Dim OPENB As Double
    Dim T1, T2, T3, T4, T5, T6, T7, T8 As Integer
    Dim paperWidth As Integer
    Dim Pno As Integer
    Dim VARLABEL As Integer
    Dim pb As Integer
    Dim v1 As Integer
    
    pb = 0
    v1 = 0
    
    
    VARLABEL = -40
    paperWidth = 150
        T1 = 10
        T2 = 25
        T3 = 40
        T4 = 55
        T5 = 70
        T6 = 85
        T7 = 100
        T8 = 115
        called1 = False
        called2 = False
        Dim Line As Integer
        Dim rs1 As ADODB.Recordset
        Dim rs2 As ADODB.Recordset
        Dim kkk As ADODB.Recordset
        Dim add1, cname As String
        Dim rs5 As ADODB.Recordset
        
        Set kkk = New ADODB.Recordset
        Set rs1 = New ADODB.Recordset
        Set rs2 = New ADODB.Recordset
        Set rs5 = New ADODB.Recordset
        
        If kkk.State = 1 Then kkk.close
        kkk.Open "Select * from setup1 where " & stringyear & "", con, adOpenKeyset, adLockReadOnly
        If Not kkk.BOF Then
           cname = dspace(kkk!cname)
           add1 = dspace(kkk!add1)
        End If
        
        DoEvents
        DoEvents
        CASHBOOK.pb.Visible = True
         
        
        
        Open "" + App.Path + "\vipin.txt" For Output As #1
        MaxLine = 72
        Pno = 0
        Dim d1 As Integer
        d1 = 0
header:
        If VARLABEL >= 0 Then
            Do While Line < 72
               Print #1, " "
               Line = Line + 1
            Loop
        Else
           If called1 = True Then
               Do While Line < 72
                    Print #1, " "
                    Line = Line + 1
                    
               Loop
            End If
        End If
        Line = 0
        d1 = d1 + 1
        Pno = Pno + 1
        If kkk.State = 1 Then kkk.close
        kkk.Open "Select * from setup1 where " & stringyear & "", con, adOpenKeyset, adLockReadOnly
        If Not kkk.BOF Then
            Print #1, ""
            Print #1, ""
            Print #1, Chr(27) + Chr(15) + Chr(14)
            Print #1, Tab(127); "Page No:  " & Pno
            Print #1, Tab(((paperWidth - (Len(Trim(kkk!cname)) * 2)) / 2) - 15); Chr(27) + Chr(15) + Chr(14); cname  'dspace(Trim(kkk!cname))
            'Print #1, Tab((paperWidth - (Len(Trim(kkk!add1)) * 2)) / 2); Chr(27) + Chr(15); 'add1
            Print #1, Tab((paperWidth - (Len(Trim(kkk!add1)) * 2)) / 2); Chr(27) + Chr(15); dspace(Trim(kkk!add1))
            
            Line = Line + 6
        End If
        
        If rs2.State = 1 Then rs2.close
        rs2.Open "select top 1 * from treport where " & stringyear & " and len(header)>0 order by sno ", con, adOpenKeyset, adLockOptimistic, adCmdText
        xstr = rs2!Period
        Print #1, Chr(27) + Chr(14); Tab((73 - Len(Trim(rs2!header))) / 2); Trim(rs2!header); Chr(27) + Chr(15)
        Line = Line + 1
        
        Print #1, Tab((paperWidth - Len(Trim("Period : " + Trim(xstr)))) / 2); Trim("Period : " + Trim(xstr))
        Line = Line + 1
        
        Print #1, "R E C E I P T S"; Tab(T8 + 7); "P A Y M E N T S"
        Print #1, repli("-", 143)
        Line = Line + 2
        
        Print #1, "V.TYPE"; Tab(8); "E.NO."; Tab(T2 - 3); "GEN./SUB. LEDGER ACCOUNT"; Tab(T5 + 7); "V.TYPE"; Tab(T1 + 74); "E.NO."; Tab(T2 + 72); "GEN./SUB. LEDGER ACCOUNT"
        Print #1, Tab(T2 + 5); "[  Narration  ]"; Tab(T4 + 12); "Amount"; Tab(T2 + 80); "[  Narration  ]"; Tab(T4 + 82); "Amount"; Chr(27) + Chr(72)
        Print #1, repli("-", 143)
        Line = Line + 3
        
        rs2.close
        If called1 Then
            called1 = False
            GoTo printagain0
        End If
        If VARLABEL = 1 Then
            GoTo printagain1
        End If
        If VARLABEL = 2 Then
            GoTo printagain2
        End If
        If VARLABEL = 3 Then
            GoTo printagain3
        End If
        If VARLABEL = 14 Then
            GoTo printagain4
        End If
        If VARLABEL = 5 Then
            GoTo printagain5
        End If
        If VARLABEL = 6 Then
            GoTo printagain6
        End If
        If VARLABEL = 7 Then
            GoTo printagain7
        End If
        If VARLABEL = 8 Then
            GoTo printagain8
        End If
        
        If rs1.State = 1 Then rs1.close
        rs1.Open "select top 1 * from treport where " & stringyear & " order by vdate,sno", con, adOpenKeyset, adLockReadOnly
        If Not rs1.BOF Then
            Do While Not rs1.EOF
                If Left$(Trim(rs1!Text), 10) = Trim("** Opening") Then
                    OPENB = rs1!ad
                    Exit Do
                End If
                rs1.MoveNext
            Loop
        End If
        
        
      If v1 = 0 Then
        If rs5.State = 1 Then rs5.close
        rs5.Open "select Distinct vdate from treport where " & stringyear & " order by vdate", con, adOpenStatic, adLockReadOnly
        If rs5.RecordCount > 0 Then
        If pb = 0 Then
           CASHBOOK.pb.Max = rs5.RecordCount
           pb = 1
        End If
        rs5.MoveNext
        v1 = 1
        End If
      Else
        rs5.MoveFirst
        rs5.MoveNext
      End If
        
       
       If Not rs1.BOF Then
            Dim prevdate As String
            Dim prevbal As Double
            Dim prevpayment As Double
            Dim rsr As ADODB.Recordset
            Dim rsp As ADODB.Recordset
            Dim rsj As ADODB.Recordset
            Dim rsjC As ADODB.Recordset
            Set rsr = New ADODB.Recordset
            Set rsp = New ADODB.Recordset
            Set rsj = New ADODB.Recordset
            Set rsjC = New ADODB.Recordset
            prevdate = ""
            prevpayment = 0
            prevbal = 0
            prevbal = OPENB
            Do While Not rs5.EOF
               
               If CASHBOOK.pb.value = CASHBOOK.pb.Max Then
               CASHBOOK.pb.value = 0
               Else
               CASHBOOK.pb.value = CASHBOOK.pb.value + 1
               End If
               
               If Trim(rs5!vdate) <> "" Then
                    If Trim(prevdate) <> Trim(rs5!vdate) Then
printagain1:
               Line = Line + 1
                        If Line > MaxLine - 5 Then
                           Line = Line - 1
                           VARLABEL = 1
                           GoTo header
                        End If
                        prevdate = Trim(rs5!vdate)
                        prevbal = prevbal - prevpayment
                        Print #1, Chr(27) + Chr(71); "DATE:  "; rs5!vdate; "      ***** Balance B/F *****"; Tab(T4 + 8); rsets(Trim(Format(Str(prevbal), "0.00")), 12); Chr(27) + Chr(72)
                        If CASHBOOK.CheckCash = True Then
                             If prevbal < 0 Then
                                      MsgBox "Amount Going In Credit..." & Chr(13) & "Please Check  Amount for Date :" & rs5!vdate
                                      Close #1
                                      CASHBOOK.CheckCash = False
                                      Exit Function
                             End If
                         
                       End If
                       prevpayment = 0
                    End If
               End If
               
               Set rsr = con.Execute("select * from treport where " & stringyear & " and convert(smalldatetime,vdate,103)=convert(smalldatetime,'" + Trim(prevdate) + "',103) and  dorc='C' and vtype ='R' order by vdate,sno")
               Set rsp = con.Execute("select * from treport where " & stringyear & " and convert(smalldatetime,vdate,103)=convert(smalldatetime,'" + Trim(prevdate) + "',103) and  dorc='D' and vtype='P' order by vdate,sno")
               Set rsj = con.Execute("select * from treport where " & stringyear & " and convert(smalldatetime,vdate,103)=convert(smalldatetime,'" + Trim(prevdate) + "',103) and vtype='J' and  dorc='D' order by vdate,sno")
               Set rsjC = con.Execute("select * from treport where " & stringyear & " and convert(smalldatetime,vdate,103)=convert(smalldatetime,'" + Trim(prevdate) + "',103) and vtype='J' and dorc='C' order by vdate,sno")
               Set rsi = con.Execute("select * from treport where " & stringyear & " and convert(smalldatetime,vdate,103)=convert(smalldatetime,'" + Trim(prevdate) + "',103) and vtype='S' and dorc='C' order by vdate,sno")
               Do While Not rsi.EOF
printagain2:
                        Line = Line + 2
                        If Line > MaxLine - 4 Then
                            Line = Line - 2
                           VARLABEL = 2
                           GoTo header
                        End If
                        Print #1, Tab(2); rsi!vtype; Tab(7); rsi!vno; Tab(T2 - 3); IIf((IsNull(rsi!SUBLEDGER) Or rsi!SUBLEDGER = ""), rsi!Genledger, rsi!SUBLEDGER)
                        Print #1, Tab(T2 - 3); rsi!narration; Tab(T5 - 9); rsets(Trim(Format(Str(rsi!ac), "0.00")), 12)
                        prevbal = prevbal + rsi!ac
                        If Not rsi.EOF Then
                            rsi.MoveNext
                        End If
               Loop
               Do While Not rsj.EOF And Not rsjC.EOF
printagain3:
                   Line = Line + 2
                   If Line > MaxLine - 5 Then
                           Line = Line - 2
                           VARLABEL = 3
                           GoTo header
                   End If
                   Print #1, Tab(2); rsj!vtype; Tab(7); rsj!vno; Tab(T2 - 3); IIf(IsNull(rsj!SUBLEDGER), rsj!Genledger, rsj!SUBLEDGER); Tab(5 + 72); rsjC!vtype; Tab(11 + 72); rsjC!vno; Tab(T2 + 72); IIf((IsNull(rsjC!SUBLEDGER) Or rsjC!SUBLEDGER = ""), rsjC!Genledger, rsjC!SUBLEDGER)
                   Print #1, Tab(T2 - 3); Left(rsj!narration, 35); Tab(T5 - 9); rsets(Trim(Format(Str(rsj!ad), "0.00")), 12); Tab(T2 + 72); rsjC!narration; Tab(T5 + 62); rsets(Trim(Format(Str(rsjC!ac), "0.00")), 12) '      Format(rsp!aD, "0.00")
                   prevbal = prevbal + rsj!ad
                   prevpayment = prevpayment + rsjC!ac
                   If Not rsj.EOF Then
                        rsj.MoveNext
                   End If
                   If Not rsjC.EOF Then
                        rsjC.MoveNext
                   End If
               Loop
               Do While Not rsj.EOF
printagain4:
                      Line = Line + 2
                      If Line > MaxLine - 5 Then
                           Line = Line - 2
                           VARLABEL = 4
                           GoTo header
                        End If
                        Print #1, Tab(2); rsj!vtype; Tab(7); rsj!vno; Tab(T2 - 3); IIf((IsNull(rsj!SUBLEDGER) Or rsj!SUBLEDGER = ""), rsj!Genledger, rsj!SUBLEDGER)
                        Print #1, Tab(T2 - 3); Left(rsj!narration, 35); Tab(T5 - 9); rsets(Trim(Format(Str(rsj!ad), "0.00")), 12)
                        prevbal = prevbal + rsj!ad
                        If Not rsj.EOF Then
                            rsj.MoveNext
                        End If
               Loop
               Do While Not rsjC.EOF
printagain5:
                       Line = Line + 2
                       If Line > MaxLine - 5 Then
                           Line = Line - 2
                           VARLABEL = 5
                           GoTo header
                        End If
                        Print #1, Tab(5 + 72); rsjC!vtype; Tab(11 + 72); rsjC!vno; Tab(T2 + 72); IIf((IsNull(rsjC!SUBLEDGER) Or rsjC!SUBLEDGER = ""), rsjC!Genledger, rsjC!SUBLEDGER)
                        Print #1, Tab(T2 + 72); Left(rsjC!narration, 35); Tab(T5 + 62); rsets(Trim(Format(Str(rsjC!ac), "0.00")), 12)
                        prevpayment = prevpayment + rsjC!ac
                        If Not rsjC.EOF Then
                            rsjC.MoveNext
                        End If
               Loop
               Do While Not rsr.EOF And Not rsp.EOF
printagain6:
                   Line = Line + 2
                   If Line > MaxLine - 5 Then
                      Line = Line - 2
                      VARLABEL = 6
                      GoTo header
                   End If
                   Print #1, Tab(2); rsr!vtype; Tab(7); rsr!vno; Tab(T2 - 3); IIf(IsNull(rsr!SUBLEDGER), rsr!Genledger, rsr!SUBLEDGER); Tab(5 + 72); rsp!vtype; Tab(11 + 72); rsp!vno; Tab(T2 + 72); IIf((IsNull(rsp!SUBLEDGER) Or rsp!SUBLEDGER = ""), rsp!Genledger, rsp!SUBLEDGER)
                   Print #1, Tab(T2 - 3); Left(rsr!narration, 35); Tab(T5 - 9); rsets(Trim(Format(Str(rsr!ac), "0.00")), 12); Tab(T2 + 72); rsp!narration; Tab(T5 + 62); rsets(Trim(Format(Str(rsp!ad), "0.00")), 12) '      Format(rsp!aD, "0.00")
                   prevbal = prevbal + rsr!ac
                   prevpayment = prevpayment + rsp!ad
                   If Not rsr.EOF Then
                        rsr.MoveNext
                   End If
                   If Not rsp.EOF Then
                        rsp.MoveNext
                   End If
               Loop
               
               
               
               Do While Not rsr.EOF
printagain7:
                      Line = Line + 2
                      If Line > MaxLine - 5 Then
                           Line = Line - 2
                            VARLABEL = 7
                            GoTo header
                       End If
                       Print #1, Tab(2); rsr!vtype; Tab(7); rsr!vno; Tab(T2 - 3); IIf((IsNull(rsr!SUBLEDGER) Or rsr!SUBLEDGER = ""), rsr!Genledger, rsr!SUBLEDGER)
                       Print #1, Tab(T2 - 3); Left(rsr!narration, 35); Tab(T5 - 9); rsets(Trim(Format(Str(rsr!ac), "0.00")), 12)
                       prevbal = prevbal + rsr!ac
                       If Not rsr.EOF Then
                             rsr.MoveNext
                       End If
               Loop
               Do While Not rsp.EOF
printagain8:
                      Line = Line + 2
                      If Line > MaxLine - 5 Then
                           Line = Line - 2
                           VARLABEL = 8
                           GoTo header
                        End If
                        Print #1, Tab(5 + 72); rsp!vtype; Tab(11 + 72); rsp!vno; Tab(T2 + 72); IIf((IsNull(rsp!SUBLEDGER) Or rsp!SUBLEDGER = ""), rsp!Genledger, rsp!SUBLEDGER)
                        Print #1, Tab(T2 + 72); Left(rsp!narration, 35); Tab(T5 + 62); rsets(Trim(Format(Str(rsp!ad), "0.00")), 12)
                        prevpayment = prevpayment + rsp!ad
                        If Not rsp.EOF Then
                            rsp.MoveNext
                        End If
               Loop
               
printagain0:
               Line = Line + 6
               If Line > MaxLine - 10 Then
                  Line = Line - 6
                  called1 = True
                  GoTo header
               End If
               Print #1, ""
               Print #1, Chr(27) + Chr(71); Tab(T2 + 69); "***** Balance C/F *****"; Tab(T5 + 63); rsets(Trim(Format(Str(prevbal - prevpayment), "0.00")), 12); Chr(27) + Chr(72)    'Format(Str(prevbal - prevpayment), "0.00")
               Print #1, Tab(T5 - 9); "---------------"; Tab(T5 + 61); "-------------"
               Print #1, Tab(T5 - 9); rsets(Trim(Format(Str(prevbal), "0.00")), 12); Tab(T5 + 62); rsets(Trim(Format(Str(prevbal), "0.00")), 12)
               Print #1, Tab(T5 - 9); "---------------"; Tab(T5 + 61); "-------------"
               Print #1, Tab(0); repli("-", 143)
printnext:     If Not rs5.EOF Then
                     rs5.MoveNext
               End If
            Loop
printfooter:
            Do While Line < 72
               Print #1, " "
               Line = Line + 1
            Loop
            
            Dim tempdata As ADODB.Recordset
            Set tempdata = New ADODB.Recordset
            CNSetup
            tempdata.Open "Select * from setup1 where " & stringyear & "", con, adOpenKeyset, adLockReadOnly
            Pno = Pno + 1
            If called1 Then
               GoTo printnext
            End If
        End If
        Close #1
End Function

Private Sub r1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
