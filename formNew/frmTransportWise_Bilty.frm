VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmTransportWise_Bilty 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5928
   ClientLeft      =   2796
   ClientTop       =   1620
   ClientWidth     =   8652
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5928
   ScaleWidth      =   8652
   Begin VB.CommandButton cmdServiceTax 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Print Service Tax"
      Height          =   585
      Left            =   3675
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4824
      Width           =   1830
   End
   Begin Crystal.CrystalReport cr 
      Left            =   375
      Top             =   4995
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComDlg.CommonDialog ComD 
      Left            =   75
      Top             =   4170
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Select Path"
      Height          =   465
      Left            =   9825
      TabIndex        =   9
      Top             =   870
      Width           =   690
   End
   Begin VB.TextBox txtPath 
      Height          =   390
      Left            =   9375
      TabIndex        =   8
      Top             =   870
      Width           =   390
   End
   Begin VB.CommandButton Commandreturn 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Return"
      Height          =   585
      Left            =   5625
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4824
      Width           =   1980
   End
   Begin VB.CommandButton Commandshow 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Print Transport Wise"
      Height          =   585
      Left            =   1725
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4824
      Width           =   1830
   End
   Begin VB.ListBox Glist 
      Height          =   3720
      Left            =   1728
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   780
      Width           =   3768
   End
   Begin MSMask.MaskEdBox date1 
      Height          =   345
      Left            =   1665
      TabIndex        =   3
      Top             =   270
      Width           =   1065
      _ExtentX        =   1884
      _ExtentY        =   593
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox date2 
      Height          =   345
      Left            =   3855
      TabIndex        =   4
      Top             =   270
      Width           =   1215
      _ExtentX        =   2159
      _ExtentY        =   593
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Transport  :"
      Height          =   285
      Left            =   225
      TabIndex        =   7
      Top             =   780
      Width           =   1185
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      Height          =   315
      Left            =   3015
      TabIndex        =   6
      Top             =   285
      Width           =   465
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "From The Date"
      Height          =   315
      Left            =   225
      TabIndex        =   5
      Top             =   300
      Width           =   1305
   End
End
Attribute VB_Name = "frmTransportWise_Bilty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public GWFLAG As Boolean
Dim SQLSTRING As String
Dim s1 As String
Dim con_conven As New ADODB.Connection

Private Sub Command1_Click()
    Unload Me
    ''MainMenu.Toolbar1.Visible = True
End Sub

Private Sub cmdsalesreturn_Click()
If DateDiff("d", Trim(date1.Text), Trim(date2.Text)) <= 0 Then
    MsgBox "invalid date"
    Exit Sub
End If
TRFLAG = False
genreport
PrintOption.Show
End Sub

Private Sub cmdAdd_Click()
'''ComD.ShowOpen
'''txtPath = ComD.FileName
'''
'''If txtPath <> "" Then
'''   CON.Execute "delete from conpath where " & stringyear & " and len(path)>0"
'''   CON.Execute "insert into conpath values('" & txtPath & "')"
'''End If

End Sub

Private Sub cmdServiceTax_Click()
s1 = 2
Report_coding

End Sub

Private Sub CommandReturn_Click()
''MainMenu.Toolbar1.Visible = True
Unload Me
End Sub
Private Sub Commandshow_Click()

s1 = 1
Report_coding

End Sub
Sub Report_coding()


DSNNew


trans = ""
If trans = "" Then
For I = 0 To Glist.ListCount - 1
If Glist.Selected(I) = True Then
    If trans = "" Then
       trans = "transportname = '" & Glist.List(I) & "'"
    Else
       trans = trans & " or " & "transportname = '" & Glist.List(I) & "'"
    End If
End If
Next
End If

If trans <> "" Then
   trans = "(" & trans & ")"
End If



'===============Date Comes From InvoiceA

con.Execute "delete from TransPortRpt"

If trans = "" Then
    con.Execute "insert into  TransPortRpt(Party,Bilty,Dates,Amount_Text,Amount,Vtype,setupid,fyear,inv)" & _
    " SELECT transportname,BILTYNO,BILTYDATE,FREIGHT,0,'I',setupid,fyear,invoiceno FROM INVOICEA where (len(transportname)>0 and convert(smalldatetime,biltydate,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,biltydate,103) <=convert(smalldatetime,'" + Trim(date2.Text) + "',103))"
Else
    con.Execute "insert into  TransPortRpt(Party,Bilty,Dates,Amount_Text,Amount,Vtype,setupid,fyear,inv)" & _
    " SELECT transportname,BILTYNO,BILTYDATE,FREIGHT,0,'I',setupid,fyear,invoiceno FROM INVOICEA where (len(transportname)>0 and convert(smalldatetime,biltydate,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,biltydate,103) <=convert(smalldatetime,'" + Trim(date2.Text) + "',103) and " & trans & ")"
End If

'===============Date Comes From CashA
If trans = "" Then
    con.Execute "insert into  TransPortRpt(Party,Bilty,Dates,Amount_Text,Amount,Vtype,setupid,fyear)" & _
    " SELECT transportname,BILTYNO,BILTYDATE,FREIGHT,0,'C',setupid,fyear FROM CashA where " & stringyear & " and (len(transportname)>0) and convert(smalldatetime,biltydate,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,biltydate,103) <=convert(smalldatetime,'" + Trim(date2.Text) + "',103)"
Else
    con.Execute "insert into  TransPortRpt(Party,Bilty,Dates,Amount_Text,Amount,Vtype,setupid,fyear)" & _
    " SELECT transportname,BILTYNO,BILTYDATE,FREIGHT,0,'C',setupid,fyear FROM CashA where " & stringyear & " and (len(transportname)>0) and convert(smalldatetime,biltydate,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,biltydate,103) <=convert(smalldatetime,'" + Trim(date2.Text) + "',103) and " & trans
End If






If RS.State = 1 Then RS.close
If trans = "" Then
    RS.Open "SELECT transportname,BILTYNO,BILTYDATE,FREIGHT,0,'I',invoiceno FROM INVOICEA_sp where " & stringyear & " and " & _
    "(len(transportname)>0) and convert(smalldatetime,biltydate,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,biltydate,103) <=convert(smalldatetime,'" + Trim(date2.Text) + "',103)" & _
    "", con, adOpenStatic, adLockReadOnly
Else
    RS.Open "SELECT transportname,BILTYNO,BILTYDATE,FREIGHT,0,'I',invoiceno FROM INVOICEA_sp where " & stringyear & " and " & _
    "(len(transportname)>0) and convert(smalldatetime,biltydate,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,biltydate,103) <=convert(smalldatetime,'" + Trim(date2.Text) + "',103) and " & trans & _
    "", con, adOpenStatic, adLockReadOnly
End If

Dim aaaa As Date

While RS.EOF = False
con.Execute "insert into  TransPortRpt(Party,Bilty,Dates,Amount_Text,Amount,Vtype,setupid,fyear,inv) values ('" & RS(0) & "','" & RS(1) & "',Convert(smalldatetime, '" & RS(2) & "', 103),'" & RS(3) & "','0','C'," & setupid & ",'" & session & "','" & RS!invoiceNo & "')"
RS.MoveNext
Wend





Dim cname, Add As String

If RS.State = 1 Then RS.close
RS.Open "select cname,add1 from setup1 where " & stringyear, con
If RS.EOF = False Then
   cname = RS(0)
   Add = RS(1)
End If


'===============================================================



If MsgBox("Want To View ?", vbQuestion + vbYesNo) = vbYes Then

If s1 = 1 Then
    
   

    cr.Reset
    cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    cr.ReportFileName = rptPath & "\frmBiltyRegister.rpt"
    cr.Formulas(0) = "fromdate='" & date1.Text & "'"
    cr.Formulas(1) = "todate='" & date2.Text & "'"
    cr.WindowState = crptMaximized
    cr.WindowShowPrintSetupBtn = True
    cr.WindowShowRefreshBtn = True
   
    cr.Action = 1
Else
    cr.Reset
    cr.ReportFileName = rptPath & "\frmBiltyRegister_ServiceTax.rpt"
    
    cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    cr.Formulas(0) = "fromdate='" & date1.Text & "'"
    cr.Formulas(1) = "todate='" & date2.Text & "'"
    
    
    cr.WindowState = crptMaximized
    cr.WindowShowPrintSetupBtn = True
    cr.Action = 1
End If

End If




End Sub

Private Sub date1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub
Private Sub date1_LostFocus()
    If Trim(date1.Text) <> "" Then
        If Not checkdate(Trim(date1.Text), date1) Then
            date1.SetFocus
        End If
    End If
End Sub
Private Sub date2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub
Private Sub date2_LostFocus()
    If Trim(date2.Text) <> "" Then
        If Not checkdate(Trim(date2.Text), date2) Then
            date2.SetFocus
        End If
    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{taB}"
End If

End Sub

Private Sub Form_Load()


    If main.UserName = "v" Then
       Me.cmdAdd.Visible = True
       Me.txtPath.Visible = True
    Else
       Me.cmdAdd.Visible = False
       Me.txtPath.Visible = False
    End If



  GWFLAG = True
  Dim rs1 As New ADODB.Recordset
  Dim RS As New ADODB.Recordset
  rs1.Open "select distinct Transportname from transportmaster where " & stringyear, con, adOpenDynamic, adLockReadOnly
     If Not rs1.EOF Then
        Do While Not rs1.EOF
            Me.Glist.AddItem rs1(0)
            If Not rs1.EOF Then
                rs1.MoveNext
            End If
        Loop
 End If
 RS.Open "select * from setup1 where " & stringyear, con, adOpenDynamic, adLockReadOnly
 date1.Text = RS!yarfrom
 date2.Text = RS!yarto
 Me.top = 0
 Me.Left = 0
 
 
Me.Left = Me.Left + 1000
Me.top = Me.top + 1500
 
'============================================
BackColorFrom Me
 
End Sub
Private Sub return1_Click()
    Unload Me
    ''MainMenu.Toolbar1.Visible = True
End Sub

Private Sub print_Click()

End Sub
Function genreport()

   Dim RS As ADODB.Recordset
   Dim rs1 As ADODB.Recordset
   Dim rs9 As New ADODB.Recordset
   Dim kk As ADODB.Recordset
   Dim trs As ADODB.Recordset
   Dim paperWidth As Integer
   Dim kkk As ADODB.Recordset
   Dim Tot As Double
   Dim MaxLine, Pno, Line As Integer
   Dim called1 As Boolean
   Dim Glist1 As String
   Dim ID1 As String
   Dim Gc As String
   Dim Gc1 As String
   Dim FooterYes As Boolean
   Dim NetTotal As Double
   Dim GTotal As Double
   Dim J As Integer
   NetTotal = 0
   I = 0
   GTotal = 0
   FooterYes = False
   Set kkk = New ADODB.Recordset
   Set rs1 = New ADODB.Recordset
   Set RS = New ADODB.Recordset
   Set kk = New ADODB.Recordset
   Set trs = New ADODB.Recordset
   Tot = 0
   Line = 0
   Pno = 1
   MaxLine = 72
   called1 = False
   called2 = False
   main.reportname = "Dis. Sales"
   main.reportdata
   main.repors.Find "reportname='" + Trim(main.reportname) + "'"
   MaxLine = main.repors!totalline
   If main.repors!comp = True Then
      paperWidth = Int(main.repors!totalcolumn * 1.75)
   Else
      paperWidth = main.repors!totalcolumn
   End If
   Open "" + VB.App.Path + "\vipin.txt" For Output As #1
   MaxLine = 72
   called1 = False
   Pno = 1
   paperWidth = 125
header:
        For I = 1 To main.repors!TopMargin
            Print #1, ""
            Line = Line + 1
        Next
        If FooterYes = True Then
           While Line <= 72
              Print #1, ""
              Line = Line + 1
           Wend
           Line = 0
           FooterYes = False
        End If
        If kkk.State = 1 Then kkk.close
        CNSetup
        kkk.Open "select * from setup1 where " & stringyear, con, adOpenDynamic, adLockReadOnly, adCmdText
        If Not kkk.BOF Then
            Print #1, ""
            Print #1, Chr(27) + Chr(15) + Chr(14)
            Print #1, Tab(110); "Page No:  " & Pno
            Print #1, Tab(((paperWidth - (Len(Trim(kkk!cname)) * 2)) / 2) + LEFTM - 15); Chr(27) + Chr(15) + Chr(14); Trim(kkk!cname)
            Print #1, Tab(((paperWidth - (Len(Trim(kkk!add1)) * 2)) / 2) + LEFTM); Chr(27) + Chr(15); dspace(Trim(kkk!add1))
            Line = Line + 5
         End If
         xstr = Me.date1.Text & "  To  " & Me.date2.Text
             
            Print #1, Tab(LEFTM + ((paperWidth - Len(Trim("GROUP WISE SALES")) * 2) / 2)); Chr(27) + Chr(14); Trim("TRANSPORT WISE DETAILS"); Chr(27) + Chr(15)
        Print #1, Tab(LEFTM + ((paperWidth - Len(Trim("Period : " + Trim(xstr)))) / 2)); Trim("Period : " + Trim(xstr))
         Print #1, ""
         Print #1, repli("-", paperWidth)
             Print #1, Tab(5); "Transport Name"; Tab(35); "Bilty No. "; Tab(60); "    Date"; Tab(95); "Amount"
         Print #1, repli("-", paperWidth)
         Line = Line + 6
         If called1 = True Then
            called1 = False
            GoTo printagain1
         End If
         If called2 = True Then
            called2 = False
            GoTo printagain2
         End If
         If RS.State = 1 Then RS.close
         Glist1 = ""
         ID1 = ""
         Gc = ""
         Gc1 = ""
         For J = 0 To Glist.ListCount - 1
            If Glist.Selected(J) = True Then
                Glist1 = Glist.List(J)
                If RS.State = 1 Then RS.close
                  If rs9.State = 1 Then rs9.close
                   RS.Open "SELECT  transportname, biltyno, biltydate, freight FROM INVOICEA where " & stringyear & " and biltyDATE >=cdate('" + Trim(date1.Text) + "') And biltydate <= cdate('" + Trim(date2.Text) + "')  and  transportname = '" & Glist1 & "' group by transportname,biltyno, biltydate, freight order by biltydate", con, adOpenStatic, adLockOptimistic, adCmdText
                   rs9.Open "SELECT  transportname,biltyno, biltydate, freight FROM cashA where " & stringyear & " and biltyDATE >=cdate('" + Trim(date1.Text) + "') And  biltydate <= cdate('" + Trim(date2.Text) + "')  and  transportname = '" & Glist1 & "' group by transportname,biltyno, biltydate, freight order by biltydate", con, adOpenStatic, adLockOptimistic, adCmdText
       '            rs9.Open "SELECT  BOOKS.GROUPCODE as Gcode,INVOICENO,invoicedate, sum(CashB.NETAMOUNT) as samount FROM CashB LEFT JOIN BOOKS ON CashB.BOOKCODE = BOOKS.BOOKCODE  where INVOICEDATE >=cdate('" + Trim(date1.Text) + "') And INVOICEDATE <= cdate('" + Trim(date2.Text) + "')  and  groupcode = '" & Glist1 & "' group by books.groupcode,INVOICENO,invoicedate order by INVOICENO,invoicedate", CON, adOpenStatic, adLockOptimistic, adCmdText
               If RS.RecordCount > 0 Then
                       RS.MoveFirst
                       While Not RS.EOF
                            If Gc <> RS!transportname Then
                               Gc1 = RS!transportname
                            Else
                               Gc1 = ""
                            End If
                            Print #1, Tab(5); Gc1; Tab(35); RS!biltyno; Tab(60); RS!BILTYDATE; Tab(90); rsets(Trim(Format(Str(Val(RS!freight)), "0.00")), 12)
                            Line = Line + 1
                            GTotal = GTotal + Val(RS!freight)
                           
                            Gc = RS!transportname
                            If Line > MaxLine - 10 Then
                               FooterYes = True
                               Pno = Pno + 1
                               called1 = True
                               GoTo header
printagain1:
                               called1 = False
                             End If
                             RS.MoveNext
                        Wend
                  End If
                  If GWFLAG = True Then
                      If rs9.RecordCount > 0 Then
                       rs9.MoveFirst
                       While Not rs9.EOF
                            If Gc <> rs9!transportname Then
                               Gc1 = rs9!transportname
                            Else
                               Gc1 = ""
                            End If
                            Print #1, Tab(5); Gc1; Tab(35); rs9!biltyno; Tab(60); rs9!BILTYDATE; Tab(90); rsets(Trim(Format(Str(Val(rs9!freight)), "0.00")), 12)
                            Line = Line + 1
                            GTotal = GTotal + Val(rs9!freight)
                            Gc = rs9!transportname
                            If Line > MaxLine - 10 Then
                               FooterYes = True
                               Pno = Pno + 1
                               called2 = True
                               
                               GoTo header
printagain2:
                               called2 = False
                             End If
                             rs9.MoveNext
                        Wend
                       End If
                  End If
                        
                  Print #1, ""
                  Print #1, repli("-", paperWidth)
                  Print #1, Tab(5); "Transport  Total"; Tab(90); rsets(Trim(Format(Str(GTotal), "0.00")), 12)
                  Print #1, repli("-", paperWidth)
                  Line = Line + 4
                  NetTotal = NetTotal + GTotal
                  GTotal = 0
                        
                                   
            End If
        Next J
        Print #1, repli("-", paperWidth)
        Print #1, Tab(5); "Net Total"; Tab(90); rsets(Trim(Format(Str(NetTotal), "0.00")), 12)
        Print #1, repli("-", paperWidth)
        Line = Line + 3
        While Line <= 72
          Print #1, ""
          Line = Line + 1
        Wend
        Close #1
        
End Function

Private Sub print1_Click()
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
End Sub
Function rsets(ST As String, length As Integer) As String
   
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

Private Sub txtPath_LostFocus()
If txtPath = "" Then
   con.Execute "delete from conpath where " & stringyear & " and len(path)>0"
End If
End Sub
