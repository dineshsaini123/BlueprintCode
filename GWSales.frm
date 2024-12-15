VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form groupwisesales 
   Caption         =   "Group Wise Sales"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4830
   ScaleWidth      =   9480
   Begin VB.CommandButton cmdsalesreturn 
      Caption         =   "&Sales Return "
      Height          =   405
      Left            =   5910
      TabIndex        =   4
      Top             =   1740
      Width           =   1605
   End
   Begin VB.ListBox Glist 
      Height          =   2985
      Left            =   1830
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   630
      Width           =   2925
   End
   Begin VB.CommandButton Commandshow 
      Caption         =   "&Sales"
      Height          =   405
      Left            =   5880
      TabIndex        =   3
      Top             =   1140
      Width           =   1605
   End
   Begin VB.CommandButton Commandreturn 
      Caption         =   "&Return"
      Height          =   405
      Left            =   5940
      TabIndex        =   5
      Top             =   2280
      Width           =   1605
   End
   Begin MSMask.MaskEdBox date1 
      Height          =   345
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   609
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox date2 
      Height          =   345
      Left            =   3990
      TabIndex        =   1
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VB.Label Label2 
      Caption         =   "From The Date"
      Height          =   315
      Left            =   360
      TabIndex        =   8
      Top             =   150
      Width           =   1305
   End
   Begin VB.Label Label3 
      Caption         =   "To"
      Height          =   315
      Left            =   3150
      TabIndex        =   7
      Top             =   210
      Width           =   465
   End
   Begin VB.Label Label1 
      Caption         =   "Group Code :"
      Height          =   285
      Left            =   360
      TabIndex        =   6
      Top             =   630
      Width           =   1185
   End
End
Attribute VB_Name = "groupwisesales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public GWFLAG As Boolean
Dim SQLSTRING As String

Private Sub Command1_Click()
    Unload Me
    'MainMenu.Toolbar1.Visible = True
End Sub

Private Sub cmdsalesreturn_Click()
If DateDiff("d", Trim(date1.Text), Trim(date2.Text)) <= 0 Then
    MsgBox "invalid date"
    Exit Sub
End If
GWFLAG = False
genreport
PrintOption.Show
End Sub

Private Sub Commandreturn_Click()
''MainMenu.Toolbar1.Visible = True
Unload Me
End Sub

Private Sub Commandshow_Click()
If DateDiff("d", Trim(date1.Text), Trim(date2.Text)) <= 0 Then
        MsgBox "invalid date"
        Exit Sub
End If
GWFLAG = True
genreport
PrintOption.Show
        

    
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
  GWFLAG = True
  Dim rs1 As New ADODB.Recordset
  Dim RS As New ADODB.Recordset
  rs1.Open "select * from GROUPS where " & stringyear & " ", CON, adOpenKeyset, adLockReadOnly, adCmdText
     If Not rs1.EOF Then
        Do While Not rs1.EOF
            Me.Glist.AddItem rs1(0)
            If Not rs1.EOF Then
                rs1.MoveNext
            End If
        Loop
 End If
 RS.Open "Select * from setup where " & stringyear & "", CON, adOpenKeyset, adLockReadOnly, adCmdText
 date1.Text = RS!yarfrom
 date2.Text = RS!yarto
 Me.TOP = 0
 Me.Left = 0
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
   Open "" + App.Path + "\vipin.txt" For Output As #1
   MaxLine = 72
   called1 = False
   Pno = 1
   paperWidth = 125
header:
        For I = 1 To main.repors!topmargin
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
        If kkk.State = 1 Then kkk.Close
        CNSetup
        kkk.Open "Select * from setup where " & stringyear & "", CON, adOpenKeyset, adLockReadOnly, adCmdText
        If Not kkk.BOF Then
            Print #1, ""
            Print #1, Chr(27) + Chr(15) + Chr(14)
            Print #1, Tab(110); "Page No:  " & Pno
            Print #1, Tab(((paperWidth - (Len(Trim(kkk!cname)) * 2)) / 2) + LEFTM - 15); Chr(27) + Chr(15) + Chr(14); dspace(Trim(kkk!cname))
            Print #1, Tab(((paperWidth - (Len(Trim(kkk!add1)) * 2)) / 2) + LEFTM); Chr(27) + Chr(15); dspace(Trim(kkk!add1))
            Line = Line + 5
         End If
         xstr = Me.date1.Text & "  To  " & Me.date2.Text
             
         If GWFLAG = True Then
            Print #1, Tab(LEFTM + ((paperWidth - Len(Trim("GROUP WISE SALES")) * 2) / 2)); Chr(27) + Chr(14); Trim("GROUP WISE SALES"); Chr(27) + Chr(15)
         Else
            Print #1, Tab(LEFTM + ((paperWidth - Len(Trim("GROUP WISE SALES")) * 2) / 2)); Chr(27) + Chr(14); Trim("GROUP WISE SALES RETURN"); Chr(27) + Chr(15)
         End If
         Print #1, Tab(LEFTM + ((paperWidth - Len(Trim("Period : " + Trim(xstr)))) / 2)); Trim("Period : " + Trim(xstr))
         Print #1, ""
         Print #1, repli("-", paperWidth)
         If GWFLAG = True Then
             Print #1, Tab(5); "Group Code "; Tab(35); "INVOICE NO. "; Tab(60); "    Date"; Tab(95); "NetAmount"
         Else
             Print #1, Tab(5); "Group Code "; Tab(35); "CREDIT(ITEM) NO. "; Tab(60); "    Date"; Tab(95); "NetAmount"
         End If
         Print #1, repli("-", paperWidth)
         Line = Line + 6
         If called1 = True Then
            called1 = False
            GoTo printagain1
         End If
         If RS.State = 1 Then RS.Close
         Glist1 = ""
         ID1 = ""
         Gc = ""
         Gc1 = ""
         For J = 0 To Glist.ListCount - 1
            If Glist.Selected(J) = True Then
                Glist1 = Glist.List(J)
                If RS.State = 1 Then RS.Close
               If GWFLAG = True Then
                   RS.Open "SELECT  BOOKS.GROUPCODE as Gcode,INVOICENO,invoicedate, sum(INVOICEB.NETAMOUNT) as samount FROM INVOICEB LEFT JOIN BOOKS ON INVOICEB.BOOKCODE = BOOKS.BOOKCODE and INVOICEB.fyear = BOOKS.fyear and INVOICEB.setupid = BOOKS.setupid where  invoiceb.setupid=" & main.setupid & " and invoiceb.fyear='" & main.session & "' and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" & Trim(date2.Text) & "',103)  and  groupcode = '" & Glist1 & "' group by books.groupcode,INVOICENO,invoicedate order by INVOICENO,invoicedate", CON, adOpenStatic, adLockOptimistic, adCmdText
               Else
                  RS.Open "SELECT  BOOKS.GROUPCODE as Gcode,INVOICENO,invoicedate, sum(CREDITB.NETAMOUNT) as samount FROM CREDITB LEFT JOIN BOOKS ON CREDITB.BOOKCODE = BOOKS.BOOKCODE  and INVOICEB.fyear = BOOKS.fyear and INVOICEB.setupid = BOOKS.setupid  where   invoiceb.setupid=" & main.setupid & " and invoiceb.fyear='" & main.session & "' and  convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" & Trim(date2.Text) & "',103)  and  groupcode = '" & Glist1 & "' group by books.groupcode,INVOICENO,invoicedate order by INVOICENO,invoicedate", CON, adOpenStatic, adLockOptimistic, adCmdText
               End If
               If RS.RecordCount > 0 Then
                       RS.MoveFirst
                       While Not RS.EOF
                            If Gc <> RS!gcode Then
                               Gc1 = RS!gcode
                            Else
                               Gc1 = ""
                            End If
                            Print #1, Tab(5); Gc1; Tab(35); RS!INVOICENO; Tab(60); RS!InvoiceDate; Tab(90); rsets(Trim(Format(str(RS!samount), "0.00")), 12)
                            Line = Line + 1
                            GTotal = GTotal + RS!samount
                            Gc = RS!gcode
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
                        
                        
                  Print #1, ""
                  Print #1, repli("-", paperWidth)
                  Print #1, Tab(5); "Group  Total"; Tab(90); rsets(Trim(Format(str(GTotal), "0.00")), 12)
                  Print #1, repli("-", paperWidth)
                  Line = Line + 4
                  NetTotal = NetTotal + GTotal
                  GTotal = 0
                        
                  End If
                 
            End If
        Next J
        Print #1, repli("-", paperWidth)
        Print #1, Tab(5); "Net Total"; Tab(90); rsets(Trim(Format(str(NetTotal), "0.00")), 12)
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
    c1.max = 11
    c1.Flags = &H40000 Or &H4&
    c1.ShowPrinter
    frompage = c1.frompage
    topage = c1.topage
    copies = c1.copies
    If c1.Flags And &H20 Then
        c2.ShowSave
    End If
End Sub


