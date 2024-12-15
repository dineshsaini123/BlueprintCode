VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form groupwisesales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Group Wise Sales"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   7230
   Begin VB.CommandButton cmdsalesreturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Sales Return "
      Height          =   630
      Left            =   5175
      Picture         =   "GWSales.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1305
      Width           =   1605
   End
   Begin VB.ListBox Glist 
      Height          =   3435
      Left            =   1830
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   630
      Width           =   2685
   End
   Begin VB.CommandButton Commandshow 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Sales"
      Height          =   630
      Left            =   5175
      Picture         =   "GWSales.frx":0BE4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   645
      Width           =   1605
   End
   Begin VB.CommandButton Commandreturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Return"
      Height          =   675
      Left            =   5175
      Picture         =   "GWSales.frx":17C8
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1965
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
      Left            =   3315
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
      BackStyle       =   0  'Transparent
      Caption         =   "From The Date"
      Height          =   315
      Left            =   360
      TabIndex        =   8
      Top             =   150
      Width           =   1305
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      Height          =   315
      Left            =   2970
      TabIndex        =   7
      Top             =   165
      Width           =   465
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
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

'---------
s1 = "11"
PopUpValue1 = ""
For J = 0 To Glist.ListCount - 1
If Glist.Selected(J) = True Then
  If PopUpValue1 = "" Then
     PopUpValue1 = "{invoiceBQry.GROUPCODE} = " & "'" & Glist.List(J) & "'"
  Else
     PopUpValue1 = PopUpValue1 & " or {invoiceBQry.GROUPCODE} = " & "'" & Glist.List(J) & "'"
  End If
End If
Next
'--------


If PopUpValue1 <> "" Then
PopUpValue1 = PopUpValue1 & " and " & "{invoiceBQry.INVOICEDATE}>=datevalue('" & Format(date1.Text, "MM/dd/yyyy") & "') and {invoiceBQry.INVOICEDATE}<=datevalue('" & Format(date2.Text, "MM/dd/yyyy") & "')"
Else
PopUpValue1 = "{invoiceBQry.INVOICEDATE}>=datevalue('" & Format(date1.Text, "MM/dd/yyyy") & "') and {invoiceBQry.INVOICEDATE}<=datevalue('" & Format(date2.Text, "MM/dd/yyyy") & "')"
End If



PrintOption.Show
End Sub

Private Sub Commandreturn_Click()
Unload Me
End Sub
Private Sub Commandshow_Click()

If DateDiff("d", Trim(date1.Text), Trim(date2.Text)) <= 0 Then
        MsgBox "invalid date"
        Exit Sub
End If
GWFLAG = True
genreport

s1 = "10"
PopUpValue1 = ""
For J = 0 To Glist.ListCount - 1
If Glist.Selected(J) = True Then
  If PopUpValue1 = "" Then
     PopUpValue1 = "{invoiceBQry.GROUPCODE} = " & "'" & Glist.List(J) & "'"
  Else
     PopUpValue1 = PopUpValue1 & " or {invoiceBQry.GROUPCODE} = " & "'" & Glist.List(J) & "'"
  End If
End If
Next

If PopUpValue1 <> "" Then
PopUpValue1 = PopUpValue1 & " and " & "{invoiceBQry.INVOICEDATE}>=datevalue('" & Format(date1.Text, "MM/dd/yyyy") & "') and {invoiceBQry.INVOICEDATE}<=datevalue('" & Format(date2.Text, "MM/dd/yyyy") & "')"
Else
PopUpValue1 = "{invoiceBQry.INVOICEDATE}>=datevalue('" & Format(date1.Text, "MM/dd/yyyy") & "') and {invoiceBQry.INVOICEDATE}<=datevalue('" & Format(date2.Text, "MM/dd/yyyy") & "')"
End If


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
  rs1.Open "select * from GROUPS where " & stringyear, con, adOpenDynamic, adLockReadOnly
     If Not rs1.EOF Then
        Do While Not rs1.EOF
            Me.Glist.AddItem rs1(0)
            If Not rs1.EOF Then
                rs1.MoveNext
            End If
        Loop
 End If
 RS.Open "select * from setup1 where  " & stringyear, con, adOpenDynamic, adLockReadOnly
 date1.Text = RS!yarfrom
 date2.Text = RS!yarto
 Me.Top = 0
 Me.Left = 0
 
 BackColorFrom Me
 
End Sub
Private Sub return1_Click()
    Unload Me
    'MainMenu.Toolbar1.Visible = True
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
               If GWFLAG = True Then
                   RS.Open "SELECT  BOOKS.GROUPCODE as Gcode,INVOICENO,invoicedate, sum(INVOICEB.NETAMOUNT) as samount FROM INVOICEB LEFT JOIN BOOKS ON INVOICEB.BOOKCODE = BOOKS.BOOKCODE  where  invoiceb.fyear ='" & session & "' and invoiceb.setupid=" & setupid & " and convert(smalldatetime,INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,INVOICEDATE,103) <= convert(smalldatetime,'" + Trim(date2.Text) + "',103)  and  groupcode = '" & Glist1 & "' group by books.groupcode,INVOICENO,invoicedate order by INVOICENO,invoicedate", con, adOpenStatic, adLockOptimistic
                   rs9.Open "SELECT  BOOKS.GROUPCODE as Gcode,INVOICENO,invoicedate, sum(CashB.NETAMOUNT) as samount FROM CashB LEFT JOIN BOOKS ON CashB.BOOKCODE = BOOKS.BOOKCODE  where cashb.fyear ='" & session & "' and cashb.setupid=" & setupid & " and convert(smalldatetime,INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,INVOICEDATE,103) <= convert(smalldatetime,'" + Trim(date2.Text) + "',103)  and  groupcode = '" & Glist1 & "' group by books.groupcode,INVOICENO,invoicedate order by INVOICENO,invoicedate", con, adOpenStatic, adLockOptimistic, adCmdText
                   
               Else
                  RS.Open "SELECT  BOOKS.GROUPCODE as Gcode,INVOICENO,invoicedate, sum(CREDITB.NETAMOUNT) as samount FROM CREDITB LEFT JOIN BOOKS ON CREDITB.BOOKCODE = BOOKS.BOOKCODE  where CREDITB.fyear ='" & session & "' and CREDITB.setupid=" & setupid & " and  convert(smalldatetime,INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,INVOICEDATE,103) <= convert(smalldatetime,'" + Trim(date2.Text) + "',103)  and  groupcode = '" & Glist1 & "' group by books.groupcode,INVOICENO,invoicedate order by INVOICENO,invoicedate", con, adOpenStatic, adLockOptimistic, adCmdText
               End If
               If RS.RecordCount > 0 Then
                       RS.MoveFirst
                       While Not RS.EOF
                            If Gc <> RS!gcode Then
                               Gc1 = RS!gcode
                            Else
                               Gc1 = ""
                            End If
                            Print #1, Tab(5); Gc1; Tab(35); RS!invoiceNo; Tab(60); RS!INVOICEDATE; Tab(90); rsets(Trim(Format(Str(RS!samount), "0.00")), 12)
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
                  End If
                  If GWFLAG = True Then
                      If rs9.RecordCount > 0 Then
                       rs9.MoveFirst
                       While Not rs9.EOF
                            If Gc <> rs9!gcode Then
                               Gc1 = rs9!gcode
                            Else
                               Gc1 = ""
                            End If
                            Print #1, Tab(5); Gc1; Tab(35); rs9!invoiceNo; Tab(60); rs9!INVOICEDATE; Tab(90); rsets(Trim(Format(Str(rs9!samount), "0.00")), 12)
                            Line = Line + 1
                            GTotal = GTotal + rs9!samount
                            Gc = rs9!gcode
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
                  Print #1, Tab(5); "Group  Total"; Tab(90); rsets(Trim(Format(Str(GTotal), "0.00")), 12)
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



