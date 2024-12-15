VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAdviceStatus 
   ClientHeight    =   7770
   ClientLeft      =   165
   ClientTop       =   -1095
   ClientWidth     =   13905
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   13905
   WindowState     =   2  'Maximized
   Begin VB.Frame panel 
      Height          =   8520
      Left            =   405
      TabIndex        =   0
      Top             =   180
      Width           =   12075
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   180
         ScaleHeight     =   705
         ScaleWidth      =   6135
         TabIndex        =   9
         Top             =   7650
         Width           =   6135
         Begin VB.CommandButton TESTENTRYCD 
            BackColor       =   &H00FFFFFF&
            Cancel          =   -1  'True
            Caption         =   "&Close"
            Height          =   570
            Left            =   4680
            Picture         =   "Advstat.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   75
            Width           =   1335
         End
         Begin VB.CommandButton REPORTCD 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Print"
            Height          =   570
            Left            =   3150
            Picture         =   "Advstat.frx":0BE4
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   75
            Width           =   1335
         End
         Begin VB.CommandButton savecd 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Save"
            Height          =   570
            Left            =   75
            Picture         =   "Advstat.frx":17C8
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   75
            Width           =   1335
         End
         Begin VB.CommandButton Help 
            Caption         =   "&Help"
            Height          =   435
            Left            =   -1335
            TabIndex        =   11
            Top             =   30
            Width           =   1155
         End
         Begin VB.CommandButton SSCommand2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Abandone"
            Height          =   570
            Left            =   1605
            Picture         =   "Advstat.frx":23AC
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   75
            Width           =   1335
         End
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1170
         TabIndex        =   8
         Top             =   2025
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Frame Frame1 
         Caption         =   "Advice Status"
         Height          =   615
         Left            =   210
         TabIndex        =   3
         Top             =   1020
         Width           =   8955
         Begin VB.OptionButton Option1 
            BackColor       =   &H0078CFE9&
            Caption         =   "Clear "
            Height          =   195
            Index           =   0
            Left            =   1890
            TabIndex        =   7
            Top             =   210
            Width           =   735
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H0078CFE9&
            Caption         =   "Pending"
            Height          =   195
            Index           =   1
            Left            =   3300
            TabIndex        =   6
            Top             =   210
            Width           =   1365
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H0078CFE9&
            Caption         =   "Return"
            Height          =   195
            Index           =   2
            Left            =   4980
            TabIndex        =   5
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H0078CFE9&
            Caption         =   "All"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   4
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.ComboBox cmbAgentName 
         Height          =   315
         Left            =   2340
         TabIndex        =   2
         Top             =   540
         Width           =   2925
      End
      Begin VB.ComboBox Adstatus 
         Height          =   315
         ItemData        =   "Advstat.frx":2936
         Left            =   5580
         List            =   "Advstat.frx":2943
         TabIndex        =   1
         Top             =   1995
         Width           =   1515
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
         Height          =   5850
         Left            =   180
         TabIndex        =   15
         Top             =   1695
         Width           =   11670
         _ExtentX        =   20585
         _ExtentY        =   10319
         _Version        =   393216
         FixedCols       =   0
         RowHeightMin    =   300
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSMask.MaskEdBox date1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   2340
         TabIndex        =   16
         Top             =   180
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox date2 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   4050
         TabIndex        =   17
         Top             =   180
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H0078CFE9&
         BorderWidth     =   4
         Height          =   825
         Left            =   135
         Top             =   7605
         Width           =   6225
      End
      Begin VB.Label Label3 
         Caption         =   " - To - "
         Height          =   315
         Left            =   3555
         TabIndex        =   20
         Top             =   195
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "From The Date :"
         Height          =   195
         Left            =   330
         TabIndex        =   19
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Agent :"
         Height          =   195
         Left            =   315
         TabIndex        =   18
         Top             =   615
         Width           =   1995
      End
   End
End
Attribute VB_Name = "frmAdviceStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private SortCol As Integer
Private SortType As Integer
Sub Heading()

If Grid1.Cols > 2 Then
    For I = 1 To Grid1.Rows - 1
         Grid1.Col = 0
         Grid1.Row = I
         Grid1.Text = I
    Next
    Grid1.TextMatrix(0, 0) = "SNo."
    Grid1.TextMatrix(0, 1) = "Inv. No"
    Grid1.TextMatrix(0, 2) = "Inv. Date"
    Grid1.TextMatrix(0, 3) = "Sub Ledger"
    Grid1.TextMatrix(0, 4) = "B.Amount"
    Grid1.TextMatrix(0, 5) = "Advice Status"
    Grid1.TextMatrix(0, 6) = "Remark"
    Grid1.ColWidth(0) = 500
    Grid1.ColWidth(1) = 700
    Grid1.ColWidth(2) = 1000
    Grid1.ColWidth(3) = 3000
    Grid1.ColWidth(4) = 1000
    Grid1.ColWidth(5) = 1500
    Grid1.ColWidth(6) = 2000
End If
    
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Cmdprint_Click()
End Sub

Private Sub CmdSave_Click()
    Dim rsb As New ADODB.Recordset
    Dim rsd As New ADODB.Recordset
    rsd.Open "select * from testmaster where  " & stringyear & "", CON, adOpenKeyset, adLockOptimistic
    For I = 1 To Grid1.Rows - 1
          Grid1.Col = 0
          Grid1.Row = I
          If Grid1.Text = "" Then
              Exit For
          End If
          If Grid1.Text <> "" Then
            If Rsbd.State = 1 Then Rsbd.close
            rsd.Find "name = '" & Grid1.Text & "',1,1"
            If Not rsd.EOF Then
               Grid1.Col = 2
               rsd!charges = Val(Grid1.Text)
               rsd.update
            End If
         End If
       Next I

End Sub



Private Sub adstatus_Click()
   Grid1.Text = Adstatus.Text
   'Grid1.SetFocus
End Sub

Private Sub adstatus_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      Adstatus.Visible = True
      Grid1.Text = Adstatus.Text
      If Grid1.Row < Grid1.Rows - 1 Then
          Grid1.Row = Grid1.Row + 1
          Grid1_Click
          Exit Sub
      End If
      Text1.Visible = False
      'Adstatus.Text = ""
      'Grid1.SetFocus
      'Adstatus.Move Grid1.Left + Grid1.CellLeft, Grid1.Top + Grid1.CellTop - 30, Grid1.CellWidth - 15
      'Adstatus.SetFocus
 End If
End Sub

Private Sub Adstatus1_Click()

End Sub

Private Sub cmbAgentName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If
End Sub

Private Sub date1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
   

End If
End Sub

Private Sub date1_LostFocus()
   If Trim(date1.Text) <> "" Then
        If Not checkdate(Trim(date1.Text), date1) Then
            date1.SetFocus
        End If
    End If
End Sub

Private Sub date2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If
End Sub

Private Sub date2_LostFocus()
If Trim(date2.Text) <> "" Then
    If Not checkdate(Trim(date2.Text), date2) Then
           date2.SetFocus
    End If
End If
End Sub

Private Sub Form_GotFocus()
frmAdviceStatus.WindowState = 2
End Sub

Private Sub Form_Load()
    Dim rs As New ADODB.Recordset
    'rs.Open "Select * from setup where " & stridnyear & "", CON, adOpenKeyset, adLockReadOnly, adCmdText
    rs.Open "Select * from setup", CON, adOpenKeyset, adLockReadOnly, adCmdText
    date1.Text = rs!yarfrom
    date2.Text = rs!yarto
    If rs.State = 1 Then rs.close
    '*******Agent  combo fill
    rs.Open "select  Agentname from AgentMaster where " & stringyear & " order by agentname", CON, adOpenKeyset, adLockReadOnly, adCmdText
    cmbAgentName.Clear
    If Not rs.EOF Then
       Do While Not rs.EOF
          If IsNull(rs(0)) = False Then
            Me.cmbAgentName.AddItem rs(0)
          End If
          If Not rs.EOF Then rs.MoveNext
        Loop
    End If
    If rs.State = 1 Then rs.close
    
    BackColorFrom Me
    
End Sub


Private Sub Form_Resize()
frmAdviceStatus.WindowState = 2

panel.Left = (Me.ScaleWidth - panel.Width) / 2
panel.Top = (Me.ScaleHeight - panel.Height) / 2


End Sub

Private Sub Form_Unload(cancel As Integer)
Unload Me
End Sub

Private Sub Grid1_Click()
 If Grid1.MouseRow < Grid1.FixedRows Then Exit Sub
 If Grid1.Row > 0 Then
    
    If Grid1.Col = 5 Then
         
         Text1.Visible = False
         Adstatus.Visible = True
         Adstatus.Move Grid1.Left + Grid1.CellLeft, Grid1.Top + Grid1.CellTop - 30, Grid1.CellWidth - 15
         Grid1.Col = 5
         Adstatus.Text = Grid1.Text
         Adstatus.SetFocus
    End If
    
     If Grid1.Col = 6 Then
         Adstatus.Visible = False
         Text1.Move Grid1.Left + Grid1.CellLeft, Grid1.Top + Grid1.CellTop - 30, Grid1.CellWidth - 15, Grid1.CellHeight
         Text1.Text = Grid1.Text
         Text1.Visible = True
         Text1.SetFocus
    End If
  End If
 
 
 
 

End Sub

Private Sub Grid1_DblClick()


 If Grid1.MouseRow >= Grid1.FixedRows Then Exit Sub
    I = SortCol
    SortCol = Grid1.Col
    If I <> SortCol Then
        SortType = 1
    Else
        SortType = SortType + 1
        If SortType = 3 Then SortType = 1
    End If
    With Grid1
        .Redraw = False
        .Row = 1
        .RowSel = .Rows - 1
        .Col = SortCol
        .Sort = SortType
        .Redraw = True
    End With
    Heading
    
End Sub

Private Sub Grid1_GotFocus()
 If Grid1.MouseRow >= Grid1.FixedRows Then Exit Sub
  If Grid1.Row > 0 Then
    If Grid1.Col = 5 Then
         Text1.Visible = False
         Adstatus.Visible = True
         Adstatus.Move Grid1.Left + Grid1.CellLeft, Grid1.Top + Grid1.CellTop - 30, Grid1.CellWidth - 15
         Adstatus.Text = Grid1.Text
         Adstatus.SetFocus
    End If
    If Grid1.Col = 6 Then
         Text1.Move Grid1.Left + Grid1.CellLeft, Grid1.Top + Grid1.CellTop - 30, Grid1.CellWidth - 15, Grid1.CellHeight
         Text1.Text = Grid1.Text
         Text1.Visible = True
         Text1.SetFocus
    End If
  End If

End Sub





Private Sub Grid1_RowColChange()

If Grid1.MouseRow >= Grid1.FixedRows Then Exit Sub
If Grid1.Col = 5 Then

      Adstatus.Text = Grid1.Text
      Adstatus.Visible = True
      Adstatus.Move Grid1.Left + Grid1.CellLeft, Grid1.Top + Grid1.CellTop - 30, Grid1.CellWidth - 15
      Grid1.SetFocus
Else
    
      Adstatus.Visible = False
  

End If


End Sub

Private Sub Option1_Click(Index As Integer)
Dim rsAd As ADODB.Recordset
Set rsAd = New ADODB.Recordset
Text1.Visible = False
Adstatus.Visible = False
If Option1(0).value = True Then
   If cmbAgentName.Text = "" Then
      rsAd.Open "select 1 as SNo, INVOICENO,INVOICEDATE,SUBLEDGER,BAA,AdviceStatus,AdviceRemark from Invoicea  where  " & stringyear & " and BAA <> 0 and   upper(advicestatus) = '" & UCase("Clear") & "'and  convert(smalldatetime,invoicea.invoicedate,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,invoicea.invoicedate,103) >=convert(smalldatetime,'" + Trim(date2.Text) + "',103)   order by INVOICENO, INVOICEDATE , SUBLEDGER ", CON, adOpenStatic, adLockReadOnly
   Else
      rsAd.Open "select 1 as Sno, INVOICENO,INVOICEDATE,SUBLEDGER,BAA,AdviceStatus,AdviceRemark from Invoicea  where   " & stringyear & " and BAA <> 0 and  upper(advicestatus) = '" & UCase("Clear") & "'and  convert(smalldatetime,invoicea.invoicedate,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,invoicea.invoicedate,103) >=convert(smalldatetime,'" + Trim(date2.Text) + "',103) and AgentName ='" & cmbAgentName.Text & "' order by INVOICENO, INVOICEDATE , SUBLEDGER ", CON, adOpenStatic, adLockReadOnly
   End If
   If rsAd.RecordCount > 0 Then
     Set Grid1.DataSource = rsAd
     Grid1.Refresh
     Heading
   Else
     Grid1.Clear
     Heading
   End If
End If
If Option1(1).value = True Then
    If cmbAgentName.Text = "" Then
       rsAd.Open "select 1 as SNo,INVOICENO,INVOICEDATE,SUBLEDGER,BAA,AdviceStatus,AdviceRemark from Invoicea  where  " & stringyear & " and BAA <> 0 and upper(advicestatus) = '" & UCase("Pending") & "' and convert(smalldatetime,invoicea.invoicedate,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,invoicea.invoicedate,103) >=convert(smalldatetime,'" + Trim(date2.Text) + "',103)   order by INVOICENO, INVOICEDATE , SUBLEDGER ", CON, adOpenStatic, adLockReadOnly
    Else
       rsAd.Open "select 1 as SNo,INVOICENO,INVOICEDATE,SUBLEDGER,BAA,AdviceStatus,AdviceRemark from Invoicea  where  " & stringyear & " and BAA <> 0 and upper(advicestatus) = '" & UCase("Pending") & "' and convert(smalldatetime,invoicea.invoicedate,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,invoicea.invoicedate,103) >=convert(smalldatetime,'" + Trim(date2.Text) + "',103) and AgentName ='" & cmbAgentName.Text & "'   order by INVOICENO, INVOICEDATE , SUBLEDGER ", CON, adOpenStatic, adLockReadOnly
    End If
      If rsAd.RecordCount > 0 Then
         Set Grid1.DataSource = rsAd
         Grid1.Refresh
         Heading
      Else
         Grid1.Clear
         Heading
      End If
End If



If Option1(2).value = True Then
    If cmbAgentName.Text = "" Then
       rsAd.Open "select 1 as SNo,INVOICENO,INVOICEDATE,SUBLEDGER,BAA,AdviceStatus,AdviceRemark from Invoicea  where  " & stringyear & " and BAA <> 0 and advicestatus = '" & UCase("Return") & "' and convert(smalldatetime,invoicea.invoicedate,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,invoicea.invoicedate,103) >=convert(smalldatetime,'" + Trim(date2.Text) + "',103)   order by INVOICENO, INVOICEDATE , SUBLEDGER ", CON, adOpenStatic, adLockReadOnly
    Else
       rsAd.Open "select 1 as SNo,INVOICENO,INVOICEDATE,SUBLEDGER,BAA,AdviceStatus,AdviceRemark from Invoicea  where  " & stringyear & " and BAA <> 0 and advicestatus = '" & UCase("Return") & "' and convert(smalldatetime,invoicea.invoicedate,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,invoicea.invoicedate,103) >=convert(smalldatetime,'" + Trim(date2.Text) + "',103) and AgentName ='" & cmbAgentName.Text & "'   order by INVOICENO, INVOICEDATE , SUBLEDGER ", CON, adOpenStatic, adLockReadOnly
    End If
    If rsAd.RecordCount > 0 Then
         Set Grid1.DataSource = rsAd
         Grid1.Refresh
         Heading
    Else
         Grid1.Clear
         Heading
        
    End If
End If





If Option1(3).value = True Then
    If cmbAgentName.Text = "" Then
       rsAd.Open "select 1 as SNo,INVOICENO,INVOICEDATE,SUBLEDGER,BAA,AdviceStatus,AdviceRemark from Invoicea  where  " & stringyear & " and BAA <> 0  and convert(smalldatetime,INVOICEA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,INVOICEA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)   order by INVOICENO, INVOICEDATE , SUBLEDGER ", CON, adOpenStatic, adLockReadOnly
    Else
       rsAd.Open "select 1 as SNo,INVOICENO,INVOICEDATE,SUBLEDGER,BAA,AdviceStatus,AdviceRemark from Invoicea  where  " & stringyear & " and BAA <> 0  and convert(smalldatetime,INVOICEA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,INVOICEA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) and AgentName ='" & cmbAgentName.Text & "'   order by INVOICENO, INVOICEDATE , SUBLEDGER ", CON, adOpenStatic, adLockReadOnly
    End If
    If rsAd.RecordCount > 0 Then
         Set Grid1.DataSource = rsAd
         Grid1.Refresh
         Heading
    Else
         Grid1.Clear
         Heading
        
    End If
End If





End Sub

Private Sub Option1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Grid1.SetFocus
    If Grid1.Col >= 5 Then
      Grid1.Col = 5
    End If
End If
End Sub

Private Sub REPORTCD_Click()
If DateDiff("d", Trim(date1.Text), Trim(date2.Text)) <= 0 Then
        MsgBox "invalid date"
        Exit Sub
End If
genreport
PrintOption.Show
End Sub

Private Sub savecd_Click()
    Dim rsd As New ADODB.Recordset
    rsd.Open "select * from invoicea where " & stringyear, CON, adOpenKeyset, adLockOptimistic
    For I = 1 To Grid1.Rows - 1
        Grid1.Col = 1
        Grid1.Row = I
        If Grid1.Text = "" Then
           Exit For
        End If
        If Grid1.Text <> "" Then
           rsd.MoveFirst
           rsd.Find "InvoiceNo = '" & Grid1.Text & "'", 0, adSearchForward
           If Not rsd.EOF Then
              Grid1.Col = 5
              rsd!advicestatus = Grid1.Text
              Grid1.Col = 6
              rsd!adviceremark = Grid1.Text
              rsd.update
           End If
        End If
    Next I
    MsgBox "Record is saved.... "
    
End Sub

Private Sub SSCommand2_Click()
MsgBox "Operation Not allowed.."
End Sub

Private Sub TESTENTRYCD_Click()
Unload Me
End Sub

Private Sub Text1_Change()
If Grid1.Col = 6 And Grid1.Row > 0 Then
    Grid1.Text = Text1.Text
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then

          If Grid1.Col = 6 Then
              If Grid1.Row < Grid1.Rows - 1 Then
                 Grid1.Row = Grid1.Row + 1
              End If
              Adstatus.Visible = False
              Text1.Move Grid1.Left + Grid1.CellLeft, Grid1.Top + Grid1.CellTop - 30, Grid1.CellWidth - 15, Grid1.CellHeight
              Text1.Visible = True
              Text1.Text = Grid1.Text
              Text1.SetFocus
          End If
          
          
  End If
End Sub














Function genreport()
   Dim rs As ADODB.Recordset
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
   Set rs = New ADODB.Recordset
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
         Print #1, Tab(LEFTM + ((paperWidth - Len(Trim("ADVICE STATUS")) * 2) / 2)); Chr(27) + Chr(14); Trim("ADVICE STATUS"); Chr(27) + Chr(15)
         Print #1, Tab(LEFTM + ((paperWidth - Len(Trim("Period : " + Trim(xstr)))) / 2)); Trim("Period : " + Trim(xstr))
         Print #1, ""
         Print #1, repli("-", paperWidth)
         Print #1, Tab(1); "SNo."; Tab(5); "INV.No."; Tab(15); "Date"; Tab(28); "Party Description"; Tab(73); "Bank Amount"; Tab(88); "Adv.Status"; Tab(100); "Remark"
         Print #1, repli("-", paperWidth)
         Line = Line + 6
         If called1 = True Then
            called1 = False
            GoTo printagain1
         End If
         If rs.State = 1 Then rs.close
         For J = 1 To Grid1.Rows - 1
             Grid1.Row = J
             Print #1, Tab(1); Grid1.TextMatrix(J, 0); Tab(7); Grid1.TextMatrix(J, 1); Tab(15); Grid1.TextMatrix(J, 2); Tab(28); Grid1.TextMatrix(J, 3); Tab(71); rsets(Trim(Format(Str(Grid1.TextMatrix(J, 4)), "0.00")), 12); Tab(88); Grid1.TextMatrix(J, 5); Tab(100); Left(Grid1.TextMatrix(J, 6), 30)
             Line = Line + 1
             If Line > MaxLine - 10 Then
                FooterYes = True
                Pno = Pno + 1
                called1 = True
                GoTo header
printagain1:
                called1 = False
             End If
        Next J
        Print #1, repli("-", paperWidth)
        Line = Line + 1
        While Line <= 72
          Print #1, ""
          Line = Line + 1
        Wend
        Close #1

End Function



















    
