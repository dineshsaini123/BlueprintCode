VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ChallanOther 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6045
   ClientLeft      =   255
   ClientTop       =   1545
   ClientWidth     =   9450
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   9450
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox amount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   480
      TabIndex        =   14
      Top             =   30
      Width           =   1995
   End
   Begin MSMask.MaskEdBox T1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   315
      Left            =   4080
      TabIndex        =   3
      Top             =   4230
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   556
      _Version        =   393216
      AutoTab         =   -1  'True
      PromptChar      =   "_"
   End
   Begin VB.CommandButton Commandreturn 
      Caption         =   "&Return"
      Height          =   405
      Left            =   8040
      TabIndex        =   1
      Top             =   5190
      Width           =   1065
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
      Height          =   2985
      Left            =   30
      TabIndex        =   0
      Top             =   570
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   5265
      _Version        =   393216
      AllowBigSelection=   0   'False
      FocusRect       =   0
      MergeCells      =   1
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSMask.MaskEdBox T1TEXT 
      Height          =   315
      Left            =   1350
      TabIndex        =   2
      Top             =   4230
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   556
      _Version        =   393216
      AutoTab         =   -1  'True
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox T2 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   315
      Left            =   4080
      TabIndex        =   5
      Top             =   4560
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   556
      _Version        =   393216
      AutoTab         =   -1  'True
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox T2TEXT 
      Height          =   315
      Left            =   1350
      TabIndex        =   4
      Top             =   4560
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   556
      _Version        =   393216
      AutoTab         =   -1  'True
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox T3TEXT 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   315
      Left            =   4050
      TabIndex        =   6
      Top             =   4890
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   556
      _Version        =   393216
      AutoTab         =   -1  'True
      PromptChar      =   "_"
   End
   Begin VB.Label labelbalance 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   4050
      TabIndex        =   13
      Top             =   5220
      Width           =   2700
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Balance :"
      Height          =   315
      Left            =   1350
      TabIndex        =   12
      Top             =   5220
      Width           =   2700
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   4050
      TabIndex        =   11
      Top             =   3900
      Width           =   2700
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Net Amount"
      Height          =   315
      Left            =   1350
      TabIndex        =   10
      Top             =   3900
      Width           =   2700
   End
   Begin VB.Label GROSS 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   6180
      TabIndex        =   8
      Top             =   90
      Width           =   2700
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "GROSS AMOUNT :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3480
      TabIndex        =   7
      Top             =   90
      Width           =   2700
   End
   Begin VB.Label T3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "By BANK :"
      Height          =   315
      Left            =   1350
      TabIndex        =   9
      Top             =   4890
      Width           =   2700
   End
End
Attribute VB_Name = "ChallanOther"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim CON As ADODB.Connection
Dim rs As ADODB.Recordset
Dim lastrow, lastcol As Integer
Public mrow
Sub calc()
Dim TMPPREVROW  As Integer
Dim TMPPREVCOL  As Integer
TMPPREVROW = Grid1.row
TMPPREVCOL = Grid1.col
Dim i As Integer
frmChallan.otherdiscount = 0
frmChallan.otheramount = 0
    For i = 1 To mrow
        Grid1.row = i
        If rs.State = 1 Then
            rs.Close
        End If
        If frmChallan.edit Then
                rs.Open "select * from challanCTMP", CON, adOpenKeyset, adLockReadOnly, adCmdText
        Else
                rs.Open "select * from challanEND where " & stringyear & " ", CON, adOpenKeyset, adLockReadOnly, adCmdText
        End If
        Grid1.col = 0
        If Grid1.Text <> "" Then
            rs.Find "TEXT='" + Trim(Grid1.Text) + "'"
            If Not rs.EOF Then
                If Trim(rs!DEBITORCREDIT) = Trim("Debit") Then
                    If rs!rate > 0 Then
                        Grid1.col = 1
                        frmChallan.otherdiscount = frmChallan.otherdiscount + ((frmChallan.totalamount - frmChallan.totaldiscount) * (Val(Grid1.Text) / 100))
                    Else
                        Grid1.col = 2
                        frmChallan.otherdiscount = frmChallan.otherdiscount + Val(Grid1.Text)
                    End If
                Else
                    If rs!rate > 0 Then
                        Grid1.col = 1
                        frmChallan.otheramount = frmChallan.otheramount + ((frmChallan.totalamount - frmChallan.totaldiscount) * (Val(Grid1.Text) / 100))
                    Else
                        Grid1.col = 2
                        frmChallan.otheramount = frmChallan.otheramount + Val(Grid1.Text)
                    End If
                End If
                End If
        End If
        rs.Close
    Next
    frmChallan.mna.Caption = Format(myround(frmChallan.totalamount + frmChallan.otheramount - frmChallan.totaldiscount - frmChallan.otherdiscount, 2), "0.00")
    frmChallan.mna.Caption = Format(myround(frmChallan.totalamount + frmChallan.otheramount - frmChallan.totaldiscount - frmChallan.otherdiscount, 2), "0.00")
    Me.Label3 = Format(frmChallan.totalamount + frmChallan.otheramount - frmChallan.totaldiscount - frmChallan.otherdiscount, "0.00")
    Me.labelbalance = Format(Val(Label3) + Val(Me.T1) + Val(Me.T2) - Val(Me.T3TEXT), "0.00")
      
    'Me.Label3 = Format(myround(INVOICE.totalamount + INVOICE.otheramount - INVOICE.totaldiscount - INVOICE.otherdiscount, 2), "0.00")
    'Me.labelbalance = Format(myround(Val(Label3) + Val(Me.T1) + Val(Me.T2) - Val(Me.T3TEXT)), "0.00")
Grid1.row = TMPPREVROW
Grid1.col = TMPPREVCOL


End Sub
Sub otherabandon()
        
    For i = 1 To ChallanOther.mrow
        ChallanOther.Grid1.row = i
        ChallanOther.Grid1.col = 0
                    Grid1.col = 2
                    Grid1.Text = ""
    Next

        ChallanOther.T1 = ""
        ChallanOther.T1TEXT = ""
        ChallanOther.T2 = ""
        ChallanOther.T2TEXT = ""
        ChallanOther.T3TEXT = ""
End Sub
Private Sub amount_Change()
    If Grid1.row = lastrow And Grid1.col = lastcol Then
        If Trim(amount.Text) = "" Then
           amount.Text = 0
        End If
        Grid1.Text = Format(amount.Text, "0.00")
        If Grid1.col = 1 Then
            Grid1.col = 2
            Grid1.Text = Format(((frmChallan.totalamount - frmChallan.totaldiscount) * (Val(amount.Text) / 100#)), "0.00")
            Grid1.col = 1
        End If
    End If
     
End Sub
Private Sub amount_GotFocus()
    If Grid1.col = 1 Then
        Grid1.col = 2
        Grid1.Text = Format(((frmChallan.totalamount - frmChallan.totaldiscount) * (Val(amount.Text) / 100#)), "0.00")
        Grid1.col = 1
    End If
End Sub
Private Sub amount_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
again:
calc
        If Grid1.col = 1 Then
            Grid1.col = 2
            Grid1.Text = Format(((frmChallan.totalamount - frmChallan.totaldiscount) * (Val(amount.Text) / 100#)), "0.00")
            Grid1.col = 1
        End If
        If frmChallan.edit Then
                
                rs.Open "select * from challanctmp", CON, adOpenKeyset, adLockReadOnly, adCmdText

        Else
        
        rs.Open "select * from challanEND where  " & stringyear & " ", CON, adOpenKeyset, adLockReadOnly, adCmdText
        End If
        
        Grid1.col = 0
        Grid1.row = Grid1.row + 1
        If Grid1.Text <> "" Then
           rs.Find "TEXT='" + Trim(Grid1.Text) + "'"
            If rs!rate > 0 Then
                Grid1.col = 1
            Else
                Grid1.col = 2
            End If
        Else
           
          amount.Visible = False
          rs.Close
          Me.T1TEXT.SetFocus
            HIT
            Exit Sub
        End If
        rs.Close
        Grid1_Click
    End If
 End Sub
Private Sub amount_LostFocus()
    If Grid1.row = lastrow And Grid1.col = lastcol Then
        If Trim(amount.Text) = "" Then
            amount.Text = 0
        End If
        Grid1.Text = Format(amount.Text, "0.00")
    End If
    If Grid1.row = lastrow And Grid1.col = lastcol Then
        If Trim(amount.Text) = "" Then
            amount.Text = 0
        End If
        Grid1.Text = Format(amount.Text, "0.00")
    End If
    amount.Visible = False

    'Grid1_Click
End Sub

Sub Commandreturn_Click()
    calc
    frmChallan.labelbybank = T3TEXT
    ChallanOther.Hide
    frmChallan.Enabled = True
End Sub

Private Sub Form_Activate()
calc
GROSS.Caption = frmChallan.totalamount - frmChallan.totaldiscount
GROSS.Caption = Format(GROSS.Caption, "0.00")
Grid1.Font.Size = 10
Grid1.row = 1
Grid1.col = 0
If frmChallan.edit Then
  'Set rs = New ADODB.Recordset
  rs.Open "select * from challanctmp", CON, adOpenKeyset, adLockReadOnly, adCmdText
Else
   rs.Open "select * from challanEND where  " & stringyear & " ", CON, adOpenKeyset, adLockReadOnly, adCmdText
End If
If Grid1.Text <> "" Then
    rs.Find "TEXT='" + Trim(Grid1.Text) + "'"
    If Not rs.EOF And Not rs.BOF Then
    If rs!rate > 0 Then
        Grid1.col = 1
    Else
        Grid1.col = 2
    End If
    End If
End If
rs.Close
amount.Text = Grid1.Text
amount.Left = Grid1.CellLeft '+ 25
amount.TOP = Grid1.TOP + Grid1.CellTop - 15
amount.Visible = True
amount.ZOrder
lastrow = Grid1.row
lastcol = Grid1.col
amount.SetFocus
HIT
End Sub

 Sub Form_Load()
Grid1.TOP = 500
Grid1.Left = 20
mrow = 0
Me.Left = 0
Grid1.Rows = 2
Grid1.Cols = 0
Grid1.Cols = 3

If frmChallan.edit Then
   Set rs = New ADODB.Recordset
   rs.Open "select * from challanCTMP", CON, adOpenKeyset, adLockReadOnly, adCmdText
Else
  Set rs = New ADODB.Recordset
  rs.Open "select * from challanEND where  " & stringyear & " ORDER BY PRINTORDER", CON, adOpenKeyset, adLockReadOnly, adCmdText
End If
Grid1.col = 1
Grid1.row = 0
Grid1.Text = "Rate(%)if any"
Grid1.col = 2
Grid1.Text = "Amount"
Grid1.col = 0
Grid1.row = 1
Grid1.ColWidth(0) = 4000
Grid1.ColWidth(1) = 2000
Grid1.ColWidth(2) = 2000
otherabandon
If Not rs.EOF Then
r = rs.RecordCount
    Do While Not rs.EOF
        If mrow < Grid1.row Then
            mrow = Grid1.row
        End If
        If rs!Text <> "" Then
           Grid1.col = 0
           Grid1.Text = rs!Text
        End If
        If rs!rate <> 0 Then
            Grid1.col = 1
            Grid1.Text = rs!rate
            Grid1.col = 0
        End If
        If frmChallan.edit Then
            If rs!amount <> 0 Then
                 Grid1.col = 2
                Grid1.Text = Format(rs!amount, "0.00")
            End If
       End If
        If Not rs.EOF Then
            Grid1.Rows = Grid1.Rows + 1
            Grid1.row = Grid1.row + 1
            rs.MoveNext
        End If
    Loop
End If
rs.Close
Grid1.Font.Size = 10
''/*/*/////
Grid1.row = 1
Grid1.col = 0
If frmChallan.edit Then
rs.Open "select * from challanctmp", CON, adOpenKeyset, adLockReadOnly, adCmdText
Else
rs.Open "select * from challanEND where  " & stringyear & " ", CON, adOpenKeyset, adLockReadOnly, adCmdText

End If
If Grid1.Text <> "" Then
    rs.Find "TEXT='" + Trim(Grid1.Text) + "'"
    If rs!rate > 0 Then
        Grid1.col = 1
    Else
        Grid1.col = 2
    End If
End If
rs.Close
rs.Open "select * from challana where invoiceno=" + Trim(frmChallan.I_NO.Text), CON, adOpenKeyset, adLockReadOnly, adCmdText
If Not rs.EOF Then
   ChallanOther.T1 = Format(myround(rs!txt1a, 2), "0.00")
   ChallanOther.T1TEXT = IIf(IsNull(rs!txt1), "", rs!txt1)
   ChallanOther.T2 = Format(myround(rs!txt2a, 2), "0.00")
   ChallanOther.T2TEXT = IIf(IsNull(rs!txt2), "", rs!txt2)
   ChallanOther.T3TEXT = Format(myround(rs!baa, 2), "0.00")
End If
rs.Close
amount.Text = Format(Grid1.Text, "0.00")
amount.Left = Grid1.CellLeft '+ 25
amount.TOP = Grid1.TOP + Grid1.CellTop - 15
amount.Visible = True
amount.ZOrder
lastrow = Grid1.row
lastcol = Grid1.col
End Sub
Private Sub Grid1_Click()
If rs.State = 1 Then
    rs.Close
End If
If frmChallan.edit Then
rs.Open "select * from challanctmp", CON, adOpenKeyset, adLockReadOnly, adCmdText
Else
rs.Open "select * from challanEND where  " & stringyear & " ", CON, adOpenKeyset, adLockReadOnly, adCmdText

End If

Grid1.col = 0
    If Grid1.Text <> "" Then
        rs.Find "TEXT='" + Trim(Grid1.Text) + "'"
        If rs!rate > 0 Then
            Grid1.col = 1
        Else
            Grid1.col = 2
        End If
    Else
        Me.amount.Visible = False
        Grid1.SetFocus
        Exit Sub
    End If
rs.Close
    amount.Left = Grid1.CellLeft
    amount.TOP = Grid1.TOP + Grid1.CellTop - 15
    amount.Text = Format(Grid1.Text, "0.00")
    amount.Visible = True
    amount.ZOrder
    amount.SetFocus
    HIT
    lastrow = Grid1.row
    lastcol = Grid1.col
    
     
End Sub

Private Sub Label5_Click()

End Sub

Private Sub T1_GotFocus()

VB.Screen.ActiveForm.ActiveControl.SelLength = Len(VB.Screen.ActiveControl.Text)
End Sub

Private Sub T1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    T1.Text = Format(T1.Text, "0.00")
    SendKeys ("{TAB}")
Else
If KeyAscii >= 48 And KeyAscii <= 57 Then
Else
    If KeyAscii <> 46 Then
        If KeyAscii <> 8 And KeyAscii <> 45 Then
            KeyAscii = 0
        End If
    End If
End If
End If
End Sub

Private Sub T1_LostFocus()
calc
End Sub
Private Sub T1TEXT_GotFocus()
Dim i As Integer
VB.Screen.ActiveForm.ActiveControl.SelLength = Len(VB.Screen.ActiveControl.Text)
    frmChallan.otherdiscount = 0
    frmChallan.otheramount = 0
    For i = 1 To mrow
        Grid1.row = i
        If rs.State = 1 Then
            rs.Close
        End If
        If edit Then
                rs.Open "select * from challanctmp", CON, adOpenKeyset, adLockReadOnly, adCmdText
        Else
                rs.Open "select * from challanEND where  " & stringyear & " ", CON, adOpenKeyset, adLockReadOnly, adCmdText
        End If
        Grid1.col = 0
        If Grid1.Text <> "" Then
            rs.Find "TEXT='" + Trim(Grid1.Text) + "'"
            If Not rs.EOF Then
                If Trim(rs!DEBITORCREDIT) = Trim("Debit") Then
                    If rs!rate > 0 Then
                        Grid1.col = 1
                        frmChallan.otherdiscount = frmChallan.otherdiscount + ((frmChallan.totalamount - frmChallan.totaldiscount) * (Val(Grid1.Text) / 100))
                    Else
                        Grid1.col = 2
                        frmChallan.otherdiscount = frmChallan.otherdiscount + Val(Grid1.Text)
                    End If
                Else
                    If rs!rate > 0 Then
                        Grid1.col = 1
                        frmChallan.otheramount = frmChallan.otheramount + ((frmChallan.totalamount - frmChallan.totaldiscount) * (Val(Grid1.Text) / 100))
                    Else
                        Grid1.col = 2
                        frmChallan.otheramount = frmChallan.otheramount + Val(Grid1.Text)
                    End If
                End If
                End If
        End If
        rs.Close
    Next
End Sub

Private Sub T1TEXT_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys ("{TAB}")
End If
End Sub

Private Sub T1TEXT_LostFocus()
calc
End Sub

Private Sub T2_GotFocus()
VB.Screen.ActiveForm.ActiveControl.SelLength = Len(VB.Screen.ActiveControl.Text)
End Sub

Private Sub T2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    T2.Text = Format(T2.Text, "0.00")
    SendKeys ("{TAB}")
Else
    If KeyAscii >= 48 And KeyAscii <= 57 Then
    Else
        If KeyAscii <> 46 Then
            If KeyAscii <> 8 And KeyAscii <> 45 Then
                KeyAscii = 0
            End If
        End If
    End If
End If
End Sub
Private Sub T2_LostFocus()
calc
End Sub

Private Sub T2TEXT_GotFocus()
VB.Screen.ActiveForm.ActiveControl.SelLength = Len(VB.Screen.ActiveControl.Text)
End Sub

Private Sub T2TEXT_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys ("{TAB}")
End If
End Sub

Private Sub T3TEXT_GotFocus()
VB.Screen.ActiveForm.ActiveControl.SelLength = Len(VB.Screen.ActiveControl.Text)
End Sub

Private Sub T3TEXT_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    T3TEXT.Text = Format(T3TEXT.Text, "0.00")
       Me.CommandReturn.SetFocus
Else
    If KeyAscii >= 48 And KeyAscii <= 57 Then
    Else
        If KeyAscii <> 46 Then
            If KeyAscii <> 8 Then
                KeyAscii = 0
            End If
        End If
    End If
End If
End Sub
Private Sub T3TEXT_LostFocus()
frmChallan.labelbybank = T3TEXT
 calc
End Sub
