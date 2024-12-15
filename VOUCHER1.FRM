VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Voucherform 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6105
   ClientLeft      =   255
   ClientTop       =   1695
   ClientWidth     =   9510
   ClipControls    =   0   'False
   Icon            =   "Voucher1.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "<<"
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
      Left            =   240
      TabIndex        =   30
      Top             =   5160
      Width           =   675
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">>"
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
      Left            =   990
      TabIndex        =   29
      Top             =   5160
      Width           =   675
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid1 
      Height          =   4665
      Left            =   600
      TabIndex        =   4
      Top             =   450
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   8229
      _Version        =   393216
      Cols            =   5
      MergeCells      =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.ComboBox vtype 
      Height          =   315
      ItemData        =   "Voucher1.frx":000C
      Left            =   3000
      List            =   "Voucher1.frx":0019
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   60
      Width           =   705
   End
   Begin VB.PictureBox Picture1 
      Height          =   435
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   9585
      TabIndex        =   12
      Top             =   5520
      Width           =   9645
      Begin VB.CommandButton CommandReturn 
         Caption         =   "&Return"
         Height          =   375
         Left            =   7755
         TabIndex        =   20
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton CommandPrint 
         Caption         =   "&Print"
         Height          =   375
         Left            =   6810
         TabIndex        =   19
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton Commandsearch 
         Caption         =   "&Search"
         Height          =   375
         Left            =   5865
         TabIndex        =   18
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton Commanddelete 
         Caption         =   "De&lete"
         Height          =   375
         Left            =   4950
         TabIndex        =   17
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton Commandabandon 
         Caption         =   "Aba&ndon"
         Height          =   375
         Left            =   4050
         TabIndex        =   16
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton Commandsave 
         Caption         =   "Sa&ve"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         TabIndex        =   15
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton Commandedit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   2190
         TabIndex        =   14
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton Commandmasteradd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   1260
         TabIndex        =   0
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton Commandmasterhelp 
         Caption         =   "Help"
         Height          =   375
         Left            =   200
         TabIndex        =   13
         Top             =   0
         Width           =   800
      End
   End
   Begin MSMask.MaskEdBox vno 
      Height          =   255
      Left            =   8580
      TabIndex        =   3
      Top             =   90
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox vdate 
      Height          =   285
      Left            =   5460
      TabIndex        =   2
      Top             =   60
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medit 
      Height          =   285
      Left            =   3270
      TabIndex        =   21
      Top             =   3360
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   503
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox debit 
      Height          =   285
      Left            =   780
      TabIndex        =   8
      Top             =   4110
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      AutoTab         =   -1  'True
      Format          =   "###0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox crno 
      Height          =   285
      Left            =   7200
      TabIndex        =   10
      Top             =   3330
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      AutoTab         =   -1  'True
      MaxLength       =   20
      Format          =   "#,##0"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox credit 
      Height          =   255
      Left            =   780
      TabIndex        =   7
      Top             =   4650
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      AutoTab         =   -1  'True
      Format          =   "###0.00"
      PromptChar      =   "_"
   End
   Begin VB.ComboBox SubCombo 
      Height          =   1935
      ItemData        =   "Voucher1.frx":0026
      Left            =   3450
      List            =   "Voucher1.frx":0028
      Style           =   1  'Simple Combo
      TabIndex        =   6
      Top             =   990
      Width           =   2595
   End
   Begin VB.ComboBox Gencombo 
      Height          =   1935
      Left            =   810
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   5
      Top             =   1020
      Width           =   2595
   End
   Begin MSMask.MaskEdBox totalcredit 
      Height          =   315
      Left            =   6570
      TabIndex        =   11
      Top             =   5160
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox totaldebit 
      Height          =   285
      Left            =   5490
      TabIndex        =   9
      Top             =   5160
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox DESCRIPTION1 
      Height          =   255
      Left            =   2250
      TabIndex        =   26
      Top             =   3630
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      AutoTab         =   -1  'True
      Format          =   "#,##0"
      PromptChar      =   "_"
   End
   Begin VB.TextBox DESCRIPTION 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   2100
      MaxLength       =   50
      TabIndex        =   27
      Top             =   4110
      Width           =   1725
   End
   Begin VB.TextBox tmpsubledger 
      Height          =   285
      Left            =   900
      TabIndex        =   28
      Top             =   3330
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Label Label2 
      Caption         =   "Voucher Date"
      Height          =   255
      Left            =   3930
      TabIndex        =   25
      Top             =   90
      Width           =   1275
   End
   Begin VB.Label Label3 
      Caption         =   "Voucher No."
      Height          =   255
      Left            =   7110
      TabIndex        =   24
      Top             =   90
      Width           =   1275
   End
   Begin VB.Label Total 
      Caption         =   "Total :- "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4350
      TabIndex        =   23
      Top             =   5190
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "Voucher Type"
      Height          =   255
      Left            =   1500
      TabIndex        =   22
      Top             =   60
      Width           =   1275
   End
End
Attribute VB_Name = "Voucherform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim rs1 As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim lastrow, lastcol As Integer
Dim maxrow As Integer
Dim clicksave As Boolean
Dim tmplistindex As Long
Dim edit As Boolean
Dim addmode As Boolean
Dim searchmode As Boolean

Private Sub Command1_Click()



If Not rs2.EOF Then
    If rs2.EOF And rs2.RecordCount > 0 Then
           rs2.MoveLast
           vtype = rs2!vouchertype
           vdate = rs2!voucherdate
           vno = rs2!vouchernumber
           Me.vtype_LostFocus
           Me.vdate_LostFocus
           Me.vno_LostFocus
    
     End If
     vtype = rs2!vouchertype
     vdate = rs2!voucherdate
     vno = rs2!vouchernumber
     Me.vtype_LostFocus
    Me.vdate_LostFocus
    Me.vno_LostFocus
    rs2.MoveNext

End If




End Sub

Private Sub Command2_Click()

If Not rs2.BOF Then
    rs2.MovePrevious
    If rs2.BOF And rs2.RecordCount > 0 Then
           rs2.MoveFirst
           vtype = rs2!vouchertype
           vdate = rs2!voucherdate
           vno = rs2!vouchernumber
           Me.vtype_LostFocus
           Me.vdate_LostFocus
           Me.vno_LostFocus
    
     End If
     vtype = rs2!vouchertype
     vdate = rs2!voucherdate
     vno = rs2!vouchernumber
     Me.vtype_LostFocus
     Me.vdate_LostFocus
     Me.vno_LostFocus
     
End If



End Sub

Private Sub Commandabandon_Click()
'If SAVED Then
    For I = 1 To 99
        grid1.row = I
        For j = 0 To 4
            grid1.col = j
            grid1.Text = ""
        Next
    Next
    grid1.row = 0
    grid1.col = 0
    grid1.MergeCells = flexMergeRestrictRows
       For I = 1 To grid1.Rows - 1
        If I Mod 2 = 0 Then
            grid1.row = I
            For j = 0 To 4
                grid1.col = j
                grid1.Text = "."
                grid1.MergeRow(I) = True
            Next
        End If
    Next

    grid1.Refresh
    totalcredit = ""
    totaldebit = ""
    Gencombo.Visible = False
    SubCombo.Visible = False
    credit.Visible = False
    debit.Visible = False
    crno.Visible = False
    grid1.row = 1
    grid1.col = 0
    addmode = False
   ' rs.Open "Select max(VoucherNumber) from vouchers where vouchertype='" + Trim(vtype.Text) + "' and voucherdate= CDate ('" + Format(vdate.Text, "dd/mm/yy") + "')", con, adOpenKeyset, adLockReadOnly, adCmdText
   ' If Not rs.BOF Then
   '     If Not rs(0) Then
   '         c = rs(0)
    '    Else
     '       c = 0
      '  End If
   ' Else
   '     c = 0
   ' End If
   ' rs.Close
   ' vno.Text = Str(c + 1)
    'vno = Str(Val(vno) + 1)
  '  Grid1_Click
'End If
Commandsave.Enabled = False
Commandmasteradd.Enabled = True
Commandedit.Enabled = True
Commanddelete.Enabled = True
Commandsearch.Enabled = True
CommandPrint.Enabled = True
Gencombo.Clear
SubCombo.Clear
End Sub
Private Sub Commandedit_Click()
    Gencombo.Clear
    rs.Open "Select * from gledger order by gledger", con, adOpenKeyset, adLockReadOnly, adCmdText
    If Not rs.EOF Then
        Do While Not rs.EOF
            Gencombo.AddItem rs(1)
            If Not rs.EOF Then
                rs.MoveNext
            End If
        Loop
    End If
    rs.Close
    Commandedit.Enabled = False
    edit = True
    grid1.Enabled = True
    vtype.Enabled = True
    vdate.Enabled = True
    vtype.SetFocus
    vno.Enabled = True
    Commandsave.Enabled = True
Commandmasteradd.Enabled = False
Commandedit.Enabled = False
Commanddelete.Enabled = False
Commandsearch.Enabled = False
CommandPrint.Enabled = False
 
End Sub
Private Sub Commandmasteradd_Click()
    Commandabandon_Click
    addmode = True
    Set con = New ADODB.Connection
    Set rs1 = New ADODB.Recordset
    Set rs = New ADODB.Recordset
    con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + VB.App.Path + "\" + Trim(main.directory) + "\data.mdb"
    con.Open
   ' vdate.Text = Date
    rs1.Open "Select max(VoucherNumber) from vouchers where vouchertype='" + Trim(vtype.Text) + "' and voucherdate= CDate ('" + Format(Date, "dd/mm/yy") + "')", con, adOpenKeyset, adLockReadOnly, adCmdText
    If Not rs1.BOF Then
        If Not rs1(0) Then
            c = rs1(0)
        Else
            c = 0
        End If
    Else
        c = 0
    End If
    rs1.Close
    vno.Text = Str(c + 1)
    maxrow = 0
    Gencombo.Width = Gencombo.Width + 180
    SubCombo.Width = SubCombo.Width + 180
    debit.Width = credit.Width
    debit.Height = credit.Height
    crno.Width = credit.Width
    crno.Height = credit.Height
    medit.Width = Gencombo.Width
    medit.Height = Gencombo.Height
    Gencombo.Clear
    rs.Open "Select * from gledger order by gledger", con, adOpenKeyset, adLockReadOnly, adCmdText
    If Not rs.EOF Then
        Do While Not rs.EOF
            Gencombo.AddItem rs(1)
            If Not rs.EOF Then
                rs.MoveNext
            End If
        Loop
    End If
    rs.Close
    SubCombo.Clear
    rs.Open "Select * from sledger order by subledger", con, adOpenKeyset, adLockReadOnly, adCmdText
    If Not rs.EOF Then
        Do While Not rs.EOF
            SubCombo.AddItem rs(1)
            If Not rs.EOF Then
                rs.MoveNext
            End If
        Loop
    End If
    If rs.State = 1 Then
        rs.Close
    End If
    If rs1.State = 1 Then
        rs1.Close
    End If
    grid1.row = 0
    grid1.col = 0
    grid1.SelectionMode = flexSelectionFree
    grid1.ColWidth(0, 1) = 11111
    X = grid1.MergeCol(1)
    DESCRIPTION.Width = grid1.Width - 340
    DESCRIPTION.Height = crno.Height
vtype.Enabled = True
vdate.Enabled = True
vno.Enabled = True
grid1.Enabled = True
vtype.SetFocus
Commandsave.Enabled = True
Commandmasteradd.Enabled = False
Commandedit.Enabled = False
Commanddelete.Enabled = False
Commandsearch.Enabled = False
CommandPrint.Enabled = False
End Sub

Private Sub Commandreturn_Click()
    Unload Me
    MainMenu.Toolbar1.Visible = True
End Sub
Private Sub Commandsave_Click()
    Dim SAVED As Boolean
    SAVED = False
    If MsgBox("Do you want to save it now ?", vbYesNo) = vbNo Then
        Exit Sub
    End If
If Val(totalcredit.Text) = Val(totaldebit.Text) Then 'And Me.Commandedit.Enabled = True Then
    If rs.State = 1 Then
        rs.Close
    End If
    If Not edit Then
        rs.Open "vouchers", con, adOpenKeyset, adLockOptimistic, adCmdTable
    Else
        con.Execute ("delete * from vouchers where VoucherType='" + Trim(vtype.Text) + "' and voucherdate= CDate ('" + Trim(vdate.Text) + "') and vouchernumber=" + Trim(vno.Text) + "")
        'rs.Open "Select * from vouchers where VoucherType='" + Trim(vtype.TEXT) + "' and voucherdate= CDate ('" + Trim(vdate.TEXT) + "') and vouchernumber=" + Trim(vno.TEXT) + "", CON, adOpenKeyset, adLockOptimistic, adCmdText 'and VoucherNumber=" + Val(Trim(vno.Text)) + "", con, adOpenKeyset, adLockOptimistic, adCmdText
        rs.Open "vouchers", con, adOpenKeyset, adLockOptimistic, adCmdTable
    End If
    grid1.row = I
    grid1.col = 0
    'If grid1.TEXT <> "" And Commandedit.Enabled = False Then
    '    Do While Not rs.EOF
    '        rs.Delete
    '        If Not rs.EOF Then
    '            rs.MoveNext
    '        End If
    '    Loop
    'End If
    For I = 1 To 88 Step 2
        grid1.row = I
        grid1.col = 0
        If grid1.Text <> "" Then
            rs.AddNew
            rs(0) = vtype.Text
            rs(1) = vdate.Text
            rs(2) = vno.Text
            rs(3) = grid1.Text
            grid1.col = 1
            rs(4) = grid1.Text
            grid1.col = 2
            If grid1.Text <> "" Then
                rs(5) = Val(grid1.Text)
                rs(6) = "D"
            Else
                grid1.col = 3
                rs(5) = Val(grid1.Text)
                rs(6) = "C"
            End If
            grid1.col = 4
            rs(7) = grid1.Text
            grid1.col = 0
            grid1.row = grid1.row + 1
            If grid1.Text <> "" Then
                rs(9) = grid1.Text
            End If
            SAVED = True
            rs.Update
        End If
    Next
    CommandPrint.Visible = True
    rs.Close
    MsgBox "RECORD SAVED"
Commandsave.Enabled = False
Commandmasteradd.Enabled = True
Commandedit.Enabled = True
Commanddelete.Enabled = True
Commandsearch.Enabled = True
CommandPrint.Enabled = True
Else
  '  If Commandedit.Enabled = True Then
        MsgBox "Please Check That Total Credit and Total Debit is Differ "
   ' End If
End If





End Sub

Private Sub Commandsearch_Click()
    searchmode = True
'    Commandedit.Enabled = False
 '   edit = True
  
    Set con = New ADODB.Connection
    Set rs1 = New ADODB.Recordset
    Set rs = New ADODB.Recordset
    con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + VB.App.Path + "\" + Trim(main.directory) + "\data.mdb"
    con.Open

  vno.Enabled = True
    vtype.Enabled = True
    vdate.Enabled = True
    vtype.SetFocus
 'Me.Enabled = False
   ' searchscreen.grid1.row = 0
   ' searchscreen.grid1.col = 0
   ' Call searchscreen.tempr(12, "voucher")
End Sub

Private Sub credit_Change()
    If Val(credit.Text) <> 0 Then
            grid1.Text = Val(credit.Text)
            grid1.Text = Format(grid1.Text, "0.00")
    Else
    If grid1.col <> 4 Then grid1.Text = ""
    End If
End Sub
Private Sub credit_GotFocus()
    If maxrow < grid1.row Then
            maxrow = grid1.row
    End If
    If searchmode = False Then
        HIT
    End If
        
End Sub
Private Sub credit_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Val(credit) > 0 Or Val(debit) > 0 Then
    Dim changed As Boolean
    Dim HIT As Boolean
    HIT = False
    changed = False
    RRR = grid1.row
    CCC = grid1.col
    grid1.col = 3
    totalcredit.Text = ""
    For I = 1 To maxrow
        grid1.row = I
        totalcredit.Text = Val(totalcredit.Text) + Val(grid1.Text)
    Next
    totalcredit.Text = Format(totalcredit.Text, "0.00")
    grid1.row = RRR
    grid1.col = CCC
    If grid1.col = lastcol And grid1.row = lastrow Then
        grid1.col = 0
        If Trim(grid1.col) <> "" Then
            HIT = True
        End If
        grid1.col = 3
        grid1.col = grid1.col + 1
        changed = True
    End If
    RRR = grid1.row
    CCC = grid1.col
    grid1.col = 2
    If grid1.row <> lastrow Or grid1.col <> lastcol Then
        grid1.Text = ""
    End If
    grid1.row = RRR
    grid1.col = CCC
    If changed And HIT Then
        Grid1_Click
    End If
    Else
    credit.Visible = False
    grid1.col = 2
'    SendKeys "{left}{left}"
    Grid1_Click
    End If
End If
End Sub
Private Sub credit_LostFocus()
    credit = ""
    credit.Visible = False
End Sub
Private Sub Crno_Change()
'crno = UCase(crno)
'If grid1.col = lastcol And grid1.row = lastrow Then
    grid1.Text = crno.Text
'End If
End Sub

Private Sub crno_GotFocus()
    crno.Text = grid1.Text
If searchmode = False Then
    HIT
End If

End Sub

Private Sub Crno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim changed As Boolean
    Dim HIT As Boolean
    HIT = False
    changed = False
    
    If grid1.col = lastcol And grid1.row = lastrow Then
        crno = UCase(crno)
        grid1.col = 0
        If Trim(grid1.Text) <> "" Then
            HIT = True
        End If
        grid1.row = grid1.row + 1
        changed = True
    End If
    'crno.Text = ""
    If changed And HIT Then
        Grid1_Click
    End If
End If
End Sub

Private Sub Crno_LostFocus()
    crno.Visible = False
End Sub
Private Sub debit_Change()
If Val(debit.Text) <> 0 Then
    grid1.Text = Val(debit.Text)
    grid1.Text = Format(grid1.Text, "0.00")
Else
If grid1.col <> 4 Then grid1.Text = ""
End If
End Sub
Private Sub debit_GotFocus()
    If maxrow < grid1.row Then
        maxrow = grid1.row
    End If
    If searchmode = False Then
        HIT
        End If
        
End Sub
Private Sub debit_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim changed As Boolean
    Dim HIT As Boolean
    changed = False
    HIT = False
    RRR = grid1.row
    CCC = grid1.col
    grid1.col = 2
    totaldebit.Text = ""
    For I = 1 To maxrow
        grid1.row = I
        totaldebit.Text = Val(totaldebit.Text) + Val(grid1.Text)
    Next
    totaldebit.Text = Format(totaldebit.Text, "0.00")
    grid1.row = RRR
    grid1.col = CCC
    If grid1.col = lastcol And grid1.row = lastrow Then
        grid1.col = 0
        If Trim(grid1.col) <> "" Then
            HIT = True
        End If
        grid1.col = 2
        If Val(grid1.Text) <> 0 Then
            grid1.col = 4
        Else
            grid1.col = 3
        End If
        changed = True
    End If
    RRR = grid1.row
    CCC = grid1.col
    grid1.col = 2
    If grid1.row <> lastrow Or grid1.col <> lastcol Then
        grid1.Text = ""
    End If
    grid1.row = RRR
    grid1.col = CCC
    If changed And HIT Then
        Grid1_Click
    End If
End If
End Sub

Private Sub debit_LostFocus()
RRR = grid1.row
CCC = grid1.col
grid1.row = lastrow
grid1.col = lastcol
If Trim(grid1.Text) <> "" Then
    grid1.row = RRR
    grid1.col = CCC
    debit.Visible = False
Else
'    Grid1.col = Grid1.col + 1
'    debit.Visible = False
'    Grid1_Click
End If
debit = ""
End Sub

Private Sub Description_Change()
  
   ' If grid1.col = lastcol And grid1.row = lastrow Then
   '     grid1.Text = DESCRIPTION.Text
   ' End If

End Sub

Private Sub DESCRIPTION_GotFocus()
'    grid1.col = 0
    If Trim(grid1.Text) <> Trim(".") Then
         DESCRIPTION.Text = grid1.Text
    End If
     HIT
End Sub
Private Sub Description_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DESCRIPTION = UCase(DESCRIPTION)
        grid1.Text = DESCRIPTION.Text
    Dim changed As Boolean
    Dim HIT As Boolean
    HIT = False
    changed = False
    If grid1.col = lastcol And grid1.row = lastrow Then
        grid1.col = 0
        HIT = True
        grid1.row = grid1.row + 1
        changed = True
    End If
    'DESCRIPTION.Text = ""
    If changed And HIT Then
        Grid1_Click
    End If
    End If
End Sub

Private Sub DESCRIPTION_LostFocus()
    DESCRIPTION = UCase(DESCRIPTION)
    DESCRIPTION.Visible = False
End Sub

Private Sub Form_Activate()
Command1.SetFocus

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 26 Then
   Command1_Click
  End If
  


End Sub

Private Sub Gencombo_Change()
'    Gencombo = UCase(Gencombo)
    If grid1.col = lastcol And grid1.row = lastrow Then
        grid1.Text = Gencombo.Text
        grid1.col = 1
        grid1.Text = ""
        grid1.col = 0
    End If
End Sub
Private Sub genCombo_Click()
    grid1.Text = Gencombo.Text
    grid1.col = 1
    grid1.Text = ""
    grid1.col = 0
End Sub
Private Sub Form_Load()




searchmode = False
addmode = False
vtype.Enabled = False
vdate.Enabled = False
vno.Enabled = False
Commandsearch.Enabled = True
grid1.Enabled = False
edit = False
    grid1.Top = 450
    grid1.Left = 60
    Me.Top = 0
    Me.Left = 0
    grid1.Rows = 100
    grid1.Cols = 0
    grid1.Cols = 5
    For I = 0 To 99
        grid1.RowHeight(I) = 250
    Next
    grid1.row = 0
    grid1.col = 0
    grid1.MergeCells = flexMergeRestrictRows
       For I = 1 To grid1.Rows - 1
        If I Mod 2 = 0 Then
            grid1.row = I
            For j = 0 To 4
                grid1.col = j
                grid1.Text = "."
                grid1.MergeRow(I) = True
            Next
        End If
    Next
        'Grid1.MergeCells = flexMergeRestrictRows
        grid1.row = 0
        grid1.col = 0
    grid1.ColAlignment(0) = 0
    grid1.ColAlignment(1) = 0
    grid1.ColWidth(0) = 2750
    grid1.ColWidth(1) = 2750
    grid1.ColWidth(2) = 1160
    grid1.ColWidth(3) = 1160
    grid1.ColWidth(4) = 1160
    grid1.Text = "Gen. Ledger"
    grid1.col = 1
    grid1.Text = "Sub. Ledger"
    grid1.col = 2
    grid1.Text = "Amount. (Dr.)"
    totaldebit.Left = grid1.CellLeft + 50
    grid1.col = 3
    grid1.Text = "Amount. (Cr.)"
    totalcredit.Left = grid1.CellLeft + 60
    grid1.col = 4
    grid1.Text = "C/B No."
    totalcredit.Width = crno.Width
    totaldebit.Width = totalcredit.Width
    totaldebit.Height = crno.Height + 10
    totalcredit.Height = totaldebit.Height
'   Load first voucher start
    Set con = New ADODB.Connection
    con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + VB.App.Path + "\" + Trim(main.directory) + "\data.mdb"
    con.Open

    Set rs = New ADODB.Recordset
    rs.Open "vouchers", con, adOpenKeyset, adLockOptimistic, adCmdTable
    If Not rs.BOF Then
       vtype = rs!vouchertype
       vdate = rs!voucherdate
       vno = rs!vouchernumber
       Me.vtype_LostFocus
       Me.vdate_LostFocus
       Me.vno_LostFocus
    End If
  
Set rs2 = New ADODB.Recordset
rs2.Open "Select DISTINCT vouchertype,voucherdate,vouchernumber from vouchers order by vouchertype,voucherdate,vouchernumber", con, adOpenStatic
  
  
  '  rs.Close
' Load first voucher end
End Sub
Private Sub Gencombo_GotFocus()
    grid1.col = 0
    Gencombo.Text = grid1.Text
    If Gencombo.Enabled = True Then
    If searchmode = False Then
    HIT
    End If
    End If
End Sub
Private Sub Gencombo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim HIT As Boolean
        Dim changed As Boolean
        changed = False
        HIT = False
        rs.Open "GLEDGER", con, adOpenKeyset, adLockReadOnly, adCmdTable
        If Trim(Gencombo.Text) <> "" Then
        SendKeys "{DOWN}"
           If Not rs.BOF Then
                rs.MoveFirst
                rs.Find "GLEDGER='" + Trim(Gencombo.Text) + "'"
                If rs.EOF Then
                    rs.Close
                    Exit Sub
                End If
            End If
        Else
           Gencombo.Visible = False
        End If
        rs.Close
        If grid1.col = lastcol And grid1.row = lastrow Then
            If Trim(grid1.Text) <> "" Then
                HIT = True
            End If
            grid1.col = grid1.col + 1
            changed = True
        End If
        Gencombo.Text = ""
        If changed And HIT Then
            Grid1_Click
       Else
            If changed Then
           Commandsave.SetFocus
            End If
        End If

    End If
End Sub

Private Sub Gencombo_LostFocus()
    rs.Open "GLEDGER", con, adOpenKeyset, adLockReadOnly, adCmdTable
    If Trim(Gencombo.Text) <> "" Then
        If Not rs.BOF Then
            rs.MoveFirst
            rs.Find "GLEDGER='" + Trim(Gencombo.Text) + "'"
            If rs.EOF Then
                Gencombo.Visible = True
                Gencombo.SetFocus
            Else
            
                Gencombo.Visible = False
               Me.Commandsave.SetFocus
            End If
        End If
    Else
        Gencombo.Visible = False
        
    End If
    rs.Close
End Sub
Private Sub Grid1_Click()
If grid1.Enabled = True Then
Dim check As Boolean
Dim row, col As Integer
row = grid1.row
col = grid1.col
check = True
If grid1.row Mod 2 <> 0 Then
    DESCRIPTION.Visible = False
    Dim gr As MSHFlexGrid
    Set gr = grid1
    If grid1.row >= 3 Then
        check = False
        gr.row = grid1.row - 2
        gr.col = 0
        If gr.Text <> "" Then
            gr.col = 2
            If gr.Text <> "" Then
                check = True
            Else
                gr.col = 3
                If gr.Text <> "" Then
                    check = True
                End If
            End If
        End If
    End If
    grid1.row = row
    grid1.col = col
    credit.Visible = False
    debit.Visible = False
    SubCombo.Visible = False
    crno.Visible = False
    Gencombo.Visible = False
    If check Then
        If grid1.col = 0 Then
            credit.Visible = False
            debit.Visible = False
            SubCombo.Visible = False
            crno.Visible = False
            Gencombo.Text = grid1.Text
            Gencombo.Visible = True
            Gencombo.Left = grid1.CellLeft + 40
            Gencombo.Top = grid1.Top + grid1.CellTop - 50
            Gencombo.ZOrder
            Gencombo.SetFocus
        End If
        If grid1.col = 1 Then
        If grid1.Text <> "" Then
            tmpsubledger = grid1.Text
        End If
            Gencombo.Visible = False
            credit.Visible = False
            debit.Visible = False
            crno.Visible = False
            SubCombo.Visible = True
            SubCombo.Text = grid1.Text
            SubCombo.Left = grid1.CellLeft + 40
            SubCombo.Top = grid1.Top + grid1.CellTop - 50
            SubCombo.ZOrder
            SubCombo.SetFocus
        End If
        If grid1.col = 2 Then
            Gencombo.Visible = False
            SubCombo.Visible = False
            credit.Visible = False
            crno.Visible = False
            grid1.col = 3
            If Trim(grid1.Text) = "" Then
                grid1.col = 2
                If grid1.Text <> "" Then
                debit.Text = grid1.Text
                End If
                debit.Visible = True
                debit.Left = grid1.CellLeft + 40
                debit.Top = grid1.Top + grid1.CellTop - 25
                debit.ZOrder
                debit.SetFocus
            End If
            grid1.col = 2
        End If
        If grid1.col = 3 Then
            Gencombo.Visible = False
            SubCombo.Visible = False
            debit.Visible = False
            crno.Visible = False
            grid1.col = 2
            If Trim(grid1.Text) = "" Then
                grid1.col = 3
                If grid1.Text <> "" Then
                credit.Text = grid1.Text
                End If
                credit.Visible = True
                credit.Left = grid1.CellLeft + 40
                credit.Top = grid1.Top + grid1.CellTop - 25
                credit.ZOrder
                credit.SetFocus
            End If
        End If
        If grid1.col = 4 Then
            Gencombo.Visible = False
            SubCombo.Visible = False
            debit.Visible = False
            credit.Visible = False
            crno.Text = grid1.Text
            crno.Visible = True
            crno.Left = grid1.CellLeft + 40
            crno.Top = grid1.Top + grid1.CellTop - 25
            crno.ZOrder
            crno.SetFocus
        End If
    lastrow = grid1.row
    lastcol = grid1.col
    Else
    MsgBox "Previous entry didn't complete"
    End If
Else
    If grid1.row >= 2 Then
        RRR = grid1.row
        CCC = grid1.col
        grid1.row = grid1.row - 1
        grid1.col = 0
        If Trim(grid1.Text) <> "" Then
            grid1.row = RRR
            grid1.col = CCC
            If grid1.col = 0 Then
                DESCRIPTION.Left = grid1.CellLeft + 40
                DESCRIPTION.Top = grid1.Top + grid1.CellTop - 25
                crno.Visible = False
                Gencombo.Visible = False
                SubCombo.Visible = False
                credit.Visible = False
                debit.Visible = False
                DESCRIPTION.Visible = True
                DESCRIPTION.ZOrder
                DESCRIPTION.SetFocus
            End If
        End If
        lastrow = grid1.row
        lastcol = grid1.col
    End If
End If
End If
End Sub

Private Sub SubCombo_Change()
   ' SubCombo = UCase(SubCombo)
    If grid1.col = lastcol And grid1.row = lastrow Then
        grid1.Text = SubCombo.Text
    End If
End Sub
Private Sub SubCombo_Click()
    grid1.Text = SubCombo.Text
End Sub
Private Sub SubCombo_GotFocus()
Dim X(5000) As String
    SubCombo.Clear
    Set rs = New ADODB.Recordset
    rs.Open "SLEDGER", con, adOpenKeyset, adLockReadOnly, adCmdTable
    grid1.col = 0
    rs.Find "gledger='" + Trim(grid1.Text) + "'"
    If Not rs.EOF Then
    Dim x1  As Integer
            x1 = 0
            SubCombo.Clear
         Do While Not rs.EOF
            X(x1) = rs(1)
            SubCombo.AddItem X(x1)
            If rs(1) = tmpsubledger Then
            tmplistindex = x1
            End If
            If Not rs.EOF Then
            x1 = x1 + 1
                rs.MoveNext
            End If
        Loop
    End If
    grid1.col = CCC
    grid1.col = 1
    If SubCombo.ListCount = 0 Then
        grid1.col = 2
        Grid1_Click
    End If
      '  SubCombo = tmpsubledger
      If tmplistindex >= 1 Then
SubCombo.ListIndex = tmplistindex - 1
End If

End Sub
Private Sub SubCombo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim changed As Boolean
    Dim valid As Boolean
    Dim HIT As Boolean
    changed = False
    HIT = False
    valid = False
    RRR = grid1.row
    CCC = grid1.col
    grid1.row = lastrow
    grid1.col = 0
    If rs.State = 1 Then
        rs.Close
    End If
    If Trim(SubCombo.Text) <> "" And Trim(grid1.Text) <> "" Then
        SendKeys "{down}"
        rs.Open "select * from SLEDGER where subledger='" + Trim(SubCombo.Text) + "' and gledger='" + Trim(grid1.Text) + "'", con, adOpenKeyset, adLockReadOnly, adCmdText
        If rs.BOF Or rs.EOF Then
            SubCombo.Visible = True
            SubCombo.SetFocus
        Else
            valid = True
        End If
        rs.Close
    Else
        If Trim(SubCombo.Text) = "" Then
            SubCombo.Visible = True
            SubCombo.SetFocus
        End If
    End If
    grid1.row = RRR
    grid1.col = CCC
    If valid Then
        If grid1.col = lastcol And grid1.row = lastrow Then
            grid1.col = 0
            If Trim(grid1.col) <> "" Then
                HIT = True
            End If
            grid1.col = 3
            If grid1.Text <> "" Then
            
            Else
                grid1.col = 2
            End If
            changed = True
        End If
            SubCombo.Text = ""
        If changed And HIT Then
            Grid1_Click
        End If
    End If
End If
End Sub
Private Sub SubCombo_LostFocus()
Dim RRR, CCC As Integer
    If rs.State = 1 Then
        rs.Close
    End If
    RRR = grid1.row
    CCC = grid1.col
    grid1.row = lastrow
    grid1.col = 0
    If Trim(SubCombo.Text) <> "" And Trim(grid1.Text) <> "" Then
        rs.Open "select * from SLEDGER where subledger='" + Trim(SubCombo.Text) + "' and gledger='" + Trim(grid1.Text) + "'", con, adOpenKeyset, adLockReadOnly, adCmdText
        If rs.BOF Or rs.EOF Then
            SubCombo.Visible = True
            SubCombo.SetFocus
        Else
            SubCombo.Visible = False
        End If
        rs.Close
    Else
        If Trim(SubCombo.Text) <> "" Then
            SubCombo.Visible = True
            SubCombo.SetFocus
        End If
    End If
    grid1.row = RRR
    grid1.col = CCC
    
End Sub

'Private Sub totalcredit_Change()
'    totaldebit = Format(totaldebit, "0.00")
'End Sub

Private Sub totaldebit_Change()
 '   totaldebit = Format(totaldebit, "0.00")
End Sub

Private Sub vdate_GotFocus()
HIT
End Sub

Private Sub vdate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If addmode = True Or edit = True Then
        grid1.row = 1
        grid1.col = 0
        grid1.SetFocus
        Grid1_Click
        Else
        SendKeys "{tab}"
        End If
End If
End Sub
 Sub vdate_LostFocus()
    If Trim(vdate.Text) <> "" Then
        If Not checkdate(Trim(vdate.Text), vdate) Then
            If vdate.Enabled = True Then vdate.SetFocus
            Exit Sub
        End If
    End If
    If addmode = True Then
        If rs.State = 1 Then
            rs.Close
        End If
        Set rs = New ADODB.Recordset
        rs.Open "Select max(VoucherNumber) from vouchers where vouchertype='" + Trim(vtype.Text) + "' and voucherdate= CDate ('" + Format(vdate.Text, "dd/mm/yy") + "')", con, adOpenKeyset, adLockReadOnly, adCmdText
        If Not rs.BOF Then
            If Not rs(0) Then
                c = rs(0)
            Else
                c = 0
            End If
        Else
            c = 0
        End If
        rs.Close
        vno.Text = Str(c + 1)
    End If
End Sub

Private Sub vno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys ("{TAB}")
End If
End Sub
Sub vno_LostFocus()
If Val(vno) <> 0 Then
    If rs.State = 1 Then
        rs.Close
    End If
    rs.Open "Select * from vouchers where VoucherType='" + Trim(vtype.Text) + "' and voucherdate= CDate ('" + Trim(vdate.Text) + "') and vouchernumber=" + Trim(vno.Text) + "", con, adOpenKeyset, adLockOptimistic, adCmdText 'and VoucherNumber=" + Val(Trim(vno.Text)) + "", con, adOpenKeyset, adLockOptimistic, adCmdText
    For I = 1 To 99
        grid1.row = I
        For j = 0 To 4
            grid1.col = j
            grid1.Text = ""
        Next
    Next
    
    If Not rs.BOF Then
        I = 1
       ' For i = 1 To maxrow Step 2
        Do While Not rs.EOF
            grid1.row = I
            grid1.col = 0
            grid1.Text = rs(3)
            grid1.col = 1
            grid1.Text = rs(4)
            grid1.col = 2
            If rs(6) = "D" Then
                grid1.Text = rs(5)
            Else
                grid1.col = 3
                grid1.Text = rs(5)
            End If
                grid1.col = 4
                grid1.Text = rs(7)
                grid1.row = grid1.row + 1
                grid1.col = 0
                grid1.Text = rs(9)
           I = I + 2
            If Not rs.EOF Then
                rs.MoveNext
            End If
        Loop
        totalcredit = ""
        totaldebit = ""
        Gencombo.Visible = False
        SubCombo.Visible = False
        credit.Visible = False
        debit.Visible = False
        crno.Visible = False
        grid1.row = 1
        grid1.col = 0
    End If
    rs.Close
    grid1.row = 1
    grid1.col = 0
    If searchmode = False Then
        Grid1_Click
    End If
    Else
    vno.SetFocus
    End If
    
End Sub
Private Sub vtype_Change()
 
        If rs.State = 1 Then
            rs.Close
        End If
        Set rs = New ADODB.Recordset
        rs.Open "Select max(VoucherNumber) from vouchers where vouchertype='" + Trim(vtype.Text) + "' and voucherdate= CDate ('" + Format(vdate.Text, "dd/mm/yy") + "')", con, adOpenKeyset, adLockReadOnly, adCmdText
        If Not rs.BOF Then
            If Not rs(0) Then
                c = rs(0)
            Else
                c = 0
            End If
        Else
            c = 0
        End If
        rs.Close
        vno.Text = Str(c + 1)
End Sub

Private Sub vtype_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.vdate.SetFocus
    End If
End Sub
Public Function getval()
    Grid1_Click
End Function

 Sub vtype_LostFocus()
 ' If rs.State = 1 Then
 '           rs.Close
  '      End If
   '     Set rs = New ADODB.Recordset
    '    rs.Open "Select max(VoucherNumber) from vouchers where vouchertype='" + Trim(vtype.Text) + "' and voucherdate= CDate ('" + Format(vdate.Text, "dd/mm/yy") + "')", con, adOpenKeyset, adLockReadOnly, adCmdText
     '   If Not rs.BOF Then
      '      If Not rs(0) Then
       '         c = rs(0)
        '    Else
         '       c = 0
    '        End If
  '      Else
   '         c = 0
     '   End If
      '  rs.Close
       ' vno.Text = Str(c + 1)
End Sub
