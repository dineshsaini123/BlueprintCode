VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Voucher2 
   Caption         =   "Voucher2"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   9390
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Description 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   22
      Top             =   360
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.TextBox Crno 
      Appearance      =   0  'Flat
      Height          =   280
      Left            =   6300
      TabIndex        =   21
      Top             =   360
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.ComboBox SubCombo 
      Height          =   1935
      ItemData        =   "voucher2.frx":0000
      Left            =   2730
      List            =   "voucher2.frx":0002
      Style           =   1  'Simple Combo
      TabIndex        =   17
      Top             =   930
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.ComboBox Gencombo 
      Height          =   1935
      Left            =   210
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   20
      Top             =   930
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.TextBox CrText 
      Appearance      =   0  'Flat
      Height          =   280
      Left            =   5130
      TabIndex        =   19
      Top             =   360
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.TextBox DrText 
      Appearance      =   0  'Flat
      Height          =   280
      Left            =   3930
      TabIndex        =   18
      Top             =   360
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.PictureBox Picture1 
      Height          =   435
      Left            =   210
      ScaleHeight     =   375
      ScaleWidth      =   9585
      TabIndex        =   7
      Top             =   5250
      Width           =   9645
      Begin VB.CommandButton Commandmasterhelp 
         Caption         =   "Help"
         Height          =   375
         Left            =   200
         TabIndex        =   16
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton Commandmasteradd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   1260
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
      Begin VB.CommandButton Commandsave 
         Caption         =   "Sa&ve"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         TabIndex        =   13
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton Commandabandon 
         Caption         =   "Aba&ndon"
         Height          =   375
         Left            =   4050
         TabIndex        =   12
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton Commanddelete 
         Caption         =   "De&lete"
         Height          =   375
         Left            =   4950
         TabIndex        =   11
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton Commandsearch 
         Caption         =   "&Search"
         Height          =   375
         Left            =   5880
         TabIndex        =   10
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton CommandPrint 
         Caption         =   "&Print"
         Height          =   375
         Left            =   6810
         TabIndex        =   9
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton CommandReturn 
         Caption         =   "&Return"
         Height          =   375
         Left            =   7755
         TabIndex        =   8
         Top             =   0
         Width           =   800
      End
   End
   Begin VB.ComboBox vtype 
      Height          =   315
      ItemData        =   "voucher2.frx":0004
      Left            =   2640
      List            =   "voucher2.frx":0011
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   30
      Width           =   705
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
      Height          =   3975
      Left            =   150
      TabIndex        =   0
      Top             =   660
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   7011
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin MSMask.MaskEdBox vno 
      Height          =   255
      Left            =   8220
      TabIndex        =   2
      Top             =   60
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox vdate 
      Height          =   285
      Left            =   5100
      TabIndex        =   3
      Top             =   30
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      Caption         =   "Voucher Type"
      Height          =   255
      Left            =   1140
      TabIndex        =   6
      Top             =   30
      Width           =   1275
   End
   Begin VB.Label Label3 
      Caption         =   "Voucher No."
      Height          =   255
      Left            =   6750
      TabIndex        =   5
      Top             =   60
      Width           =   1275
   End
   Begin VB.Label Label2 
      Caption         =   "Voucher Date"
      Height          =   255
      Left            =   3570
      TabIndex        =   4
      Top             =   60
      Width           =   1275
   End
End
Attribute VB_Name = "Voucher2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs As ADODB.Recordset
Dim LCC As Integer
Dim LRC As Integer
Sub GridFill()
    Dim count As Integer
    If rs.State = 1 Then
        rs.Close
    End If
    
    rs.Open "Select * from vouchers where VoucherType='" + Trim(vtype.Text) + "' and voucherdate= CDate ('" + Trim(vdate.Text) + "') and vouchernumber=" + Trim(vno.Text) + "", CON1, adOpenKeyset, adLockOptimistic, adCmdText
    
    'Grid1.Clear
    If rs.RecordCount > 0 Then
    rs.MoveFirst
    count = 1
    
    For i = 1 To 99
         For j = 0 To 4
         
             If j = 0 Then Grid1.TextMatrix(i, j) = rs(3)
             If j = 1 Then Grid1.TextMatrix(i, j) = rs(4)
             If rs(6) = "D" Then
                If j = 2 Then Grid1.TextMatrix(i, j) = rs(5)
             Else
                If j = 3 Then Grid1.TextMatrix(i, j) = rs(5)
             End If
             
             If j = 4 Then Grid1.TextMatrix(i, j) = rs(7)
             
        Next j
        j = 0
        i = i + 1
        Grid1.row = i
        Grid1.col = j
        Grid1.Text = rs(9)
        count = count + 1
        If count <= rs.RecordCount Then
           rs.MoveNext
        Else
           Exit For
        End If
      Next i
  End If



End Sub

Private Sub Commanddelete_Click()
Unload Me
Voucher2.Show
End Sub

Private Sub Commandmasteradd_Click()
    Dim rs1 As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    LRC = 1
    LCC = 0
    rs1.Open "Select max(VoucherNumber) from vouchers where vouchertype='" + Trim(vtype.Text) + "' and voucherdate= CDate ('" + Format(Date, "dd/mm/yy") + "')", CON1, adOpenKeyset, adLockReadOnly, adCmdText
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
    
   
    rs.Close
    SubCombo.Clear
    rs.Open "Select * from sledger order by subledger", CON1, adOpenKeyset, adLockReadOnly, adCmdText
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
    Grid1.row = 0
    Grid1.col = 0
    Grid1.SelectionMode = flexSelectionFree
    Grid1.ColWidth(0, 1) = 11111
    X = Grid1.MergeCol(1)
    Description.Width = Grid1.Width - 340
    Description.Height = Crno.Height
vtype.Enabled = True
vdate.Enabled = True
vno.Enabled = True
Grid1.Enabled = True
vtype.SetFocus
Commandsave.Enabled = True
Commandmasteradd.Enabled = False
Commandedit.Enabled = False
Commanddelete.Enabled = False
Commandsearch.Enabled = False
CommandPrint.Enabled = False











End Sub

Private Sub CommandPrint_Click()

GridFill
End Sub


Private Sub Crno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Grid1.Text = Crno.Text
     Grid1.row = Grid1.row + 1
     LRC = LRC + 1
     
     Grid1.col = 0
     Grid1_Click
End If
End Sub

Private Sub Crno_LostFocus()
Crno.Text = ""
End Sub

Private Sub CrText_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   
   If Val(CrText.Text) = 0 Or CrText.Text = "" Then
        Grid1.Text = ""
        Grid1.col = Grid1.col - 1
        Grid1_Click
   Else
       Grid1.Text = CrText.Text
       Grid1.col = Grid1.col + 1
       
       Grid1_Click
   End If

End If
End Sub

Private Sub CrText_LostFocus()
CrText.Text = ""
End Sub

Private Sub Description_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Grid1.Text = Description.Text
     Grid1.row = Grid1.row + 1
     LRC = LRC + 1
     Grid1.col = 0
     Grid1_Click
   End If
End Sub

Private Sub DrText_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   
   If Val(DrText.Text) = 0 Or DrText.Text = "" Then
        Grid1.Text = ""
        Grid1.col = Grid1.col + 1
        Grid1_Click
   Else
     Grid1.Text = DrText.Text
     Grid1.col = Grid1.col + 2
     Grid1_Click
   End If

End If
End Sub

Private Sub DrText_LostFocus()
DrText.Text = ""
End Sub

Private Sub Form_Load()
    Dim rs2 As New ADODB.Recordset
    Set rs = New ADODB.Recordset
    LRC = 1
    Grid1.Top = 450
    Grid1.Left = 60

    Grid1.Rows = 100
    Grid1.Cols = 5

    Me.Top = 0
    Me.Left = 0
    
    
    Grid1.row = 0
    Grid1.col = 0
    Grid1.MergeCells = flexMergeRestrictRows
       For i = 2 To Grid1.Rows - 1
        If i Mod 2 = 0 Then
            Grid1.row = i
            For j = 0 To 4
                Grid1.col = j
                Grid1.Text = "."
                Grid1.MergeRow(i) = True
            Next
        End If
    Next
    
    Grid1.row = 0
    Grid1.col = 0
    Grid1.ColAlignment(0) = 0
    Grid1.ColAlignment(1) = 0
    Grid1.ColWidth(0) = 2750
    Grid1.ColWidth(1) = 2750
    Grid1.ColWidth(2) = 1160
    Grid1.ColWidth(3) = 1160
    Grid1.ColWidth(4) = 1160
    Grid1.Text = "Gen. Ledger"
    Grid1.col = 1
    Grid1.Text = "Sub. Ledger"
    Grid1.col = 2
    Grid1.Text = "Amount. (Dr.)"
    'totaldebit.Left = Grid1.CellLeft + 50
    Grid1.col = 3
    Grid1.Text = "Amount. (Cr.)"
    'totalcredit.Left = Grid1.CellLeft + 60
    Grid1.col = 4
    Grid1.Text = "C/B No."

    
       ' totalcredit.Width = crno.Width
   ' totaldebit.Width = totalcredit.Width
   ' totaldebit.Height = crno.Height + 10
    'totalcredit.Height = totaldebit.Height
    main.OpenCon
    Set rs = New ADODB.Recordset
    rs.Open "vouchers", main.CON1, adOpenKeyset, adLockOptimistic, adCmdTable
    If Not rs.BOF Then
       vtype = rs!vouchertype
       vdate = rs!voucherdate
       vno = rs!VoucherNumber
       'Me.vtype_LostFocus
       'Me.vdate_LostFocus
       'vno_LostFocus
       vno_LostFocus
    End If

    Gencombo.Clear
    rs2.Open "Select * from gledger order by gledger", CON1, adOpenStatic, adLockOptimistic
    If Not rs2.EOF Then
        Do While Not rs2.EOF
            Gencombo.AddItem rs2(1)
            If Not rs2.EOF Then
                rs2.MoveNext
            End If
        Loop
    End If


End Sub

Private Sub Gencombo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
    If Gencombo.Text = "" Then
         MsgBox "Please Enter Gen. led......"
         Gencombo.SetFocus
         Exit Sub
    End If
    Dim rs2 As New ADODB.Recordset
    Dim rs3 As New ADODB.Recordset
    
    Grid1.Text = Gencombo.Text
    SubCombo.Clear
    rs2.Open "Select * from sledger where GLEDGER='" + Trim(Gencombo.Text) + "'", CON1, adOpenStatic, adLockReadOnly, adCmdText
       Grid1.row = LRC
       Grid1.col = 0
       If rs2.RecordCount > 0 Then
            
            rs2.MoveFirst
            Do While Not rs2.EOF
                    SubCombo.AddItem rs2(1)
                    If Not rs2.EOF Then
                        rs2.MoveNext
                    End If
            Loop
            Grid1.col = Grid1.col + 1
            Grid1_Click
         Else
           
           Grid1.col = Grid1.col + 2
           Gencombo.Visible = False
           Grid1_Click
           
         End If
           
  

                
 End If
   
   




End Sub




Private Sub Grid1_Click()
If Grid1.row Mod 2 <> 0 Then
       Description.Visible = False
        Select Case Grid1.col
    
            Case 0
                Gencombo.Text = Grid1.Text
                Gencombo.Visible = True
                SubCombo.Visible = False
                DrText.Visible = False
                CrText.Visible = False
                Crno.Visible = False
                Gencombo.Move Grid1.CellLeft + Grid1.Left, Grid1.CellTop + Grid1.Top - 15, Grid1.CellWidth
                Gencombo.SetFocus
            Case 1
                SubCombo.Text = Grid1.Text
                Gencombo.Visible = False
                Crno.Visible = False
                DrText.Visible = False
                CrText.Visible = False

                If SubCombo.ListCount > 0 Then
                        SubCombo.Text = Grid1.Text
                        SubCombo.Visible = True
                        SubCombo.Move Grid1.CellLeft + Grid1.Left, Grid1.CellTop + Grid1.Top - 15, Grid1.CellWidth
                        SubCombo.SetFocus
                End If
       
 
            Case 2
                DrText.Text = Grid1.Text
                DrText.Visible = True
                Gencombo.Visible = False
                SubCombo.Visible = False
                CrText.Visible = False
                Crno.Visible = False
       
                DrText.Move Grid1.CellLeft + Grid1.Left, Grid1.CellTop + Grid1.Top - 15, Grid1.CellWidth
                DrText.SetFocus
 
            Case 3
                CrText.Text = Grid1.Text
                CrText.Visible = True
                Gencombo.Visible = False
                SubCombo.Visible = False
                DrText.Visible = False
                Crno.Visible = False
       
                CrText.Move Grid1.CellLeft + Grid1.Left, Grid1.CellTop + Grid1.Top - 15, Grid1.CellWidth
                CrText.SetFocus
            Case 4
                Crno.Text = Grid1.Text

                Crno.Visible = True
                CrText.Visible = False
                Gencombo.Visible = False
                SubCombo.Visible = False
                DrText.Visible = False
                Crno.Move Grid1.CellLeft + Grid1.Left, Grid1.CellTop + Grid1.Top - 15, Grid1.CellWidth
                Crno.SetFocus
      

        End Select
Else
  Crno.Visible = False
  Gencombo.Visible = False
  SubCombo.Visible = False
  DrText.Visible = False
  CrText.Visible = False

  Description.Visible = True
  Description.Move Grid1.CellLeft + Grid1.Left, Grid1.CellTop + Grid1.Top
  Description.SetFocus
    


End If

    


LRC = Grid1.row




End Sub

Private Sub Grid1_RowColChange()
'LRC = Grid1.row
'Grid1.col = 0
'If Grid1.Text = "" Then
'                MsgBox "Please fillup this entry.... "
'                Gencombo.Text = Grid1.Text
'                Gencombo.Visible = True
'                SubCombo.Visible = False
'                DrText.Visible = False
'                CrText.Visible = False
'                Crno.Visible = False
'                Gencombo.Move Grid1.CellLeft + Grid1.Left, Grid1.CellTop + Grid1.Top - 15, Grid1.CellWidth
'                Gencombo.SetFocus
'                Grid1.row = LRC
'
'End If

End Sub

Private Sub SubCombo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Grid1.Text = SubCombo.Text
     Grid1.col = Grid1.col + 1
     Grid1_Click
 End If
End Sub

Private Sub SubCombo_LostFocus()
SubCombo.Clear
End Sub

Private Sub vno_LostFocus()

If Val(vno) <> 0 Then
  'GridFill
End If
  
    
End Sub
Private Sub vtype_Change()
 
        If rs.State = 1 Then
            rs.Close
        End If
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
