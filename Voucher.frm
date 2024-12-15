VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Voucherform 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8505
   ClientLeft      =   4005
   ClientTop       =   1620
   ClientWidth     =   12510
   ClipControls    =   0   'False
   Icon            =   "Voucher.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   12510
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FRMINV 
      BackColor       =   &H8000000B&
      Height          =   345
      Left            =   10320
      TabIndex        =   34
      Top             =   180
      Visible         =   0   'False
      Width           =   165
      Begin VB.ListBox LSTINV 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2580
         Left            =   90
         TabIndex        =   35
         Top             =   420
         Width           =   5655
      End
      Begin VB.Label lbllstinvtitle 
         BackColor       =   &H8000000A&
         Caption         =   "I.No     I.Date         NetAmt     RecAmt    Balance"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   36
         Top             =   210
         Width           =   5475
      End
   End
   Begin VB.ComboBox FindCombo 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3105
      Left            =   3930
      Style           =   1  'Simple Combo
      TabIndex        =   31
      Top             =   1080
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.TextBox totalcredit 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   6510
      TabIndex        =   29
      Top             =   7170
      Width           =   1425
   End
   Begin VB.TextBox totaldebit 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   5040
      TabIndex        =   28
      Top             =   7170
      Width           =   1425
   End
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
      TabIndex        =   27
      Top             =   7140
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
      Left            =   960
      TabIndex        =   26
      Top             =   7140
      Width           =   675
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid1 
      Height          =   6465
      Left            =   240
      TabIndex        =   4
      Top             =   540
      Width           =   12075
      _ExtentX        =   21299
      _ExtentY        =   11404
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
      ItemData        =   "Voucher.frx":000C
      Left            =   1710
      List            =   "Voucher.frx":0019
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   90
      Width           =   585
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   840
      ScaleHeight     =   435
      ScaleWidth      =   7725
      TabIndex        =   9
      Top             =   7860
      Width           =   7785
      Begin VB.CommandButton CommandReturn 
         Caption         =   "&Return"
         Height          =   435
         Left            =   6735
         TabIndex        =   17
         Top             =   0
         Width           =   915
      End
      Begin VB.CommandButton CommandPrint 
         Caption         =   "&Print"
         Height          =   435
         Left            =   5850
         TabIndex        =   16
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton Commandsearch 
         Caption         =   "&Search"
         Height          =   435
         Left            =   4920
         TabIndex        =   13
         Top             =   0
         Width           =   915
      End
      Begin VB.CommandButton Commanddelete 
         Caption         =   "De&lete"
         Height          =   435
         Left            =   3900
         TabIndex        =   15
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton Commandabandon 
         Caption         =   "Aba&ndon"
         Height          =   435
         Left            =   2970
         TabIndex        =   14
         Top             =   0
         Width           =   915
      End
      Begin VB.CommandButton Commandsave 
         Caption         =   "Sa&ve"
         Enabled         =   0   'False
         Height          =   435
         Left            =   2040
         TabIndex        =   11
         Top             =   0
         Width           =   915
      End
      Begin VB.CommandButton Commandedit 
         Caption         =   "&Edit"
         Height          =   435
         Left            =   1080
         TabIndex        =   10
         Top             =   0
         Width           =   915
      End
      Begin VB.CommandButton Commandmasteradd 
         Caption         =   "&Add"
         Height          =   435
         Left            =   120
         TabIndex        =   0
         Top             =   0
         Width           =   915
      End
      Begin VB.CommandButton Commandmasterhelp 
         Caption         =   "Help"
         Height          =   375
         Left            =   -1785
         TabIndex        =   12
         Top             =   0
         Width           =   800
      End
   End
   Begin MSMask.MaskEdBox vno 
      Height          =   315
      Left            =   6480
      TabIndex        =   3
      Top             =   90
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox vdate 
      Height          =   315
      Left            =   3840
      TabIndex        =   2
      Top             =   90
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medit 
      Height          =   285
      Left            =   300
      TabIndex        =   18
      Top             =   3180
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   503
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox debit 
      Height          =   285
      Left            =   1740
      TabIndex        =   8
      Top             =   4530
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      AutoTab         =   -1  'True
      Format          =   "###0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox credit 
      Height          =   255
      Left            =   360
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
      CausesValidation=   0   'False
      Height          =   1935
      ItemData        =   "Voucher.frx":0026
      Left            =   2910
      List            =   "Voucher.frx":0028
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   6
      Top             =   1290
      Width           =   2595
   End
   Begin VB.ComboBox Gencombo 
      Height          =   1935
      Left            =   300
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   5
      Top             =   1290
      Width           =   2595
   End
   Begin MSMask.MaskEdBox DESCRIPTION1 
      Height          =   255
      Left            =   360
      TabIndex        =   23
      Top             =   3750
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
      Left            =   660
      MaxLength       =   255
      TabIndex        =   24
      Top             =   4530
      Width           =   1725
   End
   Begin VB.TextBox tmpsubledger 
      Height          =   285
      Left            =   3210
      TabIndex        =   25
      Top             =   3480
      Width           =   3045
   End
   Begin VB.TextBox crnoold 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8460
      TabIndex        =   30
      Top             =   4860
      Width           =   1005
   End
   Begin MSMask.MaskEdBox crno 
      Height          =   285
      Left            =   8460
      TabIndex        =   33
      Top             =   4500
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      AutoTab         =   -1  'True
      Format          =   "###0.00"
      PromptChar      =   "_"
   End
   Begin VB.Label Label5 
      Caption         =   "Press F3 in Credit Amt to Search a Bill"
      Height          =   405
      Left            =   7380
      TabIndex        =   37
      Top             =   30
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.Label Label4 
      Caption         =   "Press F2 for Search A Voucher"
      Height          =   255
      Left            =   1740
      TabIndex        =   32
      Top             =   7200
      Width           =   2475
   End
   Begin VB.Label Label2 
      Caption         =   "Voucher Date"
      Height          =   255
      Left            =   2460
      TabIndex        =   22
      Top             =   120
      Width           =   1275
   End
   Begin VB.Label Label3 
      Caption         =   "Voucher No."
      Height          =   255
      Left            =   5160
      TabIndex        =   21
      Top             =   120
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
      Left            =   4260
      TabIndex        =   20
      Top             =   7170
      Width           =   825
   End
   Begin VB.Label Label1 
      Caption         =   "Voucher Type"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   120
      Width           =   1275
   End
End
Attribute VB_Name = "Voucherform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim CON As ADODB.Connection
Dim rs1 As ADODB.Recordset
Dim RS As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim lastrow, lastcol As Integer
Dim maxrow As Integer
Dim clicksave As Boolean
Dim tmplistindex As Long
Dim Edit As Boolean
Dim addmode As Boolean
Dim searchmode As Boolean
Dim varVtype As String
Dim varVdate As Variant
Dim VarVno As Integer
Dim flagsearch As Boolean
Dim blninvpopup As Boolean
Dim callby As String
Dim invlrow As Integer
Dim invlcol As Integer

Sub refreshTotal()
  totalcredit.Text = ""
  totaldebit.Text = ""
  For I = 1 To Grid1.Rows - 1
  If IsNumeric(Grid1.TextMatrix(I, 2)) Then
     totaldebit.Text = Val(totaldebit.Text) + Val(Grid1.TextMatrix(I, 2))
  End If
  If IsNumeric(Grid1.TextMatrix(I, 3)) Then
     totalcredit.Text = Val(totalcredit.Text) + Val(Grid1.TextMatrix(I, 3))
  End If
  
  Next
  totaldebit.Text = Format(totaldebit.Text, "0.00")
  totalcredit.Text = Format(totalcredit.Text, "0.00")



End Sub

Private Sub Command1_Click()
addmode = False
vtype.Enabled = False
vdate.Enabled = False
vno.Enabled = False
Commandsearch.Enabled = True
Grid1.Enabled = False
searchmode = True
If Not rs2.EOF Then
     rs2.MoveNext
     If rs2.EOF And rs2.RecordCount > 0 Then
           rs2.MoveLast
           Exit Sub
     End If
     totaldebit.Text = ""
     totalcredit.Text = ""
     vtype = rs2!VoucherType
     vdate = rs2!VoucherDate
     vno = rs2!VOUCHERNUMBER
     Me.vtype_LostFocus
     Me.vdate_LostFocus
     Me.vno_LostFocus
     
          

End If




End Sub

Private Sub Command2_Click()

If Not rs2.BOF Then
    addmode = False
    searchmode = True
    
    rs2.MovePrevious
    If rs2.BOF And rs2.RecordCount > 0 Then
           rs2.MoveFirst
           Exit Sub
    End If
     totaldebit.Text = ""
     totalcredit.Text = ""
 
     vtype = rs2!VoucherType
     vdate = rs2!VoucherDate
     vno = rs2!VOUCHERNUMBER
     Me.vtype_LostFocus
     Me.vdate_LostFocus
     Me.vno_LostFocus
     
End If



End Sub

Sub Commandabandon_Click()
'If SAVED Then
    For I = 1 To Grid1.Rows - 1
        Grid1.Row = I
          If I Mod 2 = 0 Then
               Grid1.Text = "."
          Else
          For J = 0 To 4
            Grid1.Col = J
            Grid1.Text = " "
            Grid1.Text = ""
            
          
        Next J
          Grid1.Col = 0
       End If
    Next I
    Grid1.Refresh
    totalcredit = ""
    totaldebit = ""
    Gencombo.Visible = False
    SubCombo.Visible = False
    credit.Visible = False
    debit.Visible = False
    crno.Visible = False
    Grid1.Row = 1
    Grid1.Col = 0
    addmode = False

Commandmasteradd.Enabled = True
Commandedit.Enabled = True
Commanddelete.Enabled = True
Commandsearch.Enabled = True
CommandPrint.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
SetButton Commandmasteradd, Commandedit, Commandsave, Commanddelete
Commandsave.Enabled = False
Edit = False
addmode = False
'Gencombo.Clear
'SubCombo.Clear
End Sub

Private Sub Commanddelete_Click()
If MsgBox("Are You Sure, Delete it ?", vbYesNo) = vbNo Then
        Exit Sub
Else
    CON.Execute "DELETE from vouchers where  vouchertype='" + Trim(vtype.Text) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" + Format(vdate, "dd/mm/yyyy") + "',103) And vouchernumber = " + Trim(vno.Text) & " and " & stringyear
    Commandabandon_Click
'    For I = 1 To 99
'        Grid1.row = I
'        If I Mod 2 = 0 Then
'
'                 Grid1.col = 0
'                Grid1.Text = "."
'
'        Else
'             For j = 0 To 4
'                Grid1.col = j
'                Grid1.Text = ""
'            Next
'        End If
'    Next


                 
End If


End Sub

Private Sub Commandedit_Click()
    searchmode = False
    Gencombo.Clear
    
    If RS.State = 1 Then RS.Close
    RS.Open "Select * from gledger where  " & stringyear & "  order by gledger", CON, adOpenDynamic, adLockReadOnly, adCmdText
    If Not RS.EOF Then
        Do While Not RS.EOF
            Gencombo.AddItem RS(1)
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.Close
    Commandedit.Enabled = False
    Edit = True
    
    varVtype = Trim(vtype)
    varVdate = Format(vdate.Text, "dd/mm/yyyy")
    VarVno = Val(vno)
    Grid1.Enabled = True
    vtype.Enabled = True
    vdate.Enabled = True
    vtype.SetFocus
    vno.Enabled = True
    Commandsave.Enabled = True
    addmode = False
    Commandmasteradd.Enabled = False
    Commandedit.Enabled = False
    Commanddelete.Enabled = False
    Commandsearch.Enabled = False
    CommandPrint.Enabled = False
    Command1.Enabled = False
    Command2.Enabled = False
    varVtype = Trim(vtype)
    varVdate = Format(vdate.Text, "dd/mm/yyyy")
    VarVno = Val(vno)
    searchmode = False
    refreshTotal
End Sub
Private Sub Commandmasteradd_Click()
blninvpopup = True
  searchmode = False
  Commandabandon_Click
  Dim tRS1 As New ADODB.Recordset
  If tRS1.State = 1 Then tRS1.Close
  tRS1.Open "SELECT * FROM VOUCHERS where  " & stringyear & "  order by vsno", CON, adOpenStatic, adLockReadOnly
  If tRS1.RecordCount <= 0 Then
      vtype.Text = "J"
  End If
    Command1.Enabled = False
    Command2.Enabled = False
    addmode = True
    Edit = False
    
    Set rs1 = New ADODB.Recordset
    Set RS = New ADODB.Recordset
    rs1.Open "Select max(VoucherNumber) as mvn from vouchers where vouchertype='" & Trim(vtype.Text) & "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,('" & Now & "'),103)  and " & stringyear, CON, adOpenDynamic, adLockReadOnly, adCmdText
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
    vno.Text = str(c + 1)
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
    RS.Open "Select * from gledger where  " & stringyear & "   order by gledger", CON, adOpenDynamic, adLockReadOnly, adCmdText
    If Not RS.EOF Then
        Do While Not RS.EOF
            Gencombo.AddItem RS(1)
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.Close
    SubCombo.Clear
    RS.Open "Select * from sledger where  " & stringyear & "   order by subledger", CON, adOpenDynamic, adLockReadOnly, adCmdText
    If Not RS.EOF Then
        Do While Not RS.EOF
            SubCombo.AddItem RS(1)
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    If RS.State = 1 Then
        RS.Close
    End If
    If rs1.State = 1 Then
        rs1.Close
    End If
    Grid1.Row = 0
    Grid1.Col = 0
    Grid1.SelectionMode = flexSelectionFree
    Grid1.ColWidth(0, 1) = 11111
    X = Grid1.MergeCol(1)
    DESCRIPTION.Width = Grid1.Width - 340
    DESCRIPTION.Height = crno.Height
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
addmode = True
End Sub

Private Sub Commandreturn_Click()
    Unload Me
End Sub

Private Sub Commandsave_Click()

If vtype = "R" Then
    For I = 1 To 30
    If (Trim(Mid(UCase(Grid1.TextMatrix(I, 0)), 1, 4)) = UCase("Bank") Or Trim(Mid(UCase(Grid1.TextMatrix(I, 0)), 1, 4)) = UCase("Cash")) Then
    If Val(Grid1.TextMatrix(I, 3)) > 0 Then
       MsgBox "Cash or Bank Must Be Debit ....", vbCritical
       Exit Sub
    End If
    End If
    Next
ElseIf vtype = "P" Then
    For I = 1 To 30
    If (Trim(Mid(UCase(Grid1.TextMatrix(I, 0)), 1, 4)) = UCase("Bank") Or Trim(Mid(UCase(Grid1.TextMatrix(I, 0)), 1, 4)) = UCase("Cash")) Then
    If Val(Grid1.TextMatrix(I, 2)) > 0 Then
       MsgBox "Cash or Bank Must Be Credit ....", vbCritical
       Exit Sub
    End If
    End If
    Next
End If



On Error Resume Next

Dim rs5 As New ADODB.Recordset
    Dim SAVED As Boolean
    SAVED = False
    If MsgBox("Do you want to save it now ?", vbYesNo) = vbNo Then
        Exit Sub
    End If
    Commandsave.Enabled = False
    Commandsearch.Enabled = True
    Commandabandon.Enabled = False
If Val(totalcredit.Text) = Val(totaldebit.Text) And totalcredit.Text <> "" And totaldebit.Text <> "" Then 'And Me.Commandedit.Enabled = True Then
    If RS.State = 1 Then
        RS.Close
    End If
    If Edit = False Then
        RS.Open "select * from vouchers where VoucherNumber<=0  and " & stringyear, CON, adOpenDynamic, adLockOptimistic
    Else
        CON.Execute "Delete from vouchers where VoucherType='" + Trim(varVtype) + "' and CONVERT(smalldatetime,voucherdate,103)= CONVERT(smalldatetime,('" + Trim(varVdate) + "'),103) and vouchernumber=" + Trim(VarVno) + "  and " & stringyear
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        RS.Open "select * from vouchers where VoucherNumber<=0  and " & stringyear, CON, adOpenDynamic, adLockPessimistic
    End If
  If addmode = True Then
     rs5.Open "Select max(VoucherNumber) from vouchers where vouchertype='" + Trim(vtype.Text) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" + Format(vdate.Text, "dd/mm/yyyy") + "',103)  and " & stringyear, CON, adOpenDynamic, adLockReadOnly, adCmdText
     If Not rs5.BOF Then
        If Not rs5(0) Then
            c = rs5(0)
        Else
            c = 0
       End If
    Else
        c = 0
    End If
    If rs5.State = 1 Then rs5.Close
    vno.Text = str(c + 1)
  End If
    Grid1.Row = I
    Grid1.Col = 0
    'If grid1.TEXT <> "" And Commandedit.Enabled = False Then
    '    Do While Not rs.EOF
    '        rs.Delete
    '        If Not rs.EOF Then
    '            rs.MoveNext
    '        End If
    '    Loop
    'End If
    For I = 1 To Grid1.Rows Step 2
        Grid1.Row = I
        Grid1.Col = 0
        If Grid1.Text <> "" Then
            RS.addNew
            RS(0) = vtype.Text
            RS(1) = vdate.Text
            RS(2) = Val(vno.Text)
            RS(3) = Grid1.Text
            Grid1.Col = 1
            If Grid1.Text = "" Then
               'rs(4) = Null
               RS(4) = ""
            Else
               RS(4) = Grid1.Text
            End If
            Grid1.Col = 2
            If Grid1.Text <> "" Then
                RS(5) = Val(Grid1.Text)
                RS(6) = "D"
            Else
                Grid1.Col = 3
                RS(5) = Val(Grid1.Text)
                RS(6) = "C"
            End If
            Grid1.Col = 4
            RS(7) = Grid1.Text
            'If Val(Grid1.Text) > 0 And addmode = True Then
            'CON.Execute "update invoicea set recamt=(ISNULL((select sum(amount) from vouchers where debitorcredit='C' and CBND='" & Grid1.Text & "' and " & stringyear & "),0)-ISNULL((select sum(amount) from vouchers where debitorcredit='D' and CBND='" & Grid1.Text & "' and " & stringyear & "),0))" & IIf(Val(Grid1.TextMatrix(I, 2)) > 0, "-" & Grid1.TextMatrix(I, 2), "+" & Grid1.TextMatrix(I, 3)) & " where invoiceno=" & Grid1.Text & " and " & stringyear
            'Else
            
            'End If
            
            Grid1.Col = 0
            Grid1.Row = Grid1.Row + 1
            If Grid1.Text <> "" Then
                RS(9) = Grid1.Text
            End If
            SAVED = True
            RS!FYear = main.session: RS!setupid = main.setupid
            RS!createdby = main.username
            RS!createdon = Now
            RS.Update
        End If
    Next
    If rs2.State = 1 Then rs2.Close
    Set rs2 = New ADODB.Recordset
    rs2.Open "Select DISTINCT vouchertype,voucherdate,vouchernumber from vouchers where  " & stringyear & "   order by vouchertype,voucherdate,vouchernumber", CON, adOpenStatic, adLockReadOnly
    CommandPrint.Visible = True
    RS.Close
    
    'MsgBox "RECORD SAVED", vbInformation
    blninvpopup = False
 'fillprtyinvoice
    
DoEvents
DoEvents

Commandsave.Enabled = False
Commandmasteradd.Enabled = True
'Commandmasteradd.SetFocus
ommandedit.Enabled = True
Commanddelete.Enabled = True
Commandsearch.Enabled = True
CommandPrint.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Commandabandon.Enabled = True

vtype.Enabled = False
vdate.Enabled = False
vno.Enabled = False
Commandsearch.Enabled = True
Grid1.Enabled = False
SAVED = False
Else
  '  If Commandedit.Enabled = True Then
        MsgBox "Please Check That Total Credit and Total Debit is Differ "
   ' End If
End If
addmode = False




End Sub

Private Sub Commandsearch_Click()
    searchmode = True
    Edit = False
    addmode = False
    Set rs1 = New ADODB.Recordset
    Set RS = New ADODB.Recordset
    vno.Enabled = True
    vtype.Enabled = True
    vdate.Enabled = True
    vtype.SetFocus
 
End Sub

Private Sub credit_Change()
    If Val(credit.Text) <> 0 Then
       Grid1.Text = Val(credit.Text)
       
       Grid1.Text = Format(Grid1.Text, "0.00")
    Else
        If Grid1.Col <> 4 And Grid1.Col <> 0 Then
           'Grid1.Text = ""
        End If
    End If
End Sub
Private Sub credit_GotFocus()
If (invlcol <> Grid1.Col And invrow <> Grid1.Col) Then
blninvpopup = True
End If
callby = "C"
fillprtyinvoice
If credit.Text = "" And Edit = False And LSTINV.ListCount > 0 And blninvpopup = True Then
FRMINV.Visible = True
FRMINV.ZOrder
LSTINV.SetFocus
Else
'MsgBox "No Related Invoice Exist"
FRMINV.Visible = False
End If

    If maxrow < Grid1.Row Then
            maxrow = Grid1.Row
    End If
    If searchmode = False Then
        'HIT
        'credit.SetFocus
        SendKeys "{END}"
        SendKeys "+{HOME}"
  End If
End Sub

Private Sub credit_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 And Edit = False Then
callby = "C"
fillprtyinvoice
If LSTINV.ListCount > 0 Then
FRMINV.Visible = True
FRMINV.ZOrder
LSTINV.SetFocus
If Grid1.TextMatrix(Grid1.Row, 4) <> "" Then
For I = 0 To LSTINV.ListCount - 1
If Val(Left(LSTINV.List(I), 9)) = Val(Grid1.TextMatrix(Grid1.Row, 4)) Then
LSTINV.Selected(I) = True
Exit For
End If
Next

End If

Else
MsgBox "No Related Invoice Exist"
End If
End If
End Sub

Private Sub credit_KeyPress(KeyAscii As Integer)


If KeyAscii = 13 Then
'   If Val(credit) > 0 Or Val(debit) > 0 Then
    Dim changed As Boolean
    Dim HIT As Boolean
    HIT = False
    changed = False
    RRR = Grid1.Row
    CCC = Grid1.Col
    Grid1.Col = 3
    totalcredit.Text = ""
    For I = 1 To maxrow Step 1
        Grid1.Row = I
        totalcredit.Text = Val(totalcredit.Text) + Val(Grid1.Text)
    Next
    totalcredit.Text = Format(totalcredit.Text, "0.00")
    Grid1.Row = RRR
    Grid1.Col = CCC
    Grid1.Col = 3
    If credit.Text = "" Then
            Grid1.Text = ""
            Grid1.Col = 2
            Grid1_Click
            Exit Sub
     End If
    If Grid1.Col = lastcol And Grid1.Row = lastrow Then
        Grid1.Col = 0
        If Trim(Grid1.Col) <> "" Then
            HIT = True
        End If
        Grid1.Col = 3
        Grid1.Col = Grid1.Col + 1
        refreshTotal
        changed = True
    End If
    RRR = Grid1.Row
    CCC = Grid1.Col
    Grid1.Col = 2
    If Grid1.Row <> lastrow Or Grid1.Col <> lastcol Then
        Grid1.Text = ""
    End If
   
    
    Grid1.Row = RRR
    Grid1.Col = CCC
  
    If RS.State <> adStateClosed Then RS.Close
    If Val(Trim(Grid1.TextMatrix(Grid1.Row, 4))) = 0 Then
        RS.Open "select (yearopening-isnull(recamt,0)) AS BALANCE from sledgeryearopening where subledger='" + Trim(Grid1.TextMatrix(Grid1.Row, 1)) + "' and gledger='" + Trim(Grid1.TextMatrix(Grid1.Row, 0)) + "' and " & stringyear, CON, adOpenForwardOnly, adLockReadOnly, adCmdText
    Else
        RS.Open "select Balance from invoicedueamt where subledger='" + Trim(Grid1.TextMatrix(Grid1.Row, 1)) + "' and genledger='" + Trim(Grid1.TextMatrix(Grid1.Row, 0)) + "' and invoiceno=" & Val(Trim(Grid1.TextMatrix(Grid1.Row, 4))) & " and " & stringyear, CON, adOpenForwardOnly, adLockReadOnly, adCmdText
    End If
    If RS.RecordCount > 0 Then
        If Val(RS!Balance) < Val(Grid1.TextMatrix(Grid1.Row, 3)) Then
          MsgBox "Enter Value is less then Credit Amount"
        Else
        'MsgBox "Invalid Bill No."
        'changed = False
        End If
    End If
    RS.Close
     Grid1.Row = Grid1.Row + 1
    Grid1.Col = 0
    
    If changed And HIT Then
        Grid1_Click
    End If
    'Else
    ''credit.Visible = False
   ' 'Grid1.col = 2
'    SendKeys "{left}{left}"
    ''Grid1_Click
'    End If
End If
End Sub
Private Sub credit_LostFocus()
    ''credit = ""
    ''credit.Visible = False
    RRR = Grid1.Row
CCC = Grid1.Col
Grid1.Row = lastrow
Grid1.Col = lastcol
If Trim(Grid1.Text) <> "" Then
    Grid1.Row = RRR
    Grid1.Col = CCC
    credit.Visible = False
Else
'  Grid1.col = Grid1.col + 1
'  debit.Visible = False
'  Grid1_Click
'Grid1.Text = credit.Text
End If
credit = ""
    
End Sub
Private Sub crno_Change()
If Val(crno.Text) <> 0 Then

    Grid1.Text = Val(crno.Text)
Else
'If Grid1.col <> 4 Then  Grid1.Text = ""
 If Grid1.Col <> 4 And Grid1.Col <> 0 Then
   ' Grid1.Text = ""
 End If
End If
'crno = UCase(crno)
'If grid1.col = lastcol And grid1.row = lastrow Then
'  grid1.Text = crno.Text
'End If
End Sub

Private Sub crno_GotFocus()
crno.Text = Grid1.Text
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
    
    If Grid1.Col = lastcol And Grid1.Row = lastrow Then
        crno = UCase(crno)
        Grid1.Col = 0
        If Trim(Grid1.Text) <> "" Then
            HIT = True
        End If
        changed = True
        If Val(crno.Text) > 0 Then
            If RS.State <> adStateClosed Then RS.Close
            
            RS.Open "select Balance from invoicedueamt where subledger='" + Trim(Grid1.TextMatrix(Grid1.Row, 1)) + "' and genledger='" + Trim(Grid1.TextMatrix(Grid1.Row, 0)) + "' and invoiceno=" & crno.Text & " and " & stringyear, CON, adOpenForwardOnly, adLockReadOnly, adCmdText
            If RS.RecordCount > 0 Then
                If Val(RS!Balance) < Val(Grid1.TextMatrix(Grid1.Row, 3)) Then
                MsgBox "Invoice Balance is less then Credit Amount"
                End If
                Else
                'MsgBox "Invalid Bill No."
                'changed = False
                End If
            RS.Close
        End If
         Grid1.Row = Grid1.Row + 1
        
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
    Grid1.Text = Val(debit.Text)
    Grid1.Text = Format(Grid1.Text, "0.00")
Else
'If Grid1.col <> 4 Then  Grid1.Text = ""
 If Grid1.Col <> 4 And Grid1.Col <> 0 Then
   ' Grid1.Text = ""
 End If
End If

End Sub
Private Sub debit_GotFocus()
If (invlcol <> Grid1.Col And invrow <> Grid1.Col) Then
blninvpopup = True
End If

callby = "D"
fillprtyinvoice
If debit.Text = "" And Edit = False And LSTINV.ListCount > 0 And blninvpopup = True Then
FRMINV.Visible = True
FRMINV.ZOrder
LSTINV.SetFocus
Else
'MsgBox "No Related Invoice Exist"
FRMINV.Visible = False
End If

     cl = Grid1.Col
     Grid1.Col = 0
   If Grid1.Text <> "" Then
     If Grid1.Text = "CASH-IN-HAND" And vtype.Text = "P" Then
        Grid1.Col = 2
        Grid1.Text = ""
        Grid1.Col = 3
        debit.Visible = False
        Grid1_Click
        
    Else
       Grid1.Col = cl
      ' debit.SetFocus
       SendKeys "{END}"
       SendKeys "+{HOME}"
       
    End If
   

    If maxrow < Grid1.Row Then
        maxrow = Grid1.Row
    End If
    If searchmode = False Then
       '; debit.SetFocus
       ' HIT
        SendKeys "{END}"
        SendKeys "+{HOME}"
    End If
  Else
   debit.Visible = False
  End If
   
 
End Sub

Private Sub debit_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 And Edit = False Then
callby = "D"
fillprtyinvoice
If LSTINV.ListCount > 0 Then
FRMINV.Visible = True
FRMINV.ZOrder
LSTINV.SetFocus
If Grid1.TextMatrix(Grid1.Row, 4) <> "" Then
For I = 0 To LSTINV.ListCount - 1
If Val(Left(LSTINV.List(I), 9)) = Val(Grid1.TextMatrix(Grid1.Row, 4)) Then
LSTINV.Selected(I) = True
Exit For
End If
Next

End If

Else
MsgBox "No Related Bills Exist"
End If
End If
End Sub

Private Sub debit_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim changed As Boolean
    Dim HIT As Boolean
    changed = False
    HIT = False
    RRR = Grid1.Row
    CCC = Grid1.Col
    Grid1.Col = 2
    totaldebit.Text = ""
    For I = 1 To maxrow
        Grid1.Row = I
        totaldebit.Text = Val(totaldebit.Text) + Val(Grid1.Text)
    Next
    totaldebit.Text = Format(totaldebit.Text, "0.00")
    Grid1.Row = RRR
    Grid1.Col = CCC
    If Grid1.Col = lastcol And Grid1.Row = lastrow Then
        Grid1.Col = 0
        If Trim(Grid1.Text) <> "" Then
            HIT = True
        End If
        
        Grid1.Col = 2
        If debit.Text = "" Then Grid1.Text = ""

        If Grid1.Text <> "" Then
            Grid1.Col = 4
        Else
            Grid1.Col = 0
            If Grid1.Text = "CASH-IN-HAND" And vtype.Text = "R" Then
                 Grid1.Col = 4
                  Grid1.Col = 3
                  If Grid1.Text = "" Then
                      debit.Visible = True
                      credit.Visible = False
                      Grid1.Col = 2
                      debit.SetFocus
                     Exit Sub
                  End If
                 
            Else
                 Grid1.Col = 2
                 If debit.Text = "" Then Grid1.Text = ""
                 Grid1.Col = 3
            End If
              
            
            
        End If
        changed = True
    End If
    RRR = Grid1.Row
    CCC = Grid1.Col
    Grid1.Col = 2
    If Grid1.Row <> lastrow Or Grid1.Col <> lastcol Then
        Grid1.Text = ""
    End If
    Grid1.Row = RRR
    Grid1.Col = CCC
    If Grid1.TextMatrix(Grid1.Row, 2) <> "" Then
    Grid1.Row = Grid1.Row + 1
    Grid1.Col = 0
    End If
       
    If changed And HIT Then
        Grid1_Click
    End If
End If
End Sub

Private Sub debit_LostFocus()
RRR = Grid1.Row
CCC = Grid1.Col
Grid1.Row = lastrow
Grid1.Col = lastcol
If Trim(Grid1.Text) <> "" Then
    Grid1.Row = RRR
    Grid1.Col = CCC
   
    debit.Visible = False
Else
     refreshTotal
    'Grid1.col = Grid1.col + 1
    'debit.Visible = False
    'Grid1_Click
End If
debit = ""
End Sub

Private Sub DESCRIPTION_Change()
  
   ' If grid1.col = lastcol And grid1.row = lastrow Then
   '     grid1.Text = DESCRIPTION.Text
   ' End If

End Sub

Private Sub DESCRIPTION_GotFocus()
 If Trim(Grid1.Text) <> Trim(".") Then
         DESCRIPTION.Text = Grid1.Text
 End If
 HIT
End Sub

Private Sub Description_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DESCRIPTION = UCase(DESCRIPTION)
        Grid1.Text = IIf(DESCRIPTION.Text = "", ".", DESCRIPTION.Text)
    Dim changed As Boolean
    Dim HIT As Boolean
    HIT = False
    changed = False
    If Grid1.Col = lastcol And Grid1.Row = lastrow Then
        Grid1.Col = 0
        HIT = True
        Grid1.Row = Grid1.Row + 1
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

Private Sub FindCombo_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
If FindCombo.Text <> "" Then
vtype.Enabled = False
   vdate.Enabled = False
   vno.Enabled = False
   vtype = Trim(Mid(FindCombo.Text, 1, 1))
   vdate = Trim(Mid(FindCombo.Text, 2, 12))
   vno = Trim(Mid(FindCombo.Text, 14, 5))
   vno_LostFocus
   FindCombo.Visible = False
 Else
   FindCombo.Visible = False
 End If
 End If
 
If KeyAscii = 27 Then FindCombo.Visible = False
   
End Sub

Private Sub Form_Activate()
If Command1.Enabled = True Then
Command1.SetFocus
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
  FindCombo.Visible = True
  FindCombo.SetFocus

  Dim rs7 As New ADODB.Recordset
  rs7.Open "Select DISTINCT vouchertype,voucherdate,vouchernumber from vouchers where  " & stringyear & "   order by vouchertype,voucherdate,vouchernumber asc", CON, adOpenStatic
  If rs7.RecordCount > 0 Then
   FindCombo.Clear
   rs7.MoveFirst
   While Not rs7.EOF
       FindCombo.AddItem rs7(0) & Space(1) & rs7(1) & Space(1) & rs7(2)
       rs7.MoveNext
   Wend
  End If
  

End If


End Sub

Private Sub Gencombo_Change()
 If Grid1.Col = lastcol And Grid1.Row = lastrow Then
        Grid1.Text = Gencombo.Text
        Grid1.Col = 0
 End If
End Sub

Private Sub genCombo_Click()
    Grid1.Text = Gencombo.Text
    If vtype.Text = "R" And Grid1.Text = "CASH-IN-HAND" Then
        Grid1.Col = 2
        DB = Grid1.Text
        Grid1.Col = 3
        CR = Grid1.Text
        If DB = "" Then
            Grid1.Col = 2
            Grid1.Text = CR
            Grid1.Col = 3
            Grid1.Text = ""
       End If
    End If
   If vtype.Text = "P" And Grid1.Text = "CASH-IN-HAND" Then
        Grid1.Col = 2
        DB = Grid1.Text
        Grid1.Col = 3
        CR = Grid1.Text
        If CR = "" Then
            Grid1.Col = 3
            Grid1.Text = DB
            Grid1.Col = 2
            Grid1.Text = ""
       End If
    End If
    Grid1.Col = 1
    Grid1.Text = ""
    Grid1.Col = 0
End Sub
Private Sub Form_Load()


blninvpopup = False
varVtype = ""
varVdate = ""
VarVno = 0
searchmode = False
addmode = False
vtype.Enabled = False
vdate.Enabled = False
vno.Enabled = False
Commandsearch.Enabled = True
Grid1.Enabled = False
Edit = False
    Grid1.TOP = 450
    Grid1.Left = 60
    Me.TOP = 0
    Me.Left = 0
    Grid1.Rows = 500
    Grid1.Cols = 0
    Grid1.Cols = 5
    For I = 0 To 499
        Grid1.RowHeight(I) = 250
    Next
    Grid1.Row = 0
    Grid1.Col = 0
    Grid1.MergeCells = flexMergeRestrictRows
       For I = 1 To Grid1.Rows - 1
        If I Mod 2 = 0 Then
            Grid1.Row = I
            For J = 0 To 4
                Grid1.Col = J
                Grid1.Text = "."
                Grid1.MergeRow(I) = True
            Next
        End If
    Next
   'Grid1.MergeCells = flexMergeRestrictRows
    Grid1.Row = 0
    Grid1.Col = 0
    Grid1.ColAlignment(0) = 0
    Grid1.ColAlignment(1) = 0
    Grid1.ColWidth(0) = 3500      '2750
    Grid1.ColWidth(1) = 4000
    Grid1.ColWidth(2) = 1160
    Grid1.ColWidth(3) = 1160
    Grid1.ColWidth(4) = 1160
    Grid1.Text = "Gen. Ledger"
    Grid1.Col = 1
    Grid1.Text = "Sub. Ledger"
    Grid1.Col = 2
    Grid1.Text = "Amount. (Dr.)"
    totaldebit.Left = Grid1.CellLeft + 50
    Grid1.Col = 3
    Grid1.Text = "Amount. (Cr.)"
    totalcredit.Left = Grid1.CellLeft + 60
    Grid1.Col = 4
    'Grid1.Text = "C/B No."
    Grid1.Text = "Bill No."
    totalcredit.Width = crno.Width
    totaldebit.Width = totalcredit.Width
    totaldebit.Height = crno.Height + 10
    totalcredit.Height = totaldebit.Height
   
    'Set CON = New ADODB.Connection
     '   CON.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + App.Path + "\" + Trim(main.directory) + "\data.mdb"
    'CON.Open
    

'******Load first voucher start
    
    Set RS = New ADODB.Recordset
    'RS.Open "Select * from vouchers where " & stringyear & " order by vouchertype,voucherdate,vouchernumber,vsno ", CON, adOpenDynamic, adLockOptimistic, adCmdText
    RS.PageSize = 2
    RS.Open "Select * from vouchers where " & stringyear & " and VoucherNumber='" & vnumbers & "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & vdates & "',103)   order by vouchertype,voucherdate,vouchernumber", CON, adOpenDynamic, adLockOptimistic
    If Not RS.BOF Then
        vtype = RS!VoucherType
        vdate = RS!VoucherDate
        vno = RS!VOUCHERNUMBER
        Me.vtype_LostFocus
        Me.vdate_LostFocus
        Me.vno_LostFocus
    End If
  
    Set rs2 = New ADODB.Recordset
     rs2.PageSize = 2
    'rs2.Open "Select DISTINCT vouchertype,voucherdate,vouchernumber from vouchers where " & stringyear & " order by vouchertype,voucherdate,vouchernumber", CON, adOpenStatic
    rs2.Open "Select DISTINCT vouchertype,voucherdate,vouchernumber from vouchers where  " & stringyear & "   order by vouchertype,voucherdate,vouchernumber", CON, adOpenStatic
    SetButton Commandmasteradd, Commandedit, Commandsave, Commanddelete
    Commandsave.Enabled = False
  '  rs.Close
' Load first voucher end



Voucherform.TOP = 1500
Voucherform.Left = 1000


ButtonPermission Commandsave, Commanddelete, Commandedit

End Sub



Private Sub Gencombo_GotFocus()
    Grid1.Col = 0
    Gencombo.Text = Grid1.Text
    If Gencombo.Enabled = True Then
    If searchmode = False Then
    HIT
    End If
    End If
End Sub

Private Sub Gencombo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 115 Then
   If MsgBox("Are You Sure, Delete it ?", vbYesNo) = vbYes Then
      If Grid1.Row >= 1 Then
           Grid1.RemoveItem Grid1.Row
           Grid1.RemoveItem Grid1.Row
           'Gencombo.Text = ""
           Gencombo.Visible = False
           Me.DESCRIPTION.Visible = False
           refreshTotal
           Grid1.SetFocus
       End If
   End If
End If
End Sub

Private Sub Gencombo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim HIT As Boolean
        Dim changed As Boolean
        changed = False
        HIT = False
        If RS.State = 1 Then RS.Close
        RS.Open "select * from gledger where " & stringyear, CON, adOpenStatic, adLockReadOnly, adCmdText
        If Trim(Gencombo.Text) <> "" Then
        'SendKeys "{DOWN}"
        If Not RS.BOF Then
                RS.MoveFirst
                RS.Find "GLEDGER='" + Trim(Gencombo.Text) + "'"
                If RS.EOF Then
                    RS.Close
                    Exit Sub
                End If
            End If
        Else
           Gencombo.Visible = False
        End If
        RS.Close
        
        If Grid1.Col = lastcol And Grid1.Row = lastrow Then
            If Trim(Grid1.Text) <> "" Then
                HIT = True
            End If
            Grid1.Col = Grid1.Col + 1
            changed = True
        End If
        Gencombo.Text = ""
        If changed And HIT Then
            Grid1_Click
       Else
         If changed Then
             Commandsave.Enabled = True
             Commandsave.SetFocus
         End If
        End If

    End If
End Sub

Private Sub Gencombo_LostFocus()
   ' rs.Open "GLEDGER", CON, adOpenDynamic, adLockReadOnly, adcmdtext
   ' If Trim(Gencombo.Text) <> "" Then
   '     If Not rs.BOF Then
   '         rs.MoveFirst
   '         rs.Find "GLEDGER='" + Trim(Gencombo.Text) + "'"
   '         If rs.EOF Then
   '             Gencombo.Visible = True
   '             Gencombo.SetFocus
   '         Else
            
  '              Gencombo.Visible = False
  '              Me.Commandsave.SetFocus
  '          End If
   '     End If
   ' Else
    '    Gencombo.Visible = False
        
    'End If
    'rs.Close
End Sub
Private Sub Grid1_Click()
If Grid1.Enabled = True Then
Dim check As Boolean
Dim Row, Col As Integer
Row = Grid1.Row
Col = Grid1.Col
check = True
If Grid1.Row Mod 2 <> 0 Then
    DESCRIPTION.Visible = False
    Dim gr As MSHFlexGrid
    Set gr = Grid1
    If Grid1.Row >= 3 Then
        check = False
        gr.Row = Grid1.Row - 2
        gr.Col = 0
        If gr.Text <> "" Then
            gr.Col = 2
            If gr.Text <> "" Then
                check = True
            Else
                gr.Col = 3
                If gr.Text <> "" Then
                    check = True
                End If
            End If
        End If
    End If
    Grid1.Row = Row
    Grid1.Col = Col
    credit.Visible = False
    debit.Visible = False
    SubCombo.Visible = False
    crno.Visible = False
    Gencombo.Visible = False
    If check Then
       If Grid1.Col = 0 Then
          credit.Visible = False
          debit.Visible = False
          SubCombo.Visible = False
          crno.Visible = False
          
          Gencombo.Text = Grid1.Text
          Gencombo.Visible = True
          Gencombo.Left = Grid1.CellLeft + 40
          Gencombo.TOP = Grid1.TOP + Grid1.CellTop - 50
          Gencombo.Width = Grid1.ColWidth(Grid1.Col)
          Gencombo.ZOrder
          Gencombo.SetFocus
       End If
        
        
        If Grid1.Col = 1 Then
            tmpsubledger = ""
            SubCombo.Text = ""
            If Grid1.Text <> "" Then
               tmpsubledger = Grid1.Text
            
            End If
            
            Gencombo.Visible = False
            credit.Visible = False
            debit.Visible = False
            crno.Visible = False
            SubCombo.Visible = True
            SubCombo.Text = Grid1.Text
            SubCombo.Left = Grid1.CellLeft + 40
            SubCombo.TOP = Grid1.TOP + Grid1.CellTop - 50
            SubCombo.ZOrder
            SubCombo.SetFocus
            SubCombo.Width = Grid1.ColWidth(Grid1.Col)
        End If
        If Grid1.Col = 2 Then
            Gencombo.Visible = False
            SubCombo.Visible = False
            credit.Visible = False
            crno.Visible = False
            Grid1.Col = 3
            If Trim(Grid1.Text) = "" Then
                Grid1.Col = 2
                If Grid1.Text <> "" Then
                    debit.Text = Grid1.Text
                End If
                debit.Visible = True
                debit.Left = Grid1.CellLeft + 40
                debit.TOP = Grid1.TOP + Grid1.CellTop - 25
                debit.Width = Grid1.ColWidth(Grid1.Col)
                debit.ZOrder
                debit.SetFocus
                
           Else
                       
              Grid1.Col = 3
          
            End If
          End If
        If Grid1.Col = 3 Then
            Gencombo.Visible = False
            SubCombo.Visible = False
            debit.Visible = False
            crno.Visible = False
            Grid1.Col = 2
            If Trim(Grid1.Text) = "" Then
                Grid1.Col = 3
                If Grid1.Text <> "" Then
                credit.Text = Grid1.Text
                End If
                credit.Visible = True
                credit.Left = Grid1.CellLeft + 40
                credit.TOP = Grid1.TOP + Grid1.CellTop - 25
                credit.Width = Grid1.ColWidth(Grid1.Col)
                
                credit.ZOrder
                credit.SetFocus
                
            End If
        End If
        If Grid1.Col = 4 Then
            Gencombo.Visible = False
            SubCombo.Visible = False
            debit.Visible = False
            credit.Visible = False
'            crno.Text = Grid1.Text
'            crno.Visible = True
'            crno.Left = Grid1.CellLeft + 40
'            crno.Top = Grid1.Top + Grid1.CellTop - 25
'            crno.Width = Grid1.ColWidth(Grid1.col)
'            crno.ZOrder
'
'            crno.SetFocus
           
        End If
    lastrow = Grid1.Row
    lastcol = Grid1.Col
    Else
    MsgBox "Previous entry didn't complete"
    End If
Else
    If Grid1.Row >= 2 Then
        RRR = Grid1.Row
        CCC = Grid1.Col
        Grid1.Row = Grid1.Row - 1
        Grid1.Col = 0
        If Trim(Grid1.Text) <> "" Then
            Grid1.Row = RRR
            Grid1.Col = CCC
            If Grid1.Col = 0 Then
                DESCRIPTION.Left = Grid1.CellLeft + 40
                DESCRIPTION.TOP = Grid1.TOP + Grid1.CellTop - 25
                DESCRIPTION.Width = Grid1.CellWidth  '' Grid1.ColWidth(Grid1.col)
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
        lastrow = Grid1.Row
        lastcol = Grid1.Col
    End If
End If
End If

End Sub


Private Sub LSTINV_DblClick()

If LSTINV.ListIndex >= 0 Then
If callby = "C" Then
credit.Text = Val(VBA.Right(LSTINV.List(LSTINV.ListIndex), 10))
Grid1.TextMatrix(Grid1.Row, 3) = Val(VBA.Right(LSTINV.List(LSTINV.ListIndex), 10))
ElseIf callby = "D" Then
Grid1.TextMatrix(Grid1.Row, 2) = Val(VBA.Right(LSTINV.List(LSTINV.ListIndex), 10))
debit.Text = Val(VBA.Right(LSTINV.List(LSTINV.ListIndex), 10))
End If
Grid1.TextMatrix(Grid1.Row, 4) = Val(VBA.Left(LSTINV.List(LSTINV.ListIndex), 9))
'grid1.SetFocus
FRMINV.Visible = False
End If

End Sub

Private Sub LSTINV_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then FRMINV.Visible = False
If KeyAscii = 13 Then LSTINV_DblClick
End Sub

Private Sub LSTINV_LostFocus()
Grid1_Click
If UCase(Grid1.TextMatrix(Grid1.Row, 0)) = "SUNDRY DEBTORS" Then
'credit.SetFocus
ElseIf UCase(Grid1.TextMatrix(Grid1.Row, 0)) = "SUNDRY CREDITORS" Then
'debit.SetFocus
End If

invlrow = Grid1.Row
invlcol = Grid1.Col
blninvpopup = False
callby = ""

End Sub

Private Sub SubCombo_Change()
   ' SubCombo = UCase(SubCombo)
    If Grid1.Col = lastcol And Grid1.Row = lastrow Then
        Grid1.Text = SubCombo.Text
    End If
End Sub
Private Sub SubCombo_Click()
    Grid1.Text = SubCombo.Text
End Sub
Private Sub SubCombo_GotFocus()

   tmplistindex = 0
   Dim X(5000) As String
   Gencombo.Visible = False
   SubCombo.Clear
   tmpsubledger.Text = Grid1.Text
   Set RS = New ADODB.Recordset
   Grid1.Col = 0
   If Grid1.Text <> "" Then
       RS.Open "Select SUBLEDGER from  SLEDGER where gledger='" + Trim(Grid1.Text) + "' and  " & stringyear & "   ORDER BY gledger,SUBLEDGER", CON, adOpenStatic, adLockReadOnly, adCmdText
       If Not RS.EOF Then
        Dim X1  As Integer
          X1 = 0
          CV = 0
          Do While Not RS.EOF
             SubCombo.AddItem RS(0)
             If RS(0) = tmpsubledger.Text Then
                 tmplistindex = CV
                 CV = X1
             End If
             X1 = X1 + 1
             RS.MoveNext
         Loop
        
    End If
    Grid1.Col = CCC
    Grid1.Col = 1
    If SubCombo.ListCount = 0 Then
        Grid1.Col = 2
        Grid1_Click
    End If
    SubCombo.Text = tmpsubledger.Text
    If CV >= 1 Then
       SubCombo.ListIndex = CV
    End If

    
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
    RRR = Grid1.Row
    CCC = Grid1.Col
    Grid1.Row = lastrow
    Grid1.Col = 0
    If RS.State = 1 Then
        RS.Close
    End If
    If Trim(SubCombo.Text) <> "" And Trim(Grid1.Text) <> "" Then
        SendKeys "{down}"
        RS.Open "select * from SLEDGER where subledger='" + Trim(SubCombo.Text) + "' and gledger='" + Trim(Grid1.Text) + "' and " & stringyear, CON, adOpenForwardOnly, adLockReadOnly, adCmdText
        If RS.BOF Or RS.EOF Then
            SubCombo.Visible = True
            SubCombo.SetFocus
        Else
            valid = True
        End If
        RS.Close
        'fillprtyinvoice
    Else
        If Trim(SubCombo.Text) = "" Then
            SubCombo.Visible = True
            SubCombo.SetFocus
       End If
    End If
    Grid1.Row = RRR
    Grid1.Col = CCC
    If valid Then
        If Grid1.Col = lastcol And Grid1.Row = lastrow Then
            Grid1.Col = 0
            If Trim(Grid1.Col) <> "" Then
                HIT = True
            End If
            Grid1.Col = 3
            If Grid1.Text <> "" Then
            
            Else
                Grid1.Col = 2
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
'    If rs.State = 1 Then
'        rs.Close
'    End If
'    RRR = Grid1.row
'    CCC = Grid1.col
'    Grid1.row = lastrow
'    Grid1.col = 0
'    If Trim(SubCombo.Text) <> "" And Trim(Grid1.Text) <> "" Then
'        rs.Open "select * from SLEDGER where subledger='" + Trim(SubCombo.Text) + "' and gledger='" + Trim(Grid1.Text) + "'", CON, adOpenDynamic, adLockReadOnly, adCmdText
'        If rs.BOF Or rs.EOF Then
'            SubCombo.Visible = True
'            SubCombo.SetFocus
'        Else
            SubCombo.Visible = False
'        End If
'        rs.Close
'    Else
'        If Trim(SubCombo.Text) <> "" Then
'            SubCombo.Visible = True
'            SubCombo.SetFocus
'        End If
'    End If
'    Grid1.row = RRR
'    Grid1.col = CCC
  SubCombo.Text = ""
End Sub

Private Sub vdate_GotFocus()
HIT
End Sub

Private Sub vdate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If addmode = True Or Edit = True Then
           Grid1.Row = 1
           Grid1.Col = 0
           Grid1.SetFocus
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
        If RS.State = 1 Then RS.Close
        Set RS = New ADODB.Recordset
        RS.Open "Select max(VoucherNumber) from vouchers where vouchertype='" + Trim(vtype.Text) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & (vdate.Text) & "',103) and " & stringyear, CON, adOpenDynamic, adLockReadOnly, adCmdText
        If Not RS.BOF Then
            If Not RS(0) Then
                c = RS(0)
            Else
                c = 0
            End If
        Else
            c = 0
        End If
        RS.Close
        vno.Text = str(c + 1)
    Else
       If varVtype = "" Then Exit Sub
       If varVdate = "" Then Exit Sub
       If searchmode = True Then Exit Sub
       If Trim(vtype.Text) <> Trim(varVtype) Or CDate(Format(varVdate, "dd/mm/yyyy")) <> Format(vdate, "dd/mm/yyyy") Then
            If RS.State = 1 Then RS.Close
            Set RS = New ADODB.Recordset
            RS.Open "Select max(VoucherNumber) from vouchers where  " & stringyear & "   and vouchertype='" + Trim(vtype.Text) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" + vdate.Text + "',103)", CON, adOpenDynamic, adLockReadOnly, adCmdText
            If Not RS.BOF Then
                If Not RS(0) Then
                    c = RS(0)
                Else
                    c = 0
                End If
            Else
                c = 0
            End If
            RS.Close
            vno.Text = str(c + 1)
       End If
   End If
End Sub

Private Sub vno_GotFocus()
       vno.SetFocus
       SendKeys "{END}"
       SendKeys "+{HOME}"
End Sub

Private Sub vno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If searchmode = True Then
    vtype.SetFocus
    Exit Sub
  End If
  SendKeys ("{TAB}")
End If
End Sub
Sub vno_LostFocus()



If Val(vno) <> 0 Then
   If RS.State = 1 Then RS.Close
   'RS.Open "Select * from vouchers where VoucherType='" + Trim(vtype.Text) + "' and voucherdate= CDate ('" + Trim(vdate.Text) + "') and vouchernumber=" + Trim(vno.Text) + "", CON, adOpenDynamic, adLockOptimistic, adCmdText 'and VoucherNumber=" + Val(Trim(vno.Text)) + "", con, adOpenDynamic, adLockOptimistic, adCmdText
   RS.Open "Select * from vouchers where  " & stringyear & "   and VoucherType='" + Trim(vtype.Text) + "' and CONVERT(smalldatetime,voucherdate,103)= CONVERT(smalldatetime,('" + Trim(vdate.Text) + "'),103)  and vouchernumber=" & Trim(vno.Text) & " order by vsno", CON, adOpenDynamic, adLockOptimistic, adCmdText   'and VoucherNumber=" + Val(Trim(vno.Text)) + "", con, adOpenDynamic, adLockOptimistic, adCmdText
   For I = 1 To Grid1.Rows - 1
       Grid1.Row = I
       If I Mod 2 = 0 Then
          Grid1.Text = "."
       Else
          For J = 0 To 4
             Grid1.Col = J
             Grid1.Text = ""
          Next J
          Grid1.Col = 0
         End If
   Next I
   If Not RS.BOF Then
      I = 1
      totaldebit.Text = 0
      totalcredit.Text = 0
    
      Do While Not RS.EOF
         Grid1.Row = I
         Grid1.Col = 0
         Grid1.Text = RS(3)
         Grid1.Col = 1
         If IsNull(RS(4)) Then
              Grid1.Text = ""
         Else
              Grid1.Text = RS(4)
         End If
         Grid1.Col = 2
         If RS(6) = "D" Then
            Grid1.Text = Format(RS(5), "0.00")
            totaldebit.Text = Val(totaldebit.Text) + RS(5)
            totaldebit.Text = Format(totaldebit.Text, "0.00")
         Else
            Grid1.Col = 3
            Grid1.Text = Format(RS(5), "0.00")
            totalcredit.Text = Val(totalcredit.Text) + RS(5)
            totalcredit.Text = Format(totalcredit.Text, "0.00")
         End If
         Grid1.Col = 4
         Grid1.Text = IIf(IsNull(RS(7)), ".", RS(7))
         Grid1.Row = Grid1.Row + 1
         Grid1.Col = 0
         Grid1.Text = IIf(IsNull(RS(9)), ".", RS(9))
         I = I + 2
         If Not RS.EOF Then
            RS.MoveNext
         End If
     Loop
     Gencombo.Visible = False
     SubCombo.Visible = False
     credit.Visible = False
     debit.Visible = False
     crno.Visible = False
     Grid1.Row = 1
     Grid1.Col = 0
  End If
  RS.Close
  Grid1.Row = 1
  Grid1.Col = 0
  If searchmode = False Then Grid1_Click
Else
  vno.SetFocus
End If

    
End Sub
Private Sub vtype_Change()
        If RS.State = 1 Then RS.Close
        Set RS = New ADODB.Recordset
        RS.Open "Select max(VoucherNumber) from vouchers where  " & stringyear & "   and vouchertype='" + Trim(vtype.Text) + "' and voucherdate= CDate ('" + Format(vdate.Text, "dd/mm/yy") + "')", CON, adOpenDynamic, adLockReadOnly, adCmdText
        If Not RS.BOF Then
           If Not RS(0) Then
                c = RS(0)
           Else
                c = 0
           End If
        Else
           c = 0
        End If
        RS.Close
        vno.Text = str(c + 1)
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

End Sub


Sub fillprtyinvoice()
''''vsprtyinvoice.Clear
'''If edit = False Then
'''If UCase(Grid1.TextMatrix(Grid1.Row, 0)) = UCase("Sundry Debtors") And Grid1.TextMatrix(Grid1.Row, 1) <> "" Then
'''LSTINV.Clear
'''If rs.State <> adStateClosed Then rs.Close
'''rs.Open "select yearopening,ISNULL(recamt,0) AS RECAMT ,(yearopening-ISNULL(recamt,0)) as balance from sledgeryearopening where subledger='" + Grid1.TextMatrix(Grid1.Row, 1) + "' and gledger='" + Grid1.TextMatrix(Grid1.Row, 0) + "' and " & stringyear, CON, adOpenForwardOnly, adLockReadOnly, adCmdText
'''If rs!balance > 0 Or IsNull(rs!balance) = True Then
'''LSTINV.AddItem 0 & Space(9 - 1) & "YearOpening" & Space(10 - Len(rs!YEAROPENING)) & rs!YEAROPENING & " " & Space(10 - Len(rs!recAmt)) & rs!recAmt & " " & Space(10 - Len(rs!balance)) & rs!balance
'''End If
'''
'''If rs.State <> adStateClosed Then rs.Close
'''rs.Open "select invoiceno,invoicedate,Netamount,Recamt,Balance from invoicedueamt where subledger='" + Grid1.TextMatrix(Grid1.Row, 1) + "' and genledger='" + Grid1.TextMatrix(Grid1.Row, 0) + "' and " & stringyear, CON, adOpenForwardOnly, adLockReadOnly, adCmdText
'''While rs.EOF = False
'''LSTINV.AddItem rs!INVOICENO & Space(9 - Len(rs!INVOICENO)) & rs!INVOICEDATE & Space(11 - Len(rs!INVOICEDATE)) & Space(10 - Len(rs!netamount)) & rs!netamount & " " & Space(10 - Len(rs!recAmt)) & rs!recAmt & " " & Space(10 - Len(rs!balance)) & rs!balance
'''rs.MoveNext
'''Wend
'''If LSTINV.ListCount > 0 Then
'''End If
'''
''''If rs.RecordCount > 0 Then
''''Set vsprtyinvoice.DataSource = rs
''''vsprtyinvoice.Refresh
''''vsprtyinvoice.ColWidth(1) = 1000
''''vsprtyinvoice.ColWidth(2) = 1000
''''vsprtyinvoice.ColWidth(3) = 1000
''''vsprtyinvoice.Visible = True
''''Else
''''vsprtyinvoice.Visible = False
''''End If
'''
'''rs.Close
'''
'''ElseIf UCase(Grid1.TextMatrix(Grid1.Row, 0)) = UCase("Sundry Creditors") And Grid1.TextMatrix(Grid1.Row, 1) <> "" Then
'''LSTINV.Clear
'''If rs.State <> adStateClosed Then rs.Close
'''rs.Open "select yearopening,ISNULL(recamt,0) AS RECAMT ,(yearopening-ISNULL(recamt,0)) as balance from sledgeryearopening where subledger='" + Grid1.TextMatrix(Grid1.Row, 1) + "' and gledger='" + Grid1.TextMatrix(Grid1.Row, 0) + "' and " & stringyear, CON, adOpenForwardOnly, adLockReadOnly, adCmdText
'''If rs!balance > 0 Or IsNull(rs!balance) = True Then
'''   LSTINV.AddItem 0 & Space(9 - 1) & "YearOpening" & Space(10 - Len(rs!YEAROPENING)) & rs!YEAROPENING & " " & Space(10 - Len(rs!recAmt)) & rs!recAmt & " " & Space(10 - Len(rs!balance)) & rs!balance
'''End If
'''
'''If rs.State <> adStateClosed Then rs.Close
'''rs.Open "select invoiceno,invoicedate,Netamount,payamt,Balance from purchasedueamt where subledger='" + Grid1.TextMatrix(Grid1.Row, 1) + "' and genledger='" + Grid1.TextMatrix(Grid1.Row, 0) + "' and " & stringyear, CON, adOpenForwardOnly, adLockReadOnly, adCmdText
'''While rs.EOF = False
'''   LSTINV.AddItem rs!INVOICENO & Space(9 - Len(rs!INVOICENO)) & rs!INVOICEDATE & Space(11 - Len(rs!INVOICEDATE)) & Space(10 - Len(rs!netamount)) & rs!netamount & " " & Space(10 - Len(rs!payAmt)) & rs!payAmt & " " & Space(10 - Len(rs!balance)) & rs!balance
'''rs.MoveNext
'''Wend
'''   If LSTINV.ListCount > 0 Then
'''End If
'''
''''If rs.RecordCount > 0 Then
''''Set vsprtyinvoice.DataSource = rs
''''vsprtyinvoice.Refresh
''''vsprtyinvoice.ColWidth(1) = 1000
''''vsprtyinvoice.ColWidth(2) = 1000
''''vsprtyinvoice.ColWidth(3) = 1000
''''vsprtyinvoice.Visible = True
''''Else
''''vsprtyinvoice.Visible = False
''''End If
'''
'''rs.Close
'''End If
'''End If
End Sub


