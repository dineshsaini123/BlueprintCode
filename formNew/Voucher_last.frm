VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Voucherform_last 
   ClientHeight    =   9876
   ClientLeft      =   276
   ClientTop       =   1716
   ClientWidth     =   12744
   ClipControls    =   0   'False
   Icon            =   "Voucher_last.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9876
   ScaleWidth      =   12744
   Begin VB.Frame panel 
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   11.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   9825
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   12570
      Begin VB.TextBox crno 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   28
         Top             =   1440
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.ComboBox Gencombo 
         Height          =   1872
         Left            =   540
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   24
         Top             =   1905
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.ComboBox SubCombo 
         CausesValidation=   0   'False
         Height          =   1872
         ItemData        =   "Voucher_last.frx":000C
         Left            =   3180
         List            =   "Voucher_last.frx":000E
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   23
         Top             =   1845
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   915
         Left            =   150
         ScaleHeight     =   912
         ScaleWidth      =   8880
         TabIndex        =   8
         Top             =   8565
         Width           =   8880
         Begin VB.CommandButton Commandmasterhelp 
            Caption         =   "Help"
            Height          =   375
            Left            =   -1290
            TabIndex        =   17
            Top             =   0
            Width           =   800
         End
         Begin VB.CommandButton Commandmasteradd 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Add"
            Height          =   735
            Left            =   90
            Picture         =   "Voucher_last.frx":0010
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   120
            Width           =   1065
         End
         Begin VB.CommandButton Commandedit 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Edit"
            Height          =   735
            Left            =   1170
            Picture         =   "Voucher_last.frx":0BF4
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   120
            Width           =   1065
         End
         Begin VB.CommandButton Commandsave 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Sa&ve"
            Enabled         =   0   'False
            Height          =   735
            Left            =   2300
            Picture         =   "Voucher_last.frx":1036
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   120
            Width           =   1065
         End
         Begin VB.CommandButton Commandabandon 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Aba&ndon"
            Height          =   735
            Left            =   3390
            Picture         =   "Voucher_last.frx":1C1A
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   120
            Width           =   1065
         End
         Begin VB.CommandButton Commanddelete 
            BackColor       =   &H00FFFFFF&
            Caption         =   "De&lete"
            Height          =   735
            Left            =   4480
            Picture         =   "Voucher_last.frx":21A4
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   120
            Width           =   1065
         End
         Begin VB.CommandButton Commandsearch 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Search"
            Height          =   735
            Left            =   5570
            Picture         =   "Voucher_last.frx":2D88
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   120
            Width           =   1065
         End
         Begin VB.CommandButton CommandPrint 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Print"
            Height          =   735
            Left            =   6660
            Picture         =   "Voucher_last.frx":396C
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   120
            Width           =   1065
         End
         Begin VB.CommandButton CommandReturn 
            BackColor       =   &H00FFFFFF&
            Caption         =   "E&xit"
            Height          =   735
            Left            =   7755
            Picture         =   "Voucher_last.frx":4550
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   120
            Width           =   1065
         End
      End
      Begin VB.ComboBox vtype 
         Height          =   315
         ItemData        =   "Voucher_last.frx":5134
         Left            =   1560
         List            =   "Voucher_last.frx":5141
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   360
         Width           =   780
      End
      Begin VB.CommandButton Command1 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   825
         TabIndex        =   5
         Top             =   7605
         Width           =   675
      End
      Begin VB.CommandButton Command2 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   105
         TabIndex        =   4
         Top             =   7605
         Width           =   675
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
         Left            =   6765
         TabIndex        =   3
         Top             =   7635
         Width           =   1380
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
         Left            =   8235
         TabIndex        =   2
         Top             =   7635
         Width           =   1380
      End
      Begin VB.ComboBox FindCombo 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   7.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   4284
         Left            =   2010
         Style           =   1  'Simple Combo
         TabIndex        =   1
         Top             =   1620
         Visible         =   0   'False
         Width           =   2835
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid1 
         Height          =   6480
         Left            =   60
         TabIndex        =   6
         Top             =   960
         Width           =   12435
         _ExtentX        =   21929
         _ExtentY        =   11430
         _Version        =   393216
         Cols            =   5
         BackColorFixed  =   7917545
         GridColorFixed  =   8438015
         MergeCells      =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSMask.MaskEdBox vno 
         Height          =   315
         Left            =   7455
         TabIndex        =   20
         Top             =   300
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   550
         _Version        =   393216
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox vdate 
         Height          =   315
         Left            =   4290
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   300
         Width           =   1245
         _ExtentX        =   2180
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medit 
         Height          =   285
         Left            =   570
         TabIndex        =   19
         Top             =   3705
         Width           =   3105
         _ExtentX        =   5461
         _ExtentY        =   508
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox debit 
         Height          =   285
         Left            =   2010
         TabIndex        =   21
         Top             =   5055
         Width           =   1185
         _ExtentX        =   2074
         _ExtentY        =   508
         _Version        =   393216
         Appearance      =   0
         AutoTab         =   -1  'True
         Format          =   "###0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox credit 
         Height          =   255
         Left            =   660
         TabIndex        =   22
         Top             =   4560
         Width           =   1155
         _ExtentX        =   2032
         _ExtentY        =   445
         _Version        =   393216
         Appearance      =   0
         AutoTab         =   -1  'True
         Format          =   "###0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox DESCRIPTION1 
         Height          =   255
         Left            =   3870
         TabIndex        =   25
         Top             =   3735
         Width           =   1155
         _ExtentX        =   2032
         _ExtentY        =   466
         _Version        =   393216
         Appearance      =   0
         AutoTab         =   -1  'True
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin VB.TextBox tmpsubledger 
         Height          =   285
         Left            =   3510
         TabIndex        =   27
         Top             =   3465
         Width           =   3045
      End
      Begin VB.TextBox DESCRIPTION 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   930
         MaxLength       =   50
         TabIndex        =   26
         Top             =   5055
         Width           =   1725
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H0078CFE9&
         BorderWidth     =   4
         Height          =   1050
         Left            =   120
         Top             =   8520
         Width           =   8970
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher Type"
         Height          =   255
         Left            =   210
         TabIndex        =   33
         Top             =   390
         Width           =   1275
      End
      Begin VB.Label Total 
         Caption         =   "Total :- "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5985
         TabIndex        =   32
         Top             =   7635
         Width           =   825
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher No."
         Height          =   300
         Left            =   6390
         TabIndex        =   31
         Top             =   345
         Width           =   1365
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher Date"
         Height          =   255
         Left            =   3105
         TabIndex        =   30
         Top             =   330
         Width           =   1275
      End
      Begin VB.Label Label4 
         Caption         =   "Press F2 for Search A Voucher"
         Height          =   255
         Left            =   3405
         TabIndex        =   29
         Top             =   7665
         Width           =   2475
      End
   End
End
Attribute VB_Name = "Voucherform_last"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim CON As ADODB.Connection
Dim rs1 As ADODB.Recordset
Dim RS As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim cmd As New ADODB.Command
Dim lastrow, lastcol As Integer
Dim maxrow As Integer
Dim clicksave As Boolean
Dim tmplistindex As Long
Dim Edit As Boolean
Dim addmode As Boolean
Dim keyEdit As Boolean
Dim rs_S As New ADODB.Recordset
Dim searchmode As Boolean
Dim varVtype As String
Dim varVdate As Variant
Dim VarVno As Integer
Dim flagsearch As Boolean
Sub refreshTotal()
  
  totalcredit.Text = ""
  totaldebit.Text = ""
  For I = 1 To grid1.rows - 1
     totaldebit.Text = Val(totaldebit.Text) + Val(grid1.TextMatrix(I, 2))
     totalcredit.Text = Val(totalcredit.Text) + Val(grid1.TextMatrix(I, 3))
  Next
  totaldebit.Text = Format(totaldebit.Text, "0.00")
  totalcredit.Text = Format(totalcredit.Text, "0.00")



End Sub
Private Sub Command1_Click()

'addmode = False
'vtype.Enabled = False
'vdate.Enabled = False
'vno.Enabled = False
'Commandsearch.Enabled = True
'Grid1.Enabled = False
'searchmode = True
'If Not rs2.EOF Then
'     rs2.MoveNext
'     If rs2.EOF And rs2.RecordCount > 0 Then
'           rs2.MoveLast
'           Exit Sub
'     End If
'     totaldebit.Text = ""
'     totalcredit.Text = ""
'     vtype = rs2!vouchertype
'     vdate = rs2!voucherDATE
'     vno = rs2!VOUCHERNUMBER
'     Me.vtype_LostFocus
'     Me.vdate_LostFocus
'     Me.vno_LostFocus
'
'
'
'End If
'



End Sub

Private Sub Command2_Click()

'If Not rs2.BOF Then
'    addmode = False
'    searchmode = True
'
'    rs2.MovePrevious
'    If rs2.BOF And rs2.RecordCount > 0 Then
'           rs2.MoveFirst
'           Exit Sub
'    End If
'     totaldebit.Text = ""
'     totalcredit.Text = ""
'
'     vtype = rs2!vouchertype
'     vdate = rs2!voucherDATE
'     vno = rs2!VOUCHERNUMBER
'     Me.vtype_LostFocus
'     Me.vdate_LostFocus
'     Me.vno_LostFocus
     
'End If



End Sub

Sub Commandabandon_Click()
'If SAVED Then
    For I = 1 To grid1.rows - 1
        grid1.Row = I
          If I Mod 2 = 0 Then
               grid1.Text = "."
          Else
          For J = 0 To 5
            grid1.Col = J
            grid1.Text = " "
            grid1.Text = ""
            
          
        Next J
          grid1.Col = 0
       End If
       
 
       
    Next I
    
    
   
    grid1.Refresh
    totalcredit = ""
    totaldebit = ""
    Gencombo.Visible = False
    SubCombo.Visible = False
    credit.Visible = False
    debit.Visible = False
    crno.Visible = False
    grid1.Row = 1
    grid1.Col = 0
    addmode = False

Commandmasteradd.Enabled = True
Commandedit.Enabled = True
Commanddelete.Enabled = True
Commandsearch.Enabled = True
CommandPrint.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
mnuMenu_ = "menujournalvoucher"
SetButton Commandmasteradd, Commandedit, Commandsave, Commanddelete
Commandsave.Enabled = False
Edit = False
addmode = False



End Sub

Private Sub Commanddelete_Click()

On Error GoTo Del


If MsgBox("Are You Sure, Delete it ?", vbYesNo) = vbNo Then
        Exit Sub
Else
    con.Execute "delete  from vouchers where " & stringyear & " and  vouchertype='" + Trim(vtype.Text) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & vdate & "',103) And vouchernumber = " + Trim(vno.Text)
    Commandabandon_Click
               
End If


Exit Sub
Del:
MsgBox "" & err.DESCRIPTION

End Sub

Private Sub Commandedit_Click()
    searchmode = False
    Gencombo.Clear
    
    Set RS = New ADODB.Recordset
    RS.Open "Select * from gledger where " & stringyear & " order by gledger", con, adOpenDynamic, adLockReadOnly, adCmdText
    If Not RS.EOF Then
        Do While Not RS.EOF
            Gencombo.AddItem RS(1)
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.close
    Commandedit.Enabled = False
    Edit = True
    
    
    mnuMenu_ = "menujournalvoucher"
    SetButton Commandmasteradd, Commandedit, Commandsave, Commanddelete

    
    
    varVtype = Trim(vtype)
    varVdate = Format(vdate.Text, "dd/mm/yyyy")
    VarVno = Val(vno)
    grid1.Enabled = True
    vtype.Enabled = True
    vdate.Enabled = True
    vtype.SetFocus
    vno.Enabled = True
    Commandsave.Enabled = True
    addmode = False
    Commandmasteradd.Enabled = False
    Commandedit.Enabled = False
    Commanddelete.Enabled = True
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
  
  searchmode = False
  Commandabandon_Click
  Dim tRS1 As New ADODB.Recordset
  
  
  If tRS1.State = 1 Then tRS1.close
  tRS1.Open "SELECT top 100 VoucherNumber FROM vouchers where " & stringyear & " order by vsno", con, adOpenStatic, adLockReadOnly
  If tRS1.RecordCount <= 0 Then
      vtype.Text = "J"
      vdate = Format(Date, "dd/MM/yyyy")
  End If
    Command1.Enabled = False
    Command2.Enabled = False
    addmode = True
    Edit = False
    
    Set rs1 = New ADODB.Recordset
    Set RS = New ADODB.Recordset
    rs1.Open "Select max(VoucherNumber) from vouchers where " & stringyear & " and vouchertype='" + Trim(vtype.Text) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & vdate & "',103)", con, adOpenDynamic, adLockReadOnly, adCmdText
    If Not rs1.BOF Then
        If Not rs1(0) Then
            c = rs1(0)
        Else
            c = 0
        End If
    Else
        c = 0
    End If
    rs1.close
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
    
    
    If rs1.State = 1 Then
        rs1.close
    End If
    
    
    
    grid1.Row = 0
    grid1.Col = 0
    
    grid1.SelectionMode = flexSelectionFree
    grid1.ColWidth(0, 1) = 11111
    X = grid1.MergeCol(1)
    
    DESCRIPTION.Width = grid1.Width - 300
    DESCRIPTION.Left = 200
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
    addmode = True
    vno.Enabled = False

End Sub

Private Sub CommandPrint_Click()
Dim sss1 As String

sss1 = vno.Text & "-" & vdate.Text & "-" & vtype.Text

con.Execute "update vouchers set  printId='" & sss1 & "' where  vouchertype='" + Trim(vtype.Text) + "' and convert(smalldatetime,voucherdate,103)= convert (smalldatetime,'" & vdate & "',103) And vouchernumber = " + Trim(vno.Text)

If MsgBox("Want to View  ? ", vbQuestion + vbYesNo) = vbYes Then
    DSNNew
    MainMenu.cr1.Reset
    MainMenu.cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    MainMenu.cr1.ReportFileName = rptPath & "\" & main.directory & "\Voucher.rpt"
    MainMenu.cr1.ReplaceSelectionFormula "{vouchers.printid}='" & sss1 & "'"
    MainMenu.cr1.WindowShowPrintBtn = True
    MainMenu.cr1.WindowShowPrintSetupBtn = True
    MainMenu.cr1.WindowShowSearchBtn = True
    MainMenu.cr1.WindowState = crptMaximized
    MainMenu.cr1.WindowShowRefreshBtn = True
    MainMenu.cr1.WindowShowExportBtn = True
    MainMenu.cr1.Action = 1

End If




End Sub

Private Sub CommandReturn_Click()
    Unload Me
    ''MainMenu.Toolbar1.Visible = True
End Sub
Private Sub Commandsave_Click()

On Error GoTo save_


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
        RS.close
    End If
    If Edit = False Then
        RS.Open "select * from vouchers where " & stringyear & " and VoucherNumber<=0", con, adOpenDynamic, adLockOptimistic
    Else
        
        con.Execute "Delete  from vouchers where " & stringyear & " and VoucherType='" + Trim(varVtype) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & varVdate & "',103) and vouchernumber=" + Trim(VarVno) + ""
        con.Execute "Delete  from vouchers where " & stringyear & " and VoucherType='" + Trim(varVtype) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & varVdate & "',103) and vouchernumber=" + Trim(VarVno) + ""
        
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        RS.Open "select * from vouchers where " & stringyear & " and VoucherNumber<=0", con, adOpenDynamic, adLockPessimistic
    End If
  If addmode = True Then
     rs5.Open "Select max(VoucherNumber) from vouchers where " & stringyear & " and vouchertype='" + Trim(vtype.Text) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & vdate & "',103)", con, adOpenDynamic, adLockReadOnly, adCmdText
     If Not rs5.BOF Then
        If Not rs5(0) Then
            c = rs5(0)
        Else
            c = 0
       End If
    Else
        c = 0
    End If
    If rs5.State = 1 Then rs5.close
    vno.Text = Str(c + 1)
   End If
  
    grid1.Row = I
    grid1.Col = 0
    For I = 1 To grid1.rows Step 2
        grid1.Row = I
        grid1.Col = 0
        If grid1.Text <> "" Then
            RS.AddNew
            RS(0) = vtype.Text
            RS(1) = vdate.Text
            RS(2) = vno.Text
            RS(3) = grid1.Text
            grid1.Col = 1
            If grid1.Text = "" Then
               RS(4) = ""
            Else
               RS(4) = grid1.Text
            End If
            grid1.Col = 2
            If grid1.Text <> "" Then
                RS(5) = Val(grid1.Text)
                RS(6) = "D"
            Else
                grid1.Col = 3
                RS(5) = Val(grid1.Text)
                RS(6) = "C"
            End If
            grid1.Col = 4
            RS(7) = grid1.Text
            
            grid1.Col = 5
            RS!UserName = grid1.Text

            
            grid1.Col = 0
            grid1.Row = grid1.Row + 1
            If grid1.Text <> "" Then
                RS(9) = grid1.Text
            End If
            
            SAVED = True
            RS!setupid = setupid
            RS!fyear = session
            
            RS.update
        End If
        
    Next
    
    Set rs2 = New ADODB.Recordset
    rs2.Open "Select DISTINCT vouchertype,voucherdate,vouchernumber from vouchers where " & stringyear & " order by vouchertype,voucherdate,vouchernumber", con, adOpenStatic, adLockReadOnly
    CommandPrint.Visible = True
    RS.close
  
    MsgBox " RECORD SAVED ... ", vbInformation
 
    
    
    DoEvents
    DoEvents
    
    Commandsave.Enabled = False
    Commandmasteradd.Enabled = True
    Commandmasteradd.SetFocus
    Commandedit.Enabled = True
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
    grid1.Enabled = False
    SAVED = False
    
    mnuMenu_ = "menujournalvoucher"
   SetButton Commandmasteradd, Commandedit, Commandsave, Commanddelete

    
    Else
    
    'If Commandedit.Enabled = True Then
    MsgBox "Please Check That Total Credit and Total Debit is Differ "
    'End If
    End If
    addmode = False



Exit Sub
save_:
MsgBox "" & err.DESCRIPTION


End Sub

Private Sub Commandsearch_Click()
    searchmode = True
    Edit = False
    vno.Enabled = True
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
       grid1.Text = Val(credit.Text)
       
       grid1.Text = Format(grid1.Text, "0.00")
    Else
        If grid1.Col <> 4 And grid1.Col <> 0 Then
           'Grid1.Text = ""
        End If
    End If
End Sub
Private Sub credit_GotFocus()
    If maxrow < grid1.Row Then
            maxrow = grid1.Row
    End If
    If searchmode = False Then
        'HIT
        'credit.SetFocus
        SendKeys "{END}"
        SendKeys "+{HOME}"
  End If
        
End Sub

Private Sub credit_KeyDown(KeyCode As Integer, Shift As Integer)
   If (KeyCode >= 48 And KeyCode <= 57) Then
      grid1.TextMatrix(grid1.RowSel, 5) = UserName
   End If

End Sub

Private Sub credit_KeyPress(KeyAscii As Integer)


If KeyAscii = 13 Then
'    If Val(credit) > 0 Or Val(debit) > 0 Then
    Dim changed As Boolean
    Dim HIT As Boolean
    HIT = False
    changed = False
    RRR = grid1.Row
    CCC = grid1.Col
    grid1.Col = 3
    totalcredit.Text = ""
    For I = 1 To maxrow Step 1
        grid1.Row = I
        totalcredit.Text = Val(totalcredit.Text) + Val(grid1.Text)
    Next
    totalcredit.Text = Format(totalcredit.Text, "0.00")
    grid1.Row = RRR
    grid1.Col = CCC
    grid1.Col = 3
    If credit.Text = "" Then
            grid1.Text = ""
            grid1.Col = 2
            Grid1_Click
            Exit Sub
     End If
    If grid1.Col = lastcol And grid1.Row = lastrow Then
        grid1.Col = 0
        If Trim(grid1.Col) <> "" Then
            HIT = True
        End If
        grid1.Col = 3
        grid1.Col = grid1.Col + 1
        refreshTotal
        changed = True
    End If
    RRR = grid1.Row
    CCC = grid1.Col
    grid1.Col = 2
    If grid1.Row <> lastrow Or grid1.Col <> lastcol Then
        grid1.Text = ""
    End If
   
    
    grid1.Row = RRR
    grid1.Col = CCC
    
    
    
    If changed And HIT Then
        Grid1_Click
    End If
    Else
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
    RRR = grid1.Row
CCC = grid1.Col
grid1.Row = lastrow
grid1.Col = lastcol
If Trim(grid1.Text) <> "" Then
    grid1.Row = RRR
    grid1.Col = CCC
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
On Error GoTo err_

If KeyAscii = 13 Then
    Dim changed As Boolean
    Dim HIT As Boolean
    HIT = False
    changed = False
    
    
    If grid1.Col = lastcol And grid1.Row = lastrow Then
        crno = UCase(crno)
        grid1.Col = 0
        If Trim(grid1.Text) <> "" Then
            HIT = True
        End If
        
        
        If (grid1.TextMatrix(grid1.RowSel, 5) = "" Or grid1.TextMatrix(grid1.RowSel, 5) = ".") Then
        grid1.TextMatrix(grid1.RowSel, 5) = UserName
        End If
        
        'Grid1.TextMatrix(Grid1.RowSel + 1, 5) = UserName
        
        grid1.Row = grid1.Row + 1
        
        changed = True
    End If
    'crno.Text = ""
    If changed And HIT Then
        Grid1_Click
    End If
End If


Exit Sub
err_:
MsgBox "" & err.DESCRIPTION


End Sub

Private Sub Crno_LostFocus()
    crno.Visible = False
End Sub
Private Sub debit_Change()
If Val(debit.Text) <> 0 Then
    grid1.Text = Val(debit.Text)
    grid1.Text = Format(grid1.Text, "0.00")
Else
'If Grid1.col <> 4 Then  Grid1.Text = ""
 If grid1.Col <> 4 And grid1.Col <> 0 Then
   ' Grid1.Text = ""
 End If
End If

End Sub
Private Sub debit_GotFocus()
   
     cl = grid1.Col
     grid1.Col = 0
   If grid1.Text <> "" Then
     If grid1.Text = "CASH-IN-HAND" And vtype.Text = "P" Then
        grid1.Col = 2
        grid1.Text = ""
        grid1.Col = 3
        debit.Visible = False
        Grid1_Click
        
    Else
       grid1.Col = cl
      ' debit.SetFocus
       SendKeys "{END}"
       SendKeys "+{HOME}"
       
    End If
   

    If maxrow < grid1.Row Then
        maxrow = grid1.Row
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
   If (KeyCode >= 48 And KeyCode <= 57) Then
      grid1.TextMatrix(grid1.RowSel, 5) = UserName
   End If
End Sub

Private Sub debit_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim changed As Boolean
    Dim HIT As Boolean
    changed = False
    HIT = False
    RRR = grid1.Row
    CCC = grid1.Col
    grid1.Col = 2
    totaldebit.Text = ""
    For I = 1 To maxrow
        grid1.Row = I
        totaldebit.Text = Val(totaldebit.Text) + Val(grid1.Text)
    Next
    totaldebit.Text = Format(totaldebit.Text, "0.00")
    grid1.Row = RRR
    grid1.Col = CCC
    If grid1.Col = lastcol And grid1.Row = lastrow Then
        grid1.Col = 0
        If Trim(grid1.Text) <> "" Then
            HIT = True
        End If
        
        grid1.Col = 2
        If debit.Text = "" Then grid1.Text = ""

        If grid1.Text <> "" Then
            grid1.Col = 4
        Else
            grid1.Col = 0
            If grid1.Text = "CASH-IN-HAND" And vtype.Text = "R" Then
                 grid1.Col = 4
                  grid1.Col = 3
                  If grid1.Text = "" Then
                      debit.Visible = True
                      credit.Visible = False
                      grid1.Col = 2
                      debit.SetFocus
                     Exit Sub
                  End If
                 
            Else
                 grid1.Col = 2
                 If debit.Text = "" Then grid1.Text = ""
                 grid1.Col = 3
            End If
              
            
            
        End If
        changed = True
    End If
    RRR = grid1.Row
    CCC = grid1.Col
    grid1.Col = 2
    If grid1.Row <> lastrow Or grid1.Col <> lastcol Then
        grid1.Text = ""
    End If
    grid1.Row = RRR
    grid1.Col = CCC
    If changed And HIT Then
        Grid1_Click
    End If
End If
End Sub

Private Sub debit_LostFocus()
RRR = grid1.Row
CCC = grid1.Col
grid1.Row = lastrow
grid1.Col = lastcol
If Trim(grid1.Text) <> "" Then
    grid1.Row = RRR
    grid1.Col = CCC
   
    debit.Visible = False
Else
     refreshTotal
End If
debit = ""
End Sub


Private Sub DESCRIPTION_GotFocus()
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
    If grid1.Col = lastcol And grid1.Row = lastrow Then
        grid1.Col = 0
        HIT = True
        grid1.Row = grid1.Row + 1
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

Command1.SetFocus

mnuMenu_ = "menujournalvoucher"
SetButton Commandmasteradd, Commandedit, Commandsave, Commanddelete

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then

If addmode = False Then
  
  FindCombo.Visible = True
  FindCombo.SetFocus

  Dim rs7 As New ADODB.Recordset
  Set rs7 = con.Execute("exec voucherQry '" & session & "','" & main.setupid & "'")
  
  If rs7.EOF = False Then
   FindCombo.Clear
   rs7.MoveFirst
   While Not rs7.EOF
       FindCombo.AddItem rs7(0) & Space(1) & rs7(1) & Space(1) & rs7(2)
       rs7.MoveNext
   Wend
  End If
  
End If

End If

If KeyCode = 27 Then Unload Me


End Sub
Private Sub Form_Resize()

panel.Left = (Me.ScaleWidth - panel.Width) / 2
panel.Top = (Me.ScaleHeight - panel.Height) / 2

End Sub

Private Sub Gencombo_Change()
 If grid1.Col = lastcol And grid1.Row = lastrow Then
        grid1.Text = Gencombo.Text
        grid1.Col = 0
 End If
End Sub

Private Sub genCombo_Click()
    grid1.Text = Gencombo.Text
    If vtype.Text = "R" And grid1.Text = "CASH-IN-HAND" Then
        grid1.Col = 2
        db = grid1.Text
        grid1.Col = 3
        cr = grid1.Text
        If db = "" Then
            grid1.Col = 2
            grid1.Text = cr
            grid1.Col = 3
            grid1.Text = ""
       End If
    End If
   If vtype.Text = "P" And grid1.Text = "CASH-IN-HAND" Then
        grid1.Col = 2
        db = grid1.Text
        grid1.Col = 3
        cr = grid1.Text
        If cr = "" Then
            grid1.Col = 3
            grid1.Text = db
            grid1.Col = 2
            grid1.Text = ""
       End If
    End If
    grid1.Col = 1
    grid1.Text = ""
    grid1.Col = 0
End Sub
Private Sub Form_Load()

Me.Top = 0
Me.Left = 0

Me.Width = 12700
Me.Height = 10300

Me.Caption = "Voucher Entry"


varVtype = ""
varVdate = ""
VarVno = 0

searchmode = False
addmode = False
vtype.Enabled = False
vdate.Enabled = False
vno.Enabled = False

Commandsearch.Enabled = True
grid1.Enabled = False
Edit = False
    'Grid1.Top = 450
    'Grid1.Left = 60
    Me.Top = 0
    Me.Left = 0
    grid1.rows = 300
    grid1.Cols = 0
    grid1.Cols = 6
    
    For I = 0 To grid1.rows - 1
       grid1.RowHeight(I) = 250
    Next
    
    grid1.Row = 0
    grid1.Col = 0
    grid1.MergeCells = flexMergeRestrictRows
    For I = 1 To grid1.rows - 1
        grid1.RowHeight(I) = 350
        If I Mod 2 = 0 Then
            grid1.Row = I
            For J = 0 To 5
                grid1.Col = J
                grid1.Text = "."
                grid1.MergeRow(I) = True
            Next
        End If
    Next
    
   
   'Grid1.MergeCells = flexMergeRestrictRows
    grid1.Row = 0
    grid1.Col = 0
    grid1.ColAlignment(0) = 0
    grid1.ColAlignment(1) = 0
    grid1.ColWidth(0) = 3000
    grid1.ColWidth(1) = 3700
    grid1.ColWidth(2) = 1400
    grid1.ColWidth(3) = 1400
    grid1.ColWidth(4) = 1800
    grid1.Text = "Gen. Ledger"
    grid1.Col = 1
    grid1.Text = "Sub. Ledger"
    grid1.Col = 2
    grid1.Text = "Amount. (Dr.)"
    'totaldebit.Left = Grid1.CellLeft + 50
    grid1.Col = 3
    grid1.Text = "Amount. (Cr.)"
    'totalcredit.Left = Grid1.CellLeft + 60
    grid1.Col = 4
    grid1.Text = "C/B No."
    
    
    grid1.TextMatrix(0, 5) = "User"
    
    
    crno.Height = grid1.CellHeight
    
    
        If (UserName = "y" Or UserName = "v" Or LCase(UserName) = "admin") Then
       grid1.ColWidth(5) = 500
     Else
        grid1.ColWidth(5) = 0
    End If

    
    'totalcredit.Width = crno.Width
    'totaldebit.Width = totalcredit.Width
    'totaldebit.Height = crno.Height + 10
    'totalcredit.Height = totaldebit.Height

    
    Set RS = New ADODB.Recordset
    RS.PageSize = 2
    RS.Open "Select top 100 * from vouchers where " & stringyear & "  order by vouchertype,voucherdate,vouchernumber", con, adOpenStatic
    If Not RS.BOF Then
        vtype = RS!vouchertype
        vdate = RS!voucherDATE
        vno = RS!VOUCHERNUMBER
        Me.vtype_LostFocus
        Me.vdate_LostFocus
        Me.vno_LostFocus
    End If
    
    
 
SubCombo.Clear
Gencombo.Clear

Set cmd = New ADODB.Command
cmd.ActiveConnection = con
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "fatch_gledger"

Set rs_S = cmd.Execute
    If Not rs_S.EOF Then
        Do While Not rs_S.EOF
            Gencombo.AddItem rs_S!gledger
            If Not rs_S.EOF Then
                rs_S.MoveNext
            End If
        Loop
    End If
    rs_S.close



Set cmd.ActiveConnection = Nothing
'=================================================
    
mnuMenu_ = "menujournalvoucher"
SetButton Commandmasteradd, Commandedit, Commandsave, Commanddelete
Commandsave.Enabled = False



Set rs_S = New ADODB.Recordset
rs_S.Open "Select SUBLEDGER,gledger  from  SLEDGER where " & stringyear & "" & _
" ORDER BY gledger,SUBLEDGER", con, adOpenKeyset, adLockOptimistic


BackColorFrom Me

End Sub



Private Sub Gencombo_GotFocus()
    grid1.Col = 0
    Gencombo.Text = grid1.Text
    If Gencombo.Enabled = True Then
    If searchmode = False Then
    HIT
    End If
    End If
End Sub

Private Sub Gencombo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 115 Then
   If MsgBox("Are You Sure, Delete it ?", vbYesNo) = vbYes Then
      If grid1.Row >= 1 Then
           grid1.RemoveItem grid1.Row
           grid1.RemoveItem grid1.Row
           'Gencombo.Text = ""
           Gencombo.Visible = False
           Me.DESCRIPTION.Visible = False
           refreshTotal
           grid1.SetFocus
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
        If RS.State = 1 Then RS.close
        RS.Open "GLEDGER", con, adOpenStatic, adLockReadOnly, adCmdTable
        If Trim(Gencombo.Text) <> "" Then
        'SendKeys "{DOWN}"
        If Not RS.BOF Then
                RS.MoveFirst
                RS.Find "GLEDGER='" + Trim(Gencombo.Text) + "'"
                If RS.EOF Then
                    RS.close
                    Exit Sub
                End If
            End If
        Else
           Gencombo.Visible = False
        End If
        RS.close
        
        If grid1.Col = lastcol And grid1.Row = lastrow Then
            If Trim(grid1.Text) <> "" Then
                HIT = True
            End If
            grid1.Col = grid1.Col + 1
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
   ' rs.Open "GLEDGER", CON, adOpenDynamic, adLockReadOnly, adCmdTable
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
If grid1.Enabled = True Then
Dim check As Boolean
Dim Row, Col As Integer
Row = grid1.Row
Col = grid1.Col
check = True
If grid1.Row Mod 2 <> 0 Then
    DESCRIPTION.Visible = False
    Dim gr As MSHFlexGrid
    Set gr = grid1
    If grid1.Row >= 3 Then
        check = False
        gr.Row = grid1.Row - 2
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
    grid1.Row = Row
    grid1.Col = Col
    credit.Visible = False
    debit.Visible = False
    SubCombo.Visible = False
    crno.Visible = False
    Gencombo.Visible = False
    If check Then
       If grid1.Col = 0 Then
          credit.Visible = False
          debit.Visible = False
          SubCombo.Visible = False
          crno.Visible = False
          
          Gencombo.Text = grid1.Text
          Gencombo.Visible = True
          Gencombo.Left = grid1.CellLeft + 30
          Gencombo.Top = grid1.Top + grid1.CellTop - 50
          Gencombo.Width = grid1.ColWidth(grid1.Col)
          
          Gencombo.ZOrder
          Gencombo.SetFocus
       End If
        
        
        If grid1.Col = 1 Then
            tmpsubledger = ""
            SubCombo.Text = ""
            If grid1.Text <> "" Then
               tmpsubledger = grid1.Text
            End If
            
            Gencombo.Visible = False
            credit.Visible = False
            debit.Visible = False
            crno.Visible = False
            SubCombo.Visible = True
            SubCombo.Text = grid1.Text
            SubCombo.Left = grid1.CellLeft + 30
            SubCombo.Top = grid1.Top + grid1.CellTop - 50
            SubCombo.ZOrder
            SubCombo.SetFocus
            SubCombo.Width = grid1.ColWidth(grid1.Col)
            'SubCombo.Height = Grid1.RowHeight(Grid1.Row)
        End If
        If grid1.Col = 2 Then
            Gencombo.Visible = False
            SubCombo.Visible = False
            credit.Visible = False
            crno.Visible = False
            grid1.Col = 3
            If Trim(grid1.Text) = "" Then
                grid1.Col = 2
                If grid1.Text <> "" Then
                    debit.Text = grid1.Text
                End If
                debit.Visible = True
                debit.Left = grid1.CellLeft + 30
                debit.Top = grid1.Top + grid1.CellTop - 20
                debit.Width = grid1.ColWidth(grid1.Col)
                debit.Height = grid1.CellHeight
                debit.ZOrder
                debit.SetFocus
                
           Else
                       
              grid1.Col = 3
          
            End If
          End If
        If grid1.Col = 3 Then
            Gencombo.Visible = False
            SubCombo.Visible = False
            debit.Visible = False
            crno.Visible = False
            grid1.Col = 2
            If Trim(grid1.Text) = "" Then
                grid1.Col = 3
                If grid1.Text <> "" Then
                credit.Text = grid1.Text
                End If
                credit.Visible = True
                credit.Left = grid1.CellLeft + 30
                credit.Top = grid1.Top + grid1.CellTop - 25
                credit.Width = grid1.ColWidth(grid1.Col)
                
                credit.Height = grid1.CellHeight
                
                credit.ZOrder
                credit.SetFocus
                
            End If
        End If
        If grid1.Col = 4 Then
            Gencombo.Visible = False
            SubCombo.Visible = False
            debit.Visible = False
            credit.Visible = False
            crno.Text = grid1.Text
            crno.Visible = True
            crno.Left = grid1.CellLeft + 30
            crno.Top = grid1.Top + grid1.CellTop - 25
            crno.Width = grid1.ColWidth(grid1.Col)
            crno.Height = grid1.CellHeight
            
            crno.ZOrder
            
            crno.SetFocus
           
        End If
    lastrow = grid1.Row
    lastcol = grid1.Col
    Else
    MsgBox "Previous entry didn't complete"
    End If
Else
    If grid1.Row >= 2 Then
        RRR = grid1.Row
        CCC = grid1.Col
        grid1.Row = grid1.Row - 1
        grid1.Col = 0
        If Trim(grid1.Text) <> "" Then
            grid1.Row = RRR
            grid1.Col = CCC
            If grid1.Col = 0 Then
                DESCRIPTION.Left = grid1.CellLeft + 30
                DESCRIPTION.Top = grid1.Top + grid1.CellTop - 25
                DESCRIPTION.Width = grid1.CellWidth  '' Grid1.ColWidth(Grid1.col)
                crno.Visible = False
                Gencombo.Visible = False
                SubCombo.Visible = False
                credit.Visible = False
                debit.Visible = False
                DESCRIPTION.Visible = True
                DESCRIPTION.ZOrder
                'DESCRIPTION.Height = Grid1.CellHeight
                DESCRIPTION.SetFocus
            End If
        End If
        lastrow = grid1.Row
        lastcol = grid1.Col
    End If
End If
End If

End Sub


Private Sub SubCombo_Change()
    
    
    If grid1.Col = lastcol And grid1.Row = lastrow Then
        grid1.Text = SubCombo.Text
    End If
    
    
    
    
End Sub
Private Sub SubCombo_Click()
    grid1.Text = SubCombo.Text
End Sub
Private Sub SubCombo_GotFocus()

   tmplistindex = 0
   Dim X(5000) As String
   Gencombo.Visible = False
   SubCombo.Clear
   tmpsubledger.Text = grid1.Text
   
   
      
    grid1.Col = 0
   If grid1.Text <> "" Then
       
       
       If rs_S.State = 1 Then rs_S.close
       rs_S.Open "Select SUBLEDGER,gledger  from  SLEDGER where " & stringyear & " and gledger='" + Trim(grid1.Text) + "'  ORDER BY gledger,SUBLEDGER", CCON, adOpenStatic, adLockReadOnly
       If Not rs_S.EOF Then
        Dim X1  As Integer
          X1 = 0
          CV = 0
          Do While Not rs_S.EOF
             
           If rs_S!gledger = Trim(grid1.Text) Then
             SubCombo.AddItem rs_S(0)
           End If
             
             
             
             If rs_S(0) = tmpsubledger.Text Then
                 tmplistindex = CV
                 CV = X1
             End If
             X1 = X1 + 1
             rs_S.MoveNext
         Loop

    End If
    
    
    grid1.Col = CCC
    grid1.Col = 1
    If SubCombo.ListCount = 0 Then
        grid1.Col = 2
        Grid1_Click
    End If
    SubCombo.Text = tmpsubledger.Text
    If CV >= 1 Then
       SubCombo.ListIndex = CV
    End If
    
     Set cmd.ActiveConnection = Nothing
    
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
    RRR = grid1.Row
    CCC = grid1.Col
    grid1.Row = lastrow
    grid1.Col = 0
    If RS.State = 1 Then
        RS.close
    End If
    
    If Trim(SubCombo.Text) <> "" And Trim(grid1.Text) <> "" Then
        SendKeys "{down}"
        RS.Open "select * from SLEDGER where " & stringyear & " and subledger='" + Trim(SubCombo.Text) + "' and gledger='" + Trim(grid1.Text) + "'", con, adOpenForwardOnly, adLockReadOnly, adCmdText
        If RS.BOF Or RS.EOF Then
            SubCombo.Visible = True
            SubCombo.SetFocus
        Else
            valid = True
        End If
        RS.close
    Else
        If Trim(SubCombo.Text) = "" Then
            SubCombo.Visible = True
            SubCombo.SetFocus
       End If
    End If
    grid1.Row = RRR
    grid1.Col = CCC
    If valid Then
        If grid1.Col = lastcol And grid1.Row = lastrow Then
            grid1.Col = 0
            If Trim(grid1.Col) <> "" Then
                HIT = True
            End If
            grid1.Col = 3
            If grid1.Text <> "" Then
            
            Else
                grid1.Col = 2
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
'        rs.Open "select * from SLEDGER where " & stringyear & " and subledger='" + Trim(SubCombo.Text) + "' and gledger='" + Trim(Grid1.Text) + "'", CON, adOpenDynamic, adLockReadOnly, adCmdText
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
Sub checkData()
    Set RS = New ADODB.Recordset
    RS.Open "Select * from setup1 where " & stringyear & " and yarfrom>datevalue('" + vdate.Text + "') and yarto<=datevalue('" + vdate.Text + "')", con, adOpenDynamic, adLockReadOnly, adCmdText
    If RS.BOF = True Then
       vdate.SetFocus
    End If
End Sub
Private Sub vdate_GotFocus()
HIT
End Sub

Private Sub vdate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        Set RS = New ADODB.Recordset
        RS.Open "setup1", con
        
        If CDate(vdate.Text) >= RS!yarfrom And CDate(vdate.Text) <= RS!yarto Then
        Else
           MsgBox "Enter Valid Date ....", vbInformation
           vdate.SetFocus
           Exit Sub
        End If
        RS.close
 
    
    
        If addmode = True Or Edit = True Then
           grid1.Row = 1
           grid1.Col = 0
           grid1.SetFocus
           'Print 1111
           Grid1_Click
        Else
           'SendKeys "{tab}"
           vno.SetFocus
        End If
    End If
End Sub
    Sub vdate_LostFocus()
         
    
    If Not IsDate(vdate.Text) Then
       vdate.SetFocus
       Exit Sub
    End If
    
    
    If addmode = True Then
        If RS.State = 1 Then RS.close
        Set RS = New ADODB.Recordset
        RS.Open "Select max(VoucherNumber) from vouchers where " & stringyear & " and vouchertype='" + Trim(vtype.Text) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & vdate & "',103)", con, adOpenDynamic, adLockReadOnly, adCmdText
        If Not RS.BOF Then
            If Not RS(0) Then
                c = RS(0)
            Else
                c = 0
            End If
        Else
            c = 0
        End If
        RS.close
        vno.Text = Str(c + 1)
    Else
       If varVtype = "" Then Exit Sub
       If varVdate = "" Then Exit Sub
       If searchmode = True Then Exit Sub
       If Trim(vtype.Text) <> Trim(varVtype) Or CDate(Format(varVdate, "dd/mm/yyyy")) <> Format(vdate, "dd/mm/yyyy") Then
            If RS.State = 1 Then RS.close
            Set RS = New ADODB.Recordset
            RS.Open "Select max(VoucherNumber) from vouchers where " & stringyear & " and vouchertype='" + Trim(vtype.Text) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & vdate & "',103)", con, adOpenDynamic, adLockReadOnly, adCmdText
            If Not RS.BOF Then
                If Not RS(0) Then
                    c = RS(0)
                Else
                    c = 0
                End If
            Else
                c = 0
            End If
            RS.close
            vno.Text = Str(c + 1)
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

If grid1.Enabled = True Then
grid1.SetFocus
SendKeys "{UP}"
End If

If Val(vno) <> 0 Then
   
   If RS.State = 1 Then RS.close
   RS.Open "Select top 1000 * from vouchers where " & stringyear & " and VoucherType='" + Trim(vtype.Text) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & vdate & "',103)  and vouchernumber=" & Trim(vno.Text) & " order by vsno", con, adOpenDynamic, adLockOptimistic, adCmdText  'and VoucherNumber=" + Val(Trim(vno.Text)) + "", con, adOpenDynamic, adLockOptimistic, adCmdText
     
     For I = 1 To grid1.rows - 1
       grid1.Row = I
       If I Mod 2 = 0 Then
          grid1.Text = "."
       Else
          For J = 0 To 5
             grid1.Col = J
            grid1.Text = ""
          Next J
          grid1.Col = 0
         End If
         
         
        
         
   Next I

   

   totaldebit.Text = 0
   totalcredit.Text = 0

   
   If Not RS.BOF Then
      
      I = 1
    
      Do While Not RS.EOF
      
      
      
      
        grid1.Row = I
        grid1.Col = 0
        grid1.Text = RS(3)
        grid1.Col = 1
         If IsNull(RS(4)) Then
              grid1.Text = ""
         Else
              grid1.Text = RS(4)
         End If
         grid1.Col = 2
         If RS(6) = "D" Then
            grid1.Text = Format(RS(5), "0.00")
            totaldebit.Text = Val(totaldebit.Text) + RS(5)
            totaldebit.Text = Format(totaldebit.Text, "0.00")
         Else
            grid1.Col = 3
            grid1.Text = Format(RS(5), "0.00")
            totalcredit.Text = Val(totalcredit.Text) + RS(5)
            totalcredit.Text = Format(totalcredit.Text, "0.00")
         End If
         grid1.Col = 4
         grid1.Text = IIf(IsNull(RS(7)), ".", RS(7))
         
         grid1.Col = 5
         grid1.Text = IIf(IsNull(RS!UserName), ".", RS!UserName)
         
         grid1.Row = grid1.Row + 1
         grid1.Col = 0
         grid1.Text = RS(9) & ""
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
     grid1.Row = 1
     grid1.Col = 0
  End If
  RS.close
  grid1.Row = 1
  grid1.Col = 0
  
  If searchmode = False Then Grid1_Click
     'Print 1
  Else
     vno.SetFocus
  End If

End Sub
Private Sub vtype_Change()
        If RS.State = 1 Then RS.close
        Set RS = New ADODB.Recordset
        RS.Open "Select max(VoucherNumber) from vouchers where " & stringyear & " and vouchertype='" + Trim(vtype.Text) + "' and voucherdate= CDate ('" + Format(vdate.Text, "dd/mm/yy") + "')", con, adOpenDynamic, adLockReadOnly, adCmdText
        If Not RS.BOF Then
           If Not RS(0) Then
                c = RS(0)
           Else
                c = 0
           End If
        Else
           c = 0
        End If
        RS.close
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

End Sub

