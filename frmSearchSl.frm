VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmSearchSl 
   BackColor       =   &H00C0FFC0&
   ClientHeight    =   8535
   ClientLeft      =   3465
   ClientTop       =   1560
   ClientWidth     =   15075
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   15075
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "&View Balance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5700
      TabIndex        =   8
      Top             =   60
      Width           =   1215
   End
   Begin VB.OptionButton Option_Narr 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Narration"
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
      Left            =   4020
      TabIndex        =   6
      Top             =   660
      Width           =   1635
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6960
      TabIndex        =   5
      Top             =   60
      Width           =   1155
   End
   Begin VB.OptionButton Option_Gn 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Genledger"
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
      Left            =   4020
      TabIndex        =   4
      Top             =   360
      Width           =   1635
   End
   Begin VB.OptionButton Option_Sl 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Subledger"
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
      Left            =   4020
      TabIndex        =   3
      Top             =   60
      Value           =   -1  'True
      Width           =   1635
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   8595
      Left            =   240
      TabIndex        =   2
      Top             =   1380
      Width           =   11775
      _cx             =   20770
      _cy             =   15161
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16761024
      ForeColor       =   0
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16761024
      BackColorAlternate=   16761024
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.TextBox txtItem 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   420
      Width           =   3675
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
      Left            =   8760
      TabIndex        =   9
      Top             =   180
      Visible         =   0   'False
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
      Left            =   8760
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   -60
      TabIndex        =   7
      Top             =   1020
      Width           =   21855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Gen. Ledger/Subledger"
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
      Left            =   300
      TabIndex        =   1
      Top             =   60
      Width           =   3555
   End
End
Attribute VB_Name = "frmSearchSl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub fillGrid()
    
    Dim f As New ADODB.Recordset
    
    If f.State = 1 Then f.Close
    s = "SELECT [gledger],[Category],[SLF] FROM [GLEDGER] where gledger like '" & txtItem.Text & "%' and " & stringyear & " ORDER BY [gledger]"
    
     
    f.Open s, CON
    Set vs.DataSource = f
    
End Sub
Sub fillNarr()
    
    Dim f As New ADODB.Recordset
    
    If f.State = 1 Then f.Close
    s = "SELECT top 500 [VoucherType],[VoucherNumber],[VoucherDate],sum([Amount]),DebitorCredit,DESCRIPTION as [Narration] FROM [VOUCHERS] where DebitorCredit='D' and DESCRIPTION like '%" & txtItem.Text & "%' and  " & stringyear & " group BY [VoucherNumber],[VoucherDate],DebitorCredit,DESCRIPTION,VoucherType"
    
      
    f.Open s, CON
    Set vs.DataSource = f
    
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub


Private Sub Command1_Click()

Screen.MousePointer = vbHourglass

d10 = 0
CON.Execute "DELETE from subledgertrail where  len(subledger)>0"
CON.Execute ("DELETE from treport where len(genledger)>0")

For K = 1 To vs.Rows - 1
  If vs.TextMatrix(K, 1) <> "" Then
     FatchClosing vs.TextMatrix(K, 2), vs.TextMatrix(K, 1), date1.Text, date2.Text, vs.Rows - 1
  End If
Next

fillGridSL

Screen.MousePointer = vbDefault

End Sub

Private Sub Label3_Click()

End Sub

Private Sub Form_Load()

If rs.State = 1 Then rs.Close
rs.Open "Select * from setup where " & stringyear & "", CON, adOpenStatic, adLockReadOnly, adCmdText
If rs.EOF = False Then
date1.Text = rs!yarfrom
date2.Text = rs!yarto
End If

End Sub

Private Sub Option_Gn_Click()
   fillGrid
End Sub

Private Sub Option_Narr_Click()
fillNarr
End Sub

Private Sub Option_Sl_Click()
   fillGridSL
End Sub

Private Sub txtItem_Change()

If Option_Gn.Value = True Then
   fillGrid
ElseIf Option_Sl.Value = True Then
  fillGridSL
Else
  fillNarr
End If

End Sub
Sub fillGridSL()
    
    Dim f As New ADODB.Recordset
    If f.State = 1 Then f.Close
    
    s = "SELECT  dbo.SLEDGER.SUBLEDGER,dbo.SLEDGER.gledger, dbo.GLEDGER.Category,dbo.SLEDGER.Balance " & _
      " FROM dbo.SLEDGER  INNER JOIN dbo.GLEDGER ON dbo.SLEDGER.gledger = dbo.GLEDGER.gledger AND dbo.SLEDGER.setupid = dbo.GLEDGER.setupid AND " & _
      " dbo.sledger.FYear = dbo.gledger.FYear where dbo.SLEDGER.SUBLEDGER like '" & txtItem.Text & "%' and SLEDGER.fyear='" & main.session & "' and SLEDGER.setupid='" & main.setupid & "' order by dbo.SLEDGER.SUBLEDGER"
    
    f.Open s, CON
    Set vs.DataSource = f
     
    
End Sub

Private Sub vs_DblClick()

If Option_Narr.Value = True Then

If vs.TextMatrix(vs.RowSel, 1) <> "" Then
  vtypes = vs.TextMatrix(vs.RowSel, 1)
  vnumbers = vs.TextMatrix(vs.RowSel, 2)
  vdates = vs.TextMatrix(vs.RowSel, 3)
  Voucherform.Show 1
End If

End If

End Sub
