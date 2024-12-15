VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmProductMaster 
   Caption         =   "Product Master"
   ClientHeight    =   10095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14940
   Icon            =   "frmProductMaster.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10095
   ScaleWidth      =   14940
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport cr 
      Left            =   12720
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.ComboBox cboWarehouse 
      Height          =   315
      Left            =   2025
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   2550
      Width           =   3015
   End
   Begin VB.TextBox txtProductName 
      Height          =   285
      Left            =   3360
      TabIndex        =   4
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox txtMRP 
      Height          =   285
      Left            =   2040
      MaxLength       =   5
      TabIndex        =   5
      Top             =   2220
      Width           =   1275
   End
   Begin VB.ComboBox cboTypeProdcut 
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Top             =   960
      Width           =   4275
   End
   Begin VB.Frame buttonFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   5280
      TabIndex        =   13
      Top             =   1860
      Width           =   8895
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   6240
         Picture         =   "frmProductMaster.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   120
         Width           =   1230
      End
      Begin VB.CommandButton cmdAdd_1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   45
         Picture         =   "frmProductMaster.frx":0BF0
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   135
         Width           =   1230
      End
      Begin VB.CommandButton cmdSave_2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   1275
         Picture         =   "frmProductMaster.frx":17D4
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   135
         Width           =   1230
      End
      Begin VB.CommandButton cmdDelete_3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   2520
         Picture         =   "frmProductMaster.frx":23B8
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   135
         Width           =   1230
      End
      Begin VB.CommandButton cmdEdit_4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   3750
         Picture         =   "frmProductMaster.frx":2F9C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   135
         Width           =   1230
      End
      Begin VB.CommandButton cmdExit_12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   7500
         Picture         =   "frmProductMaster.frx":33A9
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   1230
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   4980
         Picture         =   "frmProductMaster.frx":3F8D
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   135
         Width           =   1230
      End
   End
   Begin VB.TextBox txtPcode 
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Top             =   600
      Width           =   1275
   End
   Begin VB.TextBox txtRulling 
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Top             =   1380
      Width           =   3015
   End
   Begin VB.TextBox txtPages 
      Height          =   285
      Left            =   2040
      MaxLength       =   5
      TabIndex        =   3
      Top             =   1800
      Width           =   1275
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   6975
      Left            =   240
      TabIndex        =   19
      Top             =   3000
      Width           =   11070
      _cx             =   19526
      _cy             =   12303
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   8388672
      BackColorFixed  =   9961471
      ForeColorFixed  =   0
      BackColorSel    =   15787206
      ForeColorSel    =   -2147483640
      BackColorBkg    =   15787206
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmProductMaster.frx":4B71
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
      ExplorerBar     =   7
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
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
   Begin VB.Label Label5 
      Caption         =   "Warehouse "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   225
      TabIndex        =   21
      Top             =   2625
      Width           =   1590
   End
   Begin VB.Label header 
      BackColor       =   &H8000000D&
      Caption         =   "     Product Master"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   18
      Top             =   60
      Width           =   10755
   End
   Begin VB.Label Label1 
      Caption         =   "MRP :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   240
      TabIndex        =   17
      Top             =   2220
      Width           =   1875
   End
   Begin VB.Label Label1 
      Caption         =   "Type of Product :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   16
      Top             =   960
      Width           =   1875
   End
   Begin VB.Label Label1 
      Caption         =   "Product Code :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   600
      Width           =   1875
   End
   Begin VB.Label Label1 
      Caption         =   "Ruling :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   1380
      Width           =   1875
   End
   Begin VB.Label Label1 
      Caption         =   "Pages :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   7
      Left            =   240
      TabIndex        =   10
      Top             =   1800
      Width           =   1875
   End
End
Attribute VB_Name = "frmProductMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit As Boolean

Private Sub cboTypeProdcut_Click()
fillGrid
End Sub

Private Sub cmdAdd_1_Click()
    Dim o As Object
    
    For Each o In Me
    
      If TypeOf o Is textbox Then
         o.Text = ""
      End If
    
    Next
    
   cmdEdit_4.Enabled = False
   cmdDelete_3.Enabled = False
   cmdSave_2.Enabled = True
   txtPcode.Enabled = True
   
   If txtPcode.Enabled = True Then
   
   txtPcode.SetFocus
   
   fillGrid
    
   End If
End Sub

Private Sub cmdDelete_3_Click()

If MsgBox("Want To Delete ?", vbQuestion + vbYesNo) = vbYes Then
   
   CON.BeginTrans
   CON.Execute "delete from copyMaster where " & stringyear & " and  BookNo ='" & txtPcode & "'"
   CON.CommitTrans
   fillGrid
   cboTypeProdcut.Text = ""
   txtPcode.Text = ""
   txtProductName = ""
   txtMRP = ""
   txtRulling = ""
   txtPcode.Enabled = True
   txtPcode.SetFocus
   
End If
cmdEdit_4.Enabled = False
cmdDelete_3.Enabled = False

End Sub

Private Sub cmdEdit_4_Click()
  Edit = True
  cmdEdit_4.Enabled = False
  cmdSave_2.Enabled = True
  cmdDelete_3.Enabled = True
  cboTypeProdcut.SetFocus
  txtPcode.Enabled = False
End Sub

Private Sub cmdExit_12_Click()
Unload Me
End Sub

Private Sub Cmdprint_Click()
    cr.Reset
    cr.Connect = constr
    cr.ReportFileName = App.Path & "\REPORTS\ItemList1.rpt"
    cr.Connect = "filedsn=chitradsn;uid=sa;pwd=sidc;"
    cr.WindowState = crptMaximized
    cr.WindowShowPrintSetupBtn = True
    cr.Action = 1

End Sub

Private Sub cmdSave_2_Click()



If cboTypeProdcut.Text = "" Then
   MsgBox "Plz. Select Product type...", vbCritical
   cboTypeProdcut.SetFocus
   Exit Sub
End If

If txtPcode.Text = "" Then
   MsgBox "Plz. Enter Product code...", vbCritical
   txtPcode.SetFocus
   Exit Sub
End If

'If txtProductName.Text = "" Then
'   MsgBox "Plz. Enter Product name...", vbCritical
'   txtProductName.SetFocus
'   Exit Sub
'End If

If txtRulling.Text = "" Then
   MsgBox "Plz. Specify rulling...", vbCritical
   txtRulling.SetFocus
   Exit Sub
End If

If txtMRP.Text = "" Then
   MsgBox "Plz. Enter MRP...", vbCritical
   txtMRP.SetFocus
   Exit Sub
End If



On Error GoTo save:


If Edit = False Then

    CON.BeginTrans
    
    CON.Execute "exec insertData_copyMaster '" & cboTypeProdcut & "','" & txtPcode & "','" & txtProductName & "','" & _
    txtRulling & "','" & txtPages & "','" & txtMRP & "','" & book & "','" & main.username & "','" & main.username & "','" & main.session & "'," & main.setupid & ""
    
    CON.CommitTrans
    
    MsgBox "Data Saved ...", vbInformation
Else

    CON.BeginTrans
    
    CON.Execute "exec UpdateData_copyMaster '" & cboTypeProdcut & "','" & txtPcode & "','" & txtProductName & "','" & _
    txtRulling & "','" & txtPages & "','" & txtMRP & "','" & book & "','" & main.username & "','" & main.username & "','" & main.session & "'," & main.setupid & ""
    
    CON.CommitTrans
    
    
    MsgBox "Data Modify ...", vbInformation
    
    
End If

fillGrid
cmdEdit_4.Enabled = False
Edit = False
    
Call cmdAdd_1_Click



Exit Sub
save:
CON.RollbackTrans
If Err.Number = "-2147217900" Then
   MsgBox " Duplicate ...", vbCritical
   txtPcode.SetFocus
End If

End Sub

Private Sub cmdSearch_Click()
 'popuplist10 "select TypeofProduct,NoofPages as Pages,Rulling,rate as MRP,Opening,ProductQuality from copyMaster", CON
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   SendKeys "{tab}"
End If
End Sub

Private Sub Form_Load()
  fillcombo cboTypeProdcut, "groupname", "GROUPS", CON
  header(0).Top = MainMenu.Top + 60
header(0).Left = MainMenu.Left
header(0).Width = MainMenu.Width
cmdDelete_3.Enabled = False
cmdEdit_4.Enabled = False
fillGrid

If RS.State = 1 Then RS.close
RS.Open "select * from Warehouse order by Warehouse", CON, adOpenKeyset
While RS.EOF = False
  cboWarehouse.AddItem RS(0)
  RS.MoveNext
Wend
 
cboWarehouse.ListIndex = 0


ButtonPermission cmdSave_2, cmdDelete_3, cmdEdit_4

End Sub

Private Sub Form_Unload(cancel As Integer)
  '''MainMenu.Toolbar1.Visible = True
End Sub

Private Sub Text4_Change()

End Sub
Sub fillGrid()
    
   vs.Cols = 6
    
   Dim f As New ADODB.Recordset
   If f.State = 1 Then f.close
   If Me.cboTypeProdcut.Text = "" Then
      f.Open "select BookNo,NoofPages as Pages,Rulling,rate as MRP,TypeofProduct,Opening,ProductQuality from copyMaster order by BookNo", CON
   Else
      f.Open "select BookNo,NoofPages as Pages,Rulling,rate as MRP,TypeofProduct,Opening,ProductQuality from copyMaster where " & stringyear & " and  TypeofProduct='" & cboTypeProdcut.Text & "'  order by BookNo", CON
   End If
   
   Set vs.DataSource = f
   
   vs.Cell(flexcpFontBold, 0, 0) = True
   vs.Cell(flexcpFontBold, 0, 1) = True
   vs.Cell(flexcpFontBold, 0, 2) = True
   vs.Cell(flexcpFontBold, 0, 3) = True
   vs.Cell(flexcpFontBold, 0, 4) = True
   vs.Cell(flexcpFontBold, 0, 5) = True
   'vs.Cell(flexcpFontBold, 0, 6) = True
   
   
   
   vs.FormatString = "BookNo|NoofPages|Rulling|MRP|TypeofProduct|Opening|"
   
   
   vs.ColWidth(0) = 1000
   vs.ColWidth(1) = 1000
   vs.ColWidth(2) = 2000
   vs.ColWidth(3) = 800
   vs.ColWidth(4) = 3500
   vs.ColWidth(5) = 1000
   vs.ColWidth(6) = 1000
   
   
End Sub
Private Sub txtPcode_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then

bookSearch

End If

End Sub
Sub bookSearch()
   
If RS.State = 1 Then RS.close
RS.Open "select * from copymaster where " & stringyear & " and  bookno='" & txtPcode & "'", CON
If RS.EOF = False Then
   
   'TypeofProduct,Opening,ProductQuality
   
   txtPcode = RS!bookNo & ""
   txtPages = RS!NoOfPages & ""
   txtRulling = RS!rulling & ""
   txtMRP = RS!rate & ""
   cboTypeProdcut.Text = RS!TypeofProduct & ""
   txtProductName = RS!ProductQuality & ""
   cmdEdit_4.Enabled = True
   cmdDelete_3.Enabled = True
   cmdSave_2.Enabled = False

End If

End Sub
Private Sub vs_Click()
  txtPcode = vs.TextMatrix(vs.RowSel, 0)
   'txtProductName = vs.TextMatrix(vs.RowSel, 1)
   txtPages = vs.TextMatrix(vs.RowSel, 1)
   txtRulling = vs.TextMatrix(vs.RowSel, 2)
   txtMRP = vs.TextMatrix(vs.RowSel, 3)
   cboTypeProdcut.Text = vs.TextMatrix(vs.RowSel, 4)
   txtProductName = vs.TextMatrix(vs.RowSel, 5)
   cmdEdit_4.Enabled = True
   cmdDelete_3.Enabled = True
   cmdSave_2.Enabled = False
End Sub
Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      If vs.Col = 5 Then
         SendKeys "{down}"
         CON.Execute "update copyMaster set opening=" & Val(vs.TextMatrix(vs.RowSel, 5)) & " where bookno='" & (vs.TextMatrix(vs.RowSel, 0)) & "' and " & stringyear
      End If
   End If
End Sub
