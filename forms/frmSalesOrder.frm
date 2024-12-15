VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmSalesOrder 
   Caption         =   "Product Order"
   ClientHeight    =   9405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12885
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9405
   ScaleWidth      =   12885
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7200
      TabIndex        =   47
      Top             =   3540
      Width           =   3810
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   7140
      TabIndex        =   42
      Top             =   900
      Width           =   4275
      Begin VB.OptionButton Option2 
         Caption         =   "Direct"
         Height          =   255
         Left            =   1440
         TabIndex        =   44
         Top             =   240
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Bank"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   1275
      End
   End
   Begin VB.ComboBox cboMode 
      Height          =   315
      ItemData        =   "frmSalesOrder.frx":0000
      Left            =   1680
      List            =   "frmSalesOrder.frx":0013
      TabIndex        =   40
      Top             =   2460
      Width           =   2595
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H0000C0C0&
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
      Height          =   375
      Left            =   3015
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   600
      Width           =   360
   End
   Begin VB.CommandButton cmdPrv 
      BackColor       =   &H0000C0C0&
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
      Height          =   375
      Left            =   1275
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   600
      Width           =   360
   End
   Begin VB.ComboBox cboItem 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Left            =   1260
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   18
      Top             =   4470
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Frame buttonFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   1290
      TabIndex        =   10
      Top             =   8025
      Width           =   8385
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Print"
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
         Left            =   5850
         Picture         =   "frmSalesOrder.frx":0045
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdAdd_1 
         BackColor       =   &H00FFFFC0&
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
         Picture         =   "frmSalesOrder.frx":0C29
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   135
         Width           =   1140
      End
      Begin VB.CommandButton cmdSave_2 
         BackColor       =   &H00FFFFC0&
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
         Left            =   1200
         Picture         =   "frmSalesOrder.frx":180D
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   135
         Width           =   1140
      End
      Begin VB.CommandButton cmdDelete_3 
         BackColor       =   &H00FFFFC0&
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
         Left            =   2340
         Picture         =   "frmSalesOrder.frx":23F1
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   1185
      End
      Begin VB.CommandButton cmdEdit_4 
         BackColor       =   &H00FFFFC0&
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
         Left            =   3525
         Picture         =   "frmSalesOrder.frx":2FD5
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   135
         Width           =   1140
      End
      Begin VB.CommandButton cmdExit_12 
         BackColor       =   &H00FFFFC0&
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
         Left            =   7065
         Picture         =   "frmSalesOrder.frx":33E2
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   135
         Width           =   1230
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00FFFFC0&
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
         Left            =   4680
         Picture         =   "frmSalesOrder.frx":3FC6
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   135
         Width           =   1140
      End
   End
   Begin VB.TextBox txtRemarks 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7185
      TabIndex        =   9
      Top             =   2910
      Width           =   3810
   End
   Begin VB.TextBox txtFrieght 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7740
      TabIndex        =   8
      Top             =   2220
      Width           =   3270
   End
   Begin VB.TextBox txtWeight 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7740
      TabIndex        =   7
      Top             =   1875
      Width           =   3270
   End
   Begin VB.TextBox txtBundles 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7740
      TabIndex        =   6
      Top             =   1560
      Width           =   3270
   End
   Begin VB.TextBox txtTansp 
      Appearance      =   0  'Flat
      Height          =   600
      Left            =   1680
      TabIndex        =   5
      Top             =   1470
      Width           =   3435
   End
   Begin VB.TextBox txtAgent 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1680
      TabIndex        =   4
      Top             =   2115
      Width           =   2580
   End
   Begin VB.TextBox txtSlipNo 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   1680
      TabIndex        =   3
      Top             =   630
      Width           =   1365
   End
   Begin VB.TextBox txtRR 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   2
      Top             =   1125
      Width           =   3405
   End
   Begin VB.TextBox txtAdd 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1680
      TabIndex        =   1
      Top             =   2835
      Width           =   3435
   End
   Begin VB.TextBox txtAgentCode 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   4260
      TabIndex        =   0
      Top             =   2115
      Width           =   825
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   3510
      Left            =   300
      TabIndex        =   21
      Top             =   3870
      Width           =   10725
      _cx             =   18918
      _cy             =   6191
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   16744576
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   325
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
   Begin MSMask.MaskEdBox txtSlipDate 
      Height          =   285
      Left            =   3960
      TabIndex        =   22
      Top             =   675
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   285
      Left            =   7185
      TabIndex        =   45
      Top             =   3240
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks :"
      Height          =   240
      Index           =   1
      Left            =   6420
      TabIndex        =   48
      Top             =   3540
      Width           =   765
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Dis. Date :"
      Height          =   240
      Left            =   6420
      TabIndex        =   46
      Top             =   3240
      Width           =   780
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000013&
      Caption         =   "Other Details :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   7200
      TabIndex        =   41
      Top             =   2640
      Width           =   2445
   End
   Begin VB.Label header 
      BackColor       =   &H8000000D&
      Caption         =   "    Sales Order"
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
      Left            =   0
      TabIndex        =   39
      Top             =   0
      Width           =   10755
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   6825
      TabIndex        =   38
      Top             =   6405
      Width           =   1410
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Qty :"
      Height          =   240
      Left            =   5925
      TabIndex        =   37
      Top             =   6405
      Width           =   825
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Transport :"
      Height          =   240
      Index           =   0
      Left            =   6360
      TabIndex        =   36
      Top             =   2970
      Width           =   765
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000013&
      Caption         =   "Document Through :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   0
      Left            =   7140
      TabIndex        =   35
      Top             =   600
      Width           =   2445
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Bank :"
      Height          =   240
      Left            =   7200
      TabIndex        =   34
      Top             =   1590
      Width           =   540
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Date :"
      Height          =   240
      Left            =   3555
      TabIndex        =   33
      Top             =   675
      Width           =   780
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Billing Address :"
      Height          =   240
      Left            =   345
      TabIndex        =   32
      Top             =   1470
      Width           =   1050
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Agent :"
      Height          =   240
      Index           =   1
      Left            =   360
      TabIndex        =   31
      Top             =   2100
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Order No. :"
      Height          =   195
      Left            =   345
      TabIndex        =   30
      Top             =   675
      Width           =   990
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Party Name :"
      Height          =   240
      Left            =   345
      TabIndex        =   29
      Top             =   1140
      Width           =   1095
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Mode :"
      Height          =   240
      Left            =   345
      TabIndex        =   28
      Top             =   2475
      Width           =   1095
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Recd By :"
      Height          =   240
      Left            =   345
      TabIndex        =   27
      Top             =   2820
      Width           =   1095
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   7350
      TabIndex        =   26
      Top             =   7530
      Width           =   1050
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Qty :"
      Height          =   240
      Left            =   6630
      TabIndex        =   25
      Top             =   7530
      Width           =   825
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount :"
      Height          =   240
      Left            =   9060
      TabIndex        =   24
      Top             =   7530
      Width           =   1095
   End
   Begin VB.Label lblAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   10020
      TabIndex        =   23
      Top             =   7530
      Width           =   1140
   End
End
Attribute VB_Name = "frmSalesOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim kk1 As String
''Dim rss As New ADODB.Recordset
'Dim b1 As Boolean
'Dim editValue As Boolean
'
'
'Private Sub cboHA_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then SendKeys "{tab}"
'End Sub
'
'Private Sub cboItem_DblClick()
'
'   If kk1 = "department" Then
'      txtDep = cboItem.Text
'      cboItem.Visible = False
'      Command5.SetFocus
'      txtCollegeId = ""
'      txtCollege = ""
'      txtAdd1 = ""
'      txtAdd2 = ""
'      txtTeacher = ""
'      txtTeacherId = ""
'   End If
'
'End Sub
'
'Private Sub cboItem_KeyPress(KeyAscii As Integer)
'   If KeyAscii = 13 Then
'
'   If kk1 = "department" Then
'      txtDep = cboItem.Text
'      cboItem.Visible = False
'      Command5.SetFocus
'      txtCollegeId = ""
'      txtCollege = ""
'      txtAdd1 = ""
'      txtAdd2 = ""
'      txtTeacher = ""
'      txtTeacherId = ""
'   End If
'
'
'  End If
'End Sub
'
'Private Sub cmdAdd_1_Click()
'    cmdSave_2.Enabled = True
'    txtSlipDate.SetFocus
'    cmdEdit_4.Enabled = False
'    cmdDelete_3.Enabled = False
'    editValue = False
'
'    Dim o As Object
'    For Each o In Me
'
'    'txtAgent
'    'txtAgentCode
'
'    If (UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("txtAgent")) Or UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("txtAgentCode"))) Then
'       If TypeOf o Is textbox Then
'           o.Text = ""
'       End If
'    End If
'
'    Next
'    lblAmt = ""
'    lblQty = ""
'
'    maxVoucher
'
'
'    editValue = False
'
'    vs.Clear
'    formatVSGrid
'
'
'End Sub
'
'
'
'Private Sub cmdDelete_3_Click()
'
'If txtSlipNo <> "" Then
'
'If MsgBox("Want To Delete ?", vbQuestion + vbYesNo) = vbYes Then
'
'CON.BeginTrans
'If rs.State = 1 Then rs.Close
'rs.Open "select * from  [SpcChallan] where [VoucherNo]='" & txtSlipNo & "'", CON
'If rs.EOF = False Then
'   CON.Execute "delete from  [SpcChallan] where [VoucherNo]='" & txtSlipNo & "'"
'   CON.Execute "delete from  [SpcChallanItem] where [VoucherNo]='" & txtSlipNo & "'"
'End If
'CON.CommitTrans
'
'rs1.Requery
'
'End If
'
'Call cmdAdd_1_Click
'
'End If
'
'
'End Sub
'
'Private Sub cmdEdit_4_Click()
'cmdEdit_4.Enabled = False
'cmdDelete_3.Enabled = True
'cmdSave_2.Enabled = True
'editValue = True
'End Sub
'
'Private Sub cmdExit_12_Click()
'Unload Me
'End Sub
'
'Private Sub cmdNext_Click()
'
'On Error GoTo aaa1:
'
''--------------------------------------
''formatVSGrid
'rs1.MoveFirst
'rs1.Find "VoucherNo='" & Trim(txtSlipNo.Text) & "'"
'rs1.MoveNext
'txtSlipNo.Text = rs1(0)
'SearchData
''--------------------------------------
'
'Exit Sub
'
'aaa1:
'
'MsgBox "Record not found ...", vbCritical
'
'
'End Sub
'
'Private Sub cmdNext_LostFocus()
'b1 = False
'End Sub
'
'Private Sub cmdPrv_Click()
'
'On Error GoTo aaa1:
'
''--------------------------------------
'
'
'rs1.MoveFirst
'rs1.Find "VoucherNo='" & Trim(txtSlipNo.Text) & "'"
'rs1.MovePrevious
'txtSlipNo.Text = rs1(0)
'SearchData
'
''--------------------------------------
'Exit Sub
'aaa1:
'
'MsgBox "Record not found ...", vbCritical
'
'End Sub
'Private Sub cmdSave_2_Click()
'
'
'
'If Val(lblAmt) = 0 Then
'   MsgBox "Enter Product In The Grid ...", vbInformation
'   vs.SetFocus
'End If
'
'Dim K As Integer
'
'
'
'
'CON.BeginTrans
'
'If editValue = True Then
'
'If rs.State = 1 Then rs.Close
'rs.Open "select * from  [SpcChallan] where [VoucherNo]='" & txtSlipNo & "'", CON
'If rs.EOF = False Then
'   CON.Execute "delete from  [SpcChallan] where [VoucherNo]='" & txtSlipNo & "'"
'   CON.Execute "delete from  [SpcChallanItem] where [VoucherNo]='" & txtSlipNo & "'"
'   editValue = False
'End If
'
'End If
'
'
'
'CON.Execute "INSERT INTO  [SpcChallan]" & _
'           "([VoucherNo]" & _
'           ",[Dated]" & _
'           ",[AgentID]" & _
'           ",[GRNo]" & _
'           ",[GRDate]" & _
'           ",[Transport]" & _
'           ",[Bundles]" & _
'           ",[Weight]" & _
'           ",[Freight]" & _
'           ",[StationID]" & _
'           ",[BookSellerId]" & _
'           ",[Add]" & _
'           ",[Remark]" & _
'           ",[Amount],[substation])" & _
'     "Values" & _
'           "('" & Trim(txtSlipNo) & "'" & _
'           ",'" & txtSlipDate.Text & "'" & _
'           ",'" & txtAgentCode.Text & "'" & _
'           ",'" & txtRR.Text & "'" & _
'           ",'" & txtRRDate.Text & "','" & Trim(txtTansp.Text) & "','" & Trim(txtBundles.Text) & "'" & _
'           ",'" & Trim(txtWeight.Text) & "','" & Trim(txtFrieght) & "','" & Trim(txtCityId) & "'" & _
'           ",'" & Trim(txtGRID.Text) & "','" & Trim(txtAdd.Text) & "','" & Trim(txtRemarks.Text) & "'," & Val(lblAmt) & ",'" & txtSubStation.Text & "')"
'
'
'
'            For I = 1 To vs.Rows - 1
'            If vs.TextMatrix(I, 1) <> "" Then
'            CON.Execute "INSERT INTO  [SpcChallanItem]" & _
'                       "([VoucherNo]" & _
'                       ",[Dated]" & _
'                       ",[Serial]" & _
'                       ",[ProductID]" & _
'                       ",[Qty]" & _
'                       ",[Rate]" & _
'                       ")" & _
'            " Values" & _
'                       "('" & txtSlipNo & "'," & _
'                       "'" & txtSlipDate & "'," & _
'                       "" & vs.TextMatrix(I, 0) & "," & _
'                       "'" & vs.TextMatrix(I, 7) & "'," & _
'                       "" & vs.TextMatrix(I, 4) & "," & _
'                       "'" & vs.TextMatrix(I, 5) & "'" & _
'                       ")"
'            End If
'            Next
'
'
'
'cmdSave_2.Enabled = False
'rs1.Requery
''MsgBox "Date Saved ....", vbInformation
'
'CON.CommitTrans
'
'cmdAdd_1_Click
'
'
'
'End Sub
'
'
'
'Private Sub cmdSearch_Click()
'   tblNo = 11
'   frmSearchItem.Show
'End Sub
'
'Private Sub cmdSearch_GotFocus()
'   If PopUpValue1 <> "" Then
'
'      txtSlipNo = PopUpValue1
'      SearchData
'
'      txtSlipDate.SetFocus
'
'      PopUpValue1 = ""
'      PopUpValue2 = ""
'      PopUpValue3 = ""
'
'   End If
'
'End Sub
'
'''Sub searchTeacher()
'''
'''
'''If rs1.State = 1 Then rs1.Close
'''rs1.Open "select QTC_TeacherID,[QTC_CCollegeID],[QTC_College],[QTC_TAdd1],[QTC_TAdd2],[QTC_TDistrict]," & _
'''"[QTC_TeachState],[QTC_Department],[QTC_CAdd1],[QTC_CAdd2],[QTC_District],[QTC_State],[QTC_TeacherName] from QryTeacherCollege " & _
'''" where QTC_TeacherID ='" & txtTeacherId & "'", CON, adOpenDynamic, adLockReadOnly
'''
'''If rs1.EOF = False Then
'''
'''   txtCollegeId = rs1![QTC_CCollegeID]
'''   txtCollege = rs1![QTC_College] & ""
'''   txtTeacher = rs1![QTC_TeacherName] & ""
'''
'''   txtRadd = rs1![QTC_TAdd1] & "  " & rs1![QTC_TAdd2] & vbCrLf & rs1![QTC_TDistrict] & ", " & rs1![QTC_TeachState]
'''   txtDep = rs1![QTC_Department] & ""
'''
'''   txtAdd1 = rs1![QTC_CAdd1] & " " & rs1![QTC_CAdd2]
'''   txtAdd2 = rs1![QTC_District] & ", " & rs1![QTC_State]
'''
'''End If
'''
'''End Sub
'Sub formatVSGrid()
'
'    vs.Clear
'
'    vs.Rows = 2
'    vs.Cols = 8
'
'    vs.FormatString = "SNo.|Item Code|Item Name|Quantity|M.R.P.|Net Rate|Net Amount"
'
'    For K = 0 To 6
'        vs.Cell(flexcpFontBold, 0, K) = True
'    Next
'
'    For K = 0 To 6
'        vs.Cell(flexcpForeColor, 0, K) = vbWhite
'    Next
'
'
'    vs.ColWidth(0) = 800
'    vs.ColWidth(1) = 1200
'    vs.ColWidth(2) = 3000
'    vs.ColWidth(3) = 1050
'    vs.ColWidth(4) = 1200
'    vs.ColWidth(5) = 1200
'    vs.ColWidth(6) = 1500
'    vs.ColWidth(7) = 0
'
'
'End Sub
'
'Private Sub Form_Load()
'
''DoEvents
''DoEvents
''DoEvents
'
''Screen.MousePointer = vbHourglass
'
'
''formDisplaySetting Me
''frmBack Me
'
'formatVSGrid
'
'header(0).TOP = MainMenu.TOP + 60
'header(0).Left = MainMenu.Left
'header(0).Width = MainMenu.Width
'
'
'
''txtSlipDate.Text = Format(Date, "dd/MM/yyyy")
''txtRRDate.Text = Format(Date, "dd/MM/yyyy")
'
''maxVoucher
''cmdEdit_4.Enabled = False
''cmdDelete_3.Enabled = False
'
''Me.TOP = Me.TOP + 50
'
'
''If rs1.State = 1 Then rs1.Close
''rs1.Open "SELECT [VoucherNo] From  [challanMain]", CON, adOpenDynamic, adLockReadOnly
'
''Screen.MousePointer = vbDefault
'
'
'
'
'
'End Sub
'Sub maxVoucher()
'   If rs.State = 1 Then rs.Close
'   rs.Open "Select max([VoucherNo]) from SpcChallan", CON
'   If IsNull(rs(0)) Then
'        txtSlipNo = "0000000" & 1 & "/" & Mid(Year(Date), 3)
'    Else
'        txtSlipNo = Format(Val(Mid(rs(0), 2)) + 1, "00000000")
'        txtSlipNo = txtSlipNo & "/" & Mid(Year(Date), 3)
'    End If
'
'End Sub
'
'Private Sub txtAdNo_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then SendKeys "{tab}"
'End Sub
'Private Sub txtDep_GotFocus()
'
'    kk1 = "department"
'    cboItem.Clear
'    If rs.State = 1 Then rs.Close
'    rs.Open "select Department FROM  [Department] order by Department", CON
'    While rs.EOF = False
'          cboItem.AddItem rs(0)
'          rs.MoveNext
'    Wend
'    cboItem.Visible = True
'    cboItem.ZOrder
'    cboItem.Width = txtDep.Width
'    cboItem.Left = txtDep.Left
'    cboItem.TOP = txtDep.TOP
'    cboItem.SetFocus
'
'End Sub
'
'
'Private Sub txtAgent_GotFocus()
'
'   txtAgent.SelLength = 30
'
'   If PopUpValue1 <> "" Then
'
'
'      txtAgentCode = PopUpValue1
'      txtAgent = PopUpValue2
'      txtAdd = PopUpValue3
'
'      txtRR.SetFocus
'
'      PopUpValue1 = ""
'      PopUpValue2 = ""
'      PopUpValue3 = ""
'
'   End If
'
'End Sub
'
'Private Sub txtAgent_KeyDown(KeyCode As Integer, Shift As Integer)
'
'  txtAgent = ""
'  txtAgentCode = ""
'  txtAdd = ""
'
'
'  tblNo = 12
'  frmSearchItem.Show
'
'End Sub
'
'Private Sub txtAgent_KeyPress(KeyAscii As Integer)
'  If KeyAscii = 13 Then SendKeys "{tab}"
'End Sub
'Private Sub txtDisDate_GotFocus()
'txtDisDate.SelLength = 25
'End Sub
'
'Private Sub txtDisDate_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then SendKeys "{tab}"
'End Sub
'
'Private Sub txtPAdd_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then SendKeys "{tab}"
'End Sub
'Private Sub txtRadd_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then SendKeys "{tab}"
'End Sub
'Private Sub txtBundles_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then SendKeys "{tab}"
'End Sub
'Private Sub txtFrieght_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then SendKeys "{tab}"
'End Sub
'Private Sub txtGRSent_GotFocus()
'
'txtGRSent.SelLength = 30
'
'If PopUpValue1 <> "" Then
'
'   txtGRSent = PopUpValue2
'   txtGRID = PopUpValue1
'   txtAdd = PopUpValue3
'
'   txtStation.SetFocus
'
'   PopUpValue1 = ""
'   PopUpValue2 = ""
'
'End If
'
'End Sub
'
'Private Sub txtGRSent_KeyDown(KeyCode As Integer, Shift As Integer)
'
'  txtGRSent = ""
'  txtGRID = ""
'
'  tblNo = 14
'  frmSearchItem.Show
'End Sub
'
'Private Sub txtGRSent_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then SendKeys "{tab}"
'End Sub
'
'Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then SendKeys "{tab}"
'End Sub
'Private Sub txtRR_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then SendKeys "{tab}"
'End Sub
'
'Private Sub txtRRDate_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then SendKeys "{tab}"
'End Sub
'
'Private Sub txtSlipDate_GotFocus()
'txtSlipDate.SelLength = 25
'End Sub
'Private Sub txtSlipDate_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then SendKeys "{tab}"
'End Sub
'
'Private Sub txtSlipNo_KeyPress(KeyAscii As Integer)
'   If KeyAscii = 13 Then
'      vs.Clear
'      formatVSGrid
'      SearchData
'      SendKeys "{tab}"
'   End If
'End Sub
'Sub SearchData()
'
'
' If rs.State = 1 Then rs.Close
' rs.Open "SELECT [VoucherNo]" & _
'      ",[Dated]" & _
'      ",[AgentID]" & _
'      ",[GRNo]" & _
'      ",[GRDate]" & _
'      ",[Bundles]" & _
'      ",[Transport]" & _
'      ",[Weight]" & _
'      ",[Freight]" & _
'      ",[StationID]" & _
'      ",[BookSellerId]" & _
'      ",[Add]" & _
'      ",[Remark]" & _
'      ",[BookSeler]" & _
'      ",[BookSellerId]" & _
'      ",[Amount]" & _
'      ",[Rep],city,substation " & _
'  "From  [challanMain] where VoucherNo='" & txtSlipNo & "'", CON, adOpenForwardOnly, adLockReadOnly
'  If rs.EOF = False Then
'
'        formatVSGrid
'
'        cmdEdit_4.Enabled = True
'        cmdDelete_3.Enabled = True
'        cmdSave_2.Enabled = False
'
'        txtSlipNo = rs![VoucherNo]
'        txtSlipDate.Text = rs!dated
'        txtAgentCode.Text = rs![AgentID]
'        txtRR.Text = rs![GRNo]
'        txtRRDate.Text = rs![GRDate]
'        txtTansp.Text = rs!Transport
'        txtBundles.Text = rs!bundles
'        txtWeight.Text = rs!weight
'        txtFrieght.Text = rs!freight
'        txtCityId.Text = rs!StationID
'        txtStation.Text = rs!CITY
'
'        txtGRSent.Text = rs![BookSeler]
'        txtAdd.Text = rs!Add
'        txtRemarks.Text = rs!Remark
'        lblAmt = rs!amount
'        txtAgent.Text = rs!rep
'        txtGRID = rs!BookSellerId
'        txtSubStation.Text = rs!substation
'
'
'
'
'      I = 1
'      Dim sum, amt As Long
'      sum = 0
'      amt = 0
'
'      vs.Rows = 2
'
'      If rs.State = 1 Then rs.Close
'      rs.Open "" & _
'      "SELECT [ProductID]" & _
'          ",[VoucherNo]" & _
'          ",[Serial]" & _
'          ",[Qty]" & _
'          ",[Product]" & _
'          ",[NikeName]" & _
'          ",[Type],Rate" & _
'      " From  [SpcChallanQryDetails] where VoucherNo='" & txtSlipNo & "' order by Serial", CON, adOpenForwardOnly, adLockReadOnly
'      While rs.EOF = False
'          vs.Rows = vs.Rows + 1
'          vs.TextMatrix(I, 0) = rs!serial
'          vs.TextMatrix(I, 1) = rs!NikeName
'          vs.TextMatrix(I, 2) = rs!Type
'          vs.TextMatrix(I, 3) = rs!Product
'          vs.TextMatrix(I, 4) = rs!qty
'          vs.TextMatrix(I, 5) = rs!rate
'          vs.TextMatrix(I, 6) = (rs!qty * rs!rate)
'          vs.TextMatrix(I, 7) = rs!productId
'          I = I + 1
'          sum = sum + rs!qty
'          amt = amt + (rs!qty * rs!rate)
'
'          rs.MoveNext
'      Wend
'
'      lblQty.Caption = sum & " "
'
'      lblAmt.Caption = amt & " "
'
'
'
' End If
'
'End Sub
'Private Sub txtStation_Change()
'  popupvalue4 = txtStation.Text
'End Sub
'
'Private Sub txtStation_GotFocus()
'HIT
'If PopUpValue1 <> "" Then
'
'    txtStation = PopUpValue2
'    txtCityId = PopUpValue1
'
'    txtSubStation.SetFocus
'End If
'
'PopUpValue1 = ""
'PopUpValue2 = ""
'popupvalue4 = ""
'
'End Sub
'
'Private Sub txtStation_KeyPress(KeyAscii As Integer)
'   If KeyAscii = 13 Then SendKeys "{tab}"
'End Sub
'
'Private Sub txtStation_KeyUp(KeyCode As Integer, Shift As Integer)
'If KeyCode = 27 Then Exit Sub
'If KeyCode = 13 Then Exit Sub
'tblNo = 6
'frmSearchItem.Show
'End Sub
'Private Sub txtSubStation_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then SendKeys "{tab}"
'End Sub
'
'Private Sub txtTansp_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then SendKeys "{tab}"
'End Sub
'
'Private Sub txtWeight_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then SendKeys "{tab}"
'End Sub
'
'Private Sub vs_GotFocus()
'
'  If PopUpValue1 = "" Then Exit Sub
'
'  If vs.col = 1 Then
'
'  vs.TextMatrix(vs.RowSel, 3) = PopUpValue1
'  vs.TextMatrix(vs.RowSel, 6) = PopUpValue2
'  vs.TextMatrix(vs.RowSel, 1) = PopUpValue3
'  vs.TextMatrix(vs.RowSel, 2) = popupvalue4
'
'  SendKeys "{right}"
'  SendKeys "{right}"
'  SendKeys "{right}"
'  'SendKeys "{right}"
'
'
'  ElseIf vs.col = 5 Then
'
''  vs.TextMatrix(vs.RowSel, 5) = PopUpValue1
''  vs.TextMatrix(vs.RowSel, 7) = PopUpValue2
'
'  ElseIf vs.col = 6 Then
'
'
'
'
'  End If
'
'
'   PopUpValue1 = ""
'   PopUpValue2 = ""
'   PopUpValue3 = ""
'   popupvalue4 = ""
'
'End Sub
'
'Private Sub VS_KeyDown(KeyCode As Integer, Shift As Integer)
'''If KeyCode = 113 Then
'''
'''If vs.col = 1 Then
'''
'''
'''popuplist2 "select Product,ProductId,[NikeName],[Type] from Product", CON
'''
'''ElseIf vs.col = 5 Then
'''
'''
'''
'''End If
'''
'''Else
'''
'''  vs.Editable = flexEDKbdMouse
'''End If
'''
'
'
'End Sub
'Sub Total()
'    lblAmt = 0
'    lblQty = 0
'
'    For K = 1 To vs.Rows - 1
'      lblQty = (Val(lblQty) + Val(vs.TextMatrix(K, 4)))
'      lblAmt = (Val(lblAmt) + Val(vs.TextMatrix(K, 6)))
'    Next
'
'    lblAmt = lblAmt & " "
'    lblQty = lblQty & " "
'
'End Sub
'Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)
' If KeyCode = 13 Then
'
'    If vs.col = 1 Then
'
'       If rs.State = 1 Then rs.Close
'       rs.Open "select Product,ProductId,[NikeName],[Type],RateTag from Product " & _
'       " where NikeName='" & vs.TextMatrix(vs.RowSel, 1) & "'", CON
'       If rs.EOF = False Then
'            vs.TextMatrix(vs.RowSel, 3) = rs(0)
'            vs.TextMatrix(vs.RowSel, 7) = rs(1)
'            vs.TextMatrix(vs.RowSel, 5) = rs!RateTag
'            vs.TextMatrix(vs.RowSel, 1) = rs(2)
'            vs.TextMatrix(vs.RowSel, 2) = rs(3)
'            SendKeys "{right}"
'            SendKeys "{right}"
'            SendKeys "{right}"
'       End If
'
'
'    End If
'
'    If vs.col = 4 Then
'     If Len(vs.TextMatrix(vs.RowSel, 4)) > 0 Then
'        vs.Editable = flexEDNone
'        vs.TextMatrix(vs.RowSel, 6) = (Val(vs.TextMatrix(vs.RowSel, 5)) * Val(vs.TextMatrix(vs.RowSel, 4)))
'
'        SendKeys "{right}"
'     End If
'
'
'     If Val(vs.TextMatrix(vs.RowSel, 4)) > 0 Then
'
'        vs.TextMatrix(vs.RowSel, 0) = vs.row
'        SendKeys "{home}"
'        SendKeys "{down}"
'
'        If CheckRaws(vs) = True Then vs.Rows = vs.Rows + 1
'
'        Total
'        vs.Editable = flexEDNone
'
'     End If
'
'    End If
' End If
'End Sub
'Private Sub vs_LeaveCell()
'
' vs.Select vs.RowSel, vs.col, vs.RowSel, vs.col
' vs.CellBorder &H8000000F, 2, 2, 1, 1, 1, 1
'
'
' If vs.col = 4 Then
'    vs.Editable = flexEDNone
' End If
'
'End Sub
'
'Private Sub vs_SelChange()
'
'
'
'     'cmdAdd.Visible = False
'      vs.Select vs.RowSel, vs.col, vs.RowSel, vs.col
'      vs.CellBorder vbGreen, 2, 2, 2, 2, 1, 1
'
'
'     If vs.col = 4 Then
'        vs.Editable = flexEDKbdMouse
'
'     Else
'        vs.Editable = flexEDNone
'     End If
'
'End Sub
'
'
'
'
'
'Private Sub Form_Unload(Cancel As Integer)
' 'MainMenu.Toolbar1.Visible = True
'End Sub
'
