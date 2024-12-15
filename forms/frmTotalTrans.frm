VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTotalTrans 
   Caption         =   "Total Transaction"
   ClientHeight    =   8145
   ClientLeft      =   2025
   ClientTop       =   1785
   ClientWidth     =   11415
   Icon            =   "frmTotalTrans.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8145
   ScaleWidth      =   11415
   Begin VB.TextBox txtNo 
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
      Left            =   6660
      TabIndex        =   9
      Top             =   495
      Width           =   2880
   End
   Begin VB.CommandButton cmdref 
      BackColor       =   &H00FFFFC0&
      Caption         =   "&Show"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   9900
      Picture         =   "frmTotalTrans.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   90
      Width           =   1170
   End
   Begin VB.TextBox txtSearchText 
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
      Left            =   6660
      TabIndex        =   3
      Top             =   180
      Width           =   2880
   End
   Begin VB.ComboBox cboSearchType 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmTotalTrans.frx":0BF0
      Left            =   4140
      List            =   "frmTotalTrans.frx":0C00
      TabIndex        =   2
      Top             =   135
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker txtfDate 
      Height          =   330
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   582
      _Version        =   393216
      Format          =   79167489
      CurrentDate     =   39795
   End
   Begin MSComCtl2.DTPicker txttDate 
      Height          =   330
      Left            =   1800
      TabIndex        =   1
      Top             =   135
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   582
      _Version        =   393216
      Format          =   79167489
      CurrentDate     =   39795
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   6990
      Left            =   45
      TabIndex        =   8
      Top             =   990
      Width           =   10065
      _cx             =   17754
      _cy             =   12330
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
      BackColorFixed  =   16761024
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16761992
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
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
      Rows            =   200
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmTotalTrans.frx":0C23
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "No:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5670
      TabIndex        =   10
      Top             =   495
      Width           =   285
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Search Text:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5670
      TabIndex        =   7
      Top             =   180
      Width           =   990
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Search Type :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3195
      TabIndex        =   6
      Top             =   180
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "to"
      Height          =   240
      Left            =   1395
      TabIndex        =   5
      Top             =   180
      Width           =   375
   End
End
Attribute VB_Name = "frmTotalTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdref_Click()
   
Dim rs_ As New ADODB.Recordset
Dim search_ As String

search_ = ""



search_ = "(dates>=convert(smalldatetime,'" & txtfDate.value & "',103) and dates<=convert(smalldatetime,'" & txttDate.value & "',103))"



If cboSearchType.Text <> "" Then
   search_ = search_ & " and " & cboSearchType.Text & " like '%" & txtSearchText.Text & "%'"
End If

If txtNo.Text <> "" Then
   search_ = search_ & " and [No] like '%" & txtNo.Text & "%'"
End If





If rs_.State = 1 Then rs_.close
If search_ = "" Then
   rs_.Open "select UserName,No,vtype,desc_,dates,id FROM logtbl ORDER BY username,DATES desc,Id desc", con
Else
   rs_.Open "select UserName,No,vtype,desc_,dates,id FROM logtbl where " & search_ & " ORDER BY username,DATES desc,Id desc", con
End If

If rs_.EOF = False Then
   Set vs.DataSource = rs_
End If

vs.FormatString = "|UserName|No|Vtype|Description|Dates|Id"

vs.ColWidth(0) = 100
vs.ColWidth(1) = 1200
vs.ColWidth(2) = 1200
vs.ColWidth(3) = 2000
vs.ColWidth(4) = 2500
vs.ColWidth(5) = 1400






   
   
End Sub

Private Sub Form_Load()

Me.Width = 11160
Me.Height = 8715

txtfDate.value = Format(Date, "dd/MM/yyyy")
txttDate.value = Format(Date, "dd/MM/yyyy")


End Sub
