VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmpaper 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Paper Issue/Transfer Entry "
   ClientHeight    =   9432
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   13620
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9432
   ScaleWidth      =   13620
   Visible         =   0   'False
   Begin Crystal.CrystalReport cr 
      Left            =   9420
      Top             =   7950
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.ComboBox cboToGodown 
      Height          =   315
      Left            =   7215
      TabIndex        =   30
      Top             =   1620
      Width           =   4005
   End
   Begin VB.ComboBox cboFromGodown 
      Height          =   315
      Left            =   1665
      TabIndex        =   29
      Top             =   1620
      Width           =   3960
   End
   Begin VB.TextBox txtRem 
      Height          =   285
      Left            =   1665
      MaxLength       =   100
      TabIndex        =   26
      Top             =   2100
      Width           =   9330
   End
   Begin VB.TextBox txtChallanNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7230
      MaxLength       =   15
      TabIndex        =   3
      Top             =   1140
      Width           =   1170
   End
   Begin VB.TextBox txtTo_Gdid 
      Height          =   285
      Left            =   11220
      MaxLength       =   50
      TabIndex        =   5
      Top             =   1620
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.ComboBox type1 
      Height          =   315
      Left            =   -480
      TabIndex        =   21
      Text            =   "R"
      Top             =   225
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   840
      Left            =   360
      TabIndex        =   13
      Top             =   7650
      Width           =   8580
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   7320
         Picture         =   "pstmst.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   45
         Width           =   1215
      End
      Begin VB.CommandButton cancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   6105
         Picture         =   "pstmst.frx":0BE4
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   45
         Width           =   1215
      End
      Begin VB.CommandButton ok 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&OK"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   1245
         Picture         =   "pstmst.frx":17C8
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   45
         Width           =   1215
      End
      Begin VB.CommandButton Printcmd 
         Caption         =   "&Print"
         Height          =   465
         Left            =   12090
         TabIndex        =   18
         Top             =   180
         Width           =   915
      End
      Begin VB.CommandButton search 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   4890
         Picture         =   "pstmst.frx":23AC
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   45
         Width           =   1215
      End
      Begin VB.CommandButton delete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   3675
         Picture         =   "pstmst.frx":2F90
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   45
         Width           =   1215
      End
      Begin VB.CommandButton Edit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   2460
         Picture         =   "pstmst.frx":3B74
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   45
         Width           =   1215
      End
      Begin VB.CommandButton Add 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   30
         Picture         =   "pstmst.frx":3FB6
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   45
         Width           =   1215
      End
   End
   Begin VB.TextBox ptyname2 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   555
      Left            =   -165
      MaxLength       =   50
      TabIndex        =   1
      Text            =   "Press F2 for Firm"
      Top             =   135
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtSNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1665
      MaxLength       =   50
      TabIndex        =   0
      Top             =   1140
      Width           =   1290
   End
   Begin VB.TextBox txtFrom_Gdid 
      Height          =   285
      Left            =   5745
      MaxLength       =   50
      TabIndex        =   6
      Top             =   1620
      Visible         =   0   'False
      Width           =   360
   End
   Begin MSComCtl2.DTPicker txtSNDate 
      Height          =   315
      Left            =   4200
      TabIndex        =   2
      Top             =   1140
      Width           =   1410
      _ExtentX        =   2498
      _ExtentY        =   550
      _Version        =   393216
      Format          =   507052033
      CurrentDate     =   38372
   End
   Begin VB.TextBox txtfirmid 
      Enabled         =   0   'False
      Height          =   360
      Left            =   -75
      TabIndex        =   7
      Top             =   180
      Width           =   240
   End
   Begin MSComCtl2.DTPicker txtChallanDate 
      Height          =   315
      Left            =   9105
      TabIndex        =   4
      Top             =   1110
      Width           =   1410
      _ExtentX        =   2477
      _ExtentY        =   550
      _Version        =   393216
      Format          =   507052033
      CurrentDate     =   38372
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   4980
      Left            =   240
      TabIndex        =   31
      Top             =   2520
      Width           =   13095
      _cx             =   23098
      _cy             =   8784
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   16251308
      ForeColorFixed  =   4210752
      BackColorSel    =   14286267
      ForeColorSel    =   16744448
      BackColorBkg    =   16251308
      BackColorAlternate=   16777215
      GridColor       =   255
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   8
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   -1  'True
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
      Editable        =   1
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
      BackStyle       =   0  'Transparent
      Caption         =   "Press F2 For Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   240
      Left            =   1665
      TabIndex        =   32
      Top             =   900
      Width           =   2070
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0078CFE9&
      BorderWidth     =   4
      Height          =   930
      Left            =   300
      Top             =   7590
      Width           =   8685
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   2160
      Width           =   1290
   End
   Begin VB.Label chdate 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   8475
      TabIndex        =   25
      Top             =   1155
      Width           =   555
   End
   Begin VB.Label challan1 
      BackStyle       =   0  'Transparent
      Caption         =   "Manual Ch. No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   5670
      TabIndex        =   24
      Top             =   1140
      Width           =   1605
   End
   Begin VB.Label lblto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   240
      Left            =   6165
      TabIndex        =   23
      Top             =   1665
      Width           =   375
   End
   Begin VB.Label pscustomerid 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5685
      TabIndex        =   22
      Top             =   1155
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Press F4 For Delete Record...."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   240
      Left            =   420
      TabIndex        =   12
      Top             =   8670
      Width           =   2685
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "option"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   405
      Left            =   270
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label challan 
      BackStyle       =   0  'Transparent
      Caption         =   "Challan No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   255
      TabIndex        =   10
      Top             =   1140
      Width           =   1470
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Printer Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1635
      Width           =   1710
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   3555
      TabIndex        =   8
      Top             =   1170
      Width           =   555
   End
End
Attribute VB_Name = "frmpaper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rschallan As ADODB.Recordset
Dim rsparty As ADODB.Recordset
Dim rspaperst As ADODB.Recordset
Public mode As String
Public orow As Integer

Dim v1 As Double
Public nrow As Integer
Sub disab()
    Me.grid1.Enabled = False
End Sub
Sub grid_ini()

grid1.Clear
grid1.Cols = 11
grid1.rows = 100

grid1.TextMatrix(0, 0) = "SN"
grid1.TextMatrix(0, 1) = "Paper Description"

If type1.text = "D" Then
  
  grid1.TextMatrix(0, 3) = ""
  grid1.TextMatrix(0, 4) = ""
  grid1.TextMatrix(0, 5) = "Reams"
  grid1.TextMatrix(0, 6) = "Sheets"
  grid1.TextMatrix(0, 7) = "Weight"
  grid1.TextMatrix(0, 9) = "No Of Reals"
  grid1.TextMatrix(0, 2) = ""
  grid1.TextMatrix(0, 10) = "Ream Wt."
  
  
  grid1.ColWidth(0) = 400
  grid1.ColWidth(2) = 0
  grid1.ColWidth(1) = 4500
  grid1.ColWidth(3) = 0
  grid1.ColWidth(4) = 0
  grid1.ColWidth(5) = 1200
  grid1.ColWidth(6) = 1200
  grid1.ColWidth(8) = 0
  grid1.ColWidth(9) = 1100
  grid1.ColWidth(10) = 1000
  
Else
  
 grid1.TextMatrix(0, 2) = "BillNo/Date"
 grid1.TextMatrix(0, 3) = ""
 grid1.TextMatrix(0, 4) = ""
 grid1.TextMatrix(0, 5) = "Reams"
 grid1.TextMatrix(0, 6) = "Sheets"
 grid1.TextMatrix(0, 7) = "Weight"
 grid1.TextMatrix(0, 8) = "ReceiptNo"
 grid1.TextMatrix(0, 9) = "No Of Reals"
 grid1.TextMatrix(0, 10) = "Ream Wt."
 
 grid1.ColWidth(0) = 400
 grid1.ColWidth(2) = 1200
 grid1.ColWidth(1) = 4500
 grid1.ColWidth(3) = 0
 grid1.ColWidth(4) = 0
 grid1.ColWidth(5) = 1200
 grid1.ColWidth(6) = 1200
 grid1.ColWidth(5) = 1200
 grid1.ColWidth(6) = 1200
 grid1.ColWidth(7) = 1200
 grid1.ColWidth(8) = 950
 grid1.ColWidth(9) = 1100
 grid1.ColWidth(10) = 1000
 

End If
 
 
 
For h1 = 1 To 7
  grid1.Cell(flexcpFontSize, 0, h1) = 11
Next
 
 
End Sub
Private Sub Add_Click()

delete.Enabled = False
'Me.search.Enabled = False
Me.cancel.Enabled = True


mode = ""
Set rschallan = New ADODB.Recordset
'sq = "select max(convert(int,sn_no)) from paperstatement where " & stringyear & " and PaperTrans_Deliv = '" + type1.Text + "'"
sq = "select max(convert(int,sn_no)) from paperstatement where " & stringyear & ""
rschallan.Open sq, con, adOpenStatic
If rschallan.RecordCount > 0 Then
   txtsno.text = IIf(IsNull(rschallan(0)), 1, rschallan(0) + 1)
Else
   txtsno.text = 1
End If
rschallan.close

Clearvalue
mode = "add"
grid1.Enabled = True
ok.Enabled = True
adddisab
txtsno.SetFocus

End Sub


Private Sub billid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   party_id.SetFocus
End If
End Sub

Private Sub cancel_Click()
'''Me.pscustomerid.Caption = ""
'''txtSNo.Text = ""
'''txtChallanNo.Text = ""
'''txtFrom_Gdid = ""
'''txtTo_Gdid = ""
'''Clearvalue
'''addenab
'''disab
'''
'''txtSNo.SetFocus
'''
'''Add.SetFocus

DSNNew

cr.Reset
'cr.ReportFileName = rptPath & "/PaperRecSlip.rpt"
cr.ReportFileName = rptPath & "/PaperRecChallan.rpt"

cr.ReplaceSelectionFormula "{paperstatement.sn_no}=" & txtsno.text & ""
cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
cr.WindowShowPrintSetupBtn = True
cr.WindowState = crptMaximized
cr.Action = 1




End Sub

Private Sub cboFromGodown_Click()
  txtFrom_Gdid = cboFromGodown.ItemData(cboFromGodown.ListIndex)
End Sub
Private Sub cboFromGodown_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
  If cboToGodown.Visible = True Then
     cboToGodown.SetFocus
  Else
     txtrem.SetFocus
  End If
  End If
End Sub
Private Sub cboToGodown_Click()
txtTo_Gdid = cboToGodown.ItemData(cboToGodown.ListIndex)
End Sub

Private Sub cboToGodown_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
     txtrem.SetFocus
  End If

End Sub

Private Sub chdatevalue_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error Resume Next
   If KeyCode = 13 Then
      cboFromGodown.SetFocus
   End If
End Sub

Private Sub chlno_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then
popuplist1 "Select distinct cint(Challan_No) as [Challan No],date1 as Chl_Date ,Customer_id from paperstatement where " & stringyear & " and firm_id = '" + Me.txtfirmid.text + "' and recdel = '" + type1.text + "' order by cint(challan_no) asc", con
End If


End Sub
Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub Delete_Click()
Dim rsbilmast As ADODB.Recordset
Dim rsbiltrans As ADODB.Recordset
Set rsbilmast = New ADODB.Recordset
Set rsbiltrans = New ADODB.Recordset
X = MsgBox("Are you sure you wish to delete the selected challan ", 4, "Confirmation")
If X = 6 Then
   sq = "delete  from paperstatement where " & stringyear & " and SN_no = '" + Me.txtsno.text + "' and PaperTrans_Deliv='" & type1.text & "' and FromGodown='" & cboFromGodown.text & "'"
   con.Execute sq
   'cancel_Click
End If

End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
'   billid.SetFocus
If txtChallan.Visible = True Then
   txtChallan.SetFocus
 Else
   Me.cboFromGodown.SetFocus
End If
End If
End Sub

Private Sub Edit_Click()

mode = "edit"
Me.ok.Enabled = True
delete.Enabled = True
Me.search.Enabled = True
Me.cancel.Enabled = True

End Sub

Private Sub Form_Activate()
txtfirmid.text = "chitra"
Call Add_Click
txtChallanDate.value = Date
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
sendkeys "{tab}"
End If
End Sub
Private Sub Form_Load()

 Me.Left = 100
 Me.top = 100
 
 Me.Width = 13500
 Me.Height = 9705


  

txtfirmid.Visible = False
grid_ini
Frame2.Enabled = False
addenab
type1_Click
txtfirmid.text = "chitra"
txtSNDate.value = Date
challan.Caption = "Serial No"

BackColorFrom Me
txtSNDate.value = Format(Date, "dd/MM/yyyy")
txtChallanDate.value = Format(Date, "dd/MM/yyyy")



Set RS = New ADODB.Recordset
'RS.Open "select Godwn,id from Godownmaster where " & stringyear & " and Binder_Printer='p' order by Godwn", con
RS.Open "select Godwn,id from Godownmaster where " & stringyear & " and len(Godwn)>3 order by Godwn", con
While RS.EOF = False
    
    cboFromGodown.AddItem RS(0)
    cboFromGodown.ItemData(cboFromGodown.NewIndex) = RS!id
    cboToGodown.AddItem RS(0)
    cboToGodown.ItemData(cboToGodown.NewIndex) = RS!id
    
    RS.MoveNext
Wend



'========================================================================
Dim rs2_ As New ADODB.Recordset
'========================================================================
s = ""
If rs2_.State = 1 Then rs2_.close
rs2_.Open "select * from PaperMakeMaster order by papermaker_name", con, adOpenStatic, adLockReadOnly

While rs2_.EOF = False
If s = "" Then

     s = rs2_!papermaker_name
     If rs2_!eco <> "" Then
       s = s & "-" & rs2_!eco
    End If
    
    If rs2_!SizeValue1 <> "" Then
       s = s & "-" & rs2_!SizeValue1 & "X" & rs2_!SizeValue2
    End If
    
    If rs2_!GSM <> "" Then
       s = s & "-" & rs2_!GSM
    End If
    s = s & "=>" & rs2_!papermaker_id
Else
    If rs2_!papermaker_name <> "" Then
       s = s & "|" & rs2_!papermaker_name
    End If
    If rs2_!eco <> "" Then
       s = s & "-" & rs2_!eco
    End If
    If (rs2_!SizeValue1 <> "" And rs2_!SizeValue2 <> "") Then
       s = s & "-" & rs2_!SizeValue1 & " CM X " & rs2_!SizeValue2 & "CM"
    ElseIf (rs2_!SizeValue1 <> "" And rs2_!SizeValue2 = "") Then
       s = s & "-" & rs2_!SizeValue1 & " CM "
    End If
    
    If rs2_!GSM <> "" Then
       s = s & "-" & rs2_!GSM & " GSM"
    End If
    
    s = s & "=>" & rs2_!papermaker_id
End If
rs2_.MoveNext
Wend


''
''st_ = ""
''
''If RS.State = 1 Then RS.close
'''RS.Open "select distinct  SizeValue1 + 'X' + SizeValue2 + ' : ' + papermaker_name + ' : ' + ptype  + ' : ' + gsm + ' GSM' + '-' + papermaker_id  from PaperMakeMaster", con
''
''While RS.EOF = False
''
''If st_ = "" Then
''st_ = RS(0)
''Else
''st_ = st_ & "|" & RS(0)
''End If
''
''RS.MoveNext
''Wend
''
grid1.ColComboList(1) = s




'======================================================================
'======================================================================


st_ = ""

If RS.State = 1 Then RS.close
RS.Open "select distinct size1 from SizeMaster where " & stringyear, con
While RS.EOF = False

If st_ = "" Then
st_ = RS(0)
Else
st_ = st_ & "|" & RS(0)
End If

RS.MoveNext
Wend

grid1.ColComboList(3) = st_


st_ = ""

If RS.State = 1 Then RS.close
RS.Open "select distinct GSM from GSMMaster where " & stringyear, con
While RS.EOF = False

If st_ = "" Then
st_ = RS(0)
Else
st_ = st_ & "|" & RS(0)
End If

RS.MoveNext
Wend

grid1.ColComboList(4) = st_



End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 115 Then
   If MsgBox("Are You Sure, Delete it ?", vbYesNo) = vbYes Then
      If grid1.Row >= 1 Then
           str_p = Trim(Mid(grid1.TextMatrix(grid1.RowSel, 1), InStr(grid1.TextMatrix(grid1.RowSel, 1), "=>") + 2))
           con.Execute "delete from paperstatement where SN_no = '" + Me.txtsno.text + "' and Pcode='" & str_p & "'"
           grid1.RemoveItem (grid1.RowSel)
           grid1.SetFocus
          End If
   End If
End If
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)


If KeyCode = 13 Then
   
   

Dim cm1, cm2, gsm_, divideBy

divideBy = 20000
   
   
   
   
   
If grid1.Col = 1 Then
   grid1.TextMatrix(grid1.RowSel, 0) = grid1.RowSel
   
   pcode = Trim(Mid(grid1.TextMatrix(grid1.RowSel, 1), InStr(grid1.TextMatrix(grid1.RowSel, 1), "=>") + 2))
   
   If rs1.State = 1 Then rs1.close
   rs1.Open "select gsm,SizeValue1,SizeValue2 from PaperMakeMaster where papermaker_id='" & pcode & "'", con
   If rs1.EOF = False Then
      GSM = Val(rs1!GSM)
      
      
      cm1 = Val(rs1!SizeValue1)
      cm2 = Val(rs1!SizeValue2)
      
      
      If (cm1 > 0 And cm2 > 0) Then
          grid1.TextMatrix(grid1.RowSel, 10) = Round(((cm1 * GSM) * cm2) / divideBy, 3)
      ElseIf (cm1 > 0 And cm2 = 0) Then
          grid1.TextMatrix(grid1.RowSel, 10) = Round((cm1 * GSM) / divideBy, 3)
      ElseIf (cm1 = 0 And cm2 = 0) Then
          grid1.TextMatrix(grid1.RowSel, 10) = 0
      End If
      
      
      grid1.TextMatrix(grid1.RowSel, 10) = Format(grid1.TextMatrix(grid1.RowSel, 10), ".000")
      
   End If
   
   If grid1.TextMatrix(grid1.RowSel, 1) <> "" Then
     sendkeys "{right}"
   End If
ElseIf grid1.Col = 2 Then

   sendkeys "{right}"
ElseIf grid1.Col = 3 Then
   
   If grid1.TextMatrix(grid1.RowSel, 3) <> "" Then
     sendkeys "{right}"
   End If
ElseIf grid1.Col = 4 Then
   If grid1.TextMatrix(grid1.RowSel, 4) <> "" Then
     sendkeys "{right}"
   End If
ElseIf grid1.Col = 5 Then
   sendkeys "{right}"
ElseIf grid1.Col = 6 Then
   sendkeys "{right}"
ElseIf grid1.Col = 7 Then

   Dim wt As Double
   Dim wtreams As Double
   Dim reams
   
   vv1 = 0
   wt = 0
   wtreams = 0
   
   wt = IIf(grid1.TextMatrix(grid1.RowSel, 7) = "", 0, grid1.TextMatrix(grid1.RowSel, 7))
   wtreams = IIf(grid1.TextMatrix(grid1.RowSel, 10) = "", 0, grid1.TextMatrix(grid1.RowSel, 10))
   
   reams = IIf(grid1.TextMatrix(grid1.RowSel, 5) = "", 0, grid1.TextMatrix(grid1.RowSel, 5))
   vv1 = IIf(grid1.TextMatrix(grid1.RowSel, 6) = "", 0, grid1.TextMatrix(grid1.RowSel, 6))
   
   If (reams > 0 And vv1 > 0) Then
      reams = reams & "." & vv1
      reams = Format(reams, "0.000")
   ElseIf (reams > 0 And vv1 = 0) Then
      reams = Format(reams, "0.000")
   ElseIf (reams = 0 And vv1 > 0) Then
      reams = Format(vv1, "0.000")
   End If
   
   
   If (wt > 0 And wtreams > 0) Then
      vv1 = calcReams_(wt, wtreams)
      MsgBox "Reams Should be " & vv1
      grid1.SetFocus
   End If


   sendkeys "{right}"
ElseIf grid1.Col = 8 Then
   sendkeys "{right}"
ElseIf grid1.Col = 9 Then
   sendkeys "{home}"
   sendkeys "{down}"
   sendkeys "{right}"
   grid1.TextMatrix(grid1.Row + 1, 0) = grid1.RowSel + 1
End If
   
End If

End Sub
Function calcReams_(ByRef wt As Double, ByRef wtReam As Double) As Double
 
 
 calcReams_ = Round(wt / wtReam, 3)
 
 
End Function

Private Sub Grid1_LostFocus()
'Label6.Visible = False
End Sub



Private Sub LblGodownName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
popuplist1 "Select customer_name as [Customer Name],Add1 as Address,customer_id as [Godown Id] from customerMaster where " & stringyear & " and customer_id<>'" & party_id.text & "'", con
End If
If KeyCode = 13 And party_id.text <> "" Then
grid1.Row = 1
grid1.Col = 1
grid1.SetFocus
End If
End Sub

Private Sub ok_Click()

Dim nrow As Integer
Dim st_str As String
Dim str_p As String

If mode = Trim("edit") Then
    sq = "delete from paperstatement where " & stringyear & " and sn_no = '" + txtsno.text + "' and PaperTrans_Deliv = '" + type1.text + "' and FromGodown='" & cboFromGodown.text & "'"
    con.Execute sq
End If


If type1.text = "D" Then
    If cboFromGodown.text = "" Then
       MsgBox "Please Enter From Godown!", vbInformation
       cboFromGodown.SetFocus
       Exit Sub
    End If
    
    If cboToGodown.text = "" Then
       MsgBox "Please Enter From Godown!", vbInformation
       cboToGodown.SetFocus
       Exit Sub
    End If
    
ElseIf type1.text = "R" Then
    If Me.txtChallanNo.text = "" Then
      MsgBox "Enter challan No ..", vbInformation
      txtChallanNo.SetFocus
      Exit Sub
    End If
    
    If Me.txtFrom_Gdid.text = "" Then
      MsgBox "Enter Godown Name ..", vbInformation
      cboFromGodown.SetFocus
      Exit Sub
    End If
End If


st_str = "n"
For I = 1 To grid1.rows - 1
If grid1.TextMatrix(I, 1) <> "" Then st_str = "y"
Next

If st_str = "n" Then
   MsgBox "Plz. Fill The Grid ..", vbInformation
   Exit Sub
End If


Set RS = New ADODB.Recordset
RS.Open "paperstatement", con, adOpenDynamic, adLockPessimistic
For I = 1 To grid1.rows - 1

If grid1.TextMatrix(I, 1) <> "" Then

    RS.AddNew
    RS.Fields("PaperTrans_Deliv").value = type1.text
    RS.Fields("Challan_No") = (txtChallanNo.text)
    RS.Fields("Challan_Date") = (txtChallanDate.value)
    RS.Fields("Sn_No") = txtsno.text
    RS.Fields("Sn_Date") = txtSNDate.value
    RS.Fields("remarks") = txtrem.text

    RS.Fields("FromGodown") = cboFromGodown.text
    RS.Fields("ToGodown") = cboToGodown.text

    RS.Fields("FromGodown_id") = txtFrom_Gdid.text
    RS.Fields("ToGodown_id") = txtTo_Gdid.text
    
    RS.Fields("sno") = Val(grid1.TextMatrix(I, 0))
    RS.Fields("Paper_Make") = grid1.TextMatrix(I, 1)
    RS.Fields("Bill_Date") = grid1.TextMatrix(I, 2)
    
    
    str_p = Trim(Mid(grid1.TextMatrix(I, 1), InStr(grid1.TextMatrix(I, 1), "=>") + 2))
    If rs1.State = 1 Then rs1.close
    rs1.Open "select SizeValue1,SizeValue2,gsm from PaperMakeMaster where " & stringyear & " and papermaker_id='" & str_p & "'", con
    If rs1.EOF = False Then
        RS.Fields("size") = rs1(0) & "X" & rs1(1)
        RS.Fields("gsm") = rs1(2)
        RS.Fields("pcode") = str_p
    End If
    
    
    RS.Fields("reams") = IIf(grid1.TextMatrix(I, 5) = "", 0, grid1.TextMatrix(I, 5))
    RS.Fields("Sheets") = Val(grid1.TextMatrix(I, 6))
    
    If grid1.TextMatrix(I, 7) <> "" Then
       RS.Fields("weight") = grid1.TextMatrix(I, 7)
    End If
    
    RS.Fields("VehicleNo") = grid1.TextMatrix(I, 8)
    RS.Fields("NoOfReal") = grid1.TextMatrix(I, 9)
    RS.Fields("Reamwt") = grid1.TextMatrix(I, 10)

    RS.Fields("fyear") = session
    RS.Fields("setupid") = setupid


    RS.update

End If

Next I



ok.Enabled = False
MsgBox "Record Saved", vbInformation


End Sub
Sub search_Data()

On Error Resume Next

Set RS = New ADODB.Recordset
RS.Open "select * from paperstatement where " & stringyear & " and PaperTrans_Deliv='" & type1 & "'" & _
" and SN_No='" & txtsno.text & "' and FromGodown='" & cboFromGodown.text & "' order by sno", con, adOpenDynamic, adLockPessimistic

For I = 1 To RS.RecordCount
    Edit = True
    Edit.Enabled = True
    ok.Enabled = False
    delete.Enabled = False
    cancel.Enabled = True
    
    type1.text = RS.Fields("PaperTrans_Deliv").value
    txtChallanNo.text = RS.Fields("Challan_No")
    txtChallanDate.value = RS.Fields("Challan_Date")
    txtsno.text = RS.Fields("Sn_No")
    txtSNDate.value = RS.Fields("Sn_Date")
    txtrem.text = RS.Fields("remarks") & ""
    cboFromGodown.text = RS.Fields("FromGodown")
    cboToGodown.text = RS.Fields("ToGodown")
    txtFrom_Gdid.text = RS.Fields("FromGodown_id")
    txtTo_Gdid.text = RS.Fields("ToGodown_id")
    grid1.TextMatrix(I, 0) = RS.Fields("sno")
    grid1.TextMatrix(I, 1) = RS.Fields("Paper_Make")
    grid1.TextMatrix(I, 2) = RS.Fields("Bill_Date") & ""
    grid1.TextMatrix(I, 3) = RS.Fields("size") & ""
    grid1.TextMatrix(I, 4) = RS.Fields("gsm")
    grid1.TextMatrix(I, 5) = RS.Fields("reams")
    grid1.TextMatrix(I, 6) = RS.Fields("Sheets")
    grid1.TextMatrix(I, 7) = RS.Fields("weight")
    grid1.TextMatrix(I, 8) = RS.Fields("VehicleNo")
    grid1.TextMatrix(I, 9) = RS.Fields("NoOfReal")
    
    grid1.TextMatrix(I, 10) = RS.Fields("Reamwt") & ""
    
    RS.MoveNext
Next I



End Sub
Private Sub ptyname_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then
popuplist1 "Select customer_name as [Customer Name],add1 as Address,customer_id as [Godown Id] from customerMaster", con
End If
If KeyCode = 13 And party_id.text <> "" Then
grid1.Row = 1
grid1.Col = 1
grid1.SetFocus
End If

End Sub


Private Sub Quit_Click()
Unload Me
End Sub

Private Sub search_Click()
'If KeyCode = 113 Then

If type1 = "R" Then
   popuplist1 "Select distinct SN_No as SRNo,SN_Date as Date,FromGodown as Godown  from paperstatement " & _
   "where " & stringyear & " and PaperTrans_Deliv='" & type1 & "'", con
Else
   popuplist1 "Select distinct SN_No as SRNo,SN_Date as Date,FromGodown,ToGodown  from paperstatement " & _
   "where " & stringyear & " and PaperTrans_Deliv='" & type1 & "'", con
End If
'End If

End Sub

Private Sub Textfirmname_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
popuplist1 "Select firm_name,firm_id, city from firmMaster", con
End If
End Sub

Private Sub txtChallan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   If txtChallan.Visible = True Then
      chdatevalue.SetFocus
   End If
End If
End Sub

Private Sub txtGodownId1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   MsgBox "Please Choose the Destination Godown"
   txtGodownId1.SetFocus
   Exit Sub
End If

End Sub

Private Sub search_GotFocus()
   If PopUpValue1 <> "" Then
      Clearvalue
      txtsno.text = PopUpValue1
      search_Data
      PopUpValue1 = ""
      PopUpValue2 = "'"
      PopUpValue3 = ""
   End If
End Sub

Private Sub txtChallanDate_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then cboFromGodown.SetFocus
End Sub

Private Sub txtChallanNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtChallanDate.SetFocus
End Sub

Private Sub txtrem_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   grid1.SetFocus
   grid1.Col = 1
End If
End Sub
Private Sub txtSNDate_KeyDown(KeyCode As Integer, Shift As Integer)
If txtSNDate.Visible = True Then
If txtChallanNo.Visible = True Then
   If KeyCode = 13 Then txtChallanNo.SetFocus
Else
   cboFromGodown.SetFocus
End If
End If
End Sub
Private Sub txtsno_GotFocus()
   If PopUpValue1 <> "" Then
      Clearvalue
      txtsno.text = PopUpValue1
      cboFromGodown.text = PopUpValue3
      search_Data
      PopUpValue1 = ""
      PopUpValue2 = "'"
      PopUpValue3 = ""
   End If
End Sub
Private Sub txtsno_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then

If type1 = "R" Then
   popuplist1 "Select distinct convert(int,SN_No) as SRNo,SN_Date as Date,FromGodown as Godown,Challan_No as Manual_ChNo  from paperstatement " & _
   "where " & stringyear & " and PaperTrans_Deliv='" & type1 & "'", con
Else
   popuplist1 "Select distinct SN_No as SRNo,SN_Date as Date,FromGodown,ToGodown,Challan_No as Manual_ChNo  from paperstatement " & _
   "where " & stringyear & " and PaperTrans_Deliv='" & type1 & "'", con
End If
End If

End Sub

Private Sub txtsno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then sendkeys "{tab}"
End Sub

Private Sub type1_Click()
'Frame1.Visible = False
Label5.Visible = True
If type1.text = "R" Then
Label5.Caption = "PAPER RECEIVED VOUCHER"
ElseIf type1.text = "D" Then
Label5.Caption = "PAPER TRANSFER VOUCHER"
End If
End Sub
Private Sub type1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'Me.Textfirmname.SetFocus
End If
End Sub
Private Sub type1_LostFocus()
Frame1.Visible = False
Label5.Visible = True
If type1.text = "R" Then
Label5.Caption = "PAPER RECEIVED VOUCHER"
ElseIf type1.text = "D" Then
Label5.Caption = "PAPER TRANSFER VOUCHER"
End If
End Sub

Sub Clearvalue()
grid1.Clear
grid_ini
party_id = ""
ptyname = ""
cty = ""
billid = ""
cboFromGodown.text = ""
cboToGodown.text = ""
txtChallanNo.text = ""
txtrem.text = ""

End Sub
Sub addenab()
Me.Edit.Enabled = True
Me.delete.Enabled = True
Me.search.Enabled = True
Me.Printcmd.Enabled = True
Me.Frame2.Enabled = True
End Sub
Sub adddisab()
 Me.Edit.Enabled = False
 Me.delete.Enabled = False
 ''Me.search.Enabled = False
 Me.Printcmd.Enabled = False
End Sub


