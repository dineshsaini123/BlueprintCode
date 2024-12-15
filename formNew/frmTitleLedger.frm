VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmTitleLedger 
   Caption         =   "Title Ledger"
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   11592
   Icon            =   "frmTitleLedger.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8880
   ScaleWidth      =   11592
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FAEFC9&
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   720
      Width           =   1080
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00FAEFC9&
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   6795
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   810
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.ComboBox COMBOGENLEDGER 
      Height          =   315
      Left            =   12810
      Sorted          =   -1  'True
      TabIndex        =   5
      Text            =   "SUNDRY DEBTORS"
      Top             =   420
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.CommandButton Commandshow 
      Caption         =   "&Show"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4035
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.CommandButton Commandreturn 
      Caption         =   "&Return"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4035
      TabIndex        =   3
      Top             =   780
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.ComboBox Combosubledger 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   312
      Left            =   1440
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   2
      Top             =   300
      Width           =   5745
   End
   Begin VB.TextBox Alpha 
      Height          =   315
      Left            =   12870
      MaxLength       =   1
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox txtscid 
      Enabled         =   0   'False
      Height          =   375
      Left            =   7155
      TabIndex        =   0
      Top             =   300
      Visible         =   0   'False
      Width           =   735
   End
   Begin Crystal.CrystalReport crpt 
      Left            =   585
      Top             =   9315
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
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
      Left            =   6030
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   1110
      _ExtentX        =   1947
      _ExtentY        =   550
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Left            =   5535
      TabIndex        =   9
      Top             =   1080
      Visible         =   0   'False
      Width           =   1155
      _ExtentX        =   2053
      _ExtentY        =   550
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   6885
      Left            =   135
      TabIndex        =   10
      Top             =   1485
      Width           =   11370
      _cx             =   20055
      _cy             =   12144
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16251308
      ForeColor       =   16711680
      BackColorFixed  =   16251308
      ForeColorFixed  =   255
      BackColorSel    =   16448755
      ForeColorSel    =   16744448
      BackColorBkg    =   16251308
      BackColorAlternate=   16251308
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
      RowHeightMin    =   380
      RowHeightMax    =   0
      ColWidthMin     =   400
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
   Begin VB.Label lbl_crdr 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10335
      TabIndex        =   23
      Top             =   300
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblOp 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   8775
      TabIndex        =   22
      Top             =   300
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Press F2 For Search"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   1440
      TabIndex        =   21
      Top             =   0
      Width           =   2715
   End
   Begin VB.Label lblDrCR 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10335
      TabIndex        =   20
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Closing :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7935
      TabIndex        =   19
      Top             =   720
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Opening  :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7935
      TabIndex        =   18
      Top             =   300
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label lblClosingBal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   8775
      TabIndex        =   17
      Top             =   720
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Label lblDrTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   7215
      TabIndex        =   16
      Top             =   8430
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label lblCrTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   9840
      TabIndex        =   15
      Top             =   8430
      Width           =   1335
   End
   Begin VB.Label lblCr 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   11610
      TabIndex        =   14
      Top             =   5640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblDr 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   11610
      TabIndex        =   13
      Top             =   6000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Title Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   180
      TabIndex        =   12
      Top             =   360
      Width           =   1200
   End
   Begin VB.Label Label3 
      Caption         =   " - To - "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6615
      TabIndex        =   11
      Top             =   1080
      Visible         =   0   'False
      Width           =   585
   End
End
Attribute VB_Name = "frmTitleLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim bb As Boolean

Dim bb2 As Boolean
Dim rss As New ADODB.Recordset
Dim from_date As Date
Dim I As Integer
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim search_v As Boolean
Dim to_date As Date
Dim kk As Integer
Dim bb1 As Boolean
Dim str1 As New ADODB.Recordset
Dim din_ As Boolean

Sub vsIni()

   
End Sub

Private Sub All_Click()
If All.value = True Then
'    Call cmdShow_Click
End If

End Sub

Private Sub autho_Click()
If autho.value = True Then
'    Call cmdShow_Click
End If
End Sub



Private Sub cash_Click()
    If cash.value = True Then
'       Call cmdShow_Click
    End If
End Sub

Private Sub cboop_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtdes.SetFocus
   End If
End Sub

Private Sub cboStation_Click()
cboPartyList.Visible = True
If RS.State = 1 Then RS.close
RS.Open "select distinct(AgentName) from SLEDGER where " & stringyear & " and DISTCODE='" & cboStation.Text & "'", con
cboPartyList.Clear
While RS.EOF = False
cboPartyList.AddItem RS(0)
RS.MoveNext
Wend
End Sub

Private Sub Check1_Click()
    If Check1.value = 1 Then
       'cmdSave.Enabled = False
       cmdDel.Enabled = False
       cmdModify.Enabled = False
    Else
       cmdSave.Enabled = True
       cmdDel.Enabled = True
       cmdModify.Enabled = True
    End If
End Sub

Private Sub Check2_Click()

Dim rs_1 As New ADODB.Recordset

cboStation.Clear
cboStation1.Clear

If Check2.value = 1 Then
    
    lblStation.Caption = "State :"
    
    If rs_1.State = 1 Then rs_1.close
    rs_1.Open "select distinct(states) from SLEDGER where " & stringyear & " and states<>''", con
    While rs_1.EOF = False
    cboStation.AddItem rs_1.Fields(0).value
    cboStation1.AddItem rs_1.Fields(0).value
    rs_1.MoveNext
    Wend

ElseIf Check2.value = 0 Then
    
    lblStation.Caption = "Station :"

    If rs_1.State = 1 Then rs_1.close
    rs_1.Open "select distinct(DISTCODE) from SLEDGER where " & stringyear & " and DISTCODE<>''", con
    While rs_1.EOF = False
    cboStation.AddItem rs_1.Fields(0).value
    cboStation1.AddItem rs_1.Fields(0).value
    rs_1.MoveNext
    Wend


End If

End Sub

Private Sub cmdAson_Click()
'showDataAsOn dateason
End Sub

Private Sub cmddewali_Click()
    Dim f As New ADODB.Recordset
    If f.State = 1 Then f.close
    f.Open "select AMOUNT,text,INVOICENO from invoicec where " & stringyear & " and TEXT = '" & "DIWALI SPECIAL" & "' and AMOUNT>0", con
    While f.EOF = False
        con.Execute "update INVOICEA_sp set t2='" & f.Fields("amount").value & "' where " & stringyear & " and INVOICENO=" & f.Fields("INVOICENO").value & ""
        f.MoveNext
    Wend
    If f.State = 1 Then f.close
    f.Open "select AMOUNT,text,INVOICENO from CASHC where " & stringyear & " and TEXT = '" & "DIWALI SPECIAL" & "' and AMOUNT>0", con
    While f.EOF = False
        con.Execute "update CASHA set t2='" & f.Fields("amount").value & "' where " & stringyear & " and INVOICENO=" & f.Fields("INVOICENO").value & ""
        f.MoveNext
    Wend
    MsgBox "Data Refresh...", vbInformation
End Sub

Private Sub cmdPath_Click()
'Me.comdio.ShowOpen
'Me.txtPath.Text = Me.comdio.FileName
End Sub

Private Sub cmdExit_12_Click()
Unload Me
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()

On Error GoTo aa10

Screen.MousePointer = vbHourglass
Dim op, drcr
Dim rs1 As New ADODB.Recordset

con.Execute "delete from templedger1"



'==============================================
For I = 1 To vs.rows - 1

If vs.TextMatrix(I, 1) <> "" Then
con.Execute "INSERT INTO templedger1 (bill,dates,Billtype,party,dr,setupid,fyear,des) values('" & vs.TextMatrix(I, 0) & "','" & Format(vs.TextMatrix(I, 1), "MM/dd/yyyy") & "','" & vs.TextMatrix(I, 2) & "','" & vs.TextMatrix(I, 3) & "','" & vs.TextMatrix(I, 4) & "'," & setupid & ",'" & session & "','" & Combosubledger.Text & "') "
End If

Next

DSNNew

Sleep (300)
crpt.Reset
crpt.ReportFileName = rptPath & "/SchoolSales.rpt"
crpt.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
crpt.WindowShowPrintSetupBtn = True
crpt.WindowShowPrintBtn = True
crpt.WindowState = crptMaximized
crpt.Action = 1
Screen.MousePointer = vbDefault
Exit Sub
aa10:
MsgBox err.DESCRIPTION
End Sub
Private Sub cmdPrint1_Click()

crpt.Reset

If Check_ClosingDesc.value = 1 Then
   crpt.ReportFileName = st1 & "\" & directory & "\PartyWiseClosing_descClosing.rpt"
Else
   crpt.ReportFileName = st1 & "\" & directory & "\PartyWiseClosing.rpt"
End If

crpt.DataFiles(0) = st1 & "\" + directory + "\data.mdb"

''======================================================================
''======================================================================

If Check2.value = 0 Then

    If cboStation1.Text <> "" And txtAmount.Text = "" Then
    crpt.ReplaceSelectionFormula "{tempLedgerRpt.DISTCODE}='" & cboStation1.Text & "' and {tempLedgerRpt.Owner}<>0"
    ElseIf cboStation1.Text <> "" And txtAmount.Text <> "" Then
    If Val(txtAmount) >= 0 Then
       crpt.ReplaceSelectionFormula "{tempLedgerRpt.DISTCODE}='" & cboStation1.Text & "' and {tempLedgerRpt.Owner}<>0 and {tempLedgerRpt.Owner}>=" & txtAmount.Text & ""
    Else
       crpt.ReplaceSelectionFormula "{tempLedgerRpt.DISTCODE}='" & cboStation1.Text & "' and {tempLedgerRpt.Owner}<>0 and {tempLedgerRpt.Owner}<=" & txtAmount.Text & ""
    End If
    
    
    ElseIf cboStation1.Text = "" And txtAmount.Text <> "" Then
    
    If Val(txtAmount) >= 0 Then
       crpt.ReplaceSelectionFormula "{tempLedgerRpt.Owner}<>0 and {tempLedgerRpt.Owner}>=" & txtAmount.Text & ""
    Else
       crpt.ReplaceSelectionFormula "{tempLedgerRpt.Owner}<>0 and {tempLedgerRpt.Owner}<=" & txtAmount.Text & ""
    End If
    
    Else
    crpt.ReplaceSelectionFormula "abs({tempLedgerRpt.Owner})<>0"
    End If


ElseIf Check2.value = 1 Then


    If cboStation1.Text <> "" And txtAmount.Text = "" Then
    crpt.ReplaceSelectionFormula "{tempLedgerRpt.states}='" & cboStation1.Text & "' and {tempLedgerRpt.Owner}<>0"
    ElseIf cboStation1.Text <> "" And txtAmount.Text <> "" Then
    If Val(txtAmount) >= 0 Then
       crpt.ReplaceSelectionFormula "{tempLedgerRpt.states}='" & cboStation1.Text & "' and {tempLedgerRpt.Owner}<>0 and {tempLedgerRpt.Owner}>=" & txtAmount.Text & ""
    Else
       crpt.ReplaceSelectionFormula "{tempLedgerRpt.states}='" & cboStation1.Text & "' and {tempLedgerRpt.Owner}<>0 and {tempLedgerRpt.Owner}<=" & txtAmount.Text & ""
    End If
    
    
    ElseIf cboStation1.Text = "" And txtAmount.Text <> "" Then
    
    If Val(txtAmount) >= 0 Then
       crpt.ReplaceSelectionFormula "{tempLedgerRpt.Owner}<>0 and {tempLedgerRpt.Owner}>=" & txtAmount.Text & ""
    Else
       crpt.ReplaceSelectionFormula "{tempLedgerRpt.Owner}<>0 and {tempLedgerRpt.Owner}<=" & txtAmount.Text & ""
    End If
    
    Else
    crpt.ReplaceSelectionFormula "abs({tempLedgerRpt.Owner})<>0"
    End If



End If

''======================================================================
''======================================================================










DoEvents
MsgBox ("View")
crpt.Formulas(0) = "partyname='" & cboStation1.Text & "'"
crpt.Formulas(1) = "ason='" & dateAson.value & "'"

crpt.WindowShowPrintSetupBtn = True
crpt.WindowShowPrintBtn = True
crpt.WindowState = crptMaximized
crpt.WindowShowSearchBtn = True
crpt.Action = 1

End Sub

Private Sub cmdPrintAgentLed_Click()

DSNNew

Screen.MousePointer = vbHourglass
With crpt
 .Reset
 .ReportFileName = rptPath & "/AgentLedger.rpt"
 .Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
 .ReplaceSelectionFormula "{tempLedgerRpt.party}='" & cboParty.Text & "'"
 .WindowShowPrintSetupBtn = True
 .WindowState = crptMaximized
 .Action = 1
End With
Screen.MousePointer = vbDefault


End Sub
Private Sub cmdprintalf_Click()
 
 If txtalfa.Text = "" Then
    MsgBox "Please Enter Alphabet...", vbInformation
    Exit Sub
 End If
 
 Screen.MousePointer = vbHourglass
 CityWiseStatement
 Screen.MousePointer = vbDefault

End Sub

Private Sub cmdset_Click()
   
If RS.State = 1 Then RS.close
RS.Open "select * from pass where pass='" & cp & "'", con
If RS.EOF = True Then
MsgBox "Enter Valid Password !!", vbInformation
Exit Sub
End If
   
    
saveData
   
End Sub
Sub saveData()
   
''''   If MsgBox("Want to Set ?", vbQuestion + vbYesNo) = vbYes Then
''''
''''   Screen.MousePointer = vbHourglass
''''   'cmdShow1.Visible = True
''''
''''
''''   If sales.value = True Then
''''
''''        For J = 1 To vs.Rows - 1
''''          If vs.TextMatrix(J, 5) = True Then
''''             CON.Execute "update INVOICEA_sp set bAuthorized=" & vs.TextMatrix(J, 5) & " where " & stringyear & " and INVOICENO=" & vs.TextMatrix(J, 1) & ""
''''            Else
''''             CON.Execute "update INVOICEA_sp set bAuthorized=" & vs.TextMatrix(J, 5) & " where " & stringyear & " and INVOICENO=" & vs.TextMatrix(J, 1) & ""
''''          End If
''''        Next
''''
''''  ElseIf credit.value = True Then
''''
''''        For J = 1 To vs.Rows - 1
''''          If vs.TextMatrix(J, 5) = True Then
''''            CON.Execute "update invoicea_spRet set bAuthorized=" & vs.TextMatrix(J, 5) & " where " & stringyear & " and  INVOICENO=" & vs.TextMatrix(J, 1) & ""
''''            Else
''''            CON.Execute "update invoicea_spRet set bAuthorized=" & vs.TextMatrix(J, 5) & " where " & stringyear & " and INVOICENO=" & vs.TextMatrix(J, 1) & ""
''''          End If
''''        Next
''''
''''
''''  ElseIf cash.value = True Then
''''
''''        For J = 1 To vs.Rows - 1
''''          If vs.TextMatrix(J, 5) = True Then
''''            CON.Execute "update CASHA set BAuthorized=" & vs.TextMatrix(J, 5) & " where " & stringyear & " and INVOICENO=" & vs.TextMatrix(J, 1) & ""
''''            Else
''''            CON.Execute "update CASHA set BAuthorized=" & vs.TextMatrix(J, 5) & " where INVOICENO=" & vs.TextMatrix(J, 1) & ""
''''          End If
''''        Next
''''
''''  ElseIf crdit.value = True Then
''''
''''        For J = 1 To vs.Rows - 1
''''          If vs.TextMatrix(J, 5) = True Then
''''            CON.Execute "update cnf1a set bAuthorized=" & vs.TextMatrix(J, 5) & " where " & stringyear & " and cnn=" & vs.TextMatrix(J, 1) & ""
''''            Else
''''            CON.Execute "update cnf1a set bAuthorized=" & vs.TextMatrix(J, 5) & " where " & stringyear & " and cnn=" & vs.TextMatrix(J, 1) & ""
''''          End If
''''        Next
''''
''''
''''  ElseIf dbit.value = True Then
''''
''''        For J = 1 To vs.Rows - 1
''''          If vs.TextMatrix(J, 5) = True Then
''''            CON.Execute "update dnfa set bAuthorized=" & vs.TextMatrix(J, 5) & " where " & stringyear & " and dnn=" & vs.TextMatrix(J, 1) & ""
''''            Else
''''            CON.Execute "update dnfa set bAuthorized=" & vs.TextMatrix(J, 5) & " where " & stringyear & " and dnn=" & vs.TextMatrix(J, 1) & ""
''''          End If
''''        Next
''''
''''
''''  End If
''''
''''
''''   End If
''''
''''
'''' Screen.MousePointer = vbDefault
End Sub
Sub SearchFa()
      
      If RS.State = 1 Then RS.close
      If din_ = False Then
         RS.Open "select INVOICENO,INVOICEDATE,AgentName,netamount,BAA,t2 from INVOICEA_sp where " & stringyear & " and AgentName='" & cboParty.Text & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)", con
      Else
         RS.Open "select INVOICENO,INVOICEDATE,AgentName,netamount,BAA,t2 from INVOICEA_sp where " & stringyear & " and shipto='" & cboParty.Text & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)", con
      End If
      
      If RS.EOF = False Then
        vs1.rows = (vs1.rows + RS.RecordCount)
        For I = I To vs1.rows - 1
        If RS.EOF = False Then
           vs1.TextMatrix(I, 0) = "I"
           vs1.TextMatrix(I, 1) = RS.Fields("INVOICENO").value
           vs1.TextMatrix(I, 2) = RS.Fields("INVOICEDATE").value
           If IsNull(RS.Fields("t2").value) Then
              vs1.TextMatrix(I, 3) = "Issue"
           Else
              vs1.TextMatrix(I, 3) = "Invoice Sales" & RS.Fields("t2").value & " " & "DS"
           End If
           vs1.TextMatrix(I, 4) = Format(RS.Fields("netamount").value, "0.00")
           vs1.TextMatrix(I, 5) = Format(RS.Fields("BAA").value, "0.00")
            RS.MoveNext
         End If
        Next
      End If
'    '================
     If RS.State = 1 Then RS.close
     RS.Open "select INVOICENO,INVOICEDATE,AgentName,netamount,baa from invoicea_spRet where " & stringyear & " and AgentName='" & cboParty.Text & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)", con
     'RS.Open "select INVOICENO,INVOICEDATE,AgentName,netamount,baa from invoicea_spRet where " & stringyear & " and AgentName='" & cboParty.Text & "'", CON
     If RS.EOF = False Then
        vs1.rows = vs1.rows + RS.RecordCount
        For I = I To vs1.rows - 1
         
        If RS.EOF = False Then
         vs1.TextMatrix(I, 0) = "R"
         vs1.TextMatrix(I, 1) = RS.Fields("INVOICENO").value
         vs1.TextMatrix(I, 2) = RS.Fields("INVOICEDATE").value
         vs1.TextMatrix(I, 3) = "Return"
         vs1.TextMatrix(I, 4) = Format(RS.Fields("baa").value, "0.00")
         vs1.TextMatrix(I, 5) = Format(RS.Fields("netamount").value, "0.00")
         RS.MoveNext
       End If
    Next
    End If


    vs1.FormatString = "^Bill Type|^Bill|^Date|<Description|>Dr|>Cr"
    setWidth
End Sub
Sub CityWiseStatement()
       Dim op, drcr
       Dim s As String
       s = ""
       Dim rs1 As New ADODB.Recordset
       con.Execute "delete from templedger1 " & stringyear & ""
       If RS.State = 1 Then RS.close
       If cboStation.Text <> "" And txtalfa.Text = "" Then
       For I = 0 To cboPartyList.ListCount - 1
        If cboPartyList.Selected(I) = True Then
        If s = "" Then
          s = "AgentName " & " = " & "'" & cboPartyList.List(I) & "'"
        Else
          s = s & " or " & "AgentName " & " = " & "'" & cboPartyList.List(I) & "'"
        End If
        End If
       Next
       
       If s = "" Then
        If RS.State = 1 Then RS.close
        RS.Open "select AgentName from SLEDGER where " & stringyear & " and DISTCODE = '" & cboStation.Text & "'", con
       Else
        If RS.State = 1 Then RS.close
        RS.Open "select AgentName from SLEDGER where " & stringyear & " and " & s, con
       End If
       
       ElseIf txtalfa.Text <> "" And cboStation.Text = "" Then
        RS.Open "select AgentName from SLEDGER where " & stringyear & " and AgentName like '" + Trim(txtalfa.Text) + "%'", con
       Else
         Exit Sub
       End If
       While RS.EOF = False
           '==Code For Opening=============================================
            con.Execute "INSERT INTO templedger1 (Balance,drcr,party,billtype)  SELECT op,drcr,AgentName,'Opening' from sledger where " & stringyear & " and AgentName = '" + RS.Fields(0).value + "'   group by op,AgentName,drcr HAVING  op <> 0;"
            If rs1.State = 1 Then rs1.close
            rs1.Open "SELECT op,drcr from sledger where " & stringyear & " and AgentName = '" + RS.Fields(0).value + "'", con
            If Not IsNull(rs1.Fields(0).value) Then
               op = Val(rs1.Fields(0).value)
               drcr = rs1.Fields(1).value
            Else
               op = 0
            End If
           '==============================================
            con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party)  SELECT INVOICEDATE,'I',INVOICENO,'Invoice Sales',netamount,BAA,AgentName from INVOICEA_sp where " & stringyear & " and AgentName='" & RS.Fields(0).value & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)"
            con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party) Select INVOICEDATE,'CI',INVOICENO,'Credit Note Item',baa,netamount,AgentName from invoicea_spRet where " & stringyear & " and AgentName='" & RS.Fields(0).value & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)"
            con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party) Select INVOICEDATE,'C/M',INVOICENO,'Cash Memo',netamount,baa,AgentName from CASHA where  " & stringyear & " and AgentName='" & RS.Fields(0).value & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)"
            con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party) select dnd,'DN',dnn,'Debit Note',na,'0',PSLD from dnfa where  " & stringyear & " and psld='" & RS.Fields(0).value & "'"
            con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,cr,dr,Party) Select cnd,'CN',cnn,'Credit Note',na,'0',psld from Cnf1a where " & stringyear & " and  psld='" & RS.Fields(0).value & "'"
            con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party) Select dates,'J',Recno,Particullar,Dr,CR,PartyName from ReceiveIssueParty where " & stringyear & " and PartyName='" & RS.Fields(0).value & "' order by dates,recno"
          '===============================================================
          If op <> 0 Then
           con.Execute "Update templedger1 set Balance=" & op & ",drcr= '" & drcr & "' where " & stringyear & " and party = '" + RS.Fields(0).value + "' and Billtype<>'Opening'"
          End If
          '===============================================================
           RS.MoveNext
       Wend
       DoEvents
       MsgBox "View"
 crpt.Reset
 'crpt.ReportFileName = App.Path & "\" & directory & "\PartyLedger.rpt"
 crpt.ReportFileName = st1 & "\" & directory & "\PartyLedger.rpt"
 crpt.DataFiles(0) = st1 & "\" + directory + "\data.mdb"
 crpt.WindowShowPrintSetupBtn = True
 crpt.Formulas(0) = "partyname='" & cboStation.Text & "'"
 crpt.WindowShowPrintBtn = True
 crpt.WindowState = crptMaximized
 crpt.Action = 1
End Sub
Private Sub cmdShowClosing_Click()

bb1 = False
'showData

End Sub

Private Sub cmdupdatep_Click()
   Dim partyname
   Dim pcode
   partyname = ""
   pcode = ""
   
    
   If RS.State = 1 Then RS.close
   RS.Open "select AgentName from sledger", con
   While RS.EOF = False
       
       aa = InStr(RS(0), " ")
       partyname = Mid(RS(0), aa)
       pcode = Mid(RS(0), 1, aa)
       
       con.Execute "update  Sledger  set party='" & Trim(partyname) & "',code='" & Trim(pcode) & "' where " & stringyear & " and AgentName='" & RS(0) & "'"
       
       RS.MoveNext
       
   Wend
   
End Sub

Private Sub Command1_Click()
   
  If RS.State = 1 Then RS.close
  RS.Open "select * from pass where pass='" & cp & "'", con
  If RS.EOF = True Then
     MsgBox "Enter Valid Password !!", vbInformation
     Exit Sub
  
  Else

   Screen.MousePointer = vbHourglass
   
   On Error Resume Next
   
   For I = 1 To vsop.rows - 1
       If vsop.TextMatrix(I, 1) <> "" Then
          con.Execute "update SLEDGER set op=" & CDbl(vsop.TextMatrix(I, 2)) & ",drcr='" & vsop.TextMatrix(I, 3) & "' where " & stringyear & " and AgentName='" & vsop.TextMatrix(I, 1) & "'"
       End If
   Next
   
   Screen.MousePointer = vbDefault
   

   
End If
   

   
   
   
   
   
End Sub



Private Sub Command2_Click()
crpt.Reset
crpt.ReportFileName = st1 & "\" & directory & "\PartyWiseDrClosing.rpt"
crpt.DataFiles(0) = st1 & "\" + directory + "\data.mdb"
crpt.ReplaceSelectionFormula "{tempLedgerRpt.Offdays}='" & "1" & "' and {tempLedgerRpt.Owner}>=" & 1 & ""
DoEvents
MsgBox ("View")
crpt.Formulas(0) = "partyname='" & cboStation1.Text & "'"
crpt.WindowShowPrintSetupBtn = True
crpt.WindowShowPrintBtn = True
crpt.WindowState = crptMaximized
crpt.WindowShowSearchBtn = True
crpt.Action = 1

End Sub

Private Sub Command3_Click()
 
 If cboStation.Text = "" Then
    MsgBox "Please Select Station...", vbInformation
    Exit Sub
 End If
 
 Screen.MousePointer = vbHourglass
 CityWiseStatement
 cboPartyList.Visible = False
 Screen.MousePointer = vbDefault
End Sub

Private Sub Command4_Click()

Dim FSO As FileSystemObject
Dim f As File
Dim txt As TextStream
Dim matter As String
Dim Total As String
Dim s(1, 2) As String
Set FSO = New FileSystemObject
Dim ss As String
'
Dim s1

matter = ""

Set txt = FSO.CreateTextFile(App.Path & "\mobile.txt", True)

If RS.State = 1 Then RS.close
If Check2.value = 0 Then
RS.Open "select mobile from sledger where " & stringyear & " and distcode='" & cboStation1.Text & "'", con, adOpenKeyset, adLockReadOnly
Else
RS.Open "select mobile from sledger where " & stringyear & " and states='" & cboStation1.Text & "'", con, adOpenKeyset, adLockReadOnly
End If

While RS.EOF = False


If Len(RS(0)) > 0 Then

s1 = Split(RS(0), ",")
For I = 0 To UBound(s1)
    matter = matter & Trim(s1(I)) & vbNewLine
Next



End If
RS.MoveNext
Wend

txt.Write matter
txt.close

MsgBox "File Created ....", vbInformation

Shell App.Path & "\notepad.exe " & App.Path & "\mobile.txt", vbMaximizedFocus

End Sub

Private Sub crdit_Click()
    If crdit.value = True Then
'       Call cmdShow_Click
    End If
End Sub

Private Sub credit_Click()
    If credit.value = True Then
'       Call cmdShow_Click
    End If

End Sub

Private Sub dbit_Click()
   If dbit.value = True Then
'       Call cmdShow_Click
    End If
End Sub

Private Sub Combosubledger_GotFocus()

Dim fill_ As New ADODB.Recordset

If PopUpValue1 <> "" Then

    
    
    lblCrTotal.Caption = 0
    
    Combosubledger.Text = PopUpValue1 & " : " & PopUpValue2
    
    If frmBillList.cboParty.Text = "" Then Exit Sub
    
    Set fill_ = New ADODB.Recordset
    Set fill_ = con.Execute("exec TitleLedger '" & frmBillList.cboParty.Text & "','" & PopUpValue1 & "'")
    'Set fill_ = con.Execute("exec TitleLedger '" & PopUpValue1 & "'")
    
    setWidth
       
    k1 = 1
    For I = 1 To fill_.RecordCount
      
      
      vs.TextMatrix(k1, 0) = fill_!invoiceNo
      vs.TextMatrix(k1, 1) = fill_!invoiceDate
      vs.TextMatrix(k1, 2) = fill_!scname
      vs.TextMatrix(k1, 3) = fill_!QUANTITY
      vs.TextMatrix(k1, 4) = fill_!rate
      vs.TextMatrix(k1, 5) = fill_!amount
      vs.TextMatrix(k1, 6) = fill_!discount
      vs.TextMatrix(k1, 7) = fill_!netamount
      
      lblCrTotal.Caption = Val(lblCrTotal.Caption) + fill_!netamount
      
      vs.rows = vs.rows + 1
      k1 = k1 + 1
      
      fill_.MoveNext
      
      
      
    Next
       
    
       
    PopUpValue1 = ""
    PopUpValue2 = ""
   

End If


End Sub

Private Sub Combosubledger_KeyDown(KeyCode As Integer, Shift As Integer)

'If KeyCode = 113 Then
'
'    searchType = "party"
'    value = "SELECT ScName,ScID FROM INVOICEA where len(ScName)>0 group by ScName,ScID"
'    popuplist_client value, con
'    set_focus = True
'
'End If


If KeyCode = 113 Then
    value = "select BookCode,BookName from books where " & stringyear & "  order by BookCode"
    popuplist_client value, CCON
End If

End Sub


Private Sub Form_Activate()
' Me.WindowState = 2
  
End Sub
Private Sub Form_Load()

Me.Top = 500
Me.Left = 500

Me.Width = 11700
Me.Height = 9800

Me.Caption = frmBillList.cboParty.Text

vsIni
On Error Resume Next

kk = 1

dateAson.value = Date




FromDate.value = Date
toDate.value = Date
from_date = FromDate.value

maxId
setWidth

cboop.ListIndex = 0



If RS.State = 1 Then RS.close
RS.Open "select yarfrom,yarto from setup1 where " & stringyear & "", con
If RS.EOF = False Then
   FromDate.value = RS.Fields(0).value
   
   If (DateValue(RS!yarfrom) <= DateValue(Date) And DateValue(RS!yarto) >= DateValue(Date)) Then
      RecDates.value = Date
   Else
      RecDates.value = RS.Fields(1).value
   End If
   
End If

Me.Top = 50
Me.Left = 50


Opening.Tab = 1


If RS.State = 1 Then RS.close
RS.Open "select * from setup1 where " & stringyear & "", con, adOpenDynamic, adLockReadOnly, adCmdTable
If RS.EOF = False Then
    date1.Text = RS!yarfrom
    date2.Text = RS!yarto
End If



bb1 = False


fetchTab2

BackColorFrom Me

Screen.MousePointer = vbDefault

End Sub
Sub setsecurity()
   
If LCase(strledger) <> "cp" Then
   cmdShow1.Visible = False
   MsgBox "Enter Valid Password !!", vbInformation
   Exit Sub
Else
  
  
  
  saveData
   
End If
   
End Sub
Private Sub Form_Resize()
'panel.Left = (Me.ScaleWidth - panel.Width) / 2
'panel.Top = (Me.ScaleHeight - panel.Height) / 2

End Sub

Private Sub Opening_Click(PreviousTab As Integer)
      
     ' Screen.MousePointer = vbHourglass
      
      
      Dim closing As Double
      
      
      closing = 0
      
      If Opening.Tab = 0 Then
         
'         Call cmdShow_Click
         
      ElseIf Opening.Tab = 2 Then
       
        
'
      
      
      End If
      
      
      
      'Screen.MousePointer = vbDefault
      
End Sub
Sub fetchTab2()

        Screen.MousePointer = vbHourglass

        Dim fillVs As New ADODB.Recordset
        If fillVs.State = 1 Then fillVs.close
        'fillvs.Open "select DISTCODE as City,AgentName as Party,op,drcr from closing where gledger='SUNDRY DEBTORS'", con
        fillVs.Open "SELECT SLEDGER.DISTCODE,SLEDGER.AgentName,SLEDGER.OP,SLEDGER.drcr,(Sum(templedger1.Dr)-Sum(templedger1.Cr)) AS bal1 FROM SLEDGER LEFT JOIN templedger1 ON SLEDGER.AgentName = templedger1.Party where " & stringyear & " and  gledger='SUNDRY DEBTORS' GROUP BY SLEDGER.AgentName,SLEDGER.DISTCODE,[SLEDGER.OP], SLEDGER.drcr, SLEDGER.gledger", con

        If fillVs.EOF = False Then
            vsop.rows = fillVs.RecordCount
            For I = 1 To vsop.rows - 1
              vsop.TextMatrix(I, 0) = fillVs(0) & ""
              vsop.TextMatrix(I, 1) = fillVs(1)
              vsop.TextMatrix(I, 2) = Format(fillVs(2), "0.00")
              vsop.TextMatrix(I, 3) = fillVs(3) & ""

              If Not IsNull(fillVs(4)) Then

                     If vsop.TextMatrix(I, 3) = "Cr" Then
                         vsop.TextMatrix(I, 4) = ((-1 * (vsop.TextMatrix(I, 2))) + fillVs(4))
                         If vsop.TextMatrix(I, 4) < 0 Then
                            vsop.TextMatrix(I, 4) = Format((-1 * Val(vsop.TextMatrix(I, 4))), "0.00")
                            vsop.TextMatrix(I, 5) = "Cr"
                         Else
                            vsop.TextMatrix(I, 5) = "Dr"
                            vsop.TextMatrix(I, 4) = Format(Val(vsop.TextMatrix(I, 4)), "0.00")
                         End If

                     Else
                         vsop.TextMatrix(I, 4) = ((Val(vsop.TextMatrix(I, 2))) + fillVs(4))
                         If vsop.TextMatrix(I, 4) < 0 Then
                            vsop.TextMatrix(I, 4) = Format((-1 * Val(vsop.TextMatrix(I, 4))), "0.00")
                            vsop.TextMatrix(I, 5) = "Cr"
                         Else
                            vsop.TextMatrix(I, 5) = "Dr"
                            vsop.TextMatrix(I, 4) = Format(Val(vsop.TextMatrix(I, 4)), "0.00")
                         End If


                     End If
              End If


              fillVs.MoveNext
            Next
        End If

        vsop.Cols = 6
        vsop.TextMatrix(0, 0) = "City"
        vsop.TextMatrix(0, 1) = "Party"
        vsop.TextMatrix(0, 2) = "Opening"
        vsop.TextMatrix(0, 3) = "Dr/Cr"
        vsop.TextMatrix(0, 4) = "Closing"
        vsop.TextMatrix(0, 5) = "Dr/Cr"


        vsop.ColWidth(0) = 1800
        vsop.ColWidth(1) = 3600
        vsop.ColWidth(2) = 1200
        vsop.ColWidth(3) = 500
        vsop.ColWidth(4) = 1200
        vsop.ColWidth(5) = 500

        Screen.MousePointer = vbDefault



End Sub
Private Sub Option1_Click()
   If Option1.value = True Then
      bill.Visible = True
   Else
      partydrcr.Visible = False
      bill.Visible = True
   End If
End Sub

Private Sub Option2_Click()
   If Option2.value = 1 Then
      txtadmin.Visible = True
      Label14.Visible = True
   Else
      txtadmin.Visible = False
      Label14.Visible = False
   End If
End Sub

Private Sub party_Click()
   
   If party.value = True Then
      bill.Visible = False
      frmReceiveFromParty.Show
   Else
      partydrcr.Visible = False
      bill.Visible = True
   End If
   
   frmReceiveFromParty.Top = 800

End Sub
Private Sub RecDates_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 13 Then
        set_focus = False
        cboParty.SetFocus
     End If
End Sub
Private Sub SSTab1_DblClick()
   RecDates.SetFocus
End Sub

Private Sub sales_Click()
    If sales.value = True Then
'       Call cmdShow_Click
    End If
End Sub

Private Sub selectAll_Click()
If selectAll.value = 1 Then
    For I = 0 To cboPartyList.ListCount - 1
        cboPartyList.Selected(I) = True
    Next
Else
   For I = 0 To cboPartyList.ListCount - 1
    cboPartyList.Selected(I) = False
   Next
End If
End Sub

Private Sub txtadmin_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   setsecurity
   'pass.Visible = False
End If
End Sub

Private Sub txtdes_GotFocus()
  txtdes.BackColor = &HFFFFC0
End Sub

Private Sub txtdes_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     txtQty.SetFocus
  End If
End Sub

Private Sub txtdes_LostFocus()
    txtdes.BackColor = &HFFFFFF
End Sub

Private Sub txtOp_GotFocus()
txtOp.BackColor = &HFFFFC0
End Sub
Private Sub txtParty_GotFocus()
   If PopUpValue1 <> "" Then
      txtParty.Text = PopUpValue1
   End If
End Sub
Private Sub txtParty_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 113 Then
       value = "select AgentName from INVOICEA_sp where " & stringyear & "  order by AgentName"
       popuplistModel10 value, con
    End If
End Sub
Private Sub txtParty_LostFocus()
  PopUpValue1 = ""
End Sub
Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
   If Val(txtQty.Text) = 0 Then
      txtQty.SetFocus
      Exit Sub
   End If
   If cmdSave.Enabled = True Then
      cmdSave.SetFocus
   End If
   End If
End Sub
Private Sub txtRem_LostFocus()
  If cboParty.Text <> "" Then
  If MsgBox("Want To Change Remarks ?", vbQuestion + vbYesNo) = vbYes Then
      con.Execute "update sledger set PartyRemarks = '" & txtRem.Text & "' where " & stringyear & " and AgentName='" & cboParty.Text & "'"
     
  End If
  End If
End Sub
Private Sub Unautho_Click()
If Unautho.value = True Then
'    Call cmdShow_Click
End If
End Sub
Private Sub vs1_KeyDown(KeyCode As Integer, Shift As Integer)



Screen.MousePointer = vbHourglass
If KeyCode = 13 Then
If vs1.TextMatrix(vs1.RowSel, 0) = "I" Then
         inviceNo = vs1.TextMatrix(vs1.RowSel, 1)
         frmBookIssueSp.Show
ElseIf vs1.TextMatrix(vs1.RowSel, 0) = "R" Then
         inviceNo = vs1.TextMatrix(vs1.RowSel, 1)
         'MainMenu.Toolbar1.Visible = False
         frmBookIssueSp_Ret.Show
'ElseIf credit.value = True Then
'   If vs1.Col = 1 Then
'         inviceNo = vs1.TextMatrix(vs1.RowSel, 1)
'         'MainMenu.Toolbar1.Visible = False
'         Critnote.Show
'   End If
'ElseIf crdit.value = True Then
'   If vs1.Col = 1 Then
'         inviceNo = vs1.TextMatrix(vs1.RowSel, 1)
'         'MainMenu.Toolbar1.Visible = False
'         Creditnotefile.Show
'   End If
'ElseIf dbit.value = True Then
'   If vs1.Col = 1 Then
'         inviceNo = vs1.TextMatrix(vs1.RowSel, 1)
'         'MainMenu.Toolbar1.Visible = False
'         Debitnotefile.Show
'   End If
End If
End If
Screen.MousePointer = vbDefault
End Sub

Private Sub vs_DblClick()

If vs.TextMatrix(vs.RowSel, 0) <> "" Then
         inviceNo = vs.TextMatrix(vs.RowSel, 0)
         s1 = 1
         invoice.Show
'ElseIf vs.TextMatrix(vs.RowSel, 2) = "SalesReturn" Then
'         inviceNo = vs.TextMatrix(vs.RowSel, 0)
'         Critnote.Show
End If

End Sub

Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then

'If vs.Col = 4 Then
'   SendKeys "{down}"
'End If

If vs.TextMatrix(vs.RowSel, 0) <> "" Then
         inviceNo = vs.TextMatrix(vs.RowSel, 0)
         s1 = 1
         invoice.Show
'ElseIf vs.TextMatrix(vs.RowSel, 2) = "SalesReturn" Then
'         inviceNo = vs.TextMatrix(vs.RowSel, 0)
'         Critnote.Show
End If



End If
End Sub
Sub CalculateTotalDrCr()
On Error Resume Next
Dim Balance As Long
Dim dr1, cr1, prbal
Dim Str
Str = ""
dr1 = 0
cr1 = 0
txtClosing.Text = 0
txtcr.Text = 0
If RS.State = 1 Then RS.close
RS.Open "select Op,drcr from SLEDGER where " & stringyear & " and AgentName='" & cboParty.Text & "'", con
If RS.EOF = False Then
txtOp.Text = Format(RS.Fields(0).value, "0.00")
If UCase(RS.Fields("drcr").value) = UCase("dr") Then
cboop.Text = "Dr"
Else
cboop.Text = "Cr"
End If
Else
txtOp.Text = 0
End If
If cboop.Text = "Dr" Then
dr1 = (Val(txtOp.Text) + Val(vs1.TextMatrix(1, 4)))
cr1 = Val(vs1.TextMatrix(1, 5))
Else
cr1 = (Val(txtOp.Text) + Val(vs1.TextMatrix(1, 5)))
dr1 = Val(vs1.TextMatrix(1, 4))
End If
prbal = dr1 - cr1
If prbal < 0 Then
vs1.TextMatrix(1, 6) = Format(-1 * prbal, "0.00")
vs1.TextMatrix(1, 7) = "Cr"
Else
vs1.TextMatrix(1, 6) = Format(prbal, "0.00")
vs1.TextMatrix(1, 7) = "Dr"
End If
For I = 1 To vs1.rows - 1
If vs1.TextMatrix(I, 0) <> "" Then
txtClosing.Text = (Val(txtClosing.Text) + Val(vs1.TextMatrix(I, 4)))
txtcr.Text = (Val(txtcr.Text) + Val(vs1.TextMatrix(I, 5)))
'-----Balance---------------
If I >= 2 Then
dr1 = Val(vs1.TextMatrix(I, 4))
cr1 = (-1 * Val(vs1.TextMatrix(I, 5)))
bal = dr1 + cr1
If Str = "Cr" Then
bal = prbal + bal
Else
bal = prbal + bal
End If
If bal < 0 Then
vs1.TextMatrix(I, 6) = Format(-1 * bal, "0.00")
vs1.TextMatrix(I, 7) = "Cr"
Else
vs1.TextMatrix(I, 6) = Format(bal, "0.00")
vs1.TextMatrix(I, 7) = "Dr"
End If
prbal = bal
Str = vs1.TextMatrix(I, 7)
End If
'---------------------------
End If
Next
txtClosing.Text = Format(txtClosing.Text, "0.00")
txtcr.Text = Format(txtcr.Text, "0.00")
If cboop.Text = "Dr" Then
txtClosing.Text = Format((CDbl(txtClosing.Text) + CDbl(txtOp.Text)), "0.00")
Else
txtcr.Text = Format((CDbl(txtcr.Text) + CDbl(txtOp.Text)), "0.00")
End If
txtBalance.Text = (Val(txtClosing.Text) - Val(txtcr.Text))
If Val(txtBalance.Text) < 1 Then
txtBalance.Text = (-1 * Val(txtBalance.Text))
closingcr.Text = "Cr"
Else
closingcr.Text = "Dr"
End If
txtBalance.Text = Format(txtBalance.Text, "0.00")

End Sub
Sub SaveDatainTempledger()

Dim d1 As Date



con.Execute "delete  from templedger1 where " & stringyear & ""
For I = 1 To vs1.rows - 1
If vs1.TextMatrix(I, 1) <> "" Then
con.Execute "INSERT INTO  templedger1(Party,dates,Billtype,Bill,Des,Dr,Cr,Balance,drcr,setupid,fyear)  values('" & Trim(cboParty) & "','" & Format(vs1.TextMatrix(I, 2), "MM/dd/yyyy") & "','" & vs1.TextMatrix(I, 0) & "', " & vs1.TextMatrix(I, 1) & ",'" & vs1.TextMatrix(I, 3) & "' ," & vs1.TextMatrix(I, 4) & "," & vs1.TextMatrix(I, 5) & "," & Val(vs1.TextMatrix(I, 6)) & ",'" & vs1.TextMatrix(I, 7) & "'," & setupid & ",'" & session & "')"
End If
Next

Dim ff As New ADODB.Recordset
If ff.State = 1 Then ff.close
ff.Open "select Billtype,bill,dates,des,dr,cr,Balance,drcr from templedger1 where " & stringyear & "  order by dates,bill", con
vs1.rows = ff.RecordCount + 1
For J = 1 To vs1.rows - 1
 If ff.EOF = False Then
     vs1.TextMatrix(J, 0) = ff.Fields(0).value
     vs1.TextMatrix(J, 1) = ff.Fields(1).value
     vs1.TextMatrix(J, 2) = ff.Fields(2).value
     vs1.TextMatrix(J, 3) = ff.Fields(3).value
     vs1.TextMatrix(J, 4) = Format(ff.Fields(4).value, "0.00")
     vs1.TextMatrix(J, 5) = Format(ff.Fields(5).value, "0.00")
     vs1.TextMatrix(J, 6) = Format(ff.Fields(6).value, "0.00")
     ff.MoveNext
 End If
Next

End Sub
Private Sub cboParty_GotFocus()

Dim ph_rs As New ADODB.Recordset
cboParty.BackColor = &HFFFFC0


I = 1
If PopUpValue1 <> "" Then
cboParty.Text = PopUpValue1
End If




End Sub

Private Sub cboParty_KeyDown(KeyCode As Integer, Shift As Integer)

Dim dr, cr As Double

If KeyCode = 27 Then Unload Me


If KeyCode = 113 Then
'-------------------------------
    din_ = False
    value = "select  rep as Reprasentative,Add1 As Address from rep order by rep"
    popuplistModel10 value, CON_blue
    set_focus = True
End If

If KeyCode = 115 Then
'-------------------------------
    din_ = True
    value = "SELECT Shipto,Shipto_City As City,Shipto_district as District,Shipto_States as States FROM INVOICEA_sp where len(Shipto)>0"
    popuplistModel10 value, con
    set_focus = True
End If



If KeyCode = 13 Then
If cboParty.Text = "" Then
  cboParty.SetFocus
  Exit Sub
End If

dataSearchingrid
cmdPrint.Enabled = True



dr = 0
cr = 0

For I = 1 To vs1.rows - 1
  dr = dr + Val(vs1.TextMatrix(I, 4))
  cr = cr + Val(vs1.TextMatrix(I, 5))
Next

drLebel.Caption = Format(dr, "0.00")
CrLebel.Caption = Format(cr, "0.00")
    
'txtdes.SetFocus
    
End If


If KeyCode = 116 Then
vs1.SetFocus
For J = 1 To vs1.rows - 1
   SendKeys "{down}"
   vs1.Row = J
Next
End If


End Sub
Sub dataSearchingrid()

Screen.MousePointer = vbHourglass
I = 1


If PopUpValue1 <> "" Then
vs1.Clear
vs1.rows = 1
fillGrid
End If
If cboParty.Text <> "" Then
SaveDatainTempledger
CalculateTotalDrCr
End If
setWidth
PopUpValue1 = ""
Screen.MousePointer = vbDefault

End Sub
Private Sub cboParty_LostFocus()



cboParty.BackColor = &HFFFFFF
PopUpValue1 = ""
PopUpValue3 = ""
PopUpValue2 = ""

End Sub
Sub DelFunction()
    Dim Del As New ADODB.Recordset
    If Del.State = 1 Then Del.close
    Set Del = con.Execute("delete from ReceiveIssueParty where " & stringyear & " and RecNo=" & txtRecno.Text & "")
End Sub
Private Sub cmdDel_Click()
  
   If MsgBox("Do U Want To Delete ?", vbQuestion + vbYesNo) = vbYes Then
       DelFunction
       fillGrid
       dataSearchingrid
       Call cmdRefresh_Click
       cmdModify.Enabled = False
       cmdDel.Enabled = False
   End If
End Sub
Private Sub cmdMain_Click()
If strledger = "cp" Then
If Val(txtQty.Text) > 0 And txtdes.Text <> "" And cboParty.Text <> "" Then
   If MsgBox("Want To Save & Exit ?", vbQuestion + vbYesNo) = vbYes Then
          SaveMain
          Call cmdRefresh_Click
          fillGrid
          cmdModify.Enabled = False
          cmdDel.Enabled = False
          cboParty.SetFocus
          dataSearchingrid
          Unload Me
          Exit Sub
   End If
End If
End If
If MsgBox("Want To Exit ?", vbQuestion + vbYesNo) = vbYes Then
  Unload Me
End If
End Sub
Sub setWidth()
    
    vs.Clear
    vs.Cols = 8
    
    vs.FormatString = "^Bill No|^Bill Dates|SchoolName|Qty|Rate|>Amount|>Discount|>NetAmount"
    vs.ColWidth(0) = 1000
    vs.ColWidth(1) = 1000
    vs.ColWidth(2) = 4000
    vs.ColWidth(3) = 900
    vs.ColWidth(4) = 900
    vs.ColWidth(5) = 900
    vs.ColWidth(6) = 700
    vs.ColWidth(7) = 900
    
   DoEvents

End Sub
Private Sub cmdModify_Click()


'''''''''''On Error GoTo aa1
''''''''''If MsgBox("Do U Want To Update ?", vbQuestion + vbYesNo) = vbYes Then
'''''''''''DelFunction
''''''''''CON.Execute "update ReceiveIssueParty set Dr=0,cr=0 where " & stringyear & " and RecNo=" & txtRecno.Text & ""
''''''''''
'''''''''''------------------------
''''''''''Set RS = New ADODB.Recordset
''''''''''RS.Open "select * from ReceiveIssueParty where " & stringyear & " and RecNo=" & txtRecno.Text & "", CON, adOpenDynamic, adLockOptimistic
''''''''''If RS.EOF = False Then
'''''''''''maxId
'''''''''''RS.AddNew
''''''''''RS.Fields("RecNo").value = txtRecno.Text
''''''''''RS.Fields("Dates").value = RecDates.value
''''''''''RS.Fields("PartyName").value = cboParty.Text
''''''''''RS.Fields("Particullar").value = txtdes.Text
''''''''''If Receive.value = True Then
''''''''''RS.Fields("Dr").value = Val(txtQty.Text)
''''''''''Else
''''''''''RS.Fields("Cr").value = Val(txtQty.Text)
''''''''''End If
''''''''''RS.update
''''''''''End If
'''''''''''------------------------
''''''''''
'''''''''''SaveMain
''''''''''
''''''''''
''''''''''fillGrid
''''''''''CalculateTotalDrCr
''''''''''setwidth
''''''''''Call cmdRefresh_Click
''''''''''vs1.SetFocus
''''''''''For I = 1 To vs1.Rows - 1
''''''''''SendKeys "{down}"
''''''''''Next
''''''''''
''''''''''cmdModify.Enabled = False
''''''''''cmdDel.Enabled = False
''''''''''End If
'Exit Sub
'aa1:
'MsgBox "Record not Save !!", vbCritical
End Sub
Private Sub cmdRefresh_Click()
 
 
 Dim o As Object
 txtQty.Text = ""
 set_focus = False
 maxId
 cmdModify.Enabled = False
 cmdDel.Enabled = False
 cmdSave.Enabled = True
 
 
 Screen.MousePointer = vbDefault
 bb2 = False

End Sub
Private Sub cmdSave_Click()

'''''''''On Error GoTo aa:
'''''''''
'''''''''
'''''''''
'''''''''If cboParty.Text = "" Then
'''''''''MsgBox "Please Select Party Name !!", vbInformation
'''''''''Exit Sub
'''''''''End If
'''''''''
'''''''''If txtQty.Text = "" Then
'''''''''MsgBox "Please Enter Amount!!", vbInformation
'''''''''txtQty.SetFocus
'''''''''Exit Sub
'''''''''End If
'''''''''
'''''''''
'''''''''If RS.State = 1 Then RS.close
'''''''''RS.Open "select * from pass where pass='" & cp & "'", CON
'''''''''If RS.EOF = True Then
'''''''''MsgBox "Enter Valid Password !!", vbInformation
'''''''''Exit Sub
'''''''''End If
'''''''''
'''''''''If MsgBox("Do U Want To Save ?", vbInformation + vbYesNo) = vbYes Then
'''''''''aa1:
'''''''''SaveMain
'''''''''
'''''''''cboParty.SetFocus
'''''''''
'''''''''Call cmdRefresh_Click
'''''''''fillGrid
'''''''''
'''''''''cmdModify.Enabled = False
'''''''''cmdDel.Enabled = False
''''''''''----------------
'''''''''dataSearchingrid
''''''''''---------------
'''''''''
'''''''''
'''''''''End If
'''''''''Exit Sub
'''''''''aa:
'''''''''maxId
'''''''''GoTo aa1

End Sub
Sub SaveMain()
   
'''''''   maxId
'''''''    Set RS = New ADODB.Recordset
'''''''    RS.Open "select * from ReceiveIssueParty where " & stringyear & " and RecNo=" & txtRecno.Text & "", CON, adOpenDynamic, adLockOptimistic
'''''''    If RS.EOF = True Then
'''''''       maxId
'''''''       RS.AddNew
'''''''       RS.Fields("RecNo").value = txtRecno.Text
'''''''       RS.Fields("Dates").value = RecDates.value
'''''''       RS.Fields("PartyName").value = cboParty.Text
'''''''       RS.Fields("Particullar").value = txtdes.Text
'''''''       If Receive.value = True Then
'''''''          RS.Fields("Dr").value = Val(txtQty.Text)
'''''''        Else
'''''''          RS.Fields("Cr").value = Val(txtQty.Text)
'''''''       End If
'''''''
'''''''    RS.update
'''''''    End If
End Sub
Sub search()
''''' If set_focus = True Then Exit Sub
''''' On Error Resume Next
'''''
'''''
'''''
'''''    If rss.State = 1 Then rss.close
'''''    rss.Open "select * from sledger where " & stringyear & " and AgentName=" & txtParty.Text & "", CON, adOpenDynamic, adLockOptimistic
'''''    If rss.EOF = 1 Then
'''''       txtRem.Text = RS.Fields("PartyRemarks").value & ""
'''''    End If
'''''
'''''
'''''
'''''
''''' If vs1.TextMatrix(vs1.RowSel, 0) = "J" Then
'''''    If RS.State = 1 Then RS.close
'''''    RS.Open "select * from ReceiveIssueParty where " & stringyear & " and RecNo=" & vs1.TextMatrix(vs1.RowSel, 1) & "", CON, adOpenDynamic, adLockOptimistic
'''''    If RS.EOF = False Then
'''''       txtRecno.Text = RS.Fields("RecNo").value
'''''       RecDates.value = RS.Fields("Dates").value
'''''       cboParty.Text = RS.Fields("PartyName").value
'''''       txtdes.Text = RS.Fields("Particullar").value
'''''
'''''
'''''       If RS.Fields("Dr").value > 0 Then
'''''          Receive.value = True
'''''          txtQty.Text = RS.Fields("Dr").value
'''''        Else
'''''          Issue.value = True
'''''          txtQty.Text = RS.Fields("Cr").value
'''''       End If
'''''      End If
'''''   cmdSave.Enabled = False
'''''   cmdModify.Enabled = True
'''''   cmdDel.Enabled = True
'''''  Else
'''''   cmdModify.Enabled = False
'''''   cmdDel.Enabled = False
'''''   cmdSave.Enabled = True
'''''   txtdes.Text = ""
'''''   txtQty.Text = ""
'''''  End If
End Sub
Private Sub cmdSearch_Click()
Frame1.Visible = True
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  
  
  
  If KeyCode = 27 Then
   If cboPartyList.Visible = True Then
      cboPartyList.Visible = False
      Exit Sub
   End If
  End If
  
  
  
  If KeyCode = 116 Then
  If bb2 = False Then
    vs1.SetFocus
    For I = 1 To vs1.rows - 1
    SendKeys "{down}"
    Next
    bb2 = True
  Else
    Call cmdRefresh_Click
    cboParty.SetFocus
    bb2 = False
  End If
  Exit Sub
  End If
  
  
  
  If KeyCode = 112 Then
     txtdes.SetFocus
     Exit Sub
  End If
   If KeyCode = 27 Then
        'If RS.State = 1 Then RS.close
        'RS.Open "select * from pass where pass='" & cp & "'", CON
        'If RS.EOF = True Then
        '  If MsgBox("Want To Exit ?", vbQuestion + vbYesNo) = vbYes Then
        '   Unload Me
        '   End If
        'Exit Sub
        'End If
        
        If Val(txtQty.Text) > 0 And txtdes.Text <> "" And cboParty.Text <> "" Then
        If MsgBox("Want To Save & Exit ?", vbQuestion + vbYesNo) = vbYes Then
            SaveMain
            Call cmdRefresh_Click
            fillGrid
            cmdModify.Enabled = False
            cmdDel.Enabled = False
            cboParty.SetFocus
            dataSearchingrid
            Unload Me
            Exit Sub
        End If
        End If
      If MsgBox("Want To Exit ?", vbQuestion + vbYesNo) = vbYes Then
         Unload Me
       End If
      ElseIf KeyCode = 13 Then
      ElseIf KeyCode = 113 Then
         kk = False
   End If
End Sub
Sub fillGrid()
    
   
    '==============
    SearchFa
    '==============
    setWidth
End Sub
Sub maxId()
  Dim rr As New ADODB.Recordset
  Set rr = New ADODB.Recordset
  rr.Open "select max(RecNo) from ReceiveIssueParty where " & stringyear & " ", con
  If IsNull(rr.Fields(0).value) Then
     txtRecno.Text = 1
     Else
     txtRecno.Text = rr.Fields(0).value + 1
  End If
End Sub

Private Sub Todate_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
    from_date = FromDate.value
    to_date = toDate.value
    fillGrid
    Frame1.Visible = False
 End If
End Sub

Private Sub txtQty_GotFocus()
   txtQty.BackColor = &HFFFFC0
End Sub

Private Sub txtQty_LostFocus()
  txtQty.BackColor = &HFFFFFF
End Sub
Private Sub txtRecno_KeyPress(KeyAscii As Integer)
   On Error Resume Next
  

  
  If KeyAscii = 13 Then
  
     If RS.State = 1 Then RS.close
     RS.Open "select * from receiveissueparty where " & stringyear & " and recno=" & txtRecno.Text & "", con
     If RS.EOF = False Then
      cboParty.Text = RS!partyname
      PopUpValue3 = cboParty.Text
      
      RecDates.value = RS.Fields("Dates").value
      txtdes.Text = RS.Fields("Particullar").value
      'txtRem.Text = RS.Fields("Remarks").Value
      If RS.Fields("Dr").value > 0 Then
          Receive.value = True
          txtQty.Text = RS.Fields("Dr").value
      Else
          Issue.value = True
          txtQty.Text = RS.Fields("Cr").value
      End If
      dataSearchingrid
     Else
       vs1.Clear
       setWidth
       txtQty.Text = ""
       txtdes.Text = ""
       cboParty.Text = ""
       txtOp.Text = ""
       txtBalance.Text = ""
     End If
  End If
End Sub
Private Sub txtSlipNo_GotFocus()
 txtSlipNo.BackColor = &HFFFFC0
End Sub
Private Sub txtSlipNo_LostFocus()
txtSlipNo.BackColor = &HFFFFFF
End Sub
Private Sub vs1_Click()
 search
End Sub
Private Sub vs1_DblClick()
set_focus = False
End Sub




Private Sub vs1_SelChange()
 search
End Sub

Private Sub vsop_Click()
If vsop.Col = 0 Then
   vsop.Editable = flexEDNone
ElseIf vsop.Col = 1 Then
   vsop.Editable = flexEDNone
ElseIf vsop.Col = 2 Then
   vsop.Editable = flexEDKbdMouse
ElseIf vsop.Col = 3 Then
   vsop.Editable = flexEDKbdMouse
End If
  

End Sub

Private Sub vsop_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    I = 1
    cboParty.Text = vsop.TextMatrix(vsop.RowSel, 1)
    PopUpValue2 = cboParty.Text
    vs1.Clear
    fillGrid
    SaveDatainTempledger
    CalculateTotalDrCr
    setWidth
    PopUpValue1 = ""
    Opening.Tab = 1
End If
End Sub


