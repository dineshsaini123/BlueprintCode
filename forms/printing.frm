VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmbill 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Printing Order"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7650
   ScaleWidth      =   14190
   Begin VB.TextBox TextPaperSize 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   10080
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtPcode 
      Height          =   285
      Left            =   10860
      TabIndex        =   27
      Top             =   1080
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.ComboBox binder_name 
      Height          =   315
      Left            =   7020
      TabIndex        =   2
      Top             =   240
      Width           =   3975
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   4095
      Left            =   60
      TabIndex        =   4
      Top             =   2100
      Width           =   13875
      _cx             =   24474
      _cy             =   7223
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   12582847
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483645
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
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   15
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"printing.frx":0000
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
      Begin VB.TextBox txtNote 
         Appearance      =   0  'Flat
         Height          =   225
         Left            =   13440
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   3960
         Width           =   195
      End
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   11640
      Top             =   6360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFD7AE&
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   6300
      Width           =   8865
      Begin VB.CommandButton cmdPrint_Slip 
         BackColor       =   &H00FFFFFF&
         Caption         =   "P&rint Slip"
         Height          =   585
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   165
         Width           =   1065
      End
      Begin VB.CommandButton cancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cancel"
         Height          =   585
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   165
         Width           =   1065
      End
      Begin VB.CommandButton ok 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         Height          =   585
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   165
         Width           =   1065
      End
      Begin VB.CommandButton Printcmd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Print"
         Height          =   585
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   165
         Width           =   1065
      End
      Begin VB.CommandButton delete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         Height          =   585
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   165
         Width           =   1065
      End
      Begin VB.CommandButton Edit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Edit"
         Height          =   585
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   165
         Width           =   1065
      End
      Begin VB.CommandButton Add 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add"
         Height          =   585
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   165
         Width           =   1065
      End
      Begin VB.CommandButton CommandQuit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "E&xit"
         Height          =   585
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   165
         Width           =   1065
      End
      Begin VB.CommandButton cmdBillCancel 
         Caption         =   "&Order Cancel"
         Height          =   585
         Left            =   12240
         TabIndex        =   14
         Top             =   165
         Width           =   75
      End
   End
   Begin VB.TextBox txtOrdNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1020
      MaxLength       =   50
      TabIndex        =   0
      Top             =   240
      Width           =   1410
   End
   Begin MSComCtl2.DTPicker txtOrdDate 
      Height          =   315
      Left            =   3840
      TabIndex        =   1
      Top             =   255
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   556
      _Version        =   393216
      CalendarBackColor=   16776960
      Format          =   65208321
      CurrentDate     =   38372
   End
   Begin VB.Label lblPaper_det 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9540
      TabIndex        =   28
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Total :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9120
      TabIndex        =   18
      Top             =   6420
      Width           =   675
   End
   Begin VB.Label lblAdd 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   7020
      TabIndex        =   17
      Top             =   660
      Width           =   3915
   End
   Begin VB.Label txtTSheet 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   10500
      TabIndex        =   16
      Top             =   6360
      Width           =   825
   End
   Begin VB.Label txtTReam 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   9660
      TabIndex        =   15
      Top             =   6360
      Width           =   825
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Press F4 Key to delete a record"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   240
      Left            =   120
      TabIndex        =   13
      Top             =   7260
      Width           =   2805
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Press F2 Search For English"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   -240
      TabIndex        =   12
      Top             =   5340
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Press F3 Search For"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   -360
      TabIndex        =   11
      Top             =   5340
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Paper Size"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   7980
      TabIndex        =   10
      Top             =   1140
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Orderl Date"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   315
      Left            =   2880
      TabIndex        =   9
      Top             =   270
      Width           =   1380
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Printer Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   5910
      TabIndex        =   8
      Top             =   285
      Width           =   1470
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Order No."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   90
      TabIndex        =   7
      Top             =   270
      Width           =   1155
   End
End
Attribute VB_Name = "frmbill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public mrpeat As Boolean
Public orderchk As Boolean
Public partchk As Boolean
Public bindchk As Boolean
Dim RS As New ADODB.Recordset
Dim page_sum As Double
Public gridchk As Boolean
Dim sheet, ream, westage
Public mode As String
Dim flag As Boolean
Sub paperstatementcalc()

''''''calculate the opbalance paper start
'''''Dim Trptrs As New ADODB.Recordset
'''''Dim rsrec As New ADODB.Recordset
'''''Dim rsdel As New ADODB.Recordset
'''''Dim rsbill As New ADODB.Recordset
'''''Dim rsbillmast As New ADODB.Recordset
'''''Dim rsbillTRANS As New ADODB.Recordset
'''''Dim rsbillTRANSprint As New ADODB.Recordset
'''''Dim rsbill1 As New ADODB.Recordset
'''''Dim rsextra As New ADODB.Recordset
'''''Set rsrec = New ADODB.Recordset
'''''Set rsdel = New ADODB.Recordset
'''''Set rsbill = New ADODB.Recordset
'''''Set rsbillmast = New ADODB.Recordset
'''''Set rsbillTRANS = New ADODB.Recordset
'''''Set rsbillTRANSprint = New ADODB.Recordset
'''''Set rsbill1 = New ADODB.Recordset
'''''Set rsextra = New ADODB.Recordset
'''''' find the previouse bill no start
'''''Dim rsbilmast As ADODB.Recordset
'''''Set rsbilmast = New ADODB.Recordset
''''''fid = Trim(Me.Textfirmid.Text)
'''''If Me.pscustomerid.Text = "" Or TextPaperSize = "" Then
'''''   Exit Sub
'''''End If
'''''
''''''sq = "Select * from billmaster where firm_id = '" + Me.Textfirmid.Text + "' and pscustomerid  = '" + Me.pscustomerid.Text + "' and papersize1 = '" + TextPaperSize + "' and cint(bill_id) <= " + bill_no + " order by cint(bill_id) desc"
'''''sq = "Select * from billmaster where pscustomerid  = '" + Me.pscustomerid.Text + "' and papersize1 = '" + TextPaperSize + "' and pstatementno <= " + Str(pstno) + " order by cint(pstatementno) desc"
'''''rsbilmast.Open sq, CON, adOpenKeyset, adLockReadOnly
'''''If rsbilmast.RecordCount > 1 Then
'''''    If Not rsbilmast.EOF Then
'''''        rsbilmast.MoveNext
'''''        If rsbilmast.Fields("pstatementno") <> "0" Then
'''''        previousbillno.Text = rsbilmast.Fields("bill_id")
'''''        previousbilldate = rsbilmast.Fields("dat")
'''''        Else
'''''        previousbillno.Text = "1"
'''''        previousbilldate = CDate("01/01/2000")
'''''        End If
'''''    End If
'''''Else
''''''Exit Sub
'''''    If Not rsbilmast.BOF Then
'''''    rsbilmast.MoveFirst
'''''    previousbillno.Text = rsbilmast.Fields("pstatementno")
'''''    previousbilldate = rsbilmast.Fields("dat")
'''''    End If
'''''End If
''''''rsextra.Open "select * from extrainfo", con, adOpenDynamic, adLockOptimistic
''''''con.Execute "Delete * from extrainfo"
''''''rsextra.Open "select * from extrainfo", con, adOpenDynamic, adLockOptimistic
''''''rsextra.AddNew
''''''rsextra.Fields("prvbno") = previousbillno.Text
''''''rsextra.Fields("prvbdate") = previousbilldate.value
''''''rsextra.Update
''''''rsextra.Close
'''''' find the previouse bill no end
'''''Balance = 0
'''''opbalancereams = 0
'''''opbalancesheets = 0
'''''CON.Execute "Delete * from billtransprint"
''''''sq = "SELECT sum(reams) as recreams,sum(sheets) as recsheets  FROM paperstatement   WHERE firm_ID = '" + Textfirmid + "' and  date1 <= cdate('" + Trim(DTPicker1) + "') and recdel='R' and papersize1 = '" + textpapersize + "' AND PSCUSTOMERID = '" + pscustomerid + "' AND (VAL(BILL_ID) < " + Str(Val(Me.bill_no)) + " AND BILL_ID <> '0')"
'''''
'''''sq = "SELECT sum(reams) as recreams,sum(sheets) as recsheets  FROM paperstatement WHERE date1 <= cdate('" + Trim(DTPicker1) + "') and recdel='R' and papersize1 = '" + TextPaperSize + "' AND PSCUSTOMERID = '" + pscustomerid + "' AND trim(BILL_ID) <> '0' and pstatementno < " + Str(pstno)
'''''If rsrec.State = 1 Then rsrec.close
'''''rsrec.Open sq, CON, adOpenDynamic, adLockReadOnly, adCmdText
''''''rsdel.Open "SELECT sum(reams) as delreams1, sum(sheets) as delsheets1 FROM paperstatement   WHERE firm_ID = '" + Textfirmid + "' and date1 < cdate('" + Trim(DTPicker1) + "') and recdel='D' and papersize1 = '" + TextPaperSize + "' AND PSCUSTOMERID = '" + pscustomerid + "'AND (VAL(BILL_ID) < " + Str(Val(Me.bill_no)) + " AND BILL_ID <> '0')", con, adOpenDynamic, adLockReadOnly, adCmdText
'''''
'''''rsdel.Open "SELECT sum(reams) as delreams1, sum(sheets) as delsheets1 FROM paperstatement   WHERE  date1 <= cdate('" + Trim(DTPicker1) + "') and recdel='D' and papersize1 = '" + TextPaperSize + "' AND PSCUSTOMERID = '" + pscustomerid + "' AND trim(BILL_ID) <> '0' and pstatementno < " + Str(pstno), CON, adOpenDynamic, adLockReadOnly, adCmdText
'''''If rsbill1.State = 1 Then rsbill1.close
''''''rsbill1.Open "SELECT sum(reams+wreams) as billreams1, sum(sheets+wsheets) as billsheets1 FROM billtrans  WHERE firm_ID = '" + Textfirmid + "' and billdate <= cdate('" + Trim(DTPicker1) + "') and papersize1 = '" + TextPaperSize + "' AND PSCUSTOMERID = '" + pscustomerid + "' AND VAL(BILL_ID) < " + Str(Val(Me.bill_no)), con, adOpenDynamic, adLockReadOnly, adCmdText
'''''
'''''rsbill1.Open "SELECT sum(reams+wreams) as billreams1, sum(sheets+wsheets) as billsheets1 FROM billtrans  WHERE  billdate <= cdate('" + Trim(DTPicker1) + "') and papersize1 = '" + TextPaperSize + "' AND PSCUSTOMERID = '" + pscustomerid + "' and (val(pstatementno) > 0 and val(pstatementno) < " + pstno + ") and bill_id <> '" + Trim(Str(Val(Me.bill_no))) + "'", CON, adOpenDynamic, adLockReadOnly, adCmdText
'''''     DoEvents
'''''    If rsrec!recsheets > 0 Then opbalancesheets = rsrec!recsheets
'''''    TMPreams = 0
'''''    TMPsheets = 0
'''''    If rsdel!delsheets1 > 0 Then opbalancesheets = opbalancesheets - rsdel!delsheets1
'''''    If rsbill1!billsheets1 > 0 Then opbalancesheets = opbalancesheets - rsbill1!billsheets1
'''''    TMPsheets = opbalancesheets
'''''     If rsrec!recreams > 0 Then opbalancereams = rsrec!recreams
'''''    If rsdel!delreams1 > 0 Then opbalancereams = opbalancereams - rsdel!delreams1
'''''    If rsbill1!billreams1 > 0 Then opbalancereams = opbalancereams - rsbill1!billreams1
'''''    If opbalancesheets < 0 Then
'''''        TMPreams = opbalancereams + Int(opbalancesheets / 500)
'''''        TMPsheets = opbalancesheets - (Int(opbalancesheets / 500) * 500)
'''''    Else
'''''    If opbalancesheets > 499 Then
'''''        TMPreams = opbalancereams + Int(opbalancesheets / 500)
'''''        TMPsheets = opbalancesheets - (Int(opbalancesheets / 500) * 500)
'''''        Else
'''''        TMPreams = opbalancereams
'''''        TMPsheets = opbalancesheets
'''''            End If
'''''    End If
'''''    If TMPreams < 0 Then
'''''        TMPreams = TMPreams + 1
'''''        TMPsheets = TMPsheets - 500
'''''       End If
'''''       ' tmpreams = opbalancereams
'''''       ' tmpsheets = opbalancesheets
'''''
'''''    opbalancereams = TMPreams
'''''    opbalancesheets = TMPsheets
'''''    'calculate the opbalance paper start
'''''     'CON.Execute "insert into billtransprint select * from billtrans where firm_id = '" + Textfirmid + "' and bill_id = '" + Me.bill_no + "'"
'''''    CON.Execute "insert into billtransprint(bill_id,billdate,papersize1,ft,recdel ,reams,sheets,pscustomerid,FIRM_id) select paperstatement.Challan_no as bill_id,paperstatement.date1 as billdate, paperstatement.papersize1 as papersize1, paperstatement.ft as ft, paperstatement.recdel as recdel , paperstatement.reams as reams, paperstatement.sheets as sheets, paperstatement.pscustomerid as pscustomerid, paperstatement.FIRM_id as FIRM_id from paperstatement where papersize1 ='" + TextPaperSize + "' AND PSCUSTOMERID = '" + pscustomerid + "'  and date1 <=cdate('" + Trim(DTPicker1) + "') AND (trim(BILL_ID) = '" + Trim(Str(Me.bill_no)) + "' OR BILL_ID = '0') "
'''''    'CON.Execute "insert into billtransprint select * from billtrans where bill_id = '" + Trim(Me.bill_no) + "'"
'''''    If previousbilldate = DTPicker1 Then
'''''      '   con.Execute "insert into billtransprint select paperstatement.Challan_no as bill_id,paperstatement.date1 as billdate, paperstatement.papersize1 as papersize1, paperstatement.ft as ft, paperstatement.recdel as recdel , paperstatement.reams as reams, paperstatement.sheets as sheets, paperstatement.pscustomerid as pscustomerid, paperstatement.FIRM_id as FIRM_id from paperstatement where firm_ID = '" + Textfirmid + "' and papersize1 ='" + TextPaperSize + "' AND PSCUSTOMERID = '" + pscustomerid + "'  and date1 <=cdate('" + Trim(DTPicker1) + "') AND (BILL_ID = '" + Str(Me.bill_no) + "' OR BILL_ID = '0') "
'''''
'''''      ''*********** Dinesh Change ****************
'''''
'''''          CON.Execute "insert into billtransprint(bill_id,billdate,papersize1,ft,recdel ,reams,sheets,pscustomerid,FIRM_id) select paperstatement.Challan_no as bill_id,paperstatement.date1 as billdate, paperstatement.papersize1 as papersize1, paperstatement.ft as ft, paperstatement.recdel as recdel , paperstatement.reams as reams, paperstatement.sheets as sheets, paperstatement.pscustomerid as pscustomerid, paperstatement.FIRM_id as FIRM_id from paperstatement where papersize1 ='" + TextPaperSize + "' AND PSCUSTOMERID = '" + pscustomerid + "'  and date1 <=cdate('" + Trim(DTPicker1) + "') AND (trim(BILL_ID) = '" + Trim(Str(Me.bill_no)) + "' OR BILL_ID = '0') "
'''''
'''''    Else
'''''           '  con.Execute "insert into billtransprint select paperstatement.Challan_no as bill_id,paperstatement.date1 as billdate, paperstatement.papersize1 as papersize1, paperstatement.ft as ft, paperstatement.recdel as recdel , paperstatement.reams as reams, paperstatement.sheets as sheets, paperstatement.pscustomerid as pscustomerid, paperstatement.FIRM_id as FIRM_id from paperstatement where firm_ID = '" + Textfirmid + "' and papersize1 ='" + TextPaperSize + "' AND PSCUSTOMERID = '" + pscustomerid + "' and (date1 >cdate('" + Trim(previousbilldate) + "') and date1 <=cdate('" + Trim(DTPicker1) + "')) AND (BILL_ID = '" + Str(Trim(Me.bill_no)) + "' OR BILL_ID = '0') "
'''''         CON.Execute "insert into billtransprint(bill_id,billdate,papersize1,ft,recdel ,reams,sheets,pscustomerid,FIRM_id) select paperstatement.Challan_no as bill_id,paperstatement.date1 as billdate, paperstatement.papersize1 as papersize1, paperstatement.ft as ft, paperstatement.recdel as recdel , paperstatement.reams as reams, paperstatement.sheets as sheets, paperstatement.pscustomerid as pscustomerid, paperstatement.FIRM_id as FIRM_id from paperstatement where papersize1 ='" + TextPaperSize + "' AND PSCUSTOMERID = '" + pscustomerid + "' and (date1 >cdate('" + Trim(previousbilldate) + "') and date1 <=cdate('" + Trim(DTPicker1) + "')) AND (trim(BILL_ID) = '" + Trim(Str(Me.bill_no)) + "' OR BILL_ID = '0') "
'''''    End If
'''''     DoEvents
'''''' update the masters start
'''''tmpreamsrec = 0
'''''tmpsheetsrec = 0
'''''tmpreamsdel = 0
'''''tmpsheetsdel = 0
'''''If rsrec.State = 1 Then rsrec.close
'''''rsrec.Open "SELECT sum(reams) as recreams,sum(sheets) as recsheets  FROM  billtransprint   WHERE recdel='R' and papersize1 = '" + TextPaperSize + "'", CON, adOpenDynamic, adLockReadOnly, adCmdText
'''''If rsdel.State = 1 Then rsdel.close
'''''rsdel.Open "SELECT sum(reams) as delreams,sum(sheets) as delsheets  FROM  billtransprint   WHERE recdel='D' and papersize1 = '" + TextPaperSize + "'", CON, adOpenDynamic, adLockReadOnly, adCmdText
'''''DoEvents
'''''If rsrec!recsheets > 499 Then
''''''    tmpreamsrec = rsrec!recreams + Int(rsrec!sheets / 500)
''''' '   tmpsheetsrec = rsrec!recsheets - (Int(rsrec!sheets / 500) * 500)
'''''    tmpreamsrec = rsrec!recreams + Int(rsrec!recsheets / 500)
'''''    tmpsheetsrec = rsrec!recsheets - (Int(rsrec!recsheets / 500) * 500)
'''''Else
'''''    If rsrec!recreams > 0 Then tmpreamsrec = rsrec!recreams
'''''    If rsrec!recsheets > 0 Then tmpsheetsrec = rsrec!recsheets
'''''End If
'''''If rsdel!delsheets > 499 Then
'''''    'tmpreamsdel = rsdel!delreams + Int(rsdel!sheets / 500)
'''''    'tmpsheetsdel = rsdel!delsheets - (Int(rsdel!sheets / 500) * 500)
'''''    tmpreamsdel = rsdel!delreams + Int(rsdel!delsheets / 500)
'''''    tmpsheetsdel = rsdel!delsheets - (Int(rsdel!delsheets / 500) * 500)
'''''Else
'''''    If rsdel!delreams > 0 Then tmpreamsdel = rsdel!delreams
'''''    If rsdel!delsheets > 0 Then tmpsheetsdel = rsdel!delsheets
'''''End If
'''''   '  rsbillmast.Open "select * from billmaster where firm_id = '" + Textfirmid + "' and bill_id = '" + Me.bill_no + "'", con, adOpenDynamic, adLockPessimistic
'''''     rsbillmast.Open "select * from billmaster where  bill_id = '" + Me.bill_no + "' AND firm_id = '" & frmbill.Textfirmid & "'", CON, adOpenDynamic, adLockPessimistic
'''''     If rsbillmast.RecordCount > 0 Then
'''''     rsbillmast!totaldelreams = tmpreamsdel
'''''     rsbillmast!totaldelsheets = tmpsheetsdel
'''''     rsbillmast!totalrecreams = tmpreamsrec
'''''     rsbillmast!totalrecsheets = tmpsheetsrec
'''''     rsbillmast!openingreams = opbalancereams
'''''     rsbillmast!openingsheets = opbalancesheets
'''''     rsbillmast!previousebillid = previousbillno.Text
'''''     rsbillmast!previousebilldate = CDate(previousbilldate)
'''''     balreams = 0
'''''     balsheets = 0
'''''     balsheets = opbalancesheets + tmpsheetsrec - tmpsheetsdel - rsbillmast!totalsht
'''''     balreams = opbalancereams + tmpreamsrec - tmpreamsdel - rsbillmast!TOTALRIM
'''''End If
'''''     If balsheets < 0 Then
'''''        TMPreamsBAL = balreams + Int(balsheets / 500)
'''''        TMPsheetsBAL = balsheets - (Int(balsheets / 500) * 500)
'''''    Else
'''''        If balsheets > 499 Then
'''''            TMPreamsBAL = balreams + Int(balsheets / 500)
'''''            TMPsheetsBAL = balsheets - (Int(balsheets / 500) * 500)
'''''        Else
'''''
'''''
'''''        TMPreamsBAL = balreams
'''''        TMPsheetsBAL = balsheets
'''''        End If
'''''    End If
'''''          If TMPreamsBAL < 0 Then
'''''        TMPreamsBAL = TMPreamsBAL + 1
'''''        TMPsheetsBAL = TMPsheetsBAL - 500
'''''        End If
'''''       rsbillmast!balancereams = TMPreamsBAL
'''''       rsbillmast!balancesheets = TMPsheetsBAL
'''''       rsbillmast.update
'''''     sq = "select * from billtrans where bill_id = '" + Trim(Me.bill_no) + "' AND firm_id = '" & frmbill.Textfirmid + "'"
'''''    rsbillTRANS.Open sq, CON, adOpenDynamic, adLockPessimistic
'''''
'''''     'con.Execute "insert into billtransprint select * from billtrans where bill_id = '" + Trim(Me.bill_no) + "'"
'''''        Do While Not rsbillTRANS.EOF
'''''     rsbillTRANS!totaldelreams = tmpreamsdel
'''''     rsbillTRANS!totaldelsheets = tmpsheetsdel
'''''     rsbillTRANS!totalrecreams = tmpreamsrec
'''''     rsbillTRANS!totalrecsheets = tmpsheetsrec
'''''     rsbillTRANS!balancereams = TMPreamsBAL
'''''     rsbillTRANS!balancesheets = TMPsheetsBAL
'''''     rsbillTRANS!TOTALRIM = rsbillmast!TOTALRIM
'''''     rsbillTRANS!totalsht = rsbillmast!totalsht
'''''     rsbillTRANS!openingreams = opbalancereams
'''''     rsbillTRANS!openingsheets = opbalancesheets
'''''     rsbillTRANS!previousebillid = previousbillno.Text
'''''     rsbillTRANS!previousebilldate = CDate(previousbilldate)
'''''     rsbillTRANS.update
'''''     If Not rsbillTRANS.EOF Then
'''''        rsbillTRANS.MoveNext
'''''     End If
'''''Loop
'''''
'''''
'''''
'''''
'''''
'''''
'''''     rsbillTRANSprint.Open "select * from BILLTRANSPRINT ", CON, adOpenDynamic, adLockPessimistic
'''''     Do While Not rsbillTRANSprint.EOF
'''''     rsbillTRANSprint!totaldelreams = tmpreamsdel
'''''     rsbillTRANSprint!totaldelsheets = tmpsheetsdel
'''''     rsbillTRANSprint!totalrecreams = tmpreamsrec
'''''     rsbillTRANSprint!totalrecsheets = tmpsheetsrec
'''''     rsbillTRANSprint!balancereams = TMPreamsBAL
'''''     rsbillTRANSprint!balancesheets = TMPsheetsBAL
'''''     rsbillTRANSprint!TOTALRIM = rsbillmast!TOTALRIM
'''''     rsbillTRANSprint!totalsht = rsbillmast!totalsht
'''''     rsbillTRANSprint!openingreams = opbalancereams
'''''     rsbillTRANSprint!openingsheets = opbalancesheets
'''''     rsbillTRANSprint!previousebillid = previousbillno.Text
'''''     rsbillTRANSprint!previousebilldate = CDate(previousbilldate)
'''''     rsbillTRANSprint.update
'''''     If Not rsbillTRANSprint.EOF Then
'''''        rsbillTRANSprint.MoveNext
'''''     End If
'''''Loop
''''''update the masters close
''''' If rsbill1.State = 1 Then rsbill1.close
''''' If previousbilldate = DTPicker1 Then
''''' '         rsbill1.Open "select * from paperstatement where firm_ID = '" + Textfirmid + "' and papersize1 ='" + TextPaperSize + "' AND PSCUSTOMERID = '" + pscustomerid + "'  and date1 <=cdate('" + Trim(DTPicker1) + "') AND (BILL_ID = '" + Str(Me.bill_no) + "' OR BILL_ID = '0') ", con, adOpenDynamic, adLockPessimistic
'''''
'''''         rsbill1.Open "select * from paperstatement where  papersize1 ='" + TextPaperSize + "' AND PSCUSTOMERID = '" + pscustomerid + "'  and date1 <=cdate('" + Trim(DTPicker1) + "') AND (trim(BILL_ID) = '" + Trim(Str(Me.bill_no)) + "' OR BILL_ID = '0') ", CON, adOpenDynamic, adLockPessimistic
'''''    Else
'''''        'rsbill1.Open "select * from paperstatement where firm_ID = '" + Textfirmid + "' and papersize1 ='" + TextPaperSize + "' AND PSCUSTOMERID = '" + pscustomerid + "' and (date1 >cdate('" + Trim(previousbilldate) + "') and date1 <=cdate('" + Trim(DTPicker1) + "')) AND (BILL_ID = '" + Str(Me.bill_no) + "' OR BILL_ID = '0') ", con, adOpenDynamic, adLockPessimistic
'''''
'''''         rsbill1.Open "select * from paperstatement where papersize1 ='" + TextPaperSize + "' AND PSCUSTOMERID = '" + pscustomerid + "' and (date1 >cdate('" + Trim(previousbilldate) + "') and date1 <=cdate('" + Trim(DTPicker1) + "')) AND (trim(BILL_ID) = '" + Trim(Str(Me.bill_no)) + "' OR BILL_ID = '0') ", CON, adOpenDynamic, adLockPessimistic
'''''    End If
'''''If rsbill1.RecordCount > 0 Then
'''''     Do While Not rsbill1.EOF
'''''     rsbill1!bill_id = Str(Trim(bill_no))
'''''     rsbill1!pstatementno = pstno
'''''     rsbill1.update
'''''     If Not rsbill1.EOF Then
'''''      rsbill1.MoveNext
'''''     End If
'''''Loop
'''''End If
''''''frmbill.cr1.ReportFileName = App.Path & "\paperstat.rpt"
'''''tmpnewfile = App.Path & "\firmlogo.tif"
'''''        tmpfile = App.Path & "\" & Trim(firmpictfilename)
''''''kamal        FileCopy tmpfile, tmpnewfile
'''''        DoEvents
'''''        DoEvents
'''''        DoEvents
'''''        DoEvents
End Sub
Sub try()
gridchk = True
vs.Row = 1
vs.Col = 8
If vs.Text = "" Then
gridchk = False
vs.SetFocus
End If
End Sub
Sub part()
partchk = True
 'If Trim(Me.party_id.Text = "") Then
 'MsgBox "Please fill the Party "
 'Me.party_id.SetFocus
 'partchk = False
 'End If
 End Sub
Sub order()
'orderchk = True
'If Me.order_no.Text = "" Then
'MsgBox " Please Fill the Order No"
'Me.order_no.SetFocus
'orderchk = False
'End If
End Sub
Sub grid_ini()

    
    Me.vs.Clear
    
    Me.vs.Cols = 17
    
    
    Me.vs.Rows = 100
    Me.vs.ColWidth(0) = 700
    Me.vs.ColWidth(1) = 2200
    Me.vs.ColWidth(2) = 700
    Me.vs.ColWidth(3) = 600
    Me.vs.ColWidth(4) = 600
    Me.vs.ColWidth(5) = 600
    Me.vs.ColWidth(6) = 600
    Me.vs.ColWidth(7) = 0
    Me.vs.ColWidth(8) = 700
    Me.vs.ColWidth(9) = 750
    Me.vs.ColWidth(10) = 700
    Me.vs.ColWidth(11) = 650
    
    Me.vs.ColWidth(12) = 700
    Me.vs.ColWidth(13) = 750
    Me.vs.ColWidth(14) = 1500
    Me.vs.ColWidth(15) = 0
    Me.vs.ColWidth(16) = 1700
    
    
    Me.vs.TextMatrix(0, 0) = "Code"
    Me.vs.TextMatrix(0, 1) = "Particulars"
    Me.vs.TextMatrix(0, 2) = "Qty"
    Me.vs.TextMatrix(0, 3) = "Inner"
    Me.vs.TextMatrix(0, 4) = "Text"
    Me.vs.TextMatrix(0, 5) = "Exam."
    Me.vs.TextMatrix(0, 6) = "Supp."
    Me.vs.TextMatrix(0, 7) = ""
    Me.vs.TextMatrix(0, 8) = "T.Page"
    Me.vs.TextMatrix(0, 9) = "DIV(8/16)"
    Me.vs.TextMatrix(0, 10) = "T.From"
    Me.vs.TextMatrix(0, 11) = "Wast(%)"
    Me.vs.TextMatrix(0, 12) = "Reams"
    Me.vs.TextMatrix(0, 13) = "Sheet"
    Me.vs.TextMatrix(0, 14) = "Binder"
    Me.vs.TextMatrix(0, 15) = ""
    Me.vs.TextMatrix(0, 16) = "Paper Size"
    
   
   For k1 = 0 To vs.Cols - 1
     vs.Cell(flexcpFontSize, k1) = 11
   Next
       
    
   'For k1 = 1 To vs.Rows - 1
   'For cl = 0 To vs.Cols - 1
   '    a1 = k1 Mod 2
   '    If a1 = 0 Then
   '       vs.Cell(flexcpBackColor, k1, cl) = &HFFFFD9
   '     Else
   '       vs.Cell(flexcpBackColor, vs.RowSel, cl) = &HE8FFDF
   '    End If
   'Next
   'Next
    
End Sub
Sub adddisab()
Me.Edit.Enabled = False
Me.Printcmd.Enabled = False
Me.cmdBillCancel.Enabled = False
End Sub
Private Sub Add_Click()

Clearvalue

mode = ""

If RS.State = 1 Then RS.close
RS.Open "select max(ord_no) from OrderPrint_main where " & stringyear, CON
If Not IsNull(RS(0)) Then
   txtOrdNo.Text = RS(0) + 1
Else
   txtOrdNo.Text = 1
End If

txtOrdNo.SetFocus

End Sub

Private Sub bill_no_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
popuplist1 "Select bill_id as [Bill No],dat as [Date] from billmaster where categories='Main' order by cint(bill_id) asc", CON
End If
End Sub

Private Sub bill_no_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

'Dim rsbillmst As ADODB.Recordset
'Set rsbillmst = New ADODB.Recordset
'
'If KeyAscii = 13 Then
'
'mode = ""
'
'If bill_no <> "" Then
'billupdate
''If Checkmanual.value <> 1 Then
'calc
''End If
''If KeyAscii = 13 And bill_no <> "" Then
'If mode = "add" Then
'If rsbillmst.State = 1 Then rsbillmst.close
'       sq = "Select * from billmaster where firm_id = '" + Me.Textfirmid.Text + "' and bill_id = '" + bill_no + "'"
'        rsbillmst.Open sq, CON, adOpenKeyset, adLockReadOnly
'         If rsbillmst.RecordCount > 0 Then
'            MsgBox "This Bill no is already exists"
'            bill_no.SetFocus
'        Else
'            DTPicker1.SetFocus
'        End If
'End If
'End If
'
'End If
End Sub

Private Sub bill_no_LostFocus()
'Dim rsbillmst As ADODB.Recordset
'Set rsbillmst = New ADODB.Recordset
'If bill_no <> "" Then
'billupdate
''If Checkmanual.value <> 1 Then
''calc
''End If
''If KeyAscii = 13 And bill_no <> "" Then
'If mode = "add" Then
'If rsbillmst.State = 1 Then rsbillmst.close
'       sq = "Select * from billmaster where firm_id = '" + Me.Textfirmid.Text + "' and bill_id = '" + bill_no + "'"
'        rsbillmst.Open sq, CON, adOpenKeyset, adLockReadOnly
'         If rsbillmst.RecordCount > 0 Then
'            MsgBox "This Bill no is already exists"
'            bill_no.SetFocus
'        Else
'            DTPicker1.SetFocus
'        End If
'End If
'
'End If

End Sub
Private Sub Binder_id_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
If Binder_id.Text = "" Then
MsgBox "Please Choose the Printers"
Binder_id.SetFocus
Exit Sub
Else
'Me.vs.Col = 1

If vs.Enabled = True Then
vs.SetFocus
End If
'SendKeys "{tab}"
End If
'Dim rsparty As ADODB.Recordset
'    Set rsparty = New ADODB.Recordset
'    sq = "Select binder_id,binder_name from binderMaster where binder_id = '" + Me.Binder_id.Text + "'"
'    rsparty.Open sq, con, adOpenKeyset, adLockReadOnly
'        If rsparty.RecordCount > 0 Then
'        Me.binder_name.Text = rsparty.Fields("binder_name")
'        Me.vs.SetFocus
'        Me.vs.Col = 1
'                    Else
'        MsgBox "Invalid Binder"
'        Me.Binder_id.SetFocus
'        Me.Binder_id.Text = ""
'        Exit Sub
'        End If
'    End If
    End If
End Sub

Private Sub Binder_id_LostFocus()
Label10.Visible = False
End Sub
Private Sub binder_name_Click()
westage = 0
If RS.State = 1 Then RS.close
RS.Open "select Address,westage from Godownmaster  where godwn='" & binder_name & "'", CON, adOpenKeyset, adLockReadOnly
If RS.EOF = False Then
   lblAdd.Caption = RS(0) & ""
   westage = RS(1)
Else
   lblAdd.Caption = ""
End If

End Sub

Private Sub binder_name_GotFocus()
''Label10.Visible = True
'If PopUpValue1 <> "" Then
'
'
' Me.binder_name = PopUpValue1
'
' If Me.binder_name = "" Then Exit Sub
'
' If rs.State = 1 Then rs.close
' rs.Open "select customer_id from CustomerMaster where customer_name='" & Me.binder_name.Text & "'"
' If rs.EOF = False Then
'    Me.Binder_id.Text = rs.Fields(0).value
' End If
'
'
' Me.binder_name.Enabled = True
'
'End If
'PopUpValue1 = ""
'PopUpValue2 = ""
End Sub

Private Sub binder_name_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 113 Then
'   popuplist1 "Select customer_name as [Printers Name],add1 as Address from CustomerMaster where Category='PRINTER'", CON
'End If
End Sub

Private Sub binder_name_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

    If binder_name.Text = "" Then
       MsgBox "Please Choose the Printers", vbInformation
       Me.binder_name.SetFocus
       Exit Sub
    End If
    
    SendKeys "{tab}"


End If

End Sub

Private Sub cancel_Click()

mode = ""
Me.binder_name.Text = ""
grid_ini


End Sub

Private Sub cmdBillCancel_Click()
   If Me.txtOrdNo.Text = "" Then Exit Sub
   
   If RS.State = 1 Then RS.close
   RS.Open "select * from billmaster where firm_id = '" + "Chitra" + "' and bill_id = '" + Me.txtOrdNo.Text + "'", CON
   If RS.EOF = True Then
      MsgBox "Bill No Already Exist !!", vbInformation
      Exit Sub
   End If
   
   If RS.State = 1 Then RS.close
   RS.Open "select * from billmaster where firm_id = '" + "Chitra" + "' and bill_id = '" + Me.txtOrdNo.Text + "'", CON, adOpenDynamic, adLockOptimistic
   If RS.EOF = False Then
   If MsgBox("Want To Order Cancel", vbQuestion + vbYesNo) = vbYes Then
      RS!OrderCancel = "Yes"
      RS.update
   End If
   End If
End Sub

Private Sub Command1_Click()
flag = True
frmbook.Show
End Sub

Private Sub Command2_Click()
flag = True
printingmaster.Option3.value = True
printingmaster.mname.Caption = "Size Master"
printingmaster.Show
End Sub
Private Sub Command3_Click()
flag = True
CustomerMaster.Show
End Sub
Private Sub Command4_Click()
flag = True
BinderMaster.Show
End Sub
Private Sub CommandPAPERSTATEMENT_Click()
''''paperstatementcalc
''''        If Textfirmid.Text = "MITTAL" Then
''''        frmbill.cr1.ReportFileName = App.Path & "\mpaperstat.rpt"
''''        ElseIf Textfirmid.Text = "DAYAL" Then
''''        frmbill.cr1.ReportFileName = App.Path & "\dpaperstat.rpt"
''''        Else
''''        frmbill.cr1.ReportFileName = App.Path & "\paperstat.rpt"
''''        End If
''''frmbill.cr1.DataFiles(0) = ""
''''frmbill.cr1.DataFiles(1) = ""
''''frmbill.cr1.DataFiles(2) = ""
''''frmbill.cr1.DataFiles(3) = ""
''''frmbill.cr1.DataFiles(4) = ""
''''frmbill.cr1.DataFiles(0) = App.Path & "\" + Trim(main.directory) + "\data.mdb"
''''frmbill.cr1.DataFiles(1) = App.Path & "\" + Trim(main.directory) + "\data.mdb"
''''frmbill.cr1.DataFiles(2) = App.Path & "\" + Trim(main.directory) + "\data.mdb"
''''frmbill.cr1.DataFiles(3) = App.Path & "\" + Trim(main.directory) + "\data.mdb"
''''{billMASTER.bill_id} = "151" AND {BILLMASTER.FIRM_ID} = "NEERAJ"
''''MRSF = "{billMASTER.firm_id} = '" & frmbill.Textfirmid & "' and  {billMASTER.bill_id} = '" & frmbill.bill_no & "'"
''''MRSF = ""
''''frmbill.cr1.ReplaceSelectionFormula (MRSF)
''''frmbill.cr1.Destination = 1
''''frmbill.cr1.Action = 1
End Sub

Private Sub cmdPrint_Slip_Click()
cr1.Reset
cr1.ReportFileName = App.Path & "/report_paper/BinderSlip.rpt"
cr1.Connect = "filedsn=chitradsn;uid=sa;pwd=sidc;"
cr1.ReplaceSelectionFormula "{orderPrint_Main.ord_no}='" & txtOrdNo.Text & "'"
cr1.WindowShowPrintSetupBtn = True
cr1.WindowState = crptMaximized
cr1.Action = 1

End Sub

Private Sub Delete_Click()

'''Dim rsbilmast As ADODB.Recordset
'''Dim rsbiltrans As ADODB.Recordset
'''Dim rsbill1 As ADODB.Recordset
'''Set rsbill1 = New ADODB.Recordset
'''Set rsbilmast = New ADODB.Recordset
'''Set rsbiltrans = New ADODB.Recordset
'''Dim fid As String
'''fid = Trim(Me.Textfirmid.Text)

X = MsgBox("Are you sure you wish to delete the selected Bill ", 4, "Confirmation")
If X = 6 Then
   sq = "delete  from OrderPrint_Main where Ord_No = '" + txtOrdNo.Text + "' and " & stringyear
   CON.Execute sq
   sq = "delete  from OrderPrint_Det where Ord_No = '" + txtOrdNo.Text + "' and " & stringyear
   CON.Execute sq
   Call Add_Click
End If


'''sq = "delete * from billmaster where firm_id = '" + fid + "' and bill_id = '" + Me.bill_no.Text + "'"
'''CON.Execute sq
'''sq = "delete * from billtrans where firm_id = '" + fid + "' and bill_id = '" + Me.bill_no.Text + "'"
'''CON.Execute sq
'''rsbill1.Open "select * from paperstatement where papersize1 ='" + TextPaperSize + "' AND PSCUSTOMERID = '" + pscustomerid + "' AND trim(BILL_ID) = '" + Trim(Me.bill_no) + "'", CON, adOpenDynamic, adLockPessimistic
'''If rsbill1.RecordCount > 0 Then
'''     Do While Not rsbill1.EOF
'''        rsbill1!bill_id = "0"
'''        rsbill1!pstatementno = "0"
'''     rsbill1.update
'''     If Not rsbill1.EOF Then
'''        rsbill1.MoveNext
'''     End If
'''Loop
'''End If
'''
'''
'''
'''
'''
'''
'''
'''
'''End If
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
binder_name.Enabled = True
Me.binder_name.SetFocus
End If
End Sub

Private Sub DTPicker1_LostFocus()

'Dim dt1, dt2 As Date
'Dim bill1 As Long
'Dim bill2 As Long
'Dim rsdate1 As ADODB.Recordset
'Set rsdate1 = New ADODB.Recordset
'Dim rsdate2 As ADODB.Recordset
'Set rsdate2 = New ADODB.Recordset
'If rsdate1.State = 1 Then rsdate1.close
'
'rsdate1.Open "select max(val(bill_id)) as bill1 from billmaster where val(bill_id) < " + Str(Val(bill_no)) + "", CON, adOpenStatic, adLockOptimistic
'If rsdate2.State = 1 Then rsdate2.close
'rsdate2.Open "select min(val(bill_id)) as bill2 from billmaster where val(bill_id) > " + Str(Val(bill_no)) + "", CON, adOpenStatic, adLockOptimistic
'If rsdate1!bill1 > 0 Then bill1 = rsdate1!bill1
'If rsdate2!bill2 > 0 Then bill2 = rsdate2!bill2 Else bill2 = Str(Val(bill_no))
'If rsdate1.State = 1 Then rsdate1.close
'rsdate1.Open "select * from  billmaster  where val(bill_id) = " + Str(bill1) + "", CON, adOpenStatic, adLockOptimistic
'If rsdate1.RecordCount > 0 Then
'dt1 = rsdate1!dat
'End If
'If rsdate2.State = 1 Then rsdate2.close
'rsdate2.Open "select * from  billmaster  where val(bill_id) = " + Str(bill2) + "", CON, adOpenStatic, adLockOptimistic
'If rsdate2.RecordCount > 0 Then
'dt2 = rsdate2!dat
'Else
'dt2 = dt1 + 365
'End If
'If Me.DTPicker1.value < dt1 Or Me.DTPicker1.value > dt2 Then
'If bill2 <> Str(Val(bill_no)) Then
'MsgBox "You cannot create back date invoice"
'
''Me.DTPicker1.SetFocus
'Me.bill_no.SetFocus
'Exit Sub
'End If
'End If
End Sub

Private Sub Edit_Click()
Me.ok.Enabled = True
delete.Enabled = True
cancel.Enabled = True
Me.Edit.Enabled = False
mode = "edit"
ok.SetFocus
End Sub

Private Sub Form_Activate()
Add_Click
End Sub
Sub maxNo()
    If RS.State = 1 Then RS.close
    RS.Open "select max(val(bill_id)) from billmaster", CON, adOpenDynamic, adLockOptimistic
    If Not IsNull(RS.Fields(0).value) Then
       bill_no.Text = RS.Fields(0).value
    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{tab}"
End If
End Sub
Private Sub Form_Load()

 Me.Left = 100
 Me.Top = 100
 Me.Width = 14000
 Me.Height = 8500


grid_ini

bkfont = "e"
fid = "chitra"


st_ = ""

binder_name.Clear
If RS.State = 1 Then RS.close
RS.Open "select Godwn as [Binder Name],Address from Godownmaster where Binder_Printer='p' order by Godwn", CON, adOpenStatic, adLockReadOnly
While RS.EOF = False
  binder_name.AddItem RS(0)
  RS.MoveNext
Wend



If RS.State = 1 Then RS.close
RS.Open "select Godwn as [Binder Name],Address from Godownmaster where Binder_Printer='b' order by Godwn", CON, adOpenStatic, adLockReadOnly
While RS.EOF = False
  If st_ = "" Then
     st_ = RS(0)
  Else
     st_ = st_ & "|" & RS(0)
  End If
  
  RS.MoveNext
Wend

'-------------------------------------------------
s = ""
If RS.State = 1 Then RS.close
RS.Open "select * from PaperMakeMaster where " & stringyear & " order by papermaker_name", CON, adOpenStatic, adLockReadOnly
While RS.EOF = False

If s = "" Then
    s = RS!papermaker_name & " : " & RS!eco & " : " & RS!SizeValue1 & "X" & RS!SizeValue2 & " : " & RS!GSM & "=>" & RS!papermaker_id
Else
    s = s & "|" & RS!papermaker_name & " : " & RS!eco & " : " & RS!SizeValue1 & "X" & RS!SizeValue2 & " : " & RS!GSM & "=>" & RS!papermaker_id
End If

RS.MoveNext
Wend

vs.ColComboList(16) = s




txtOrdDate.value = Format(Date, "dd/MM/yyyy")
vs.ColComboList(14) = st_

BackColorFrom Me

End Sub


Private Sub Godown_Click()
  If Godown.value = 1 Then
     Label2.Caption = "GoDown Id"
  Else
     Label2.Caption = "Printer Id"
  End If
End Sub
Private Sub txtOrdDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then binder_name.SetFocus
End Sub

Private Sub txtOrdNo_GotFocus()
 If PopUpValue1 <> "" Then
    txtOrdNo.Text = PopUpValue1
    SearchData
    PopUpValue1 = ""
 End If
End Sub
Private Sub txtOrdNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
   popuplist1 "Select distinct Ord_No,Ord_Date from OrderPrint_Main where " & stringyear & " order by Ord_No", CON
End If

If KeyCode = 13 Then txtOrdDate.SetFocus
End Sub

Private Sub vs_DblClick()
'' If vs.TextMatrix(vs.RowSel, 20) <> "" Then
''    txtNote.Visible = True
''    txtNote.Text = vs.TextMatrix(vs.RowSel, 18)
''    txtNote.SetFocus
'' End If
End Sub
Sub addTotalReam()

Dim ream_ As Long
Dim sheet_ As Long
Dim per_

ream_ = 0
sheet_ = 0
per_ = 0

For I = 1 To vs.Rows - 1
  If vs.TextMatrix(I, 0) <> "" Then
     ream_ = ream_ + Val(vs.TextMatrix(I, 12))
     sheet_ = sheet_ + Val(vs.TextMatrix(I, 13))
  End If
Next



If sheet_ > 499 Then
   per_ = Int(sheet_ / 500)
   sheet_ = sheet_ - per_ * 500
End If


txtTReam = ream_ + per_
txtTSheet = sheet_


End Sub
Private Sub vs_GotFocus()
'''vs.TextMatrix(1, 0) = 1
''
''If Me.vs.Col = 1 Then
''
''    If PopUpValue1 <> "" Then
''    If bkfont = "h" Then
''     vs.CellFontName = hindi
''     vs.CellFontSize = 14
''    End If
''
''
''     Me.vs.Col = 1
''     Me.vs.Text = PopUpValue1
''
''
''
''    Dim rsbook As ADODB.Recordset
''    Set rsbook = New ADODB.Recordset
''    If rsbook.State = 1 Then rsbook.close
''    sq = "select * from bookmaster where BookNo= '" + PopUpValue1 + "'"
''
''    rsbook.Open sq, CON, adOpenKeyset, adLockReadOnly
''        If rsbook.RecordCount > 0 Then
''            'vs.Col = vs.Col + 1
''            If bkfont = "h" Then
''                vs.CellFontName = hindi
''                vs.CellFontSize = 14
''            End If
''            'vs.Text = rsbook!book_info
''
''
''            vs.Col = 1: If rsbook!book <> "" Then vs.Text = rsbook!book
''            vs.Col = 3: If rsbook!book_unit <> "" Then vs.Text = rsbook!book_unit
''            vs.Col = 4: If rsbook!websheet <> "" Then vs.Text = rsbook!websheet
''            vs.Col = 6: If rsbook!atrate <> "" Then vs.Text = rsbook!atrate
''            vs.Col = 11: If rsbook!wastage <> "" Then vs.Text = rsbook!wastage
''            vs.Col = 15: If rsbook!patrate <> "" Then vs.Text = rsbook!patrate
''            vs.Col = 19: If rsbook!BookNo <> "" Then vs.Text = rsbook!BookNo
''            vs.Col = 21: If rsbook!book_size <> "" Then vs.Text = rsbook!book_size
''
''            vs.Col = 2: If rsbook!Lemination <> "" Then vs.Text = rsbook!Lemination
''            vs.Col = 18: If rsbook!Brand <> "" Then vs.Text = rsbook!Brand
''            SendKeys "{right}"
''
''        vs.Col = 17: vs.Text = bkfont
''        End If
''            'Label10.Visible = False
''            Label11.Visible = False
''            Label12.Visible = False
''            'Label14.Visible = False
''            Label16.Visible = False
''            vs.Col = 2
''            vs.SetFocus
''    End If
''
''Else
''    If PopUpValue1 <> "" Then
''    vs.Text = PopUpValue1
''    End If
''End If
''
''
''If Me.vs.Col = 2 Then
''    If PopUpValue1 <> "" Then
''    If bkfont = "e" Then
''        vs.CellFontName = english
''        vs.CellFontSize = 10
''    End If
''    End If
''End If
''
''
''
''
''If vs.Col = 1 Then
'''Label10.Visible = True
''Label11.Visible = True
''Label12.Visible = True
'''Label14.Visible = True
''Label16.Visible = True
''End If
''
'''PopUpValue1 = ""
'''PopUpValue2 = ""
''PopUpValue3 = ""

End Sub
Function cheqePrinter(P1 As String, I As Integer, bno As String) As Boolean
   
   Dim ss1 As New ADODB.Recordset
   
   Select Case I
   
   Case 1
   
     If ss1.State = 1 Then ss1.close
     ss1.Open "select * from bookmaster where Inn_Printer='" & binder_name.Text & "' and bookno='" & bno & "'", CON, adOpenDynamic, adLockReadOnly
     If ss1.EOF = False Then
        cheqePrinter = True
     Else
        cheqePrinter = False
     End If
     
   Case 2
   
     
     If ss1.State = 1 Then ss1.close
     ss1.Open "select * from bookmaster where text_Printer='" & binder_name.Text & "' and bookno='" & bno & "'", CON, adOpenDynamic, adLockReadOnly
     If ss1.EOF = False Then
        cheqePrinter = True
     Else
        cheqePrinter = False
     End If
     
   
   Case 3
   
   
     If ss1.State = 1 Then ss1.close
     ss1.Open "select * from bookmaster where Exam_Printer='" & binder_name.Text & "' and bookno='" & bno & "'", CON, adOpenDynamic, adLockReadOnly
     If ss1.EOF = False Then
        cheqePrinter = True
     Else
        cheqePrinter = False
     End If
   
   
   
   Case 4
   
     
     If ss1.State = 1 Then ss1.close
     ss1.Open "select * from bookmaster where Supp_Printer='" & binder_name.Text & "' and bookno='" & bno & "'", CON, adOpenDynamic, adLockReadOnly
     If ss1.EOF = False Then
        cheqePrinter = True
     Else
        cheqePrinter = False
     End If
     
   
   Case 5
   
     
     If ss1.State = 1 Then ss1.close
     ss1.Open "select * from bookmaster where Title_Printer='" & binder_name.Text & "' and bookno='" & bno & "'", CON, adOpenDynamic, adLockReadOnly
     If ss1.EOF = False Then
        cheqePrinter = True
     Else
        cheqePrinter = False
     End If
   
   
   
   End Select

End Function

Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)
''
''
''
''If KeyCode = 114 Then
''   vs.CellFontName = hindi
''   vs.CellFontSize = 14
''   vs.Col = 2
''   vs.CellFontName = english
''   vs.Col = 17: vs.Text = "e"
''   vs.Col = 1
''   bkfont = "h"
''
''
''
''ElseIf KeyCode = 113 Then
''vs.CellFontName = english
''vs.Col = 2
''vs.CellFontName = english
''vs.Col = 17: vs.Text = "e"
''vs.Col = 1
''bkfont = "e"
''
''End If
''End If
''
''If (Me.vs.Col = 1 And KeyCode = 13) Then
''
''    ItemCode.Visible = True
''    ItemCode.SetFocus
''    ItemCode.ZOrder
''
''    ItemCode.Text = vs.Text
''    ItemCode.Top = vs.Top + vs.CellTop
''    ItemCode.Left = vs.CellLeft + 200
''    ItemCode.Width = vs.ColWidth(vs.Col)
''
''
''    Exit Sub
''
''End If
''
''
''
''

If KeyCode = 115 Then
   If MsgBox("Are You Sure, Delete it ?", vbYesNo) = vbYes Then
      If vs.Row > 1 Then
         vs.RemoveItem vs.Row
         addTotalReam
         vs.SetFocus
      Else
      MsgBox "You cannnot delete first and fixed row you can edit it", vbInformation
      End If
   End If
End If

''
'''If vs.Col = 1 And KeyCode = 13 Then vs.Col = 3: Exit Sub
'''If vs.Col = 2 And KeyCode = 13 Then vs.Col = 3: Exit Sub
''
''If vs.Col = 3 And KeyCode = 13 Then calc: vs.Col = 4: Exit Sub
''If vs.Col = 4 And KeyCode = 13 Then calc: vs.Col = 5: Exit Sub
''If vs.Col = 5 And KeyCode = 13 Then calc: vs.Col = 9: Exit Sub
''If vs.Col = 9 And KeyCode = 13 Then calc: vs.Col = 10: Exit Sub
''If vs.Col = 10 And KeyCode = 13 Then calc: vs.Col = 11: Exit Sub
''If vs.Col = 11 And KeyCode = 13 Then calc: vs.Col = 17: Exit Sub
''
''
''If vs.Col = 21 And KeyCode = 13 Then
''calc
''
''    If vs.Row = vs.Rows - 1 Then
''    vs.AddItem "", vs.Row + 1
''    vs.Row = vs.Row + 1
''    'vs.TextMatrix(r + 1, 0) = r + 1
''    vs.TextMatrix(vs.Row, 17) = "e"
''    End If
''    ST = Me.vs.Row
''    Me.vs.RowHeight(ST) = 300
''
''    SendKeys "{DOWN}"
''    vs.Row = r + 1
''
''    If vs.TextMatrix(vs.RowSel, 0) = "" Then
''       vs.Col = 0
''       Exit Sub
''    End If
''
''    'For i = 0 To 14
''    '    SendKeys "{LEFT}"
''    'Next i
''
''End If
''
''If vs.Col = 1 Then
''    Label11.Visible = True
''    Label12.Visible = True
''    Label16.Visible = True
''End If
''
''If vs.Col <> 1 Then
''    Label11.Visible = False
''    Label12.Visible = False
''    Label16.Visible = False
''End If
''

End Sub
Sub Ream_Sheet()

Dim Tot, wastage_per, Form, wream


wastage_per = 0

Form = Val(vs.TextMatrix(vs.RowSel, 10))
quan = Val(vs.TextMatrix(vs.RowSel, 2))
wastage_per = Val(vs.TextMatrix(vs.RowSel, 11))

Tot = Form * quan
Tot = Tot / 1000

If Val(wastage_per) > 0 Then
   a1 = (Tot * wastage_per / 100)
   wream = Int(a1)
   wsheet = Round((a1 - Int(a1)) * 500)
   'a1 = (quan * wastage_per / 100)
End If

  
If Tot > 0 Then
   
   'ream = tot / 1000
   'Tot = Tot / 1000
   ream = Int(Tot)
   sheet = Round((Tot - Int(Tot)) * 500, 0)
   sheet = sheet + wsheet
   
   If sheet > 499 Then
      wream = Int(sheet / 500)
      sheet = sheet - (wream * 500)

      ream = ream + wream
   Else
      ream = ream + wream
   End If
 
   
 
   
   
   
  
End If






End Sub

Sub calc()
''prow = vs.Row
'''grid_ini
''Dim unit, unit1 As Double
''Dim rate, atrate As Double
''Dim amt As Double
''Dim Tot, quan, sheet, totrim As Double
''Dim wrim, wper, wsheet, wtemp As Double
''Dim lent As Integer
''tmpplt = 0
''vs.Col = 5: quan = Val(vs.Text)
''vs.Col = 11: per = Val(vs.Text)
''vs.Col = 6: atrate = Val(vs.Text)
'''manual calculations start
'''If Checkmanual.value <> 1 Then
'''  tmp = quan / 1100
'''  If tmp < 1 Then tmp = 1
'''  If Int(tmp) = tmp Then rate = atrate * tmp
'''    If Int(tmp) < tmp Then
'''        If (tmp - Int(tmp)) <= 0.5 Then rate = atrate * (Int(tmp) + 0.5)
'''        If (tmp - Int(tmp)) > 0.5 Then rate = atrate * (Int(tmp) + 1)
'''    End If
'''   vs.Col = 7: vs.Text = rate
'''    tmpplt = quan / 11000
'''    If Int(tmpplt) < tmpplt Then
'''        tmpplt = Int(tmpplt) + 1
'''   End If
'''   vs.Col = 14: vs.Text = tmpplt
'''End If
''  vs.Col = 3: If Val(vs.Text) >= 1 Then unit1 = Val(vs.Text) Else unit1 = 1
''  unit = Val(vs.Text)
''  vs.Col = 7: rate = Val(vs.Text)
''
'''If Checkmanual.value <> 1 Then
'''tmpplate1 = 0
'''If Int(unit) < unit Then
'''X = unit - Int(unit)
'''If X <= 0.5 Then tmpplate1 = Int(unit) + 1
'''If X > 0.5 Then tmpplate1 = Int(unit) + 2
'''Else
'''tmpplate1 = unit
'''End If
'''vs.Col = 14: tmpplt = vs.Text * tmpplate1
'''vs.Col = 14: vs.Text = tmpplt
'''Else
'''vs.Col = 14:  tmpplt = vs.Text
'''End If
''  If quan > 0 And quan <= 1100 Then
''  amt = Round(Val(tmpplt) * rate, 0)
''  Else
''  amt = Round(unit1 * rate, 0)
''  End If
''  vs.Col = 8: vs.Text = amt
''vs.Col = 15
''pltamt = Val(tmpplt) * Val(vs.Text)
''vs.Col = 16: vs.Text = pltamt
''
''  wastflag = "K"
''  tmpwast = 0
''
''  If per <> 15 Then
''  tmpwast = per * quan / 100
''  Else
''  tmpwast = 15
''  End If
''  tmpact = 0
''  tmpact = per * 1000 / 100
''  If tmpwast > 0 And tmpwast <= tmpact Then
''  wastflag = "N"
''  wast = 15
''  vs.Col = 11: vs.Text = 15
''  End If
''  per = 0
''
''  Tot = unit * quan
''  If Tot > 0 Then
''    Tot = Tot / 1000
''    sheet = (Tot - Int(Tot)) * 1000 / 2 'ghgfhfh
''
''    vs.Col = 11: wrim = Val(vs.Text)
''    If wastflag = "N" Then
''    tmptotw1 = 0
'' tmptotw1 = tmpplt * wast
''
''    wsheet = Round(tmptotw1, 0)
''   tmpwreams = 0
''tmpwsheets = 0
''If wsheet > 499 Then
''per = Int(wsheet / 500)
''wsheet = wsheet - per * 500
''End If
''    Else
''    per = Tot * wrim / 100
''    wsheet = (per - Int(per)) * 1000 / 2
''    End If
''  Else
''    If Tot > 500 Then
''        tmptot = Tot
''        sheet = Tot - 500
''        Tot = 1
''        vs.Col = 11: wrim = Val(vs.Text)
''         If wastflag = "N" Then
''    per = unit * wast
''    Else
''        per = tmptot * wrim / 100
''        End If
''        If per > 500 Then
''            wsheet = per - 1
''            per = 1
''        Else
''        wsheet = Round(per, 0)
''        per = 0
''        End If
''    Else
''    sheet = Tot
''     vs.Col = 11: wrim = Val(vs.Text)
''      If wastflag = "N" Then
''    per = unit * wast
''    Else
''        per = Tot * wrim / 100
''        End If
''        Tot = 0
''        wsheet = per
''        per = 0
''    End If
''  End If
''  vs.Col = 9
''  vs.Text = Int(Tot)
''  vs.Col = 10
''  vs.Text = Round(sheet, 0)
''
''vs.Col = 12: vs.Text = Int(per)
''vs.Col = 13: vs.Text = Round(wsheet)
''Dim amount, reams, sheets, plate
''Dim ptamt As Double
''reams = 0
''sheets = 0
''amount = 0
''wreams = 0
''wsheets = 0
''plate = 0
''ptamt = 0
''
''vs.Col = 8
''For i = 1 To vs.Rows - 1
''vs.Row = i
''vs.Col = 8
''amount = amount + Val(vs.Text)
''vs.Col = 9
''reams = Round(reams + Val(vs.Text), 0)
''vs.Col = 10
''sheets = Round(sheets + Val(vs.Text), 0)
''vs.Col = 12
''wreams = Round(wreams + Val(vs.Text), 0)
''vs.Col = 13
''wsheets = Round(wsheets + Val(vs.Text), 0)
''vs.Col = 14
''plate = plate + Val(vs.Text)
''vs.Col = 16
''ptamt = ptamt + Val(vs.Text)
''Next i
''Me.total = Round(amount, 0)
''
'''''Me.totalplate = Round(plate, 0)
''
''Me.pltamt = Round(ptamt, 0)
''TMPreams = 0
''TMPsheets = 0
''If sheets > 499 Then
''TMPreams = Int(sheets / 500)
''TMPsheets = TMPreams * 500
''End If
''Me.totalream = Round(reams, 0) + Round(TMPreams, 0)
''Me.totalsht = Round(sheets, 0) - Round(TMPsheets, 0)
''tmpwreams = 0
''tmpwsheets = 0
''If wsheets > 499 Then
''tmpwreams = Int(wsheets / 500)
''tmpwsheets = tmpwreams * 500
''End If
''
''Me.totalwream = Round(wreams, 0) + Round(tmpwreams, 0)
''Me.totalwsht = Round(wsheets, 0) - Round(tmpwsheets, 0)
''
''tmpgreams = 0
''tmpgsheets = 0
''greams = 0
''gsheets = 0
''greams = Round(Val(Me.totalream), 0) + Round(Val(Me.totalwream), 0)
''gsheets = Round(Val(Me.totalsht), 0) + Round(Val(Me.totalwsht), 0)
''If gsheets > 499 Then
''tmpgreams = Int(gsheets / 500)
''tmpgsheets = tmpgreams * 500
''End If
''Me.gtotalreams = greams + tmpgreams
''Me.gtotalsheets = gsheets - tmpgsheets
''Me.gtotalamt = Round(Val(total) + Val(pltamt), 0)
''vs.Row = prow
''
'''If vs.TextMatrix(vs.RowSel, 0) = "" Then
'''   vs.Col = 1
'''End If
''
'''Me.total = Val(total) + amt
''
End Sub
Private Sub vs_KeyPress(KeyAscii As Integer)




If vs.Col <> 1 Then
'Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
'Label14.Visible = False
Label16.Visible = False
End If
If vs.Col = 1 Then
'Label10.Visible = True
Label11.Visible = True
Label12.Visible = True
'Label14.Visible = True
Label16.Visible = True
End If



If vs.Col = 1 Or vs.Col = 20 Then
        If KeyAscii = 8 Then
        If Len(Trim(vs.Text)) <> 0 Then
                vs.Text = Left(vs.Text, (Len(vs.Text) - 1))
        End If
            'ElseIf (KeyAscii = 13) Or (KeyAscii >= 32 And KeyAscii <= 126) Then
            ElseIf (KeyAscii >= 32 And KeyAscii <= 126) Then
            vs.Text = vs.Text + Chr(KeyAscii)
        End If
End If


If vs.Col = 2 Then

        If Len(Trim(vs.Text)) = 80 Then
    KeyAscii = 0
    End If
        If KeyAscii = 8 Then
            If Len(Trim(vs.Text)) <> 0 Then
            vs.Text = Left(vs.Text, (Len(vs.Text) - 1))
        End If
            ElseIf (KeyAscii <> 13) Or (KeyAscii >= 32 And KeyAscii <= 126) Then
            vs.Text = vs.Text + Chr(KeyAscii)
        End If


End If

If vs.Col = 4 Then
    If KeyAscii = 119 Or KeyAscii = 87 Then
       vs.Text = "Web"
       web.Visible = False
    End If
    If KeyAscii = 115 Or KeyAscii = 83 Then
       vs.Text = "Sheet"
       web.Visible = False
    End If

End If

If (vs.Col = 3 Or vs.Col = 4 Or vs.Col = 5 Or vs.Col = 6 Or vs.Col = 7 Or vs.Col = 11 Or vs.Col = 12 Or vs.Col = 13 Or vs.Col = 14 Or vs.Col = 15) Then
    If Len(Trim(vs.Text)) = 10 Then
    KeyAscii = 0
    End If
    If KeyAscii = 8 Then
        If Len(Trim(vs.Text)) <> 0 Then
                vs.Text = Left(vs.Text, (Len(vs.Text) - 1))
        End If
    End If
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        Else
        If KeyAscii <> 46 Then
            If KeyAscii <> 8 Then
                KeyAscii = 0
            End If
        End If
    End If
    If KeyAscii <> 8 Then
        vs.Text = vs.Text + Chr(KeyAscii)
    End If
End If



End Sub
Function addPage() As Double
   

page_sum = 0
   
For P1 = 3 To 6
   If Val(vs.TextMatrix(vs.RowSel, P1)) > 0 Then
      page_sum = page_sum + Val(vs.TextMatrix(vs.RowSel, P1))
   End If
Next

addPage = page_sum

End Function
Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)

On Error Resume Next

If KeyCode = 13 Then
   
If vs.Col = 0 Then
   
   If RS.State = 1 Then RS.close
   RS.Open "select book,book_unit,DivideValue,HeadData1,HeadData2,HeadData3,HeadData4,HeadData5,bookfont from BookMaster " & _
   " where BookNo='" & vs.TextMatrix(vs.RowSel, 0) & "'", CON
   If RS.EOF = False Then
      vs.TextMatrix(vs.RowSel, 0) = UCase(vs.TextMatrix(vs.RowSel, 0))
      
      vs.TextMatrix(vs.RowSel, 1) = UCase(RS!book)
      If RS!bookfont = "h" Then
          vs.Cell(flexcpFontName, vs.RowSel, 1) = hindi
          vs.Cell(flexcpFontSize, vs.RowSel, 1) = 14
          vs.TextMatrix(vs.RowSel, 15) = "h"
       Else
          vs.Cell(flexcpFontName, vs.RowSel, 1) = english
          vs.Cell(flexcpFontSize, vs.RowSel, 1) = 10
          vs.TextMatrix(vs.RowSel, 15) = "e"
       End If
       
      
      vs.TextMatrix(vs.RowSel, 3) = Val(RS!HeadData1) & ""
      vs.TextMatrix(vs.RowSel, 4) = Val(RS!HeadData2) & ""
      vs.TextMatrix(vs.RowSel, 5) = Val(RS!HeadData3) & ""
      vs.TextMatrix(vs.RowSel, 6) = Val(RS!HeadData4) & ""
      vs.TextMatrix(vs.RowSel, 7) = Val(RS!HeadData5) & ""
      vs.TextMatrix(vs.RowSel, 8) = (Val(RS!HeadData1) + Val(RS!HeadData2) + Val(RS!HeadData3) + Val(RS!HeadData4) + Val(RS!HeadData5))
      vs.TextMatrix(vs.RowSel, 9) = Val(RS!DivideValue)
      vs.TextMatrix(vs.RowSel, 10) = Round(Val(vs.TextMatrix(vs.RowSel, 8)) / Val(vs.TextMatrix(vs.RowSel, 9)), 2)
      
      SendKeys "{right}"
      SendKeys "{right}"
   End If

ElseIf vs.Col = 2 Then
      SendKeys "{right}"
ElseIf vs.Col = 3 Then
      SendKeys "{right}"
      vs.TextMatrix(vs.RowSel, 8) = addPage
ElseIf vs.Col = 4 Then
      SendKeys "{right}"
      vs.TextMatrix(vs.RowSel, 8) = addPage
ElseIf vs.Col = 5 Then
      SendKeys "{right}"
      vs.TextMatrix(vs.RowSel, 8) = addPage
ElseIf vs.Col = 6 Then
      SendKeys "{right}"
      SendKeys "{right}"
      vs.TextMatrix(vs.RowSel, 8) = addPage
ElseIf vs.Col = 7 Then
      SendKeys "{right}"
      SendKeys "{right}"
      vs.TextMatrix(vs.RowSel, 8) = addPage

ElseIf vs.Col = 9 Then
      SendKeys "{right}"
      SendKeys "{right}"
      vs.TextMatrix(vs.RowSel, 11) = ""
      vs.TextMatrix(vs.RowSel, 11) = westage
      vs.TextMatrix(vs.RowSel, 10) = Round(Val(vs.TextMatrix(vs.RowSel, 8)) / Val(vs.TextMatrix(vs.RowSel, 9)), 2)
      Ream_Sheet
ElseIf vs.Col = 10 Then
      SendKeys "{right}"
ElseIf vs.Col = 11 Then
      SendKeys "{right}"
      Ream_Sheet
      vs.TextMatrix(vs.RowSel, 12) = ream
      vs.TextMatrix(vs.RowSel, 13) = sheet
      SendKeys "{right}"
      SendKeys "{right}"
ElseIf vs.Col = 14 Then
      
    If vs.TextMatrix(vs.RowSel, 14) <> "" Then
       'SendKeys "{home}"
       'SendKeys "{down}"
       SendKeys "{right}"
       addTotalReam
    End If
    
ElseIf vs.Col = 16 Then
      
    If vs.TextMatrix(vs.RowSel, 15) <> "" Then
      SendKeys "{home}"
      SendKeys "{down}"
      addTotalReam
    End If
    
End If


End If





End Sub

Private Sub vs_SelChange()

If (vs.Col = 0 Or vs.Col = 2 Or vs.Col = 3 Or vs.Col = 4 Or vs.Col = 5 Or vs.Col = 6 Or vs.Col = 7 Or vs.Col = 9 Or vs.Col = 11 Or vs.Col = 14 Or vs.Col = 16) Then
   vs.Editable = flexEDKbdMouse
ElseIf vs.Col = 19 Then
   vs.Editable = flexEDKbdMouse
ElseIf vs.Col = 20 Then
   vs.Editable = flexEDKbdMouse
ElseIf vs.Col = 21 Then
   vs.Editable = flexEDKbdMouse
ElseIf vs.Col = 17 Then
   vs.Editable = flexEDKbdMouse
Else
   vs.Editable = flexEDNone
End If




End Sub

Private Sub ItemCode_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
  SendKeys "{down}"
  PopUpValue1 = Trim(Mid(ItemCode.Text, InStr(ItemCode.Text, "=>") + 2))
  Call SearchData
  
ElseIf KeyAscii = 27 Then
  ItemCode.Visible = False
End If


End Sub
Private Sub ok_Click()
 
 
If binder_name.Text = "" Then
   MsgBox "Please Choose the Printers", vbInformation
   Me.binder_name.SetFocus
   Exit Sub
End If
    
'If TextPaperSize.Text = "" Then
'   MsgBox "Please Select Paper Size ...", vbInformation
'   Me.TextPaperSize.SetFocus
'   Exit Sub
'End If
 
 
 
 
If mode = Trim("edit") Then
    sq = "delete  from OrderPrint_Main where Ord_No = '" + txtOrdNo.Text + "' and " & stringyear
    CON.Execute sq
    sq = "delete  from OrderPrint_Det where Ord_No = '" + txtOrdNo.Text + "' and " & stringyear
    CON.Execute sq
End If
   
If RS.State = 1 Then RS.close
RS.Open "select * from OrderPrint_main where Ord_No = '" + txtOrdNo.Text + "' and " & stringyear, CON, adOpenDynamic, adLockOptimistic
If RS.EOF = True Then
   RS.AddNew
End If

RS!Ord_No = txtOrdNo.Text
RS!Ord_Date = txtOrdDate.value
RS!PrinterName = binder_name.Text
RS!TotalReam = Val(txtTReam.Caption)
RS!TotalSheet = Val(txtTSheet.Caption)
RS!PaperSize = Trim(TextPaperSize)
RS!OrderCancel = "n"
RS!papercode = Trim(txtPcode.Text)
RS!Address = Trim(lblAdd.Caption)

RS!fyear = session
RS!setupid = setupid
RS.update


If RS.State = 1 Then RS.close
RS.Open "select * from OrderPrint_det where Ord_No = '" + txtOrdNo.Text + "' and " & stringyear, CON, adOpenDynamic, adLockOptimistic
For K = 1 To vs.Rows - 1
If vs.TextMatrix(K, 1) <> "" Then

    RS.AddNew
    RS!Ord_No = txtOrdNo.Text
    RS!BCode = vs.TextMatrix(K, 0)
    RS!qty = Val(vs.TextMatrix(K, 2))
    RS![Inner] = Val(vs.TextMatrix(K, 3))
    RS!Text = Val(vs.TextMatrix(K, 4))
    RS!exam = Val(vs.TextMatrix(K, 5))
    RS!supp = Val(vs.TextMatrix(K, 6))
    RS!Title = Val(vs.TextMatrix(K, 7))
    RS!tpage = Val(vs.TextMatrix(K, 8))
    RS!DivdeBy = Val(vs.TextMatrix(K, 9))
    If vs.TextMatrix(K, 10) <> "" Then
       RS!TForm = vs.TextMatrix(K, 10)
    End If
    RS!WastPer = Val(vs.TextMatrix(K, 11))
    RS!TotalReam = Val(vs.TextMatrix(K, 12))
    RS!TotalSheet = Val(vs.TextMatrix(K, 13))
    RS!Binder = vs.TextMatrix(K, 14)
    RS!Hindi_English = vs.TextMatrix(K, 15)
    
    aa = Mid(vs.TextMatrix(K, 16), InStr(vs.TextMatrix(K, 16), ":") + 1)
    aa1 = Mid(aa, 1, 5)
    'aa1 = InStr(aa, ":")
    'aa1 = Mid(aa, 1, InStr(aa2, " : ") - 1)
    
    RS!PaperSize = vs.TextMatrix(K, 16)
    RS!Size = Trim(aa1)
    RS!pcode = Mid(vs.TextMatrix(K, 16), InStr(vs.TextMatrix(K, 16), "=>") + 2)
    
    RS!fyear = session
    RS!setupid = setupid
    RS.update
End If
Next

MsgBox "Data Saved ....", vbInformation
ok.Enabled = False
Edit.Enabled = True

End Sub
Sub SearchData()



grid_ini
   
If RS.State = 1 Then RS.close
RS.Open "select * from OrderPrint_main where Ord_No = '" + txtOrdNo.Text + "' and " & stringyear, CON
If RS.EOF = False Then
    ok.Enabled = False
    Edit.Enabled = True
    
    delete.Enabled = False
    cancel.Enabled = False
    
    
    txtPcode.Text = RS!papercode & ""
    
    txtOrdNo.Text = RS!Ord_No
    txtOrdDate.value = RS!Ord_Date
    binder_name.Text = RS!PrinterName
    txtTReam.Caption = RS!TotalReam
    txtTSheet.Caption = RS!TotalSheet
    TextPaperSize = RS!PaperSize
    
    If rs1.State = 1 Then rs1.close
    rs1.Open "select address from Godownmaster where Godwn ='" & binder_name.Text & "'", CON, adOpenKeyset, adLockReadOnly
    If rs1.EOF = False Then
       lblAdd.Caption = rs1!Address & ""
    End If
    
    If rs1.State = 1 Then rs1.close
    rs1.Open "select * from PaperMakeMaster where papermaker_id ='" & txtPcode & "'", CON
    If rs1.EOF = False Then
       lblPaper_det.Caption = "Paper Name : " & rs1!papermaker_name & vbCrLf & "Paper Type : " & rs1!ptype & vbCrLf & "Real/Sheets : " & rs1!Size & vbCrLf & "Qulality && G.S.M. : " & rs1!eco & " - " & rs1!GSM
    End If
    
    
    
End If




If RS.State = 1 Then RS.close
RS.Open "select * from OrderPrint_det where Ord_No = '" + txtOrdNo.Text + "' and " & stringyear, CON
For K = 1 To vs.Rows - 1
    
If RS.EOF = False Then
    vs.TextMatrix(K, 0) = RS!BCode
    
    'rs1.MoveFirst
    'rs1.Find "bookcode='" & RS!BCode & "'"
    
    If rs1.State = 1 Then rs1.close
    rs1.Open "select book,bookfont from BookMaster where BookNo='" & RS!BCode & "' and " & stringyear, CON, adOpenKeyset, adLockReadOnly

    If rs1.EOF = False Then
        
        If RS!Hindi_English = "h" Then
             vs.Cell(flexcpFontName, K, 1) = hindi
             vs.Cell(flexcpFontSize, K, 1) = 14
             vs.TextMatrix(K, 15) = "h"
         Else
             vs.Cell(flexcpFontName, K, 1) = english
             vs.Cell(flexcpFontSize, K, 1) = 10
             vs.TextMatrix(K, 15) = "e"
         End If
         vs.TextMatrix(K, 1) = rs1!book
         
     End If
    
    
    
    vs.TextMatrix(K, 2) = RS!qty
    vs.TextMatrix(K, 3) = RS![Inner]
    vs.TextMatrix(K, 4) = RS!Text
    vs.TextMatrix(K, 5) = RS!exam
    vs.TextMatrix(K, 6) = RS!supp
    vs.TextMatrix(K, 7) = RS!Title
    vs.TextMatrix(K, 8) = RS!tpage
    vs.TextMatrix(K, 9) = RS!DivdeBy
    vs.TextMatrix(K, 10) = RS!TForm & ""
    vs.TextMatrix(K, 11) = RS!WastPer
    vs.TextMatrix(K, 12) = RS!TotalReam
    vs.TextMatrix(K, 13) = RS!TotalSheet
    vs.TextMatrix(K, 14) = RS!Binder
    
    vs.TextMatrix(K, 16) = RS!PaperSize & ""
    
    RS.MoveNext
End If

Next

addTotalReam


End Sub


Private Sub order_no_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   TextPaperSize.SetFocus
End If
End Sub

Private Sub order_no_LostFocus()
 'Binder_id.SetFocus
End Sub

Private Sub party_id_GotFocus()
'''Label10.Visible = True
''''Label10.Caption = "Press F2 For search"
'''If PopUpValue1 <> "" Then
'''  Me.party_id.Text = PopUpValue1
'''  Me.Party_name.Text = Trim(PopUpValue2) + " " + Trim(PopUpValue3)
'''   End If
'''PopUpValue1 = ""
'''PopUpValue2 = ""
'''PopUpValue3 = ""
End Sub
Private Sub party_id_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
   popuplist1 "Select customer_id as [Printer Id],customer_name as [Printer Name],city from customerMaster where Category='Printer'", CON
ElseIf KeyCode = 13 Then
   TextPaperSize.SetFocus
End If
End Sub
Private Sub Party_id_KeyPress(KeyAscii As Integer)
''If KeyAscii = 13 Then
''If party_id.Text = "" Then
''MsgBox "Please Choose the Printers ...", vbInformation
''party_id.SetFocus
''Exit Sub
''Else
''Dim rsparty As ADODB.Recordset
''Set rsparty = New ADODB.Recordset
''sq = "Select customer_id,customer_name,city,pscustomerid from customerMaster where customer_id = '" + Me.party_id.Text + "'"
''rsparty.Open sq, CON, adOpenKeyset, adLockReadOnly
''    If rsparty.RecordCount > 0 Then
''    Me.Party_name.Text = Trim(rsparty.Fields("customer_name")) + " " + Trim(rsparty.Fields("city"))
''    If rsparty!pscustomerid <> "" Then Me.pscustomerid = rsparty!pscustomerid Else Me.pscustomerid = rsparty!customerid
''    'order_no.SetFocus
''    Else
''    MsgBox "Invalid Party/Customer .. ", vbInformation
''    Me.party_id.SetFocus
''    Me.party_id.Text = ""
''    Exit Sub
''    End If
''End If
''End If
End Sub

Private Sub party_id_LostFocus()
Label10.Visible = False
'order_no.SetFocus
End Sub
Private Sub Printcmd_Click()

On Error GoTo abc:

Dim reem, sheet, a1, per
CON.Execute "delete from tmps_LEDGER1"

sheet = 0

If rs1.State = 1 Then rs1.close
rs1.Open "select  papersize,sum(TotalReam),sum(TotalSheet) size from OrderPrint_Det where Ord_No = '" + txtOrdNo.Text + "' and " & stringyear & " group by papersize"
While rs1.EOF = False
   
   a1 = rs1(2)
   ream = Int(rs1(1))
   
   
   
   If a1 > 499 Then
      per = Int(a1 / 500)
      r_sheet = a1 - per * 500
      ream = ream + per
      sheet = sheet + r_sheet
   Else
      sheet = a1
   End If
    
   
    
   CON.Execute "insert into tmps_LEDGER1(address1,address2,address3) values('" & Mid(rs1(0), 1, Len(rs1(0)) - 4) & "','" & ream & "','" & sheet & "')"
  
   rs1.MoveNext
   
Wend

DoEvents
DoEvents
DoEvents


Screen.MousePointer = vbHourglass

cr1.Reset
cr1.ReportFileName = App.Path & "/report_paper/printorder.rpt"
cr1.Connect = "filedsn=chitradsn;uid=sa;pwd=sidc;"
cr1.ReplaceSelectionFormula "{orderPrint_Main.ord_no}='" & txtOrdNo.Text & "'"
cr1.WindowShowPrintSetupBtn = True
cr1.WindowState = crptMaximized
cr1.Action = 1

Screen.MousePointer = vbDefault


    Exit Sub
abc:
    MsgBox "" & Err.DESCRIPTION



End Sub

Private Sub commandQuit_Click()
'MainMenu.bill = False
Unload Me
End Sub

Private Sub search_Click()
'popuplist1 "Select bill_id as Bill_No,total as Total_Amount from billmaster where firm_id = '" + Me.Textfirmid.Text + "'", CON
End Sub
Private Sub billupdate()

'''ok.Enabled = False
'''Edit.Enabled = True
'''Printcmd.Enabled = True
'''
'''Dim rsbilmast As ADODB.Recordset
'''Dim rsbiltrans As ADODB.Recordset
'''Dim rsparty As ADODB.Recordset
'''Dim rsbinder As ADODB.Recordset
'''Dim rssize As ADODB.Recordset
'''Set rsbilmast = New ADODB.Recordset
'''Set rsbiltrans = New ADODB.Recordset
'''Set rsparty = New ADODB.Recordset
'''Set rsbinder = New ADODB.Recordset
'''Set rssize = New ADODB.Recordset
'''Dim i, j As Integer
'''Dim fid As String
'''
'''totalwsht = 0
'''toalwream = 0
'''
'''
'''If mode <> "add" Then
'''
'''If bill_no <> "" Then
'''sq = "Select * from billmaster where bill_id = '" + bill_no + "' and categories='Main' order by cint(bill_id) asc"
'''rsbilmast.Open sq, CON, adOpenKeyset, adLockReadOnly
'''
'''If rsbilmast.RecordCount = 0 Then
'''   Exit Sub
'''End If
'''
'''If rsbilmast.RecordCount > 0 Then
'''    Me.bill_no.Text = rsbilmast.Fields("bill_id")
'''
'''   'If rsbilmast.Fields("pscustomerid") <> "" Then Me.pscustomerid.Text = rsbilmast.Fields("pscustomerid")
'''   'If rsbilmast.Fields("pstatementno") <> "" Then Me.pstno.Text = rsbilmast.Fields("pstatementno")
'''
'''
'''pltamt = rsbilmast.Fields("totalplate")
'''total = rsbilmast.Fields("totalprint")
'''gtotalamt = rsbilmast.Fields("total")
'''totalwream = rsbilmast.Fields("totalwrim")
'''totalwsht = rsbilmast.Fields("totalwsht")
'''gtotalreams = rsbilmast.Fields("totalrim")
'''totalplate = rsbilmast.Fields("totalnoplate")
'''gtotalsheets = rsbilmast.Fields("totalsht")
'''totalream = rsbilmast.Fields("totalnoream")
'''totalsht = rsbilmast.Fields("totalnosht")
'''Me.DTPicker1.value = rsbilmast.Fields("dat")
'''    'Me.order_no.Text = rsbilmast.Fields("order_no")
'''    Me.Binder_id.Text = rsbilmast.Fields("Binder_id")
'''
'''    If rsbilmast.Fields("papersize1") <> "" Then Me.TextPaperSize.Text = rsbilmast.Fields("papersize1")
'''    sq = "Select customer_name from CustomerMaster where customer_id = '" + rsbilmast.Fields("binder_id") + "'"
'''    rsbinder.Open sq, CON, adOpenKeyset, adLockReadOnly
'''    If rsbinder.RecordCount > 0 Then
'''        Me.binder_name.Text = rsbinder.Fields("customer_name")
'''        'Me.Textfirmid.Text = rsbilmast.Fields("firm_id")
'''    End If
''''       If rsbilmast.Fields("calcmanual") = True Then Checkmanual.value = 1 Else Checkmanual.value = 0
'''If rsbilmast!previousebilldate <> "" Then previousbilldate.value = rsbilmast!previousebilldate
'''If Val(pstno) > 0 Then CommandPAPERSTATEMENT.Enabled = True Else CommandPAPERSTATEMENT.Enabled = False
'''
'''If rsbilmast!previousebillid <> "" Then previousbillno = rsbilmast!previousebillid
'''    sq = ""
'''    sq = "Select * from billtrans where trim(bill_id) = '" + Trim(rsbilmast.Fields("bill_id")) + "'"
'''    rsbiltrans.Open sq, CON, adOpenKeyset, adLockReadOnly
'''    If rsbiltrans.RecordCount > 0 Then
'''        grid_ini
'''        i = 1
'''        Me.vs.Rows = rsbiltrans.RecordCount + 1
'''        Do While Not rsbiltrans.EOF
'''
'''
'''            Me.vs.TextMatrix(i, 0) = rsbiltrans.Fields("BookNo")
'''            Me.vs.Row = i
'''            'Me.vs.Col = 1
'''            'Me.vs.CellFontSize = 14
'''
'''            If rsbiltrans!fontbk = "h" Then
'''            Me.vs.CellFontName = hindi
'''             vs.CellFontSize = 14
'''            Else
'''            Me.vs.CellFontName = english
'''            End If
'''
'''            Me.vs.TextMatrix(i, 1) = rsbiltrans.Fields("particulars")
'''            'Me.vs.Col = 2
'''
'''           ' If rsbiltrans!fontbk = "h" Then
'''            'Me.vs.CellFontName = hindi
'''            'Else
'''            Me.vs.CellFontName = english
'''            'End If
'''            Dim rs1 As New ADODB.Recordset
'''
'''            If rsbiltrans.Fields("bookinfo") <> "" Then
'''
'''                If rs1.State = 1 Then rs1.close
'''                rs1.Open "select Binder_name from BinderMaster where Binder_id='" & rsbiltrans.Fields("bookinfo") & "'", CON
'''                If rs1.EOF = False Then
'''                   Me.vs.Text = rs1.Fields("Binder_name").value & ""
'''                End If
'''
'''            End If
'''
'''            Me.vs.Col = 3
'''            Me.vs.Text = rsbiltrans.Fields("unit")
'''            Me.vs.Col = 4
'''            Me.vs.Text = rsbiltrans.Fields("type")
'''            Me.vs.Col = 5
'''            Me.vs.Text = rsbiltrans.Fields("quantity")
'''            Me.vs.Col = 6
'''            Me.vs.Text = rsbiltrans.Fields("atrate")
'''            Me.vs.Col = 7
'''            Me.vs.Text = rsbiltrans.Fields("rate")
'''            Me.vs.Col = 8
'''            Me.vs.Text = rsbiltrans.Fields("amt")
'''            Me.vs.Col = 9
'''            Me.vs.Text = rsbiltrans.Fields("reams")
'''            Me.vs.Col = 10
'''            Me.vs.Text = rsbiltrans.Fields("sheets")
'''            Me.vs.Col = 11
'''            Me.vs.Text = rsbiltrans.Fields("wastage")
'''            Me.vs.Col = 12
'''            Me.vs.Text = rsbiltrans.Fields("wreams")
'''            Me.vs.Col = 13
'''            Me.vs.Text = rsbiltrans.Fields("wsheets")
'''            Me.vs.Col = 14
'''            Me.vs.Text = rsbiltrans.Fields("plate")
'''            Me.vs.Col = 15
'''            Me.vs.Text = rsbiltrans.Fields("patrate")
'''            Me.vs.Col = 16
'''
'''            Me.vs.Text = rsbiltrans.Fields("pamt")
'''
'''            Me.vs.Col = 17
'''            Me.vs.Text = rsbiltrans.Fields("inner") & ""
'''
'''            Me.vs.Col = 18
'''            Me.vs.TextMatrix(i, 18) = rsbiltrans.Fields("PaparMake").value & ""
'''
'''            Me.vs.Col = 19
'''            Me.vs.TextMatrix(i, 19) = rsbiltrans.Fields("supp") & ""
'''
'''            Me.vs.Col = 20
'''            Me.vs.TextMatrix(i, 20) = rsbiltrans.Fields("Note").value & ""
'''
'''            Me.vs.Col = 21
'''            Me.vs.TextMatrix(i, 21) = rsbiltrans.Fields("PAPERsize1").value & ""
'''
'''            Me.vs.Col = 22
'''            Me.vs.TextMatrix(i, 22) = rsbiltrans.Fields("bookno").value & ""
'''
'''
'''
'''''            vs.TextMatrix(i, 23) = rsbiltrans.Fields("inn_Printer").value & ""
'''''            vs.TextMatrix(i, 24) = rsbiltrans.Fields("text_Printer").value & ""
'''''            vs.TextMatrix(i, 25) = rsbiltrans.Fields("Exam_Printer").value & ""
'''''            vs.TextMatrix(i, 26) = rsbiltrans.Fields("Supp_Printer").value & ""
'''''            vs.TextMatrix(i, 27) = rsbiltrans.Fields("Title_Printer").value & ""
'''
'''
'''
'''
'''
'''            If Not rsbiltrans.EOF Then
'''                rsbiltrans.MoveNext
'''                i = i + 1
'''            End If
'''        Loop
'''    End If
'''Else
'''    MsgBox "Bill not found .. ", vbInformation
'''End If
'''End If
'''End If
End Sub

Private Sub Textfirmname_GotFocus()
  ref
End Sub
Sub ref()
mode = ""

frmbill.Enabled = True



Me.Add.SetFocus

'End If
PopUpValue1 = ""
PopUpValue2 = ""
PopUpValue3 = ""

End Sub
Private Sub Textfirmname_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
popuplist1 "Select firm_name,firm_id, firmpictfilename from firmMaster", CON
End If
End Sub

Private Sub textpapersize_GotFocus()

If PopUpValue1 <> "" Then
   TextPaperSize.Text = PopUpValue1
   txtPcode.Text = PopUpValue6
   
   lblPaper_det.Caption = "Paper Name : " & PopUpValue2 & vbCrLf & "Paper Type : " & PopUpValue3 & vbCrLf & "Real/Sheets : " & popupvalue4 & vbCrLf & "Qulality && G.S.M. : " & popupvalue5
   
   PopUpValue1 = ""
   PopUpValue2 = ""
   PopUpValue3 = ""
   popupvalue4 = ""
   

End If

End Sub

Private Sub textpapersize_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
'   popuplist1 "Select size1 as [Size],size_info as [Size Info] from sizeMaster where size1 <> ''", CON


   
  searchType = "paper"
   value = "Select SizeValue1 + 'X'+ SizeValue2 as [Paper Size]," & _
   "papermaker_name as [Paper Name],PType,Size as [Sheet/Real],Eco + ' - ' + GSM  as [Quality & GSM],papermaker_Id as Code from " & _
   " papermakemaster where papermaker_id <> '' order by SizeValue1"
    popuplistModel10 value, CON
       

End If


End Sub
Sub Clearvalue()

ok.Enabled = True
Edit.Enabled = False
delete.Enabled = False
cancel.Enabled = False

vs.Clear
grid_ini

txtPcode = ""
lblPaper_det.Caption = ""
Me.txtOrdNo.Text = ""
TextPaperSize = ""
lblAdd.Caption = ""
binder_name = ""

txtTReam = ""
txtTSheet = ""

PopUpValue1 = ""
PopUpValue2 = ""
PopUpValue3 = ""
popupvalue4 = ""

End Sub
Private Sub TextPaperSize_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   vs.SetFocus
   vs.Col = 0
End If

''If TextPaperSize.Text = "" Then
''  MsgBox "Please Choose the Paper Size ..", vbInformation
''  TextPaperSize.SetFocus
''Exit Sub
''Else
''Dim rssize As ADODB.Recordset
''Set rssize = New ADODB.Recordset
''sq = "Select size1,size_info from SizeMaster where size1 = '" + Me.TextPaperSize.Text + "'"
''rssize.Open sq, CON, adOpenKeyset, adLockReadOnly
''    If rssize.RecordCount > 0 Then
''    Me.TextPaperSize.Text = Trim(rssize.Fields("size1"))
''    Me.binder_name.Enabled = True
''    Me.binder_name.SetFocus
''    Else
''    MsgBox "Invalid Paper Size ... ", vbInformation
''    Me.TextPaperSize.SetFocus
''    Me.TextPaperSize.Text = ""
''    Exit Sub
''    End If
''End If
''End If

End Sub

Private Sub TextPaperSize_LostFocus()
'''Label10.Visible = False
''If TextPaperSize.Text = "" Then
'''MsgBox "Please Choose the Paper Size"
'''textpapersize.SetFocus
''Exit Sub
''Else
''Dim rssize As ADODB.Recordset
''Set rssize = New ADODB.Recordset
''sq = "Select size1,size_info from SizeMaster where size1 = '" + Me.TextPaperSize.Text + "'"
''rssize.Open sq, CON, adOpenKeyset, adLockReadOnly
''    If rssize.RecordCount > 0 Then
''    Me.TextPaperSize.Text = Trim(rssize.Fields("size1"))
''    binder_name.Enabled = True
''    binder_name.SetFocus
''    Else
''    MsgBox "Invalid Paper Size ... ", vbInformation
''    Me.TextPaperSize.SetFocus
''    Me.TextPaperSize.Text = ""
''    Exit Sub
''    End If
''End If

End Sub
Private Sub txtNote_GotFocus()
             
'vs.TextMatrix(vs.RowSel, 20) = txtNote.Text
'txtNote.Text = ""
'txtNote.Visible = False
'vs.Col = 20

Call vs_GotFocus


End Sub

Private Sub txtNote_KeyDown(KeyCode As Integer, Shift As Integer)
       If KeyCode = 27 Then
          If txtNote.Visible = True Then
             vs.TextMatrix(vs.RowSel, 20) = txtNote.Text
             txtNote.Text = ""
             txtNote.Visible = False
             vs.Col = 20
             vs.SetFocus
          End If
       End If
End Sub

Private Sub VSFlexGrid1_Click()

End Sub
