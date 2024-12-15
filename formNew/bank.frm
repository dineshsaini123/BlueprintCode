VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBank 
   Caption         =   "Bank"
   ClientHeight    =   7752
   ClientLeft      =   168
   ClientTop       =   456
   ClientWidth     =   14724
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7752
   ScaleWidth      =   14724
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboPrint 
      Height          =   315
      ItemData        =   "bank.frx":0000
      Left            =   5130
      List            =   "bank.frx":000A
      TabIndex        =   28
      Top             =   6435
      Width           =   1425
   End
   Begin VB.TextBox txtsrno 
      Height          =   285
      Left            =   8745
      TabIndex        =   25
      Top             =   810
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Print With Party"
      Height          =   435
      Left            =   8010
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6375
      Width           =   1140
   End
   Begin Crystal.CrystalReport cr 
      Left            =   9285
      Top             =   5730
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileUseRptNumberFmt=   -1  'True
      PrintFileUseRptDateFmt=   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Print"
      Height          =   435
      Left            =   6780
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6390
      Width           =   1215
   End
   Begin VB.TextBox txtsno 
      Height          =   315
      Left            =   1575
      TabIndex        =   0
      Top             =   150
      Width           =   1455
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Exit"
      Height          =   615
      Left            =   5595
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5475
      Width           =   1110
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancel"
      Height          =   615
      Left            =   4500
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5475
      Width           =   1110
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Delete"
      Height          =   615
      Left            =   3405
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5475
      Width           =   1110
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Update"
      Height          =   615
      Left            =   2310
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5475
      Width           =   1110
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Save"
      Height          =   615
      Left            =   1215
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5475
      Width           =   1110
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Refresh"
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5475
      Width           =   1110
   End
   Begin VSFlex7DAOCtl.VSFlexGrid vs5 
      Height          =   4035
      Left            =   90
      TabIndex        =   6
      Top             =   1335
      Width           =   9420
      _cx             =   16616
      _cy             =   7117
      _ConvInfo       =   1
      Appearance      =   2
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
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
      Rows            =   50
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
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
      Editable        =   1
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.TextBox txtAmount 
      Height          =   285
      Left            =   5325
      TabIndex        =   4
      Top             =   510
      Width           =   2685
   End
   Begin VB.TextBox txtDdno 
      Height          =   315
      Left            =   5325
      TabIndex        =   3
      Top             =   150
      Width           =   2700
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1575
      TabIndex        =   1
      Top             =   480
      Width           =   1455
      _ExtentX        =   2582
      _ExtentY        =   550
      _Version        =   393216
      Format          =   155582465
      CurrentDate     =   39305
   End
   Begin VB.ComboBox comboPartyName 
      Height          =   315
      Left            =   5325
      TabIndex        =   5
      Top             =   810
      Width           =   3465
   End
   Begin VB.ComboBox comboNameOfTheBank 
      Height          =   315
      Left            =   1545
      TabIndex        =   2
      Top             =   870
      Width           =   2550
   End
   Begin MSComCtl2.DTPicker fromDate 
      Height          =   315
      Left            =   7815
      TabIndex        =   21
      Top             =   5715
      Width           =   1275
      _ExtentX        =   2265
      _ExtentY        =   572
      _Version        =   393216
      Format          =   155582465
      CurrentDate     =   39305
   End
   Begin MSComCtl2.DTPicker toDate 
      Height          =   315
      Left            =   7815
      TabIndex        =   23
      Top             =   6030
      Width           =   1275
      _ExtentX        =   2265
      _ExtentY        =   550
      _Version        =   393216
      Format          =   155582465
      CurrentDate     =   39305
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Print By"
      Height          =   195
      Left            =   5145
      TabIndex        =   27
      Top             =   6210
      Width           =   720
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Press Delete Button For Delete Row"
      Height          =   255
      Left            =   285
      TabIndex        =   26
      Top             =   6345
      Width           =   3915
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      Height          =   195
      Left            =   6945
      TabIndex        =   24
      Top             =   6060
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   1275
      Left            =   6765
      Top             =   5595
      Width           =   2460
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      Height          =   195
      Left            =   6945
      TabIndex        =   22
      Top             =   5850
      Width           =   345
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Id"
      Height          =   195
      Left            =   195
      TabIndex        =   20
      Top             =   285
      Width           =   135
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Party Name"
      Height          =   195
      Left            =   4485
      TabIndex        =   17
      Top             =   870
      Width           =   825
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   195
      Left            =   4605
      TabIndex        =   16
      Top             =   510
      Width           =   795
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name of the Bank"
      Height          =   195
      Left            =   165
      TabIndex        =   15
      Top             =   930
      Width           =   1350
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "D.D.No."
      Height          =   195
      Left            =   4605
      TabIndex        =   14
      Top             =   210
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   195
      Left            =   195
      TabIndex        =   13
      Top             =   600
      Width           =   345
   End
End
Attribute VB_Name = "frmBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim J As Integer
Private Sub cmdAdd_Click()

'Form_Load
J = 10
txtDdno.Text = ""
txtAmount.Text = ""
comboNameOfTheBank.Text = ""
comboPartyName.Text = ""
cmdAdd.Enabled = False
cmdSave.Enabled = True
cmdUpdate.Enabled = False
cmdDelete.Enabled = False
cmdCancel.Enabled = True
cmdExit.Enabled = True
Me.comboNameOfTheBank.SetFocus
vs5.Clear
setWidth
Call MaxNo
'showGridData
End Sub
Sub MaxNo()
If RS.State = 1 Then RS.close
RS.Open "select max(sno) from bankstm where " & stringyear, con
If IsNull(RS(0)) Then
txtsno.Text = 1
Else
txtsno.Text = (Val(RS(0)) + 1)
End If
End Sub

Private Sub cmdCancel_Click()

txtsno.Text = ""
txtDdno.Text = ""
txtAmount.Text = ""
comboNameOfTheBank.Text = ""
comboPartyName.Text = ""
cmdAdd.Enabled = True
cmdSave.Enabled = True
cmdUpdate.Enabled = False
cmdDelete.Enabled = False
cmdCancel.Enabled = True
cmdExit.Enabled = True
txtsno.SetFocus
End Sub

Private Sub cmdDelete_Click()
'If rs.State = 1 Then rs.Close
'rs.Open "select * from bankstm", con, adOpenDynamic, adLockOptimistic
If MsgBox("Want To Delete ?", vbInformation + vbYesNo) = vbYes Then
con.Execute "delete * from bankstm where " & stringyear & " and sno=" & txtsno.Text & ""
showGridData
End If
End Sub
Private Sub cmdexit_Click()
Unload Me
End Sub
Sub saveData()

If RS.State = 1 Then RS.close
RS.Open "select * from bankstm where " & stringyear & " and sno=" & txtsno.Text & "", con, adOpenDynamic, adLockOptimistic
If RS.EOF = True Then
For I = 1 To vs5.rows - 1
If vs5.TextMatrix(I, 1) <> "" Then
    RS.AddNew
        RS.Fields("nameOfTheBank") = vs5.TextMatrix(I, 1)
            RS.Fields("dated") = vs5.TextMatrix(I, 2)
                RS.Fields("ddNo").value = vs5.TextMatrix(I, 3)
                    RS.Fields("partyName") = vs5.TextMatrix(I, 4)
                RS.Fields("sno") = Val(txtsno.Text)
            RS.Fields("amount") = Val(vs5.TextMatrix(I, 5))
        RS.Fields("setupid") = setupid
    RS.Fields("fyear") = session
    RS.update
End If
Next
End If


comboNameOfTheBank.Clear
If RS.State = 1 Then RS.close
RS.Open "select distinct(nameOfTheBank) from bankstm where " & stringyear, con, adOpenDynamic, adLockOptimistic
While RS.EOF = False
If Not IsNull(RS.Fields("nameOfTheBank").value) Then
comboNameOfTheBank.AddItem RS.Fields("nameOfTheBank").value
End If
RS.MoveNext
Wend



'Call cmdAdd_Click

End Sub
Private Sub cmdPrint_Click()

DSNNew

cr.Reset
cr.ReportFileName = rptPath & "\BankReg.rpt"
cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
'cr.DataFiles(0) = st1 + "\" + Trim(main.directory) & "\data.mdb"
If cboPrint.ListIndex = 0 Then
cr.ReplaceSelectionFormula "{BILTYRETURNREGISTER.sno}=" & txtsno.Text & ""
Else
cr.ReplaceSelectionFormula "{BILTYRETURNREGISTER.dated}>=datevalue('" & Format(FromDate.value, "MM/dd/yyyy") & "') and {BILTYRETURNREGISTER.dated}<=datevalue('" & Format(toDate.value, "MM/dd/yyyy") & "')"
End If
cr.WindowShowPrintSetupBtn = True
cr.WindowShowPrintBtn = True
cr.WindowState = crptMaximized
cr.Action = 1
End Sub
Private Sub cmdSave_Click()
saveData

MsgBox "Record Saved", vbInformation
DTPicker1.SetFocus

cmdAdd.Enabled = True
cmdSave.Enabled = False
cmdUpdate.Enabled = False
cmdDelete.Enabled = False
txtAmount.Text = ""
txtDdno.Text = ""
comboNameOfTheBank.Text = ""
comboPartyName.Text = ""
'showGridData
'Call cmdAdd_Click
'maxNo
End Sub
Private Sub cmdUpdate_Click()
If MsgBox("Want To Update ?", vbInformation + vbYesNo) = vbYes Then

con.BeginTrans
con.Execute "delete from bankstm where " & stringyear & " and sno=" & txtsno.Text & ""
saveData
con.CommitTrans


End If

End Sub
Private Sub comboPartyName_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    J = 0
    If vs5.TextMatrix(1, 1) = "" Then
    J = 1
    Else
    
    For I = 0 To vs5.rows - 1
    If vs5.TextMatrix(I, 1) <> "" Then
    J = J + 1
    End If
    Next
    
    End If
 
    vs5.TextMatrix(J, 0) = J
    vs5.TextMatrix(J, 1) = comboNameOfTheBank.Text
    vs5.TextMatrix(J, 2) = DTPicker1.value
    vs5.TextMatrix(J, 3) = txtDdno.Text
    vs5.TextMatrix(J, 4) = comboPartyName.Text
    vs5.TextMatrix(J, 5) = txtAmount.Text
    J = J + 1
    
    
    
    If txtAmount.Text = "" Then
       cmdSave.SetFocus
    Else
        txtDdno.Text = ""
        comboNameOfTheBank.Text = ""
        comboPartyName.Text = ""
        txtAmount.Text = ""
        DTPicker1.SetFocus
    End If
    
 End If

End Sub
Private Sub Command1_Click()


DSNNew

cr.Reset
cr.ReportFileName = rptPath & "\BankReg1.rpt"
cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
''cr.DataFiles(0) = st1 + "\" + Trim(main.directory) & "\data.mdb"
If cboPrint.ListIndex = 0 Then
   cr.ReplaceSelectionFormula "{BILTYRETURNREGISTER.sno}=" & txtsno.Text & ""
Else
   cr.ReplaceSelectionFormula "{BILTYRETURNREGISTER.dated}>=datevalue('" & Format(FromDate.value, "MM/dd/yyyy") & "') and {BILTYRETURNREGISTER.dated}<=datevalue('" & Format(toDate.value, "MM/dd/yyyy") & "')"
End If
cr.WindowState = crptMaximized
cr.WindowShowPrintSetupBtn = True
cr.WindowShowPrintBtn = True
cr.Action = 1


End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("txtAmount")) Then
   SendKeys "{tab}"
End If
End If
End Sub

Private Sub Form_Load()
cboPrint.ListIndex = 0
MaxNo
txtDdno.Text = ""
txtAmount.Text = ""
comboNameOfTheBank.Text = ""
comboPartyName.Text = ""
cmdAdd.Enabled = True
cmdSave.Enabled = True
cmdUpdate.Enabled = False
cmdDelete.Enabled = False
cmdCancel.Enabled = True
cmdExit.Enabled = True
setWidth
showGridData

comboPartyName.Clear
If RS.State = 1 Then RS.close
RS.Open "select distinct(subledger) from SLEDGER", con
While RS.EOF = False
comboPartyName.AddItem RS(0)
RS.MoveNext
Wend
FromDate.value = Date
toDate.value = Date
DTPicker1.value = Date
J = 1



comboNameOfTheBank.Clear
If RS.State = 1 Then RS.close
RS.Open "select distinct(nameOfTheBank) from bankstm where " & stringyear, con, adOpenDynamic, adLockOptimistic
While RS.EOF = False
If Not IsNull(RS.Fields("nameOfTheBank").value) Then
comboNameOfTheBank.AddItem RS.Fields("nameOfTheBank").value
End If
RS.MoveNext
Wend


BackColorFrom Me

End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  comboPartyName.SetFocus
 End If

End Sub
Private Sub txtsno_GotFocus()
If PopUpValue1 <> "" Then
showGridData
End If
End Sub

Private Sub txtsno_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
 value = "select Sno as EntryId,dated from bankstm group by sno,dated"
 popuplistModel10 value, con
 
End If
End Sub
Private Sub txtsno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If RS.State = 1 Then RS.close
RS.Open "select * from bankstm where " & stringyear & " and sno=" & txtsno.Text & "", con, adOpenDynamic, adLockOptimistic
If RS.EOF = False Then
End If
End If


End Sub
Sub setWidth()
vs5.FormatString = "S.No.|Name of the Bank|Date|DD NO.|Party Name|Amount"
vs5.ColWidth(0) = 500
vs5.ColWidth(1) = 1500
vs5.ColWidth(2) = 800
vs5.ColWidth(3) = 1000
vs5.ColWidth(4) = 4000
vs5.ColWidth(5) = 1200
vs5.ColWidth(6) = 0

End Sub

Sub showGridData()

If PopUpValue1 = "" Then Exit Sub

vs5.Clear
setWidth
txtsno.Text = PopUpValue1
If RS.State = 1 Then RS.close
RS.Open "select * from bankstm where " & stringyear & " and convert(smalldatetime,dated,103)=convert(smalldatetime,'" & PopUpValue2 & "',103) and sno=" & PopUpValue1 & "", con, adOpenDynamic, adLockOptimistic
If RS.EOF = True Then
cmdDelete.Enabled = False
cmdUpdate.Enabled = False
cmdSave.Enabled = True

End If

If RS.EOF = False Then

cmdDelete.Enabled = True
cmdUpdate.Enabled = True
cmdSave.Enabled = False


DTPicker1.value = RS!Dated
For I = 0 To RS.RecordCount - 1
vs5.TextMatrix(I + 1, 0) = I + 1
vs5.TextMatrix(I + 1, 1) = RS.Fields("nameOfTheBank")
vs5.TextMatrix(I + 1, 2) = RS.Fields("dated")
vs5.TextMatrix(I + 1, 3) = RS.Fields("ddNo")
vs5.TextMatrix(I + 1, 4) = RS.Fields("PartyName")
vs5.TextMatrix(I + 1, 5) = RS.Fields("amount")
vs5.TextMatrix(I + 1, 6) = RS.Fields("autoNo")
RS.MoveNext
Next
End If
vs5.Refresh

comboNameOfTheBank.Clear
If RS.State = 1 Then RS.close
RS.Open "select distinct(nameOfTheBank) from bankstm where " & stringyear, con, adOpenDynamic, adLockOptimistic
While RS.EOF = False
If Not IsNull(RS.Fields("nameOfTheBank").value) Then
comboNameOfTheBank.AddItem RS.Fields("nameOfTheBank").value
End If
RS.MoveNext
Wend

PopUpValue1 = ""
PopUpValue2 = ""


End Sub
Sub search()
On Error Resume Next
If RS.State = 1 Then RS.close
RS.Open "select * from bankstm where " & stringyear & " and autoNo=" & vs5.TextMatrix(vs5.RowSel, 6) & "", con, adOpenDynamic, adLockOptimistic
If RS.EOF = False Then


txtAmount.Text = ""
txtDdno.Text = ""
comboNameOfTheBank.Text = ""
comboPartyName.Text = ""
DTPicker1.value = Date

'txtsno.SetFocus
txtsno.Text = RS!sno
txtAmount.Text = RS!amount
txtDdno.Text = RS!ddNo
comboNameOfTheBank.Text = RS!nameOfTheBank
comboPartyName.Text = RS!partyname & ""
DTPicker1.value = RS!Dated
txtsrno.Text = RS!autono
J = vs5.Row
End If
End Sub
Private Sub vs5_Click()
search
End Sub
Private Sub vs5_KeyDown(KeyCode As Integer, Shift As Integer)

If txtsrno.Text = "" Then Exit Sub

If KeyCode = 46 Then
If MsgBox("Want to delete ?", vbQuestion + vbYesNo) = vbYes Then
   con.Execute "delete from bankstm where " & stringyear & " and autoNo=" & txtsrno.Text & ""
   vs5.RemoveItem (vs5.RowSel)
   vs5.rows = vs5.rows + 1
End If
End If
End Sub
