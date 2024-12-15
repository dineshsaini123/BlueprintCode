VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPartyProfile 
   Caption         =   "Party Profile"
   ClientHeight    =   9588
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   16716
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9588
   ScaleWidth      =   16716
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboProfile 
      Height          =   288
      ItemData        =   "frmPartyProfile.frx":0000
      Left            =   4884
      List            =   "frmPartyProfile.frx":0002
      TabIndex        =   16
      Top             =   504
      Width           =   2004
   End
   Begin VB.CommandButton cmdRepQty 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Export to Excel"
      Height          =   555
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   180
      Width           =   1230
   End
   Begin VB.ComboBox cbostate 
      Height          =   288
      Left            =   4884
      TabIndex        =   2
      Top             =   180
      Width           =   2016
   End
   Begin VB.CommandButton cmdPrint1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Print -1"
      Height          =   300
      Left            =   13020
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   90
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.CommandButton cmdprintnotpad 
      Caption         =   "&Export To NotePad"
      Height          =   270
      Left            =   13260
      TabIndex        =   11
      Top             =   60
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Exit"
      Height          =   570
      Left            =   11715
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   180
      Width           =   1125
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Filter Mobile"
      Height          =   270
      Left            =   13500
      TabIndex        =   9
      Top             =   60
      Visible         =   0   'False
      Width           =   210
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   900
      Top             =   5760
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Print"
      Height          =   570
      Left            =   9255
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   180
      Width           =   1125
   End
   Begin VB.TextBox txtParty 
      Height          =   285
      Left            =   660
      TabIndex        =   1
      Top             =   480
      Width           =   3660
   End
   Begin VB.CommandButton cmdRef 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Refresh"
      Height          =   570
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   180
      Width           =   1125
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Save"
      Height          =   570
      Left            =   8070
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   180
      Width           =   1155
   End
   Begin VB.ComboBox cboCity 
      Height          =   315
      Left            =   660
      TabIndex        =   0
      Top             =   180
      Width           =   3660
   End
   Begin VSFlex7Ctl.VSFlexGrid vs1 
      Height          =   8100
      Left            =   120
      TabIndex        =   3
      Top             =   1116
      Width           =   16524
      _cx             =   29146
      _cy             =   14287
      _ConvInfo       =   1
      Appearance      =   1
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
      BackColor       =   16771022
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   16771022
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   4
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   600
      RowHeightMax    =   0
      ColWidthMin     =   1000
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
      ExplorerBar     =   7
      PicturesOver    =   0   'False
      FillStyle       =   1
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   3
      Editable        =   1
      ShowComboButton =   -1  'True
      WordWrap        =   -1  'True
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
      AllowUserFreezing=   3
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Profile"
      Height          =   252
      Index           =   9
      Left            =   4356
      TabIndex        =   17
      Top             =   552
      Width           =   768
   End
   Begin VB.Label lblTRaws 
      Height          =   240
      Left            =   135
      TabIndex        =   14
      Top             =   9270
      Width           =   2265
   End
   Begin VB.Label Label3 
      Caption         =   "State"
      Height          =   225
      Left            =   4365
      TabIndex        =   13
      Top             =   225
      Width           =   660
   End
   Begin VB.Label Label2 
      Caption         =   "Party"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "City"
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   225
      Width           =   660
   End
End
Attribute VB_Name = "frmPartyProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vs_rs As New ADODB.Recordset
Private Sub cboCity_Click()
addData
End Sub


Private Sub cboProfile_Click()
addData
End Sub

Private Sub cbostate_Click()
addData
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub
Private Sub cmdPrint1_Click()

DSNNew
cr1.Reset
cr1.ReportFileName = rptPath & "\ProfileList1.rpt"
cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=sidc;"
'cr1.DataFiles(0) = st1 + "\" + Trim(main.directory) & "\data.mdb"
If cboCity.Text <> "" Then
  cr1.ReplaceSelectionFormula "{partyProfile.DISTCODE}='" & cboCity.Text & "'"
End If
cr1.WindowShowPrintSetupBtn = True
cr1.WindowShowPrintBtn = True
cr1.WindowState = crptMaximized
cr1.Action = 1

End Sub

Private Sub cmdref_Click()
addData
End Sub

Private Sub cmdRepQty_Click()

Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim I As Integer
Dim xl As Excel.Application
Dim rs_1 As New ADODB.Recordset
Dim rs_2 As New ADODB.Recordset
Dim str_ As String

'


Screen.MousePointer = vbHourglass


If xl Is Nothing Then
   Set xl = CreateObject("Excel.Application")
End If

xl.Visible = True
Set xlBook = xl.Workbooks.Add
Set xlSheet = xlBook.Worksheets.Add

Dim c, r As Long
Dim Q1, q2, J As Double

Dim b1 As Boolean

b1 = False


c = 1
r = 1

row_ = 2
col_ = 1

 xl.Columns("A:H").ColumnWidth = 12
 J = 2
 
 If cbostate.Text <> "" Then
 xlSheet.Cells(1, 1).value = " Sate : " & cbostate.Text
 End If
 
 For I = 0 To vs1.rows - 1
     For J = 0 To vs1.Cols - 1
            xlSheet.Cells(row_, col_).value = vs1.TextMatrix(I, J)
           col_ = col_ + 1
     Next
     row_ = row_ + 1
     col_ = 1
 Next
    


Screen.MousePointer = vbDefault


Exit Sub
Screen.MousePointer = vbDefault
err:
MsgBox err.DESCRIPTION



End Sub

Private Sub cmdSave_Click()

Screen.MousePointer = vbHourglass
For I = 1 To vs1.rows - 1
If vs1.TextMatrix(I, 0) <> "" Then
con.Execute "update sledger set phone='" & vs1.TextMatrix(I, 4) & "', ContactP='" & vs1.TextMatrix(I, 5) & "'," & _
"PartyRemarks='" & vs1.TextMatrix(I, 6) & "',Mobile='" & vs1.TextMatrix(I, 7) & "',PAN='" & vs1.TextMatrix(I, 9) & "',GST='" & vs1.TextMatrix(I, 10) & "'" & _
" where subledger='" & vs1.TextMatrix(I, 0) & "'"
End If
Next
MsgBox "Data Updated !!", vbInformation
Screen.MousePointer = vbDefault

End Sub
Private Sub Command1_Click()

DSNNew

cr1.Reset
cr1.ReportFileName = rptPath & "/ProfileList.rpt"
cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
If cboCity.Text <> "" Then
  cr1.ReplaceSelectionFormula "{partyProfile.DISTCODE}='" & cboCity.Text & "'"
End If
cr1.WindowShowPrintSetupBtn = True
cr1.WindowShowPrintBtn = True
cr1.WindowState = crptMaximized
cr1.Action = 1
End Sub

Private Sub Command2_Click()


Dim ST, mob1, mob2, mob3, str1

For I = 1 To vs1.rows - 1

If vs1.TextMatrix(I, 2) <> "" Then
ST = Split(vs1.TextMatrix(I, 2), ",")


k2 = UBound(ST)

mob1 = ""
mob2 = ""
mob3 = ""

If k2 = 1 Then
    
    If Len(ST(0)) = 10 Then
       mob1 = ST(0)
    End If

ElseIf k2 = 2 Then
    
    If Len(ST(0)) = 10 Then
       mob1 = ST(0)
    End If
    
    If Len(ST(1)) = 10 Then
       mob2 = ST(1)
    End If
    
ElseIf k2 = 3 Then
    
    If Len(ST(0)) = 10 Then
       mob1 = ST(0)
    End If
    
    If Len(ST(1)) = 10 Then
       mob2 = ST(1)
    End If
    
    If Len(ST(2)) = 10 Then
       mob3 = ST(2)
    End If

End If


If Len(mob1) > 0 Then
   str1 = mob1
End If

If Len(mob2) > 0 Then
   str1 = str1 & "," & mob2
End If

If Len(mob3) > 0 Then
   str1 = str1 & "," & mob3
End If

vs1.TextMatrix(I, 5) = str1

con.Execute "update sledger set mobile ='" & str1 & "' where " & stringyear & " and subledger='" & vs1.TextMatrix(I, 0) & "'"

str1 = ""

End If

Next




End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
If MsgBox("Want To Exit ?", vbQuestion + vbYesNo) = vbYes Then
Unload Me
End If
End If
End Sub
Private Sub Form_Load()

'======================================
If RS.State = 1 Then RS.close
RS.Open "select distinct Profile_ from SLEDGER order by Profile_", con
If Not RS.EOF Then
   Do While Not RS.EOF
      If IsNull(RS(0)) = False Then
        Me.cboProfile.AddItem RS(0)
      End If
      If Not RS.EOF Then RS.MoveNext
    Loop
End If
 



If RS.State = 1 Then RS.close
RS.Open "select distinct(DISTCODE) from sledger", con
While RS.EOF = False
If Not IsNull(RS(0)) Then
cboCity.AddItem RS(0)
End If
RS.MoveNext
Wend

cboCity.Text = "Meerut"

addData



Set RS = New ADODB.Recordset
    RS.Open "select states from SLEDGER group by states", con, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
        Do While Not RS.EOF
            cbostate.AddItem RS!states
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.close


BackColorFrom Me

End Sub
Sub addData()

Screen.MousePointer = vbHourglass

vs1.Cols = 11

Dim where_ As String

where_ = "gledger='SUNDRY DEBTORS'"

If cboCity.Text <> "" Then
   where_ = where_ & " and DISTCODE='" & cboCity.Text & "'"
End If

If cbostate.Text <> "" Then
   where_ = where_ & " and states='" & cbostate.Text & "'"
End If


If cboProfile.Text <> "" Then
   where_ = where_ & " and Profile_='" & cboProfile.Text & "'"
End If



If vs_rs.State = 1 Then vs_rs.close

vs_rs.Open "select subledger as PARTY,ADDRESS1,ADDRESS2,ADDRESS3 as City,PHONE,ContactP as [CONTACT PERSON]," & _
"PartyRemarks,Mobile,Email,PAN,GST as GSTIN from sledger" & _
" where " & where_, con



Set vs1.DataSource = vs_rs

lblTRaws.Caption = "Total Records : " & vs_rs.RecordCount

DoEvents
DoEvents
DoEvents

setWidth



DoEvents
DoEvents

Screen.MousePointer = vbDefault

End Sub
Sub setWidth()

vs1.Cols = 11

vs1.ColWidth(0) = 1800
vs1.ColWidth(1) = 1500
vs1.ColWidth(2) = 1400
vs1.ColWidth(3) = 1400
vs1.ColWidth(4) = 1200
vs1.ColWidth(5) = 1200
vs1.ColWidth(6) = 1400
vs1.ColWidth(7) = 1400
vs1.ColWidth(8) = 1800
vs1.ColWidth(9) = 1200
vs1.ColWidth(10) = 1600

vs1.TextMatrix(0, 10) = "GST"

End Sub
Private Sub txtParty_Change()
vs1.Cols = 11


If vs_rs.State = 1 Then vs_rs.close
If cboCity.Text = "" Then
vs_rs.Open "select subledger as PARTY,ADDRESS1,ADDRESS2,ADDRESS3 as City,PHONE,ContactP as [CONTACT PERSON]," & _
"PartyRemarks,Mobile,Email,PAN,GST from sledger " & _
"where " & stringyear & " and gledger='SUNDRY DEBTORS' and  party like '" & txtParty.Text & "%'", con
Else
vs_rs.Open "select subledger as PARTY,ADDRESS1,ADDRESS2,ADDRESS3 as City,PHONE," & _
"ContactP as [CONTACT PERSON],PartyRemarks,Mobile,Email,PAN,GST from sledger " & _
" where " & stringyear & " and DISTCODE='" & cboCity.Text & "' and gledger='SUNDRY DEBTORS' and  party like '" & txtParty.Text & "%'", con
End If




Set vs1.DataSource = vs_rs
DoEvents
DoEvents
DoEvents




setWidth


DoEvents
DoEvents

End Sub

Private Sub txtParty_GotFocus()
If PopUpValue3 <> "" Then
txtParty.Text = PopUpValue3
PopUpValue3 = ""


If vs_rs.State = 1 Then vs_rs.close
If cboCity.Text = "" Then
vs_rs.Open "select subledger as PARTY,ADDRESS1,ADDRESS2,ADDRESS3 as City,PHONE,ContactP as [CONTACT PERSON]," & _
"PartyRemarks,Mobile,Email,PAN,GST from sledger " & _
"where " & stringyear & " and gledger='SUNDRY DEBTORS' and  subledger like '" & Trim(txtParty.Text) & "%'", con
Else
vs_rs.Open "select subledger as PARTY,ADDRESS1,ADDRESS2,ADDRESS3 as City,PHONE," & _
"ContactP as [CONTACT PERSON],PartyRemarks,Mobile,Email,PAN,GST from sledger " & _
" where " & stringyear & " and DISTCODE='" & cboCity.Text & "' and gledger='SUNDRY DEBTORS' and  subledger like '" & Trim(txtParty.Text) & "%'", con
End If

Set vs1.DataSource = vs_rs
DoEvents
DoEvents

setWidth

DoEvents
DoEvents

End If




End Sub
Private Sub txtParty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
value = "select distinct(Party),Code,subledger from SLEDGER where " & stringyear & " and gledger='SUNDRY DEBTORS' order by party"
popuplistModel10 value, con
End If

End Sub
Private Sub vs1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If vs1.Col = 4 Then
   SendKeys "{right}"
ElseIf vs1.Col = 5 Then
   SendKeys "{left}"
   SendKeys "{down}"
End If
End If
End Sub
