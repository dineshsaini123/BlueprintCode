VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmBinderLedger 
   Caption         =   "Binder Ledger ...."
   ClientHeight    =   9144
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   12876
   Icon            =   "frmBinderLedger.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9144
   ScaleWidth      =   12876
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
      Left            =   2304
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   708
      Width           =   990
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
      Left            =   1272
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   708
      Width           =   990
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
      Left            =   1305
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   0
      Top             =   336
      Width           =   5745
   End
   Begin Crystal.CrystalReport crpt 
      Left            =   450
      Top             =   9315
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   7440
      Left            =   36
      TabIndex        =   3
      Top             =   1404
      Width           =   12636
      _cx             =   22288
      _cy             =   13123
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
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
      RowHeightMin    =   320
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
      Height          =   312
      Left            =   1308
      TabIndex        =   7
      Top             =   36
      Width           =   2712
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
      Left            =   6315
      TabIndex        =   6
      Top             =   9375
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
      Left            =   8940
      TabIndex        =   5
      Top             =   9375
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Binder Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   72
      TabIndex        =   4
      Top             =   360
      Width           =   1260
   End
End
Attribute VB_Name = "frmBinderLedger"
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
RS.Open "select distinct(AgentName) from SLEDGER where " & stringyear & " and DISTCODE='" & cboStation.text & "'", con
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

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()


crpt.Reset
crpt.ReportFileName = rptPath & "/CHALLAN_bkrecList.rpt"
crpt.ReplaceSelectionFormula "{BinderBkReceive.subledger}='" & Combosubledger.text & "'"
crpt.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
crpt.WindowShowPrintSetupBtn = True
crpt.WindowShowPrintBtn = True
crpt.WindowState = crptMaximized
crpt.Action = 1



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

    If cboStation1.text <> "" And txtAmount.text = "" Then
    crpt.ReplaceSelectionFormula "{tempLedgerRpt.DISTCODE}='" & cboStation1.text & "' and {tempLedgerRpt.Owner}<>0"
    ElseIf cboStation1.text <> "" And txtAmount.text <> "" Then
    If Val(txtAmount) >= 0 Then
       crpt.ReplaceSelectionFormula "{tempLedgerRpt.DISTCODE}='" & cboStation1.text & "' and {tempLedgerRpt.Owner}<>0 and {tempLedgerRpt.Owner}>=" & txtAmount.text & ""
    Else
       crpt.ReplaceSelectionFormula "{tempLedgerRpt.DISTCODE}='" & cboStation1.text & "' and {tempLedgerRpt.Owner}<>0 and {tempLedgerRpt.Owner}<=" & txtAmount.text & ""
    End If
    
    
    ElseIf cboStation1.text = "" And txtAmount.text <> "" Then
    
    If Val(txtAmount) >= 0 Then
       crpt.ReplaceSelectionFormula "{tempLedgerRpt.Owner}<>0 and {tempLedgerRpt.Owner}>=" & txtAmount.text & ""
    Else
       crpt.ReplaceSelectionFormula "{tempLedgerRpt.Owner}<>0 and {tempLedgerRpt.Owner}<=" & txtAmount.text & ""
    End If
    
    Else
    crpt.ReplaceSelectionFormula "abs({tempLedgerRpt.Owner})<>0"
    End If


ElseIf Check2.value = 1 Then


    If cboStation1.text <> "" And txtAmount.text = "" Then
    crpt.ReplaceSelectionFormula "{tempLedgerRpt.states}='" & cboStation1.text & "' and {tempLedgerRpt.Owner}<>0"
    ElseIf cboStation1.text <> "" And txtAmount.text <> "" Then
    If Val(txtAmount) >= 0 Then
       crpt.ReplaceSelectionFormula "{tempLedgerRpt.states}='" & cboStation1.text & "' and {tempLedgerRpt.Owner}<>0 and {tempLedgerRpt.Owner}>=" & txtAmount.text & ""
    Else
       crpt.ReplaceSelectionFormula "{tempLedgerRpt.states}='" & cboStation1.text & "' and {tempLedgerRpt.Owner}<>0 and {tempLedgerRpt.Owner}<=" & txtAmount.text & ""
    End If
    
    
    ElseIf cboStation1.text = "" And txtAmount.text <> "" Then
    
    If Val(txtAmount) >= 0 Then
       crpt.ReplaceSelectionFormula "{tempLedgerRpt.Owner}<>0 and {tempLedgerRpt.Owner}>=" & txtAmount.text & ""
    Else
       crpt.ReplaceSelectionFormula "{tempLedgerRpt.Owner}<>0 and {tempLedgerRpt.Owner}<=" & txtAmount.text & ""
    End If
    
    Else
    crpt.ReplaceSelectionFormula "abs({tempLedgerRpt.Owner})<>0"
    End If



End If

''======================================================================
''======================================================================










DoEvents
MsgBox ("View")
crpt.Formulas(0) = "partyname='" & cboStation1.text & "'"
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
 .ReplaceSelectionFormula "{tempLedgerRpt.party}='" & cboParty.text & "'"
 .WindowShowPrintSetupBtn = True
 .WindowState = crptMaximized
 .Action = 1
End With
Screen.MousePointer = vbDefault


End Sub
Private Sub cmdprintalf_Click()
 
 If txtalfa.text = "" Then
    MsgBox "Please Enter Alphabet...", vbInformation
    Exit Sub
 End If
 
 Screen.MousePointer = vbHourglass
 'CityWiseStatement
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
      
      vs.Clear
      vs.Cols = 3
      
      
      If PopUpValue7 <> "" Then
         Combosubledger.text = PopUpValue7
      End If
      
      If Combosubledger.text = "" Then Exit Sub
      PopUpValue7 = ""
      
      If RS.State = 1 Then RS.close
      RS.Open "SELECT INVOICENO,INVOICEDATE,remarks FROM BinderBkReceive where SUBLEDGER='" & Combosubledger.text & "' order by INVOICENO", con
      If RS.EOF = False Then
        vs.rows = (vs.rows + RS.RecordCount)
        For I = 1 To vs.rows - 1
        If RS.EOF = False Then
           vs.TextMatrix(I, 0) = RS.Fields("INVOICENO").value
           vs.TextMatrix(I, 1) = RS.Fields("INVOICEDATE").value
           vs.TextMatrix(I, 2) = RS.Fields("remarks").value
           'I = I + 1
           RS.MoveNext
         End If
        Next
      End If


  
    
    
    vs.FormatString = "^ChallanNo No|^Challan Dates|Remarks"
    vs.ColWidth(0) = 1600
    vs.ColWidth(1) = 1600
    vs.ColWidth(2) = 3500
    
    
    DoEvents

End Sub






Private Sub Command2_Click()
crpt.Reset
crpt.ReportFileName = st1 & "\" & directory & "\PartyWiseDrClosing.rpt"
crpt.DataFiles(0) = st1 & "\" + directory + "\data.mdb"
crpt.ReplaceSelectionFormula "{tempLedgerRpt.Offdays}='" & "1" & "' and {tempLedgerRpt.Owner}>=" & 1 & ""
DoEvents
MsgBox ("View")
crpt.Formulas(0) = "partyname='" & cboStation1.text & "'"
crpt.WindowShowPrintSetupBtn = True
crpt.WindowShowPrintBtn = True
crpt.WindowState = crptMaximized
crpt.WindowShowSearchBtn = True
crpt.Action = 1

End Sub

Private Sub Command3_Click()
 
 If cboStation.text = "" Then
    MsgBox "Please Select Station...", vbInformation
    Exit Sub
 End If
 
 Screen.MousePointer = vbHourglass
 'CityWiseStatement
 cboPartyList.Visible = False
 Screen.MousePointer = vbDefault
End Sub

Private Sub Command4_Click()

Dim FSO As filesystemobject
Dim f As File
Dim txt As TextStream
Dim matter As String
Dim Total As String
Dim s(1, 2) As String
Set FSO = New filesystemobject
Dim ss As String
'
Dim s1

matter = ""

Set txt = FSO.CreateTextFile(App.Path & "\mobile.txt", True)

If RS.State = 1 Then RS.close
If Check2.value = 0 Then
RS.Open "select mobile from sledger where " & stringyear & " and distcode='" & cboStation1.text & "'", con, adOpenKeyset, adLockReadOnly
Else
RS.Open "select mobile from sledger where " & stringyear & " and states='" & cboStation1.text & "'", con, adOpenKeyset, adLockReadOnly
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
If PopUpValue1 <> "" Then
   Dim k1 As Integer
   
   Combosubledger.text = PopUpValue1
   
   SearchFa
 
   PopUpValue1 = ""
   
End If

End Sub

Private Sub Combosubledger_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then

    searchType = "party"
    value = "SELECT SUBLEDGER FROM BinderBkReceive group by SUBLEDGER"
    popuplist_client value, con
    set_focus = True


End If






End Sub

Private Sub CommandReturn_Click()
frmSchoolLedgerSP.Show
End Sub


Private Sub Form_Activate()
' Me.WindowState = 2
  
End Sub
Private Sub Form_Load()

Me.top = 100
Me.Left = 100
Me.Width = 12976
Me.Height = 10440


kk = 1

SearchFa
setWidth

Me.top = 250
Me.Left = 250


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
   
   frmReceiveFromParty.top = 800

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
      txtParty.text = PopUpValue1
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
   If Val(txtQty.text) = 0 Then
      txtQty.SetFocus
      Exit Sub
   End If
   If cmdSave.Enabled = True Then
      cmdSave.SetFocus
   End If
   End If
End Sub
Private Sub txtRem_LostFocus()
  If cboParty.text <> "" Then
  If MsgBox("Want To Change Remarks ?", vbQuestion + vbYesNo) = vbYes Then
      con.Execute "update sledger set PartyRemarks = '" & txtrem.text & "' where " & stringyear & " and AgentName='" & cboParty.text & "'"
     
  End If
  End If
End Sub
Private Sub Unautho_Click()
If Unautho.value = True Then
'    Call cmdShow_Click
End If
End Sub

Private Sub vs_DblClick()

    inviceNo = vs.TextMatrix(vs.RowSel, 0)
    s1 = 1
    Unload frmBinderRecChallan
    frmBinderRecChallan.Show


End Sub

Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then


    inviceNo = vs.TextMatrix(vs.RowSel, 0)
    s1 = 1
    Unload frmBinderRecChallan
    frmBinderRecChallan.Show


End If
End Sub
Private Sub cboParty_GotFocus()

Dim ph_rs As New ADODB.Recordset
cboParty.BackColor = &HFFFFC0


I = 1
If PopUpValue1 <> "" Then
cboParty.text = PopUpValue1
End If




End Sub

Private Sub cboParty_KeyDown(KeyCode As Integer, Shift As Integer)

Dim dr, CR As Double

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
If cboParty.text = "" Then
  cboParty.SetFocus
  Exit Sub
End If

dataSearchingrid
cmdprint.Enabled = True



dr = 0
CR = 0

For I = 1 To vs1.rows - 1
  dr = dr + Val(vs1.TextMatrix(I, 4))
  CR = CR + Val(vs1.TextMatrix(I, 5))
Next

drLebel.Caption = Format(dr, "0.00")
CrLebel.Caption = Format(CR, "0.00")
    
'txtdes.SetFocus
    
End If


If KeyCode = 116 Then
vs1.SetFocus
For J = 1 To vs1.rows - 1
   sendkeys "{down}"
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
If cboParty.text <> "" Then
'SaveDatainTempledger
'CalculateTotalDrCr
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
    Set Del = con.Execute("delete from ReceiveIssueParty where " & stringyear & " and RecNo=" & txtRecno.text & "")
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
If Val(txtQty.text) > 0 And txtdes.text <> "" And cboParty.text <> "" Then
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
    
'    vs.Clear
'    vs.Cols = 3
'
'    vs.rows = 200
'
'    vs.FormatString = "^ChallanNo No|^Challan Dates|Remarks"
'    vs.ColWidth(0) = 1200
'    vs.ColWidth(1) = 1500
'    vs.ColWidth(2) = 3500
'
'
'    DoEvents

End Sub
Private Sub cmdRefresh_Click()
 
 
 Dim o As Object
 txtQty.text = ""
 set_focus = False
 'maxId
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
     Unload Me
  End If
  
  
  
  If KeyCode = 116 Then
  If bb2 = False Then
    vs1.SetFocus
    For I = 1 To vs1.rows - 1
    sendkeys "{down}"
    Next
    bb2 = True
  Else
    Call cmdRefresh_Click
    cboParty.SetFocus
    bb2 = False
  End If
  Exit Sub
  End If
  
  
  
'  If KeyCode = 112 Then
'     txtdes.SetFocus
'     Exit Sub
'  End If
  
  If KeyCode = 27 Then
       
      Unload Me
        
  
  End If
   
   
End Sub
Sub fillGrid()
    
   
    '==============
    SearchFa
    '==============
    setWidth
End Sub


Private Sub Todate_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
    from_date = fromdate.value
    to_date = todate.value
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
     RS.Open "select * from receiveissueparty where " & stringyear & " and recno=" & txtRecno.text & "", con
     If RS.EOF = False Then
      cboParty.text = RS!partyname
      PopUpValue3 = cboParty.text
      
      RecDates.value = RS.Fields("Dates").value
      txtdes.text = RS.Fields("Particullar").value
      'txtRem.Text = RS.Fields("Remarks").Value
      If RS.Fields("Dr").value > 0 Then
          Receive.value = True
          txtQty.text = RS.Fields("Dr").value
      Else
          Issue.value = True
          txtQty.text = RS.Fields("Cr").value
      End If
      dataSearchingrid
     Else
       vs1.Clear
       setWidth
       txtQty.text = ""
       txtdes.text = ""
       cboParty.text = ""
       txtOp.text = ""
       txtBalance.text = ""
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
    cboParty.text = vsop.TextMatrix(vsop.RowSel, 1)
    PopUpValue2 = cboParty.text
    vs1.Clear
    fillGrid
    'SaveDatainTempledger
    'CalculateTotalDrCr
    setWidth
    PopUpValue1 = ""
    Opening.Tab = 1
End If
End Sub



