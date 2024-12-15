VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmBillList 
   Caption         =   "Authorization Option"
   ClientHeight    =   8910
   ClientLeft      =   1530
   ClientTop       =   885
   ClientWidth     =   15555
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8910
   ScaleWidth      =   15555
   WindowState     =   2  'Maximized
   Begin VB.Frame bill 
      Height          =   8100
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   12480
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Exit"
         Height          =   435
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   810
         Width           =   1440
      End
      Begin VB.Frame pass 
         Height          =   465
         Left            =   6900
         TabIndex        =   13
         Top             =   840
         Visible         =   0   'False
         Width           =   2730
         Begin VB.TextBox txtadmin 
            Appearance      =   0  'Flat
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   90
            PasswordChar    =   "*"
            TabIndex        =   14
            Top             =   135
            Width           =   1590
         End
         Begin VB.Label Label14 
            Caption         =   "Press Enter"
            Height          =   165
            Left            =   1770
            TabIndex        =   15
            Top             =   180
            Width           =   900
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1095
         Left            =   9750
         TabIndex        =   9
         Top             =   180
         Width           =   2595
         Begin VB.OptionButton All 
            Caption         =   "All"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   45
            TabIndex        =   12
            Top             =   720
            Width           =   2430
         End
         Begin VB.OptionButton Unautho 
            Caption         =   "Un Authorized"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   45
            TabIndex        =   11
            Top             =   435
            Value           =   -1  'True
            Width           =   2460
         End
         Begin VB.OptionButton autho 
            Caption         =   "Authorized"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   45
            TabIndex        =   10
            Top             =   195
            Width           =   2235
         End
      End
      Begin VB.Frame Frame1 
         Height          =   615
         Left            =   2700
         TabIndex        =   4
         Top             =   165
         Width           =   6960
         Begin VB.OptionButton dbit 
            Caption         =   "Debit Note"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   5280
            TabIndex        =   8
            Top             =   255
            Width           =   1485
         End
         Begin VB.OptionButton crdit 
            Caption         =   "Credit Note"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   3405
            TabIndex        =   7
            Top             =   270
            Width           =   1605
         End
         Begin VB.OptionButton sales 
            Caption         =   "Sales Bill "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   105
            TabIndex        =   6
            Top             =   270
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton credit 
            Caption         =   "Credit Note Item"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   1470
            TabIndex        =   5
            Top             =   270
            Width           =   1935
         End
      End
      Begin VB.TextBox txtParty 
         Height          =   300
         Left            =   105
         TabIndex        =   3
         Top             =   990
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.CommandButton cmdset 
         BackColor       =   &H00FFC0C0&
         Caption         =   "S&ave"
         Height          =   435
         Left            =   2700
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   810
         Width           =   1440
      End
      Begin VB.CommandButton cmdshow 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Show"
         Height          =   300
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   990
         Visible         =   0   'False
         Width           =   750
      End
      Begin MSComCtl2.DTPicker toDate 
         Height          =   315
         Left            =   1125
         TabIndex        =   16
         Top             =   630
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         _Version        =   393216
         Format          =   22478849
         CurrentDate     =   38845
      End
      Begin MSComCtl2.DTPicker FromDate 
         Height          =   300
         Left            =   1125
         TabIndex        =   17
         Top             =   300
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         _Version        =   393216
         Format          =   22478849
         CurrentDate     =   38845
      End
      Begin VSFlex7Ctl.VSFlexGrid vs 
         Height          =   6600
         Left            =   60
         TabIndex        =   18
         Top             =   1320
         Width           =   12300
         _cx             =   21696
         _cy             =   11642
         _ConvInfo       =   1
         Appearance      =   1
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
         BackColor       =   16777209
         ForeColor       =   16711680
         BackColorFixed  =   16761024
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16777166
         BackColorAlternate=   16777209
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   100
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmBillList.frx":0000
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
      Begin VB.Label Label1 
         Caption         =   "From Date "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   150
         TabIndex        =   20
         Top             =   360
         Width           =   945
      End
      Begin VB.Label Label2 
         Caption         =   "To Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   150
         TabIndex        =   19
         Top             =   690
         Width           =   945
      End
   End
End
Attribute VB_Name = "frmBillList"
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

Dim to_date As Date
Dim kk As Integer
Dim str1 As New ADODB.Recordset
Sub vsIni()
   vs.FormatString = "SNo|Bill No|Date|Party Name|>Amount|^Authorized"
   vs.ColWidth(0) = 900
   vs.ColWidth(1) = 1500
   vs.ColWidth(2) = 1100
   vs.ColWidth(3) = 5000
   vs.ColWidth(4) = 1500
   vs.ColWidth(5) = 800
End Sub

Private Sub All_Click()
If All.Value = True Then
    Call cmdshow_Click
End If

End Sub

Private Sub autho_Click()
If autho.Value = True Then
    Call cmdshow_Click
End If
End Sub



Private Sub cash_Click()
    If cash.Value = True Then
       Call cmdshow_Click
    End If
End Sub

Private Sub cboop_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtdes.SetFocus
   End If
End Sub
Private Sub cboStation_Click()
cboPartyList.Visible = True
If RS.State = 1 Then RS.Close
RS.Open "select distinct(SUBLEDGER) from SLEDGER where DISTCODE='" & cboStation.Text & "'", CON
cboPartyList.Clear
While RS.EOF = False
cboPartyList.AddItem RS(0)
RS.MoveNext
Wend
End Sub

Private Sub Check1_Click()
    If Check1.Value = 1 Then
       'cmdSave.Enabled = False
       cmdDel.Enabled = False
       cmdmodify.Enabled = False
    Else
       cmdSave.Enabled = True
       cmdDel.Enabled = True
       cmdmodify.Enabled = True
    End If
End Sub

Private Sub Check2_Click()

Dim rs_1 As New ADODB.Recordset

cboStation.Clear
cboStation1.Clear

If Check2.Value = 1 Then
    
    lblStation.Caption = "State :"
    
    If rs_1.State = 1 Then rs_1.Close
    rs_1.Open "select distinct(states) from SLEDGER where states<>''", CON
    While rs_1.EOF = False
    cboStation.AddItem rs_1.Fields(0).Value
    cboStation1.AddItem rs_1.Fields(0).Value
    rs_1.MoveNext
    Wend

ElseIf Check2.Value = 0 Then
    
    lblStation.Caption = "Station :"

    If rs_1.State = 1 Then rs_1.Close
    rs_1.Open "select distinct(DISTCODE) from SLEDGER where DISTCODE<>''", CON
    While rs_1.EOF = False
    cboStation.AddItem rs_1.Fields(0).Value
    cboStation1.AddItem rs_1.Fields(0).Value
    rs_1.MoveNext
    Wend


End If

End Sub

Private Sub cmdAson_Click()
'showDataAsOn dateason
End Sub

Private Sub cmddewali_Click()
    Dim f As New ADODB.Recordset
    If f.State = 1 Then f.Close
    f.Open "select AMOUNT,text,INVOICENO from invoicec where TEXT = '" & "DIWALI SPECIAL" & "' and AMOUNT>0", CON
    While f.EOF = False
        CON.Execute "update invoicea set t2='" & f.Fields("amount").Value & "' where INVOICENO=" & f.Fields("INVOICENO").Value & ""
        f.MoveNext
    Wend
    If f.State = 1 Then f.Close
    f.Open "select AMOUNT,text,INVOICENO from CASHC where TEXT = '" & "DIWALI SPECIAL" & "' and AMOUNT>0", CON
    While f.EOF = False
        CON.Execute "update CASHA set t2='" & f.Fields("amount").Value & "' where INVOICENO=" & f.Fields("INVOICENO").Value & ""
        f.MoveNext
    Wend
    MsgBox "Data Refresh...", vbInformation
End Sub

Private Sub cmdPath_Click()
'Me.comdio.ShowOpen
'Me.txtPath.Text = Me.comdio.FileName
End Sub

Private Sub cmdPrint_Click()
On Error GoTo aa10
Screen.MousePointer = vbHourglass
Dim op, drcr
Dim rs1 As New ADODB.Recordset
CON.Execute "delete from templedger1"

If RS.State = 1 Then RS.Close
RS.Open "select subledger from SLEDGER where subledger = '" + Trim(cboParty.Text) + "'", CON

While RS.EOF = False

'==Code For Opening=============================================
CON.Execute "INSERT INTO templedger1 (Balance,drcr,party,billtype)  SELECT op,drcr,subledger,'Opening' from sledger where subledger = '" + RS.Fields(0).Value + "'   group by op,subledger,drcr HAVING  op <> 0;"
If rs1.State = 1 Then rs1.Close
rs1.Open "SELECT op,drcr from sledger where subledger = '" + RS.Fields(0).Value + "'", CON
If Not IsNull(rs1.Fields(0).Value) Then
   op = Val(rs1.Fields(0).Value)
   drcr = rs1.Fields(1).Value
Else
   op = 0
End If

'==============================================
CON.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party)  SELECT INVOICEDATE,'I',INVOICENO,'Invoice Sales',netamount,BAA,SUBLEDGER from INVOICEA where SUBLEDGER='" & RS.Fields(0).Value & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)"
CON.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party) Select INVOICEDATE,'CI',INVOICENO,'Credit Note Item',BAA,netamount,SUBLEDGER from CREDITA where SUBLEDGER='" & RS.Fields(0).Value & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)"
CON.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party) Select INVOICEDATE,'C/M',INVOICENO,'Cash Memo',netamount,baa,SUBLEDGER from CASHA where  SUBLEDGER='" & RS.Fields(0).Value & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)"
CON.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party) select dnd,'DN',dnn,'Debit Note',na,'0',PSLD from dnfa where  psld='" & RS.Fields(0).Value & "'"
CON.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,cr,dr,Party) Select cnd,'CN',cnn,'Credit Note',na,'0',psld from Cnf1a where  psld='" & RS.Fields(0).Value & "'"
CON.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party) Select dates,'J',Recno,Particullar,Dr,CR,PartyName from ReceiveIssueParty where PartyName='" & RS.Fields(0).Value & "' order by dates,recno"
'===============================================================
If op <> 0 Then
CON.Execute "Update templedger1 set Balance=" & op & ",drcr= '" & drcr & "' where party = '" + RS.Fields(0).Value + "' and Billtype<>'Opening'"
End If
'===============================================================
Sleep (200)
RS.MoveNext
Wend

Sleep (300)
crpt.Reset
'crpt.ReportFileName = App.Path & "\" & directory & "\PartyLedger.rpt"
crpt.ReportFileName = st1 & "\" & directory & "\PartyLedger.rpt"
crpt.DataFiles(0) = st1 & "\" + directory + "\data.mdb"
crpt.WindowShowPrintSetupBtn = True
crpt.WindowShowPrintBtn = True
crpt.WindowState = crptMaximized
crpt.Action = 1
Screen.MousePointer = vbDefault
Exit Sub
aa10:
MsgBox Err.DESCRIPTION
End Sub
Private Sub cmdprint1_Click()

crpt.Reset
crpt.ReportFileName = st1 & "\" & directory & "\PartyWiseClosing.rpt"
crpt.DataFiles(0) = st1 & "\" + directory + "\data.mdb"


''======================================================================
''======================================================================

If Check2.Value = 0 Then

    If cboStation1.Text <> "" And txtamount.Text = "" Then
    crpt.ReplaceSelectionFormula "{tempLedgerRpt.DISTCODE}='" & cboStation1.Text & "' and {tempLedgerRpt.Owner}<>0"
    ElseIf cboStation1.Text <> "" And txtamount.Text <> "" Then
    If Val(txtamount) >= 0 Then
       crpt.ReplaceSelectionFormula "{tempLedgerRpt.DISTCODE}='" & cboStation1.Text & "' and {tempLedgerRpt.Owner}<>0 and {tempLedgerRpt.Owner}>=" & txtamount.Text & ""
    Else
       crpt.ReplaceSelectionFormula "{tempLedgerRpt.DISTCODE}='" & cboStation1.Text & "' and {tempLedgerRpt.Owner}<>0 and {tempLedgerRpt.Owner}<=" & txtamount.Text & ""
    End If
    
    
    ElseIf cboStation1.Text = "" And txtamount.Text <> "" Then
    
    If Val(txtamount) >= 0 Then
       crpt.ReplaceSelectionFormula "{tempLedgerRpt.Owner}<>0 and {tempLedgerRpt.Owner}>=" & txtamount.Text & ""
    Else
       crpt.ReplaceSelectionFormula "{tempLedgerRpt.Owner}<>0 and {tempLedgerRpt.Owner}<=" & txtamount.Text & ""
    End If
    
    Else
    crpt.ReplaceSelectionFormula "abs({tempLedgerRpt.Owner})<>0"
    End If


ElseIf Check2.Value = 1 Then


    If cboStation1.Text <> "" And txtamount.Text = "" Then
    crpt.ReplaceSelectionFormula "{tempLedgerRpt.states}='" & cboStation1.Text & "' and {tempLedgerRpt.Owner}<>0"
    ElseIf cboStation1.Text <> "" And txtamount.Text <> "" Then
    If Val(txtamount) >= 0 Then
       crpt.ReplaceSelectionFormula "{tempLedgerRpt.states}='" & cboStation1.Text & "' and {tempLedgerRpt.Owner}<>0 and {tempLedgerRpt.Owner}>=" & txtamount.Text & ""
    Else
       crpt.ReplaceSelectionFormula "{tempLedgerRpt.states}='" & cboStation1.Text & "' and {tempLedgerRpt.Owner}<>0 and {tempLedgerRpt.Owner}<=" & txtamount.Text & ""
    End If
    
    
    ElseIf cboStation1.Text = "" And txtamount.Text <> "" Then
    
    If Val(txtamount) >= 0 Then
       crpt.ReplaceSelectionFormula "{tempLedgerRpt.Owner}<>0 and {tempLedgerRpt.Owner}>=" & txtamount.Text & ""
    Else
       crpt.ReplaceSelectionFormula "{tempLedgerRpt.Owner}<>0 and {tempLedgerRpt.Owner}<=" & txtamount.Text & ""
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
crpt.Formulas(1) = "ason='" & dateAson.Value & "'"

crpt.WindowShowPrintSetupBtn = True
crpt.WindowShowPrintBtn = True
crpt.WindowState = crptMaximized
crpt.WindowShowSearchBtn = True
crpt.Action = 1
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


Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdset_Click()
   
'If RS.State = 1 Then RS.Close
'RS.Open "select * from pass where pass='" & cp & "'", con
'If RS.EOF = True Then
'MsgBox "Enter Valid Password !!", vbInformation
'Exit Sub
'End If
'
'   'pass.Visible = True
'   'strledger = ""
'   'txtadmin.SetFocus
'
    savedata
   
End Sub
Sub savedata()
   Dim var As Integer
   If MsgBox("Want to Set ?", vbQuestion + vbYesNo) = vbYes Then
        
   Screen.MousePointer = vbHourglass
   'cmdShow1.Visible = True
         
         
   If sales.Value = True Then
        
        For J = 1 To vs.Rows - 1
          If vs.TextMatrix(J, 5) = True Then
             var = 1
             CON.Execute "update INVOICEA set bAuthorized=" & var & " where INVOICENO=" & vs.TextMatrix(J, 1) & ""
            Else
            var = 0
             CON.Execute "update INVOICEA set bAuthorized=" & var & " where INVOICENO=" & vs.TextMatrix(J, 1) & ""
          End If
        Next
   
  ElseIf credit.Value = True Then
  
        For J = 1 To vs.Rows - 1
          If vs.TextMatrix(J, 5) = True Then
          var = 1
            CON.Execute "update CREDITA set bAuthorized=" & var & " where INVOICENO=" & vs.TextMatrix(J, 1) & ""
            Else
            var = 0
            CON.Execute "update CREDITA set bAuthorized=" & var & " where INVOICENO=" & vs.TextMatrix(J, 1) & ""
          End If
        Next
  
  
  ElseIf crdit.Value = True Then
  
        For J = 1 To vs.Rows - 1
          If vs.TextMatrix(J, 5) = True Then
          var = 1
            CON.Execute "update cnf1a set bAuthorized=" & var & " where cnn=" & vs.TextMatrix(J, 1) & ""
            Else
            var = 0
            CON.Execute "update cnf1a set bAuthorized=" & var & " where cnn=" & vs.TextMatrix(J, 1) & ""
          End If
        Next
  
  
  ElseIf dbit.Value = True Then
  
        For J = 1 To vs.Rows - 1
          If vs.TextMatrix(J, 5) = True Then
          var = 1
            CON.Execute "update dnfa set bAuthorized=" & var & " where dnn=" & vs.TextMatrix(J, 1) & ""
            Else
            var = 0
            CON.Execute "update dnfa set bAuthorized=" & var & " where dnn=" & vs.TextMatrix(J, 1) & ""
          End If
        Next
   
   
  End If
   
   
   End If
   
   
 Screen.MousePointer = vbDefault
End Sub
Private Sub cmdshow_Click()
      
Screen.MousePointer = vbHourglass
      
If sales.Value = True Then
      
      If RS.State = 1 Then RS.Close
      If txtparty.Text = "" Then
        
        If All.Value = True Then
           RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netAMOUNT,bauthorized from INVOICEA where (netamount>0 or netamount<0) and " & _
           "convert(smalldatetime,INVOICEDATE,103)>=convert(smalldatetime,'" & fromdate.Value & "',103) and " & _
           "convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & todate.Value & "',103)" & _
           "ORDER BY INVOICENO", CON
        ElseIf autho.Value = True Then
           RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netAMOUNT,bauthorized from INVOICEA " & _
           " where (netamount>0 or netamount<0) " & _
           " and convert(smalldatetime,INVOICEDATE,103)>=convert(smalldatetime,'" & fromdate.Value & "',103) " & _
           " and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & todate.Value & "',103)" & _
           " and bAuthorized=1 ORDER BY INVOICENO", CON
        Else
           RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netamount,bauthorized from INVOICEA where (netamount>0 or netamount<0) and " & _
           "convert(smalldatetime,INVOICEDATE,103)>=convert(smalldatetime,'" & fromdate.Value & "',103)" & _
           "and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & todate.Value & "',103)" & _
           " and bAuthorized=0 ORDER BY INVOICENO", CON
        End If
      
      Else
        If All.Value = True Then
           RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netAMOUNT,bauthorized from INVOICEA where (netamount>0 or netamount<0) and " & _
           "convert(smalldatetime,INVOICEDATE,103)>=convert(smalldatetime,'" & fromdate.Value & "',103) and " & _
           "convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & todate.Value & "',103)" & _
           "and SUBLEDGER='" & txtparty.Text & "' ORDER BY INVOICENO", CON
        ElseIf autho.Value = True Then
           RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netAMOUNT,bauthorized from INVOICEA " & _
           " where (netamount>0 or netamount<0) " & _
           " and convert(smalldatetime,INVOICEDATE,103)>=convert(smalldatetime,'" & fromdate.Value & "',103) " & _
           " and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & todate.Value & "',103)" & _
           " and SUBLEDGER='" & txtparty.Text & "' and bAuthorized=1 ORDER BY INVOICENO", CON
        Else
           RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netAMOUNT,bauthorized from INVOICEA " & _
           " where (netamount>0 or netamount<0) " & _
           " and convert(smalldatetime,INVOICEDATE,103)>=convert(smalldatetime,'" & fromdate.Value & "',103) " & _
           " and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & todate.Value & "',103)" & _
           " and SUBLEDGER='" & txtparty.Text & "' and bAuthorized=0 ORDER BY INVOICENO", CON
        End If
      End If
      
      
      
      If RS.EOF = False Then
        vs.Rows = RS.RecordCount + 1
        For I = 1 To vs.Rows - 1
           vs.TextMatrix(I, 0) = I
           vs.TextMatrix(I, 1) = RS.Fields(0).Value
           vs.TextMatrix(I, 2) = RS.Fields(1).Value
           vs.TextMatrix(I, 3) = RS.Fields(2).Value
           vs.TextMatrix(I, 4) = RS.Fields(3).Value
           vs.TextMatrix(I, 5) = RS.Fields(4).Value & ""
           RS.MoveNext
        Next
      Else
           vs.Clear
           vs.Rows = 2
      End If
      
End If

If credit.Value = True Then

      If RS.State = 1 Then RS.Close
      If txtparty.Text = "" Then
        
        If All.Value = True Then
           RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netAMOUNT,bauthorized from CREDITA where (netamount>0 or netamount<0) and " & _
           "convert(smalldatetime,INVOICEDATE,103)>=convert(smalldatetime,'" & fromdate.Value & "',103) and " & _
           "convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & todate.Value & "',103)" & _
           "ORDER BY INVOICENO", CON
        ElseIf autho.Value = True Then
          RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netAMOUNT,bauthorized from CREDITA " & _
           " where (netamount>0 or netamount<0) " & _
           " and convert(smalldatetime,INVOICEDATE,103)>=convert(smalldatetime,'" & fromdate.Value & "',103) " & _
           " and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & todate.Value & "',103)" & _
           " and bAuthorized=1 ORDER BY INVOICENO", CON
        Else
            RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netamount,bauthorized from CREDITA where (netamount>0 or netamount<0) and " & _
           "convert(smalldatetime,INVOICEDATE,103)>=convert(smalldatetime,'" & fromdate.Value & "',103)" & _
           "and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & todate.Value & "',103)" & _
           " and bAuthorized=0 ORDER BY INVOICENO", CON

        End If
      
      Else
        If All.Value = True Then
           RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netAMOUNT,bauthorized from CREDITA where (netamount>0 or netamount<0) and " & _
           "convert(smalldatetime,INVOICEDATE,103)>=convert(smalldatetime,'" & fromdate.Value & "',103) and " & _
           "convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & todate.Value & "',103)" & _
           "and SUBLEDGER='" & txtparty.Text & "' ORDER BY INVOICENO", CON
        ElseIf autho.Value = True Then
           RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netAMOUNT,bauthorized from CREDITA " & _
           " where (netamount>0 or netamount<0) " & _
           " and convert(smalldatetime,INVOICEDATE,103)>=convert(smalldatetime,'" & fromdate.Value & "',103) " & _
           " and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & todate.Value & "',103)" & _
           " and SUBLEDGER='" & txtparty.Text & "' and bAuthorized=1 ORDER BY INVOICENO", CON
        Else
           RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netAMOUNT,bauthorized from CREDITA " & _
           " where (netamount>0 or netamount<0) " & _
           " and convert(smalldatetime,INVOICEDATE,103)>=convert(smalldatetime,'" & fromdate.Value & "',103) " & _
           " and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & todate.Value & "',103)" & _
           " and SUBLEDGER='" & txtparty.Text & "' and bAuthorized=0 ORDER BY INVOICENO", CON
        End If
      End If
      If RS.EOF = False Then
        vs.Rows = RS.RecordCount + 1
        For I = 1 To vs.Rows - 1
           vs.TextMatrix(I, 0) = I
           vs.TextMatrix(I, 1) = RS.Fields(0).Value
           vs.TextMatrix(I, 2) = RS.Fields(1).Value
           vs.TextMatrix(I, 3) = RS.Fields(2).Value
           vs.TextMatrix(I, 4) = RS.Fields(3).Value
           vs.TextMatrix(I, 5) = RS.Fields(4).Value & ""
           RS.MoveNext
        Next
      Else
           vs.Clear
           vs.Rows = 2
      End If
End If

      
'==================

''If cash.Value = True Then
''
''      If RS.State = 1 Then RS.Close
''      If txtParty.Text = "" Then
''
''        If All.Value = True Then
''           RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netamount,BAuthorized from CASHA where ((netamount-BAA)>0 or (netamount-BAA)<0) and (INVOICEDATE>=datevalue('" & Format(FromDate.Value, "dd/MM/yy") & "') and INVOICEDATE<=datevalue('" & Format(toDate.Value, "dd/MM/yy") & "')) ORDER BY INVOICENO", con
''        ElseIf autho.Value = True Then
''           RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netamount,bauthorized from CASHA where ((netamount-BAA)>0 or (netamount-BAA)<0) and (INVOICEDATE>=datevalue('" & Format(FromDate.Value, "dd/MM/yy") & "') and INVOICEDATE<=datevalue('" & Format(toDate.Value, "dd/MM/yy") & "')) and bAuthorized=true ORDER BY INVOICENO", con
''        Else
''           RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netamount,bauthorized from CASHA where ((netamount-BAA)>0 or (netamount-BAA)<0) and (INVOICEDATE>=datevalue('" & Format(FromDate.Value, "dd/MM/yy") & "') and INVOICEDATE<=datevalue('" & Format(toDate.Value, "dd/MM/yy") & "')) and bAuthorized=False ORDER BY INVOICENO", con
''        End If
''
''      Else
''        If All.Value = True Then
''         RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netamount,bauthorized from CASHA where ((netamount-BAA)>0 or (netamount-BAA)<0) and (INVOICEDATE>=datevalue('" & Format(FromDate.Value, "dd/MM/yy") & "') and INVOICEDATE<=datevalue('" & Format(toDate.Value, "dd/MM/yy") & "')) and SUBLEDGER='" & txtParty.Text & "' ORDER BY INVOICENO", con
''        ElseIf autho.Value = True Then
''         RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netamount,bauthorized from CASHA where ((netamount-BAA)>0 or (netamount-BAA)<0) and (INVOICEDATE>=datevalue('" & Format(FromDate.Value, "dd/MM/yy") & "') and INVOICEDATE<=datevalue('" & Format(toDate.Value, "dd/MM/yy") & "')) and SUBLEDGER='" & txtParty.Text & "' and bAuthorized=true ORDER BY INVOICENO", con
''        Else
''         RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netamount,bauthorized from CASHA where ((netamount-BAA)>0 or (netamount-BAA)<0) and (INVOICEDATE>=datevalue('" & Format(FromDate.Value, "dd/MM/yy") & "') and INVOICEDATE<=datevalue('" & Format(toDate.Value, "dd/MM/yy") & "')) and SUBLEDGER='" & txtParty.Text & "' and bAuthorized=false ORDER BY INVOICENO", con
''        End If
''      End If
''
''
''
''
''      If RS.EOF = False Then
''        vs.Rows = RS.RecordCount + 1
''        For i = 1 To vs.Rows - 1
''           vs.TextMatrix(i, 0) = i
''           vs.TextMatrix(i, 1) = RS.Fields(0).Value
''           vs.TextMatrix(i, 2) = RS.Fields(1).Value
''           vs.TextMatrix(i, 3) = RS.Fields(2).Value
''           vs.TextMatrix(i, 4) = RS.Fields(3).Value
''           vs.TextMatrix(i, 5) = RS.Fields(4).Value
''           RS.MoveNext
''        Next
''      Else
''           vs.Clear
''           vs.Rows = 2
''      End If
''
''End If
''
''
'''================================

If crdit.Value = True Then
      
'      If rs.State = 1 Then rs.Close
'      If txtParty.Text = "" Then
'
'        If All.Value = True Then
'           rs.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netAMOUNT,bauthorized from casha where (netamount>0 or netamount<0) and " & _
'           "convert(smalldatetime,INVOICEDATE,103)>=convert(smalldatetime,'" & FromDate.Value & "',103) and " & _
'           "convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & toDate.Value & "',103)" & _
'           "ORDER BY INVOICENO", CON
'        ElseIf autho.Value = True Then
'          rs.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netAMOUNT,bauthorized from casha " & _
'           " where (netamount>0 or netamount<0) " & _
'           " and convert(smalldatetime,INVOICEDATE,103)>=convert(smalldatetime,'" & FromDate.Value & "',103) " & _
'           " and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & toDate.Value & "',103)" & _
'           " and bAuthorized=1 ORDER BY INVOICENO", CON
'        Else
'          rs.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netAMOUNT,bauthorized from casha " & _
'           " where (netamount>0 or netamount<0) " & _
'           " and convert(smalldatetime,INVOICEDATE,103)>=convert(smalldatetime,'" & FromDate.Value & "',103) " & _
'           " and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & toDate.Value & "',103)" & _
'           " and bAuthorized=0 ORDER BY INVOICENO", CON
'        End If
'
'      Else
'        If All.Value = True Then
'           rs.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netAMOUNT,bauthorized from casha where (netamount>0 or netamount<0) and " & _
'           "convert(smalldatetime,INVOICEDATE,103)>=convert(smalldatetime,'" & FromDate.Value & "',103) and " & _
'           "convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & toDate.Value & "',103)" & _
'           "and SUBLEDGER='" & txtParty.Text & "' ORDER BY INVOICENO", CON
'        ElseIf autho.Value = True Then
'           rs.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netAMOUNT,bauthorized from casha " & _
'           " where (netamount>0 or netamount<0) " & _
'           " and convert(smalldatetime,INVOICEDATE,103)>=convert(smalldatetime,'" & FromDate.Value & "',103) " & _
'           " and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & toDate.Value & "',103)" & _
'           " and SUBLEDGER='" & txtParty.Text & "' and bAuthorized=1 ORDER BY INVOICENO", CON
'        Else
'           rs.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netAMOUNT,bauthorized from casha " & _
'           " where (netamount>0 or netamount<0) " & _
'           " and convert(smalldatetime,INVOICEDATE,103)>=convert(smalldatetime,'" & FromDate.Value & "',103) " & _
'           " and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & toDate.Value & "',103)" & _
'           " and SUBLEDGER='" & txtParty.Text & "' and bAuthorized=0 ORDER BY INVOICENO", CON
'        End If
'      End If
'
'
'
'
'      If rs.EOF = False Then
'        vs.Rows = rs.RecordCount + 1
'        For i = 1 To vs.Rows - 1
'           vs.TextMatrix(i, 0) = i
'           vs.TextMatrix(i, 1) = rs.Fields(0).Value
'           vs.TextMatrix(i, 2) = rs.Fields(1).Value
'           vs.TextMatrix(i, 3) = rs.Fields(2).Value
'           vs.TextMatrix(i, 4) = rs.Fields(3).Value
'           vs.TextMatrix(i, 5) = rs.Fields(4).Value & ""
'           rs.MoveNext
'        Next
'      Else
'           vs.Clear
'           vs.Rows = 2
'      End If

End If

      
If dbit.Value = True Then


      If RS.State = 1 Then RS.Close
      If txtparty.Text = "" Then
        
        If All.Value = True Then
           RS.Open "select DNN,DND,PSLD,NA,bauthorized from DNFA where (NA>0 or NA<0) and " & _
           "convert(smalldatetime,DND,103)>=convert(smalldatetime,'" & fromdate.Value & "',103) and " & _
           "convert(smalldatetime,DND,103)<=convert(smalldatetime,'" & todate.Value & "',103)" & _
           "ORDER BY DNN", CON
        ElseIf autho.Value = True Then
          RS.Open "select DNN,DND,PSLD,NA,bauthorized from DNFA " & _
           " where (NA>0 or NA<0) " & _
           " and convert(smalldatetime,DND,103)>=convert(smalldatetime,'" & fromdate.Value & "',103) " & _
           " and convert(smalldatetime,DND,103)<=convert(smalldatetime,'" & todate.Value & "',103)" & _
           " and bAuthorized=1 ORDER BY DNN", CON
        Else
          RS.Open "select DNN,DND,PSLD,NA,bauthorized from DNFA " & _
           " where (NA>0 or NA<0) " & _
           " and convert(smalldatetime,DND,103)>=convert(smalldatetime,'" & fromdate.Value & "',103) " & _
           " and convert(smalldatetime,DND,103)<=convert(smalldatetime,'" & todate.Value & "',103)" & _
           " and bAuthorized=0 ORDER BY DNN", CON
        End If
      
      Else
        If All.Value = True Then
           RS.Open "select DNN,DND,PSLD,NA,bauthorized from DNFA " & _
           " where (NA>0 or NA<0) " & _
           " and convert(smalldatetime,DND,103)>=convert(smalldatetime,'" & fromdate.Value & "',103) " & _
           " and convert(smalldatetime,DND,103)<=convert(smalldatetime,'" & todate.Value & "',103)" & _
           " and PSLD='" & txtparty.Text & "' ORDER BY DNN", CON
        ElseIf autho.Value = True Then
           RS.Open "select DNN,DND,PSLD,NA,bauthorized from DNFA " & _
           " where (NA>0 or NA<0) " & _
           " and convert(smalldatetime,DND,103)>=convert(smalldatetime,'" & fromdate.Value & "',103) " & _
           " and convert(smalldatetime,DND,103)<=convert(smalldatetime,'" & todate.Value & "',103)" & _
           " and PSLD='" & txtparty.Text & "' and bAuthorized=1 ORDER BY DNN", CON
        Else
           RS.Open "select DNN,DND,PSLD,NA,bauthorized from DNFA " & _
           " where (NA>0 or NA<0) " & _
           " and convert(smalldatetime,DND,103)>=convert(smalldatetime,'" & fromdate.Value & "',103) " & _
           " and convert(smalldatetime,DND,103)<=convert(smalldatetime,'" & todate.Value & "',103)" & _
           " and PSLD='" & txtparty.Text & "' and bAuthorized=0 ORDER BY DNN", CON
        End If
      End If
      
      
      
      
      If RS.EOF = False Then
        vs.Rows = RS.RecordCount + 1
        For I = 1 To vs.Rows - 1
           vs.TextMatrix(I, 0) = I
           vs.TextMatrix(I, 1) = RS.Fields(0).Value
           vs.TextMatrix(I, 2) = RS.Fields(1).Value & ""
           vs.TextMatrix(I, 3) = RS.Fields(2).Value
           vs.TextMatrix(I, 4) = RS.Fields(3).Value
           vs.TextMatrix(I, 5) = RS.Fields(4).Value & ""
           RS.MoveNext
        Next
      Else
           vs.Clear
           vs.Rows = 2
      End If

End If
      
      
      
      
      
 vsIni
 
 Screen.MousePointer = vbDefault
 
End Sub

 
Private Sub cmdShow1_Click()
    
''If RS.State = 1 Then RS.Close
''RS.Open "select PartyName,Remarks from ReceiveIssueParty group by PartyName,Remarks", con
''While RS.EOF = False
''   con.Execute "update sledger set PartyRemarks='" & RS(1) & "' where subledger='" & RS(0) & "'"
''   RS.MoveNext
''Wend
    
    

    
''Dim d1, d2
''d1 = 0
''d2 = 0
''i = 1
''
''
''If Check1.Value = 0 Then
''
''    vs1.Col = 5
''
''    Dim fill As New ADODB.Recordset
''    If fill.State = 1 Then fill.Close
''    fill.Open "select RecNo,Dates,Particullar,Dr,Cr from ReceiveIssueParty order by dates,recno", con
''    If fill.EOF = False Then
''       Set vs1.DataSource = fill
''       vs1.FormatString = "^RecNo|^Dates|Description|>Dr|>Cr"
''      Else
''       Set vs1.DataSource = fill
''       vs1.FormatString = "^RecNo|^Dates|Description|>Dr|>Cr"
''     End If
''
''
''
'''==============================================
''
''For i = 1 To vs1.Rows - 1
''    d1 = d1 + Val(vs1.TextMatrix(i, 3))
''    d2 = d2 + Val(vs1.TextMatrix(i, 4))
''Next
''
''txtClosing.Text = d1 - d2
''
''If Val(txtClosing.Text) < 0 Then
''   txtClosing.Text = -1 * Val(txtClosing.Text)
''   dr.Caption = "Cr"
''   Else
''   dr.Caption = "Dr"
''End If
''
'''==============================================
''
''
''End If
''
''If Check1.Value = 1 Then
''
''   vs1.Cols = 6
''
''   SearchFa
''
''
''vs1.TextMatrix(0, 0) = "Bill Type"
''vs1.TextMatrix(0, 1) = "Bill"
''vs1.TextMatrix(0, 2) = "Date"
''vs1.TextMatrix(0, 3) = "Description"
''vs1.TextMatrix(0, 4) = "Dr"
''vs1.TextMatrix(0, 5) = "Cr"
''
''
'''vs1.ColWidth(0) = 1200
'''vs1.ColWidth(1) = 1200
'''vs1.ColWidth(2) = 1000
'''vs1.ColWidth(3) = 4500
'''vs1.ColWidth(4) = 1300
'''vs1.ColWidth(5) = 1300
''
''
'''==============================================
''
''For i = 1 To vs1.Rows - 1
''    d1 = d1 + Val(vs1.TextMatrix(i, 4))
''    d2 = d2 + Val(vs1.TextMatrix(i, 5))
''Next
''
''txtClosing.Text = d1 - d2
''
''If Val(txtClosing.Text) < 0 Then
''   txtClosing.Text = -1 * Val(txtClosing.Text)
''   dr.Caption = "Cr"
''   Else
''   dr.Caption = "Dr"
''End If
''
''txtClosing.Text = Round(txtClosing.Text, 2)
'''==============================================
''
''
''
''End If


'If kk = 1 Then
' txtRem.Visible = True
' kk = kk + 1
'Else
' txtRem.Visible = False
' kk = 1
'End If

End Sub
Sub SearchFa()
      If RS.State = 1 Then RS.Close
      RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netamount,BAA,t2 from INVOICEA where SUBLEDGER='" & cboParty.Text & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)", CON
      If RS.EOF = False Then
        vs1.Rows = (vs1.Rows + RS.RecordCount)
        For I = I To vs1.Rows - 1
        If RS.EOF = False Then
           vs1.TextMatrix(I, 0) = "I"
           vs1.TextMatrix(I, 1) = RS.Fields("INVOICENO").Value
           vs1.TextMatrix(I, 2) = RS.Fields("INVOICEDATE").Value
           If IsNull(RS.Fields("t2").Value) Then
              vs1.TextMatrix(I, 3) = "Invoice Sales"
           Else
              vs1.TextMatrix(I, 3) = "Invoice Sales" & RS.Fields("t2").Value & " " & "DS"
           End If
           vs1.TextMatrix(I, 4) = Format(RS.Fields("netamount").Value, "0.00")
           vs1.TextMatrix(I, 5) = Format(RS.Fields("BAA").Value, "0.00")
            RS.MoveNext
         End If
        Next
      End If
'    '================
     If RS.State = 1 Then RS.Close
     RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netamount,baa from CREDITA where SUBLEDGER='" & cboParty.Text & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)", CON
     If RS.EOF = False Then
        vs1.Rows = vs1.Rows + RS.RecordCount
        For I = I To vs1.Rows - 1
         
        If RS.EOF = False Then
         vs1.TextMatrix(I, 0) = "CI"
         vs1.TextMatrix(I, 1) = RS.Fields("INVOICENO").Value
         vs1.TextMatrix(I, 2) = RS.Fields("INVOICEDATE").Value
         vs1.TextMatrix(I, 3) = "Credit Note Item"
         vs1.TextMatrix(I, 4) = Format(RS.Fields("baa").Value, "0.00")
         vs1.TextMatrix(I, 5) = Format(RS.Fields("netamount").Value, "0.00")
         RS.MoveNext
       End If
    Next
    End If
    If RS.State = 1 Then RS.Close
    RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netamount,baa,t2 from CASHA where  SUBLEDGER='" & cboParty.Text & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)", CON
    If RS.EOF = False Then
     vs1.Rows = vs1.Rows + RS.RecordCount
     For I = I To vs1.Rows - 1
    If RS.EOF = False Then
      vs1.TextMatrix(I, 0) = "C/M"
      vs1.TextMatrix(I, 1) = RS.Fields("INVOICENO").Value
      vs1.TextMatrix(I, 2) = RS.Fields("INVOICEDATE").Value
      If IsNull(RS.Fields("t2").Value) Then
         vs1.TextMatrix(I, 3) = "Cash Memo"
      Else
         vs1.TextMatrix(I, 3) = "Cash Memo" & " " & RS.Fields("t2").Value & " DS"
         
      End If
      vs1.TextMatrix(I, 4) = Format(RS.Fields("netamount").Value, "0.00")
      vs1.TextMatrix(I, 5) = Format(RS.Fields("baa").Value, "0.00")
      RS.MoveNext
    End If
    Next
    End If
'===================
    If RS.State = 1 Then RS.Close
    RS.Open "select cnn,cnd,na from Cnf1a where  psld='" & cboParty.Text & "'", CON
    If RS.EOF = False Then
     vs1.Rows = vs1.Rows + RS.RecordCount
     For I = I To vs1.Rows - 1
      
    If RS.EOF = False Then
    
      vs1.TextMatrix(I, 0) = "CN"
      vs1.TextMatrix(I, 1) = RS.Fields("cnn").Value
      vs1.TextMatrix(I, 2) = RS.Fields("cnd").Value
      vs1.TextMatrix(I, 3) = "Credit Note"
      vs1.TextMatrix(I, 5) = Format(RS.Fields("na").Value, "0.00")
      vs1.TextMatrix(I, 4) = 0
      RS.MoveNext
    
    End If
    
    Next
    End If
     '===================
    If RS.State = 1 Then RS.Close
    RS.Open "select dnn,dnd,psld,na from dnfa where  psld='" & cboParty.Text & "'", CON
    If RS.EOF = False Then
     vs1.Rows = vs1.Rows + RS.RecordCount
     For I = I To vs1.Rows - 1
    If RS.EOF = False Then
      vs1.TextMatrix(I, 0) = "DN"
      vs1.TextMatrix(I, 1) = RS.Fields("dnn").Value
      vs1.TextMatrix(I, 2) = RS.Fields("dnd").Value
      vs1.TextMatrix(I, 3) = "Debit Note"
      vs1.TextMatrix(I, 4) = Format(RS.Fields("na").Value, "0.00")
      vs1.TextMatrix(I, 5) = 0
      RS.MoveNext
    End If
    Next
    End If
    vs1.FormatString = "^Bill Type|^Bill|^Date|<Description|>Dr|>Cr"
    setwidth
End Sub
Sub CityWiseStatement()
       Dim op, drcr
       Dim s As String
       s = ""
       Dim rs1 As New ADODB.Recordset
       CON.Execute "delete from templedger1"
       If RS.State = 1 Then RS.Close
       If cboStation.Text <> "" And txtalfa.Text = "" Then
       For I = 0 To cboPartyList.ListCount - 1
        If cboPartyList.Selected(I) = True Then
        If s = "" Then
          s = "SUBLEDGER " & " = " & "'" & cboPartyList.List(I) & "'"
        Else
          s = s & " or " & "SUBLEDGER " & " = " & "'" & cboPartyList.List(I) & "'"
        End If
        End If
       Next
       
       If s = "" Then
        If RS.State = 1 Then RS.Close
        RS.Open "select subledger from SLEDGER where DISTCODE = '" & cboStation.Text & "'", CON
       Else
        If RS.State = 1 Then RS.Close
        RS.Open "select subledger from SLEDGER where " & s, CON
       End If
       
       ElseIf txtalfa.Text <> "" And cboStation.Text = "" Then
        RS.Open "select subledger from SLEDGER where Subledger like '" + Trim(txtalfa.Text) + "%'", CON
       Else
         Exit Sub
       End If
       While RS.EOF = False
           '==Code For Opening=============================================
            CON.Execute "INSERT INTO templedger1 (Balance,drcr,party,billtype)  SELECT op,drcr,subledger,'Opening' from sledger where subledger = '" + RS.Fields(0).Value + "'   group by op,subledger,drcr HAVING  op <> 0;"
            If rs1.State = 1 Then rs1.Close
            rs1.Open "SELECT op,drcr from sledger where subledger = '" + RS.Fields(0).Value + "'", CON
            If Not IsNull(rs1.Fields(0).Value) Then
               op = Val(rs1.Fields(0).Value)
               drcr = rs1.Fields(1).Value
            Else
               op = 0
            End If
           '==============================================
            CON.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party)  SELECT INVOICEDATE,'I',INVOICENO,'Invoice Sales',netamount,BAA,SUBLEDGER from INVOICEA where SUBLEDGER='" & RS.Fields(0).Value & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)"
            CON.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party) Select INVOICEDATE,'CI',INVOICENO,'Credit Note Item',baa,netamount,SUBLEDGER from CREDITA where SUBLEDGER='" & RS.Fields(0).Value & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)"
            CON.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party) Select INVOICEDATE,'C/M',INVOICENO,'Cash Memo',netamount,baa,SUBLEDGER from CASHA where  SUBLEDGER='" & RS.Fields(0).Value & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)"
            CON.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party) select dnd,'DN',dnn,'Debit Note',na,'0',PSLD from dnfa where  psld='" & RS.Fields(0).Value & "'"
            CON.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,cr,dr,Party) Select cnd,'CN',cnn,'Credit Note',na,'0',psld from Cnf1a where  psld='" & RS.Fields(0).Value & "'"
            CON.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party) Select dates,'J',Recno,Particullar,Dr,CR,PartyName from ReceiveIssueParty where PartyName='" & RS.Fields(0).Value & "' order by dates,recno"
          '===============================================================
          If op <> 0 Then
           CON.Execute "Update templedger1 set Balance=" & op & ",drcr= '" & drcr & "' where party = '" + RS.Fields(0).Value + "' and Billtype<>'Opening'"
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
Sub showData()

''Dim contr As New ADODB.Connection
''Dim fillvs As New ADODB.Recordset
''Dim dr, CR As Double
''
''Dim bb As Boolean
''Screen.MousePointer = vbHourglass
''On Error GoTo aa11
''
''If Me.txtPath.Text <> "" Then
''If MsgBox("Want To Transfer Closing ?", vbQuestion + vbYesNo) = vbYes Then
''   contr.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source =" & txtPath.Text
''   contr.CursorLocation = adUseClient
''   contr.Open
''   bb = True
''   contr.Execute "update sledger set OP=" & 0 & ",drcr=''"
''End If
''End If
''
''
''
''
''If MsgBox("Want To Show Balance", vbInformation + vbYesNo) <> vbYes Then
''    Screen.MousePointer = vbDefault
''    Exit Sub
''End If
''Dim op, drcr
''Dim rs1 As New ADODB.Recordset
''CON.Execute "delete from templedger2"
''CON.Execute "INSERT INTO tempLedger2 (dates,Billtype,bill,des,dr,cr,Party)  SELECT INVOICEDATE,'I',INVOICENO,'Invoice Sales',netamount,BAA,SUBLEDGER from INVOICEA WHERE ((netamount-BAA)>0 or (netamount-BAA)<0)"
''CON.Execute "INSERT INTO tempLedger2 (dates,Billtype,bill,des,dr,cr,Party) Select INVOICEDATE,'CI',INVOICENO,'Credit Note Item',baa,netamount,SUBLEDGER from CREDITA WHERE ((netamount-BAA)>0 or (netamount-BAA)<0)"
''CON.Execute "INSERT INTO tempLedger2 (dates,Billtype,bill,des,dr,cr,Party) Select INVOICEDATE,'C/M',INVOICENO,'Cash Memo',netamount,baa,SUBLEDGER from CASHA where ((netamount-BAA)>0 or (netamount-BAA)<0)"
''CON.Execute "INSERT INTO tempLedger2 (dates,Billtype,bill,des,dr,cr,Party) select dnd,'DN',dnn,'Debit Note',na,'0',PSLD from dnfa"
''CON.Execute "INSERT INTO tempLedger2 (dates,Billtype,bill,des,cr,dr,Party) Select cnd,'CN',cnn,'Credit Note',na,'0',psld from Cnf1a"
''CON.Execute "INSERT INTO tempLedger2 (dates,Billtype,bill,des,dr,cr,Party) Select dates,'J',Recno,Particullar,Dr,CR,PartyName from ReceiveIssueParty"
''DoEvents
''DoEvents
''CON.Execute "update SLEDGER set Owner=0"
''DoEvents
''DoEvents
''
''If fillvs.State = 1 Then fillvs.Close
''fillvs.Open "select DISTCODE as City,SUBLEDGER as Party,op,drcr from SLEDGER where gledger='SUNDRY DEBTORS'", CON, adOpenDynamic, adLockOptimistic
''If fillvs.EOF = False Then
''vsop.Rows = fillvs.RecordCount
''DoEvents
''DoEvents
''abc.Caption = vsop.Rows
''For i = 1 To vsop.Rows - 1
''
''  vsop.TextMatrix(i, 0) = fillvs(0) & ""
''  vsop.TextMatrix(i, 1) = fillvs(1)
''  vsop.TextMatrix(i, 2) = Format(Round(fillvs(2), 2), "0.00")
''  vsop.TextMatrix(i, 3) = fillvs(3) & ""
''  op = 0
''  dr = 0
''  CR = 0
''  op = IIf(IsNull(fillvs(2)), 0, fillvs(2))
''
''  'If "I355  INDER BOOK AGENCY, DEHRADUN" = fillvs(1) Then
''  '  MsgBox "sdsd"
''  'End If
''
''
''  If RS.State = 1 Then RS.Close
''  RS.Open "select sum(dr),sum(cr) from tempLedger2 where party='" & fillvs(1) & "'", CON, adOpenDynamic, adLockOptimistic
''  If Not IsNull(RS(0)) Then
''     dr = RS(0)
''  End If
''
''  If Not IsNull(RS(1)) Then
''     CR = RS(1)
''  End If
''  If fillvs(3) = "Cr" Then
''    op = (-1 * fillvs(2))
''  End If
''
''  drcr = Round((op + (dr - CR)), 2)
''  If Val(drcr) < 0 Then
''     vsop.TextMatrix(i, 4) = Abs(Round((op + (dr - CR)), 2))
''     vsop.TextMatrix(i, 5) = "Cr"
''  Else
''     vsop.TextMatrix(i, 4) = Round((op + (dr - CR)), 2)
''     vsop.TextMatrix(i, 5) = "Dr"
''  End If
''
''  drcr = Format(Round(drcr, 2), "0.00")
''  If Val(drcr) < 0 Then
''  CON.Execute "update sledger set Owner=" & Round(Val(drcr), 2) & " where subledger='" & fillvs(1) & "'"
''  If bb = True Then
''      contr.Execute "update sledger set OP=" & Abs(Round(Val(drcr), 2)) & ",drcr='Cr' where code='" & Trim(Mid(fillvs(1), 1, 6)) & "'"
''  End If
''  Else
''      CON.Execute "update sledger set Owner=" & Round(Val(drcr), 2) & " where subledger='" & fillvs(1) & "'"
''  If bb = True Then
''     contr.Execute "update sledger set OP=" & Abs(Round(Val(drcr), 2)) & ",drcr='Dr' where code='" & Trim(Mid(fillvs(1), 1, 6)) & "'"
''  End If
''  End If
''
''
''
''  If Not IsNull(RS.Fields(1).Value) Then
''  If RS.Fields(1).Value = 0 Then
''  CON.Execute "update sledger set Offdays='" & "1" & "' where subledger='" & fillvs(1) & "'"
''  Else
''  CON.Execute "update sledger set Offdays='" & "2" & "' where subledger='" & fillvs(1) & "'"
''  End If
''  End If
''
''
''  fillvs.MoveNext
''  DoEvents
''  DoEvents
''  abc.Caption = abc.Caption - 1
''Next
''End If
''vsop.Cols = 6
''vsop.TextMatrix(0, 0) = "City"
''vsop.TextMatrix(0, 1) = "Party"
''vsop.TextMatrix(0, 2) = "Opening"
''vsop.TextMatrix(0, 3) = "Dr/Cr"
''vsop.TextMatrix(0, 4) = "Closing Balance"
''vsop.TextMatrix(0, 5) = "Dr/Cr"
''vsop.ColWidth(0) = 1800
''vsop.ColWidth(1) = 3200
''vsop.ColWidth(2) = 1200
''vsop.ColWidth(3) = 500
''vsop.ColWidth(4) = 1200
''vsop.ColWidth(5) = 500
''abc.Caption = ""
''bb = False
''Screen.MousePointer = vbDefault
''
''Exit Sub
''
''aa11:
''MsgBox "" & "Connection Not Created Properly !", vbInformation
''Screen.MousePointer = vbDefault

End Sub
Sub showDataAsOn(D As Date)

'''Dim contr As New ADODB.Connection
'''Dim fillvs As New ADODB.Recordset
'''Dim dr, CR As Double
'''Dim bb As Boolean
'''
'''Screen.MousePointer = vbHourglass
'''
''''On Error GoTo aa11
'''
'''If Me.txtPath.Text <> "" Then
'''If MsgBox("Want To Transfer Closing ?", vbQuestion + vbYesNo) = vbYes Then
'''   contr.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source =" & txtPath.Text
'''   contr.CursorLocation = adUseClient
'''   contr.Open
'''   bb = True
'''   contr.Execute "update sledger set OP=" & 0 & ",drcr=''"
'''End If
'''End If
'''
'''
'''
'''
'''If MsgBox("Want To Show Balance", vbInformation + vbYesNo) <> vbYes Then
'''    Screen.MousePointer = vbDefault
'''    Exit Sub
'''End If
'''Dim op, drcr
'''Dim rs1 As New ADODB.Recordset
'''CON.Execute "delete from templedger2"
'''CON.Execute "INSERT INTO tempLedger2 (dates,Billtype,bill,des,dr,cr,Party)  SELECT INVOICEDATE,'I',INVOICENO,'Invoice Sales',netamount,BAA,SUBLEDGER from INVOICEA WHERE INVOICEDATE<=datevalue('" & dateason.Value & "') and ((netamount-BAA)>0 or (netamount-BAA)<0)"
'''CON.Execute "INSERT INTO tempLedger2 (dates,Billtype,bill,des,dr,cr,Party) Select INVOICEDATE,'CI',INVOICENO,'Credit Note Item',baa,netamount,SUBLEDGER from CREDITA WHERE INVOICEDATE<=datevalue('" & dateason.Value & "') and ((netamount-BAA)>0 or (netamount-BAA)<0)"
'''CON.Execute "INSERT INTO tempLedger2 (dates,Billtype,bill,des,dr,cr,Party) Select INVOICEDATE,'C/M',INVOICENO,'Cash Memo',netamount,baa,SUBLEDGER from CASHA where INVOICEDATE<=datevalue('" & dateason.Value & "') and ((netamount-BAA)>0 or (netamount-BAA)<0)"
'''CON.Execute "INSERT INTO tempLedger2 (dates,Billtype,bill,des,dr,cr,Party) select dnd,'DN',dnn,'Debit Note',na,'0',PSLD from dnfa where DND<=datevalue('" & dateason.Value & "')"
'''CON.Execute "INSERT INTO tempLedger2 (dates,Billtype,bill,des,cr,dr,Party) Select cnd,'CN',cnn,'Credit Note',na,'0',psld from Cnf1a where cnd<=datevalue('" & dateason.Value & "')"
'''CON.Execute "INSERT INTO tempLedger2 (dates,Billtype,bill,des,dr,cr,Party) Select dates,'J',Recno,Particullar,Dr,CR,PartyName from ReceiveIssueParty where Dates<=datevalue('" & dateason.Value & "')"
'''DoEvents
'''DoEvents
'''CON.Execute "update SLEDGER set Owner=0"
'''DoEvents
'''DoEvents
'''
'''If fillvs.State = 1 Then fillvs.Close
'''fillvs.Open "select DISTCODE as City,SUBLEDGER as Party,op,drcr from SLEDGER where gledger='SUNDRY DEBTORS'", CON, adOpenDynamic, adLockOptimistic
'''If fillvs.EOF = False Then
'''vsop.Rows = fillvs.RecordCount
'''DoEvents
'''DoEvents
'''abc.Caption = vsop.Rows
'''For i = 1 To vsop.Rows - 1
'''
'''  vsop.TextMatrix(i, 0) = fillvs(0) & ""
'''  vsop.TextMatrix(i, 1) = fillvs(1)
'''  vsop.TextMatrix(i, 2) = Format(Round(fillvs(2), 2), "0.00")
'''  vsop.TextMatrix(i, 3) = fillvs(3) & ""
'''  op = 0
'''  dr = 0
'''  CR = 0
'''  op = IIf(IsNull(fillvs(2)), 0, fillvs(2))
'''
'''
'''  If RS.State = 1 Then RS.Close
'''  RS.Open "select sum(dr),sum(cr) from tempLedger2 where party='" & fillvs(1) & "'", CON, adOpenDynamic, adLockOptimistic
'''  If Not IsNull(RS(0)) Then
'''     dr = RS(0)
'''  End If
'''
'''  If Not IsNull(RS(1)) Then
'''     CR = RS(1)
'''  End If
'''  If fillvs(3) = "Cr" Then
'''    op = (-1 * fillvs(2))
'''  End If
'''
'''  drcr = Round((op + (dr - CR)), 2)
'''  If Val(drcr) < 0 Then
'''     vsop.TextMatrix(i, 4) = Abs(Round((op + (dr - CR)), 2))
'''     vsop.TextMatrix(i, 5) = "Cr"
'''  Else
'''     vsop.TextMatrix(i, 4) = Round((op + (dr - CR)), 2)
'''     vsop.TextMatrix(i, 5) = "Dr"
'''  End If
'''
'''  drcr = Format(Round(drcr, 2), "0.00")
'''  If Val(drcr) < 0 Then
'''  CON.Execute "update sledger set Owner=" & Round(Val(drcr), 2) & " where subledger='" & fillvs(1) & "'"
'''  If bb = True Then
'''      contr.Execute "update sledger set OP=" & Abs(Round(Val(drcr), 2)) & ",drcr='Cr' where code='" & Trim(Mid(fillvs(1), 1, 6)) & "'"
'''  End If
'''  Else
'''      CON.Execute "update sledger set Owner=" & Round(Val(drcr), 2) & " where subledger='" & fillvs(1) & "'"
'''  If bb = True Then
'''     contr.Execute "update sledger set OP=" & Abs(Round(Val(drcr), 2)) & ",drcr='Dr' where code='" & Trim(Mid(fillvs(1), 1, 6)) & "'"
'''  End If
'''  End If
'''
'''
'''
'''  If Not IsNull(RS.Fields(1).Value) Then
'''  If RS.Fields(1).Value = 0 Then
'''  CON.Execute "update sledger set Offdays='" & "1" & "' where subledger='" & fillvs(1) & "'"
'''  Else
'''  CON.Execute "update sledger set Offdays='" & "2" & "' where subledger='" & fillvs(1) & "'"
'''  End If
'''  End If
'''
'''
'''  fillvs.MoveNext
'''  DoEvents
'''  DoEvents
'''  abc.Caption = abc.Caption - 1
'''Next
'''End If
'''vsop.Cols = 6
'''vsop.TextMatrix(0, 0) = "City"
'''vsop.TextMatrix(0, 1) = "Party"
'''vsop.TextMatrix(0, 2) = "Opening"
'''vsop.TextMatrix(0, 3) = "Dr/Cr"
'''vsop.TextMatrix(0, 4) = "Closing Balance"
'''vsop.TextMatrix(0, 5) = "Dr/Cr"
'''vsop.ColWidth(0) = 1800
'''vsop.ColWidth(1) = 3200
'''vsop.ColWidth(2) = 1200
'''vsop.ColWidth(3) = 500
'''vsop.ColWidth(4) = 1200
'''vsop.ColWidth(5) = 500
'''abc.Caption = ""
'''bb = False
'''Screen.MousePointer = vbDefault
'''
''''Exit Sub
'''
''''aa11:
''''MsgBox "" & "Connection Not Created Properly !", vbInformation
''''Screen.MousePointer = vbDefault

End Sub

Private Sub cmdShowClosing_Click()
showData



End Sub

Private Sub cmdupdatep_Click()
   Dim partyname
   Dim pcode
   partyname = ""
   pcode = ""
   
    
   If RS.State = 1 Then RS.Close
   RS.Open "select subledger from sledger", CON
   While RS.EOF = False
       
       aa = InStr(RS(0), " ")
       partyname = Mid(RS(0), aa)
       pcode = Mid(RS(0), 1, aa)
       
       CON.Execute "update  Sledger  set party='" & Trim(partyname) & "',code='" & Trim(pcode) & "' where subledger='" & RS(0) & "'"
       
       RS.MoveNext
       
   Wend
   
End Sub

Private Sub Command1_Click()
   
  If RS.State = 1 Then RS.Close
  RS.Open "select * from pass where pass='" & cp & "'", CON
  If RS.EOF = True Then
     MsgBox "Enter Valid Password !!", vbInformation
     Exit Sub
  
  Else

   Screen.MousePointer = vbHourglass
   
   On Error Resume Next
   
   For I = 1 To vsop.Rows - 1
       If vsop.TextMatrix(I, 1) <> "" Then
          CON.Execute "update SLEDGER set op=" & CDbl(vsop.TextMatrix(I, 2)) & ",drcr='" & vsop.TextMatrix(I, 3) & "' where SUBLEDGER='" & vsop.TextMatrix(I, 1) & "'"
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

Private Sub crdit_Click()
    If crdit.Value = True Then
       Call cmdshow_Click
    End If
End Sub

Private Sub credit_Click()
    If credit.Value = True Then
       Call cmdshow_Click
    End If

End Sub

Private Sub dbit_Click()
   If dbit.Value = True Then
       Call cmdshow_Click
    End If
End Sub

Private Sub Form_Activate()
Me.WindowState = 2
End Sub

Private Sub Form_Load()

vsIni
kk = 1
AddParty




If RS.State = 1 Then RS.Close
RS.Open "Select * from setup where " & stringyear & "", CON, adOpenStatic, adLockReadOnly, adCmdText
If RS.EOF = False Then
fromdate.Value = RS!yarfrom
todate.Value = RS!yarto
End If

from_date = fromdate.Value
to_date = todate.Value







maxId
setwidth

Me.TOP = 50
Me.Left = 50


ConOpen
If RS.State = 1 Then RS.Close
RS.Open "SELECT * FROM invoicea WHERE bAuthorized=0 ORDER BY InvoiceNo", CON, adOpenKeyset, adLockReadOnly
If Not RS.EOF Then
vs.Rows = RS.RecordCount + 1
Dim I As Integer
I = 1
While Not RS.EOF
vs.TextMatrix(I, 0) = I
vs.TextMatrix(I, 1) = RS!INVOICENO
vs.TextMatrix(I, 2) = RS!InvoiceDate
vs.TextMatrix(I, 3) = RS!SUBLEDGER
vs.TextMatrix(I, 4) = RS!netamount
vs.TextMatrix(I, 5) = RS!bAuthorized
RS.MoveNext
I = I + 1
Wend
End If
Screen.MousePointer = vbDefault

End Sub
Sub setsecurity()
   
If LCase(strledger) <> "cp" Then
   cmdShow1.Visible = False
   MsgBox "Enter Valid Password !!", vbInformation
   Exit Sub
Else
  
  
  
  savedata
   
End If
   
End Sub

Private Sub Label10_Click()

End Sub

Private Sub Label15_Click()
End Sub

Private Sub Opening_Click(PreviousTab As Integer)
      
     ' Screen.MousePointer = vbHourglass
      
      
      Dim closing As Double
      
      
      closing = 0
      
      If opening.Tab = 0 Then
         Call cmdshow_Click
         
      ElseIf opening.Tab = 2 Then
        
        
       
       
       
       
        Screen.MousePointer = vbHourglass

        
        Dim fillvs As New ADODB.Recordset
        If fillvs.State = 1 Then fillvs.Close
        'fillvs.Open "select DISTCODE as City,SUBLEDGER as Party,op,drcr from closing where gledger='SUNDRY DEBTORS'", con
        fillvs.Open "SELECT SLEDGER.DISTCODE,SLEDGER.SUBLEDGER,SLEDGER.OP,SLEDGER.drcr,(Sum(templedger1.Dr)-Sum(templedger1.Cr)) AS bal1 FROM SLEDGER LEFT JOIN templedger1 ON SLEDGER.SUBLEDGER = templedger1.Party where  gledger='SUNDRY DEBTORS' GROUP BY SLEDGER.SUBLEDGER,SLEDGER.DISTCODE,[SLEDGER.OP], SLEDGER.drcr, SLEDGER.gledger", CON
        
        If fillvs.EOF = False Then
            vsop.Rows = fillvs.RecordCount
            For I = 1 To vsop.Rows - 1
              vsop.TextMatrix(I, 0) = fillvs(0) & ""
              vsop.TextMatrix(I, 1) = fillvs(1)
              vsop.TextMatrix(I, 2) = Format(fillvs(2), "0.00")
              vsop.TextMatrix(I, 3) = fillvs(3) & ""
               
              If Not IsNull(fillvs(4)) Then
                     
                     If vsop.TextMatrix(I, 3) = "Cr" Then
                         vsop.TextMatrix(I, 4) = ((-1 * (vsop.TextMatrix(I, 2))) + fillvs(4))
                         If vsop.TextMatrix(I, 4) < 0 Then
                            vsop.TextMatrix(I, 4) = Format((-1 * Val(vsop.TextMatrix(I, 4))), "0.00")
                            vsop.TextMatrix(I, 5) = "Cr"
                         Else
                            vsop.TextMatrix(I, 5) = "Dr"
                            vsop.TextMatrix(I, 4) = Format(Val(vsop.TextMatrix(I, 4)), "0.00")
                         End If
                         
                     Else
                         vsop.TextMatrix(I, 4) = ((Val(vsop.TextMatrix(I, 2))) + fillvs(4))
                         If vsop.TextMatrix(I, 4) < 0 Then
                            vsop.TextMatrix(I, 4) = Format((-1 * Val(vsop.TextMatrix(I, 4))), "0.00")
                            vsop.TextMatrix(I, 5) = "Cr"
                         Else
                            vsop.TextMatrix(I, 5) = "Dr"
                            vsop.TextMatrix(I, 4) = Format(Val(vsop.TextMatrix(I, 4)), "0.00")
                         End If
                         
                         
                     End If
              End If
              
             
              fillvs.MoveNext
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
        

      End If
      
      
      
      'Screen.MousePointer = vbDefault
      
End Sub
Private Sub Option1_Click()
   If Option1.Value = True Then
      bill.Visible = True
   Else
      partydrcr.Visible = False
      bill.Visible = True
   End If
End Sub

Private Sub Option2_Click()
   If Option2.Value = 1 Then
      txtadmin.Visible = True
      Label14.Visible = True
   Else
      txtadmin.Visible = False
      Label14.Visible = False
   End If
End Sub

Private Sub party_Click()
   
   If party.Value = True Then
      bill.Visible = False
      frmReceiveFromParty.Show
   Else
      partydrcr.Visible = False
      bill.Visible = True
   End If
   
   frmReceiveFromParty.TOP = 800

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
    If sales.Value = True Then
       Call cmdshow_Click
    End If
End Sub

Private Sub selectAll_Click()
If SelectAll.Value = 1 Then
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
   pass.Visible = False
End If
End Sub

Private Sub txtdes_GotFocus()
  txtdes.BackColor = &HFFFFC0
End Sub

Private Sub txtdes_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     txtqty.SetFocus
  End If
End Sub

Private Sub txtdes_LostFocus()
    txtdes.BackColor = &HFFFFFF
End Sub

Private Sub txtOp_GotFocus()
txtop.BackColor = &HFFFFC0
End Sub
Private Sub txtparty_GotFocus()
   If PopUpValue1 <> "" Then
      txtparty.Text = PopUpValue1
      PopUpValue1 = ""
   End If
End Sub
Private Sub txtparty_KeyDown(KeyCode As Integer, Shift As Integer)
'     If KeyCode = 113 Then
'       Value = "select SUBLEDGER from INVOICEA order by SUBLEDGER"
'       popuplist12 Value, CON
'    End If
End Sub
Private Sub txtParty_LostFocus()
  PopUpValue1 = ""
End Sub
Private Sub txtqty_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
   If Val(txtqty.Text) = 0 Then
      txtqty.SetFocus
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
     'con.Execute "update ReceiveIssueParty set Remarks = '" & txtRem.Text & "' where PartyName='" & cboParty.Text & "'"
     CON.Execute "update sledger set PartyRemarks = '" & txtrem.Text & "' where subledger='" & cboParty.Text & "'"
     
  End If
  End If
End Sub
Private Sub Unautho_Click()
If Unautho.Value = True Then
    Call cmdshow_Click
End If
End Sub
Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)
Screen.MousePointer = vbHourglass
If KeyCode = 13 Then
If sales.Value = True Then
   If vs.Col = 1 Then
         inviceNo = vs.TextMatrix(vs.RowSel, 1)
         'MainMenu.Toolbar1.Visible = False
         INVOICE.Show  '    sales
   End If
ElseIf cash.Value = True Then
   If vs.Col = 1 Then
         inviceNo = vs.TextMatrix(vs.RowSel, 1)
         'MainMenu.Toolbar1.Visible = False
         countersale.Show  '
   End If
ElseIf credit.Value = True Then
   If vs.Col = 1 Then
         inviceNo = vs.TextMatrix(vs.RowSel, 1)
         'MainMenu.Toolbar1.Visible = False
         CRITNOTE.Show
   End If
ElseIf crdit.Value = True Then
   If vs.Col = 1 Then
         inviceNo = vs.TextMatrix(vs.RowSel, 1)
         'MainMenu.Toolbar1.Visible = False
         Creditnotefile.Show
   End If
ElseIf dbit.Value = True Then
   If vs.Col = 1 Then
         inviceNo = vs.TextMatrix(vs.RowSel, 1)
         'MainMenu.Toolbar1.Visible = False
         Debitnotefile.Show
   End If
End If
End If
Screen.MousePointer = vbDefault
End Sub
Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If vs.Col = 4 Then
   SendKeys "{down}"
End If
End If
End Sub
Sub CalculateTotalDrCr()
On Error Resume Next
Dim Balance As Long
Dim dr1, cr1, prbal
Dim str
str = ""
dr1 = 0
cr1 = 0
txtClosing.Text = 0
txtCr.Text = 0
If RS.State = 1 Then RS.Close
RS.Open "select Op,drcr from SLEDGER where SUBLEDGER='" & cboParty.Text & "'", CON
If RS.EOF = False Then
txtop.Text = Format(RS.Fields(0).Value, "0.00")
If UCase(RS.Fields("drcr").Value) = UCase("dr") Then
cboop.Text = "Dr"
Else
cboop.Text = "Cr"
End If
Else
txtop.Text = 0
End If
If cboop.Text = "Dr" Then
dr1 = (Val(txtop.Text) + Val(vs1.TextMatrix(1, 4)))
cr1 = Val(vs1.TextMatrix(1, 5))
Else
cr1 = (Val(txtop.Text) + Val(vs1.TextMatrix(1, 5)))
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
For I = 1 To vs1.Rows - 1
If vs1.TextMatrix(I, 0) <> "" Then
txtClosing.Text = (Val(txtClosing.Text) + Val(vs1.TextMatrix(I, 4)))
txtCr.Text = (Val(txtCr.Text) + Val(vs1.TextMatrix(I, 5)))
'-----Balance---------------
If I >= 2 Then
dr1 = Val(vs1.TextMatrix(I, 4))
cr1 = (-1 * Val(vs1.TextMatrix(I, 5)))
bal = dr1 + cr1
If str = "Cr" Then
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
str = vs1.TextMatrix(I, 7)
End If
'---------------------------
End If
Next
txtClosing.Text = Format(txtClosing.Text, "0.00")
txtCr.Text = Format(txtCr.Text, "0.00")
If cboop.Text = "Dr" Then
txtClosing.Text = Format((CDbl(txtClosing.Text) + CDbl(txtop.Text)), "0.00")
Else
txtCr.Text = Format((CDbl(txtCr.Text) + CDbl(txtop.Text)), "0.00")
End If
txtBalance.Text = (Val(txtClosing.Text) - Val(txtCr.Text))
If Val(txtBalance.Text) < 1 Then
txtBalance.Text = (-1 * Val(txtBalance.Text))
closingcr.Text = "Cr"
Else
closingcr.Text = "Dr"
End If
txtBalance.Text = Format(txtBalance.Text, "0.00")
End Sub
Sub SaveDatainTempledger()
CON.Execute "delete * from templedger1"
For I = 1 To vs1.Rows - 1
If vs1.TextMatrix(I, 1) <> "" Then
CON.Execute "INSERT INTO  templedger1(dates,Billtype,Bill,Des,Dr,Cr,Balance,drcr)  values('" & vs1.TextMatrix(I, 2) & "','" & vs1.TextMatrix(I, 0) & "', " & vs1.TextMatrix(I, 1) & ",'" & vs1.TextMatrix(I, 3) & "' ," & vs1.TextMatrix(I, 4) & "," & vs1.TextMatrix(I, 5) & "," & Val(vs1.TextMatrix(I, 6)) & ",'" & vs1.TextMatrix(I, 7) & "')"
End If
Next
Dim ff As New ADODB.Recordset
If ff.State = 1 Then ff.Close
ff.Open "select Billtype,bill,dates,des,dr,cr,Balance,drcr from templedger1 order by dates,bill", CON
vs1.Rows = ff.RecordCount + 1
For J = 1 To vs1.Rows - 1
 If ff.EOF = False Then
     vs1.TextMatrix(J, 0) = ff.Fields(0).Value
     vs1.TextMatrix(J, 1) = ff.Fields(1).Value
     vs1.TextMatrix(J, 2) = ff.Fields(2).Value
     vs1.TextMatrix(J, 3) = ff.Fields(3).Value
     vs1.TextMatrix(J, 4) = Format(ff.Fields(4).Value, "0.00")
     vs1.TextMatrix(J, 5) = Format(ff.Fields(5).Value, "0.00")
     vs1.TextMatrix(J, 6) = Format(ff.Fields(6).Value, "0.00")
     ff.MoveNext
 End If
Next
End Sub
Private Sub cboParty_GotFocus()
Dim ph_rs As New ADODB.Recordset
cboParty.BackColor = &HFFFFC0
I = 1
If PopUpValue3 = "" Then
PopUpValue2 = cboParty.Text
End If
If PopUpValue3 <> "" Then
cboParty.Text = PopUpValue3
Set ph_rs = New ADODB.Recordset
ph_rs.Open "select phone,PartyRemarks from sledger where subledger='" & cboParty.Text & "'", CON
If ph_rs.EOF = False Then
   Phone.Caption = ph_rs(0) & ""
   txtrem.Text = ph_rs.Fields("PartyRemarks").Value & ""
Else
   Phone.Caption = ""
   txtrem.Text = ""
End If
End If
End Sub

Private Sub cboParty_KeyDown(KeyCode As Integer, Shift As Integer)
''If KeyCode = 113 Then
''Value = "select distinct(Party),Code,subledger from SLEDGER where gledger='SUNDRY DEBTORS' order by party"
''popuplist12 Value, CON
''set_focus = True
''End If
''If KeyCode = 13 Then
''
''
''If cboParty.Text = "" Then
''   cboParty.SetFocus
''   Exit Sub
''End If
''
'''txtRem.Text = ""
''dataSearchingrid
''cmdPrint.Enabled = True
'''If set_focus = True Then
'''vs1.SetFocus
'''For I = 1 To vs1.Rows - 1
'''SendKeys "{down}"
'''Next
'''Else
'''txtdes.SetFocus
'''End If
''Dim dr, CR As Double
''dr = 0
''CR = 0
''
''For i = 1 To vs1.Rows - 1
''dr = dr + Val(vs1.TextMatrix(i, 4))
''CR = CR + Val(vs1.TextMatrix(i, 5))
''Next
''drLebel.Caption = Format(dr, "0.00")
''CrLebel.Caption = Format(CR, "0.00")
''
''
''
''End If
End Sub
Sub dataSearchingrid()
Screen.MousePointer = vbHourglass
I = 1
txtdes.SetFocus
If PopUpValue3 <> "" Then
vs1.Clear
vs1.Rows = 1
fillGrid
End If
If cboParty.Text <> "" Then
SaveDatainTempledger
CalculateTotalDrCr
End If
setwidth
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
    If Del.State = 1 Then Del.Close
    Set Del = CON.Execute("delete from ReceiveIssueParty where RecNo=" & txtRecNo.Text & "")
End Sub
Private Sub cmdDel_Click()
  If RS.State = 1 Then RS.Close
  RS.Open "select * from pass where pass='" & cp & "'", CON
  If RS.EOF = True Then
     MsgBox "Enter Valid Password !!", vbInformation
     Exit Sub
  End If
   If MsgBox("Do U Want To Delete ?", vbQuestion + vbYesNo) = vbYes Then
       DelFunction
       fillGrid
       dataSearchingrid
       Call cmdRefresh_Click
       cmdmodify.Enabled = False
       cmdDel.Enabled = False
   End If
End Sub
Private Sub cmdMain_Click()
If strledger = "cp" Then
If Val(txtqty.Text) > 0 And txtdes.Text <> "" And cboParty.Text <> "" Then
   If MsgBox("Want To Save & Exit ?", vbQuestion + vbYesNo) = vbYes Then
          SaveMain
          Call cmdRefresh_Click
          fillGrid
          cmdmodify.Enabled = False
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
Sub setwidth()
'    vs1.FormatString = "^Bill Type|^Bill|^Dates|Description|>Dr|>Cr|Balance|Dr/Cr"
'    vs1.ColWidth(0) = 500
'    vs1.ColWidth(1) = 1000
'    vs1.ColWidth(2) = 1000
'    vs1.ColWidth(3) = 2400
'    vs1.ColWidth(4) = 1200
'    vs1.ColWidth(5) = 1200
'    vs1.ColWidth(6) = 1300
'    vs1.ColWidth(7) = 500
'   DoEvents
End Sub
Private Sub cmdModify_Click()
Set RS = New ADODB.Recordset
RS.Open "select * from pass where pass='" & cp & "'", CON
If RS.EOF = True Then
MsgBox "Enter Valid Password !!", vbInformation
Exit Sub
End If
'On Error GoTo aa1
If MsgBox("Do U Want To Update ?", vbQuestion + vbYesNo) = vbYes Then
'DelFunction
CON.Execute "update ReceiveIssueParty set Dr=0,cr=0 where RecNo=" & txtRecNo.Text & ""

'------------------------
Set RS = New ADODB.Recordset
RS.Open "select * from ReceiveIssueParty where RecNo=" & txtRecNo.Text & "", CON, adOpenDynamic, adLockOptimistic
If RS.EOF = False Then
'maxId
'RS.AddNew
RS.Fields("RecNo").Value = txtRecNo.Text
RS.Fields("Dates").Value = RecDates.Value
RS.Fields("PartyName").Value = cboParty.Text
RS.Fields("Particullar").Value = txtdes.Text
If Receive.Value = True Then
RS.Fields("Dr").Value = Val(txtqty.Text)
Else
RS.Fields("Cr").Value = Val(txtqty.Text)
End If
RS.Update
End If
'------------------------

'SaveMain


fillGrid
CalculateTotalDrCr
setwidth
Call cmdRefresh_Click
vs1.SetFocus
For I = 1 To vs1.Rows - 1
SendKeys "{down}"
Next

cmdmodify.Enabled = False
cmdDel.Enabled = False
End If
'Exit Sub
'aa1:
'MsgBox "Record not Save !!", vbCritical
End Sub
Private Sub cmdRefresh_Click()
 
 'con.Execute "update  INVOICEA  set  INVOICEA.t2 = & " select INVOICEC.TEXT FROM INVOICEA INNER JOIN INVOICEC ON INVOICEA.INVOICENO = INVOICEC.INVOICENO Where INVOICEC.Text = '" & "DIWALI SPECIAL" & "' And ((netamount - BAA) > 0 Or (netamount - BAA) < 0) And INVOICEC.amount > 0"
 
 
''If RS.State = 1 Then RS.Close
''RS.Open "select PartyName,Remarks from ReceiveIssueParty group by PartyName,Remarks", con
''While RS.EOF = False
''   con.Execute "update sledger set PartyRemarks='" & RS(1) & "' where subledger='" & RS(0) & "'"
''   RS.MoveNext
''Wend
''
''MsgBox "updated .."
 
 
 Dim o As Object
 txtqty.Text = ""
 'txtdes.Text = ""
 set_focus = False
 maxId
 cmdmodify.Enabled = False
 cmdDel.Enabled = False
 cmdSave.Enabled = True
 cboParty.SetFocus
 'txtRem.Visible = False
 Screen.MousePointer = vbDefault
 bb2 = False
End Sub
Private Sub cmdSave_Click()

On Error GoTo aa:



If cboParty.Text = "" Then
MsgBox "Please Select Party Name !!", vbInformation
Exit Sub
End If

If txtqty.Text = "" Then
MsgBox "Please Enter Amount!!", vbInformation
txtqty.SetFocus
Exit Sub
End If


If RS.State = 1 Then RS.Close
RS.Open "select * from pass where pass='" & cp & "'", CON
If RS.EOF = True Then
MsgBox "Enter Valid Password !!", vbInformation
Exit Sub
End If

If MsgBox("Do U Want To Save ?", vbInformation + vbYesNo) = vbYes Then
 
aa1:
 SaveMain

 fillGrid
 cmdmodify.Enabled = False
 cmdDel.Enabled = False
 '----------------
 

 dataSearchingrid
 '---------------
 

 
 

Call cmdRefresh_Click
'cmdSave.Enabled = False
End If
Exit Sub
aa:
'MsgBox "Record Not Save !!", vbCritical

maxId
GoTo aa1

End Sub
Sub SaveMain()
   
   maxId
    Set RS = New ADODB.Recordset
    RS.Open "select * from ReceiveIssueParty where RecNo=" & txtRecNo.Text & "", CON, adOpenDynamic, adLockOptimistic
    If RS.EOF = True Then
       maxId
       RS.addNew
       RS.Fields("RecNo").Value = txtRecNo.Text
       RS.Fields("Dates").Value = RecDates.Value
       RS.Fields("PartyName").Value = cboParty.Text
       RS.Fields("Particullar").Value = txtdes.Text
       If Receive.Value = True Then
          RS.Fields("Dr").Value = Val(txtqty.Text)
        Else
          RS.Fields("Cr").Value = Val(txtqty.Text)
       End If
    
    RS.Update
    End If
End Sub
Sub search()
 If set_focus = True Then Exit Sub
 On Error Resume Next
 
 
 
    If rss.State = 1 Then rss.Close
    rss.Open "select * from sledger where subledger=" & txtparty.Text & "", CON, adOpenDynamic, adLockOptimistic
    If rss.EOF = 1 Then
       txtrem.Text = RS.Fields("PartyRemarks").Value & ""
    End If

 
 
 
 If vs1.TextMatrix(vs1.RowSel, 0) = "J" Then
    If RS.State = 1 Then RS.Close
    RS.Open "select * from ReceiveIssueParty where RecNo=" & vs1.TextMatrix(vs1.RowSel, 1) & "", CON, adOpenDynamic, adLockOptimistic
    If RS.EOF = False Then
       txtRecNo.Text = RS.Fields("RecNo").Value
       RecDates.Value = RS.Fields("Dates").Value
       cboParty.Text = RS.Fields("PartyName").Value
       txtdes.Text = RS.Fields("Particullar").Value
       
       
       If RS.Fields("Dr").Value > 0 Then
          Receive.Value = True
          txtqty.Text = RS.Fields("Dr").Value
        Else
          Issue.Value = True
          txtqty.Text = RS.Fields("Cr").Value
       End If
      End If
   cmdSave.Enabled = False
   cmdmodify.Enabled = True
   cmdDel.Enabled = True
  Else
   cmdmodify.Enabled = False
   cmdDel.Enabled = False
   cmdSave.Enabled = True
   txtdes.Text = ""
   txtqty.Text = ""
  End If
End Sub
Private Sub cmdSearch_Click()
Frame1.Visible = True
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  
  
  
  If KeyCode = 27 Then
'   If cboPartyList.Visible = True Then
'      cboPartyList.Visible = False
      Exit Sub
  ' End If
  End If
  
  
  
  If KeyCode = 116 Then
  If bb2 = False Then
    vs1.SetFocus
    For I = 1 To vs1.Rows - 1
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
        If RS.State = 1 Then RS.Close
        RS.Open "select * from pass where pass='" & cp & "'", CON
        If RS.EOF = True Then
          If MsgBox("Want To Exit ?", vbQuestion + vbYesNo) = vbYes Then
           Unload Me
           End If
        Exit Sub
        End If
        If Val(txtqty.Text) > 0 And txtdes.Text <> "" And cboParty.Text <> "" Then
        If MsgBox("Want To Save & Exit ?", vbQuestion + vbYesNo) = vbYes Then
            SaveMain
            Call cmdRefresh_Click
            fillGrid
            cmdmodify.Enabled = False
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
    
    
    
    'If rsS.State = 1 Then rsS.Close
    'rsS.Open "select * from sledger where subledger=" & txtParty.Text & "", con, adOpenDynamic, adLockOptimistic
    'If rsS.EOF = False Then
    '   txtRem.Text = RS.Fields("PartyRemarks").Value & ""
    'End If

    
    
    
    Dim fill As New ADODB.Recordset
    Set fill = New ADODB.Recordset
    fill.Open "select RecNo,Dates,Particullar,Dr,Cr,Remarks from ReceiveIssueParty where PartyName='" & cboParty.Text & "' order by dates,recno", CON
    If fill.EOF = False Then
       'txtRem.Text = fill!Remarks & ""
       vs1.Rows = fill.RecordCount + 1
       For I = 1 To vs1.Rows - 1
           vs1.TextMatrix(I, 0) = "J"
           vs1.TextMatrix(I, 1) = fill.Fields(0).Value
           vs1.TextMatrix(I, 2) = fill.Fields(1).Value
           vs1.TextMatrix(I, 3) = fill.Fields(2).Value
           vs1.TextMatrix(I, 4) = Format(fill.Fields(3).Value, "0.00")
           vs1.TextMatrix(I, 5) = Format(fill.Fields(4).Value, "0.00")
              
           fill.MoveNext
       Next
    Else
    vs1.Clear
    End If

    
    '==============
    SearchFa
    '==============
    setwidth
End Sub
Sub maxId()
'  Dim rr As New ADODB.Recordset
'  Set rr = New ADODB.Recordset
'  rr.Open "select max(RecNo) from ReceiveIssueParty", con
'  If IsNull(rr.Fields(0).Value) Then
'     txtRecno.Text = 1
'     Else
'     txtRecno.Text = rr.Fields(0).Value + 1
'  End If
End Sub
Sub AddParty()
'    Dim rs_S As New ADODB.Recordset
'    If rs_S.State = 1 Then rs_S.Close
'    rs_S.Open "select distinct(BrokerName) from PartyMaster", con
'    cboParty.Clear
'    If rs_S.EOF = False Then
'       While rs_S.EOF = False
'          cboParty.AddItem rs_S.Fields(0).Value
'          rs_S.MoveNext
'       Wend
'    End If
End Sub
Private Sub Issue_Click()
    ' AddParty
End Sub
Private Sub Receive_Click()
   '  AddParty
End Sub
Private Sub todate_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
    from_date = fromdate.Value
    to_date = todate.Value
    fillGrid
    Frame1.Visible = False
 End If
End Sub

Private Sub txtQty_GotFocus()
   txtqty.BackColor = &HFFFFC0
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
'   Dim b As Boolean
'   b = val_int(txtQty, KeyAscii)
'   If b = False Then
'      KeyAscii = 0
'   End If
End Sub

Private Sub txtQty_LostFocus()
  txtqty.BackColor = &HFFFFFF
End Sub
Private Sub txtRecNo_KeyPress(KeyAscii As Integer)
''   On Error Resume Next
''
''     bb = val_int(txtRecno, KeyAscii)
''     If bb = False Then
''        KeyAscii = 0
''     End If
''
''  If KeyAscii = 13 Then
''
''     If RS.State = 1 Then RS.Close
''     RS.Open "select * from receiveissueparty where recno=" & txtRecno.Text & "", CON
''     If RS.EOF = False Then
''      cboParty.Text = RS!partyname
''      PopUpValue3 = cboParty.Text
''
''      RecDates.Value = RS.Fields("Dates").Value
''      txtdes.Text = RS.Fields("Particullar").Value
''      'txtRem.Text = RS.Fields("Remarks").Value
''      If RS.Fields("Dr").Value > 0 Then
''          Receive.Value = True
''          txtQty.Text = RS.Fields("Dr").Value
''      Else
''          Issue.Value = True
''          txtQty.Text = RS.Fields("Cr").Value
''      End If
''      dataSearchingrid
''     Else
''       vs1.Clear
''       setwidth
''       txtQty.Text = ""
''       txtdes.Text = ""
''       cboParty.Text = ""
''       txtOp.Text = ""
''       txtBalance.Text = ""
''     End If
''  End If
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
Private Sub vs1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 116 Then
    Call cmdRefresh_Click
    cboParty.SetFocus
    Exit Sub
End If

Screen.MousePointer = vbHourglass
If KeyCode = 13 Then
If vs1.TextMatrix(vs1.RowSel, 0) = "I" Then
   'If vs1.col = 1 Then
         inviceNo = vs1.TextMatrix(vs1.RowSel, 1)
         'MainMenu.Toolbar1.Visible = False
         INVOICE.Show
   'End If

ElseIf vs1.TextMatrix(vs1.RowSel, 0) = "CI" Then

   'If vs1.col = 1 Then
         inviceNo = vs1.TextMatrix(vs1.RowSel, 1)
         'MainMenu.Toolbar1.Visible = False
         CRITNOTE.Show
   'End If

ElseIf vs1.TextMatrix(vs1.RowSel, 0) = "CN" Then

   'If vs1.col = 1 Then
         inviceNo = vs1.TextMatrix(vs1.RowSel, 1)
         'MainMenu.Toolbar1.Visible = False
         Creditnotefile.Show
   'End If

ElseIf vs1.TextMatrix(vs1.RowSel, 0) = "DN" Then

   'If vs1.col = 1 Then
         inviceNo = vs1.TextMatrix(vs1.RowSel, 1)
         'MainMenu.Toolbar1.Visible = False
         Debitnotefile.Show
   'End If

ElseIf vs1.TextMatrix(vs1.RowSel, 0) = "C/M" Then

   'If vs1.col = 1 Then
         inviceNo = vs1.TextMatrix(vs1.RowSel, 1)
         'MainMenu.Toolbar1.Visible = False
         countersale.Show
   'End If

End If


End If


If KeyCode = 112 Then
   txtdes.SetFocus
End If

Screen.MousePointer = vbDefault

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
    setwidth
    PopUpValue1 = ""
    opening.Tab = 1
End If
End Sub
