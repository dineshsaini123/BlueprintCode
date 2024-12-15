VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmSLedgerView 
   Caption         =   "S. Ledger"
   ClientHeight    =   9912
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   11016
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9912
   ScaleWidth      =   11016
   Begin VB.CommandButton cmdPrintBillList 
      BackColor       =   &H00FAEFC9&
      Caption         =   "&Print Bill List"
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
      Left            =   2340
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   840
      Width           =   990
   End
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
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   840
      Width           =   990
   End
   Begin Crystal.CrystalReport crpt 
      Left            =   660
      Top             =   9720
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
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
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   840
      Width           =   990
   End
   Begin VB.ComboBox COMBOGENLEDGER 
      Height          =   315
      Left            =   12780
      Sorted          =   -1  'True
      TabIndex        =   5
      Text            =   "SUNDRY DEBTORS"
      Top             =   480
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
      Left            =   3960
      TabIndex        =   1
      Top             =   1068
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
      Left            =   4020
      TabIndex        =   4
      Top             =   1056
      Visible         =   0   'False
      Width           =   1356
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
      Left            =   1320
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   0
      Top             =   360
      Width           =   5805
   End
   Begin VB.TextBox Alpha 
      Height          =   315
      Left            =   12840
      MaxLength       =   1
      TabIndex        =   3
      Top             =   60
      Visible         =   0   'False
      Width           =   705
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
      Left            =   5940
      TabIndex        =   6
      Top             =   1140
      Visible         =   0   'False
      Width           =   1155
      _ExtentX        =   2032
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
      Left            =   5460
      TabIndex        =   7
      Top             =   1140
      Visible         =   0   'False
      Width           =   1155
      _ExtentX        =   2032
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
      Height          =   7755
      Left            =   60
      TabIndex        =   2
      Top             =   1560
      Width           =   10650
      _cx             =   18785
      _cy             =   13679
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
      RowHeightMin    =   350
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
      Left            =   10260
      TabIndex        =   20
      Top             =   360
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
      Left            =   8700
      TabIndex        =   19
      Top             =   360
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
      Left            =   1320
      TabIndex        =   18
      Top             =   60
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
      Left            =   10260
      TabIndex        =   17
      Top             =   780
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
      Left            =   7620
      TabIndex        =   16
      Top             =   780
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
      Left            =   7620
      TabIndex        =   15
      Top             =   360
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
      Left            =   8700
      TabIndex        =   14
      Top             =   780
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
      Left            =   7080
      TabIndex        =   13
      Top             =   9360
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
      Left            =   8160
      TabIndex        =   12
      Top             =   9360
      Width           =   975
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
      Left            =   11400
      TabIndex        =   11
      Top             =   5700
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
      Left            =   11400
      TabIndex        =   10
      Top             =   6060
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Sub. Ledger :"
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
      Left            =   60
      TabIndex        =   9
      Top             =   390
      Width           =   1515
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
      Left            =   6540
      TabIndex        =   8
      Top             =   1140
      Visible         =   0   'False
      Width           =   585
   End
End
Attribute VB_Name = "frmSLedgerView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As Recordset
Dim b1 As Integer
Function rsets(ST As String, length As Integer) As String
Dim kk As String
    kk = Trim(ST)
    If Len(kk) < length Then
        Do While Not Len(kk) = length
            kk = " " + kk
        Loop
    End If
    If Len(kk) > length Then
        Do While Not Len(kk) = length
            kk = Mid$(kk, 0, Len(kk) - 1)
        Loop
    End If
rsets = kk
End Function
Private Sub alpha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    sendkeys "{TAB}"
End If
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()

On Error GoTo aa10

Screen.MousePointer = vbHourglass
Dim op, drcr
Dim rs1 As New ADODB.Recordset
Dim rptid

'login.DSN


con.Execute "delete from templedger3 where Party='" & Combosubledger.text & "'"


If rs1.State = 1 Then rs1.Close
rs1.Open "select max(rptid) from templedger3", con, adOpenDynamic, adLockOptimistic
If IsNull(rs1(0)) Then
rptid = 9999
Else
rptid = rs1(0) + 1
End If


If lblOp = "" Then lblOp = 0


If lbl_crdr = "Cr" Then
   op = (lblOp * -1)
Else
   op = lblOp
End If

con.Execute "INSERT INTO templedger3 (dates,Balance,Party,drcr,fyear,setupid,rptid,billtype) values('" & fromDate_setup & "'," & op & ",'" & Combosubledger.text & "','" & Trim(lbl_crdr) & "','" & session & "','" & setupid & "','" & rptid & "','opening')"

For I = 1 To vs.rows - 1
If vs.TextMatrix(I, 1) <> "" Then
con.Execute "INSERT INTO templedger3 (dates,Billtype,bill,des,dr,cr,Party,Balance,drcr,fyear,setupid,rptid) values('" & Format(vs.TextMatrix(I, 1), "MM/dd/yyyy") & "'," & _
"'" & vs.TextMatrix(I, 2) & "'," & vs.TextMatrix(I, 0) & ",'" & vs.TextMatrix(I, 3) & "'," & Val(vs.TextMatrix(I, 4)) & "," & _
"" & Val(vs.TextMatrix(I, 5)) & ",'" & Combosubledger.text & "'," & Val(vs.TextMatrix(I, 6)) & ",'" & vs.TextMatrix(I, 7) & "','" & session & "','" & setupid & "','" & rptid & "')"
End If
Next


DSNNew


If MsgBox("Want to Send Mail", vbYesNo) = vbNo Then
    crpt.Reset
    crpt.ReportFileName = rptPath & "/PartyLedger_new.rpt"
    crpt.ReplaceSelectionFormula "{tempLedgerrpt.party}='" & Combosubledger.text & "'"
    crpt.Connect = constr
    crpt.WindowShowPrintSetupBtn = True
    crpt.WindowShowPrintBtn = True
    crpt.WindowState = crptMaximized
    crpt.Action = 1
    Screen.MousePointer = vbDefault

Else

   Screen.MousePointer = vbDefault
   popupvalue5 = rptid
   popupvalue4 = "SLLedger.rpt"
   frmSendMail.Show 1

End If




Screen.MousePointer = vbDefault
Exit Sub


aa10:
MsgBox err.Description


End Sub

Private Sub cmdPrintBillList_Click()

'login.DSN
DSNNew

crpt.Reset
crpt.ReportFileName = rptPath & "/PartyBillList.rpt"
crpt.Connect = constr
crpt.ReplaceSelectionFormula "{PartyWiseItemWiseQty.subledger}='" & Combosubledger.text & "'"
crpt.WindowShowPrintSetupBtn = True
crpt.WindowState = crptMaximized
crpt.Action = 1


End Sub

Private Sub COMBOGENLEDGER_Change()
    If RS.State = 1 Then
        RS.Close
    End If
    RS.Open "select * from sledger where " & stringyear & " and gledger='" + Trim(COMBOGENLEDGER.text) + "'", con, adOpenStatic, adLockReadOnly, adCmdText
    Combosubledger.Clear
    If Not RS.BOF Then
        Do While Not RS.EOF
            Combosubledger.AddItem Trim(RS!subledger)
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.Close
    
End Sub

Private Sub COMBOGENLEDGER_Click()
If RS.State = 1 Then
        RS.Close
End If

    RS.Open "select * from sledger where " & stringyear & " and gledger='" + Trim(COMBOGENLEDGER.text) + "'", con, adOpenStatic, adLockReadOnly, adCmdText
    Combosubledger.Clear
    If Not RS.BOF Then
        Do While Not RS.EOF
            Combosubledger.AddItem Trim(RS!subledger)
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.Close
End Sub

Private Sub COMBOGENLEDGER_KeyPress(KeyAscii As Integer)
If VB.Screen.ActiveForm.ActiveControl.SelLength > 0 Then
   sendkeys "{tab}"
   Exit Sub
End If
If KeyAscii = 13 Then
   sendkeys "{DOWN}"
   sendkeys "{tab}"
End If

End Sub

Private Sub COMBOGENLEDGER_LostFocus()
    If Trim(COMBOGENLEDGER.text) <> "" Then
        If RS.State = 1 Then RS.Close
        RS.Open "select * from gledger where " & stringyear & " and slf=true", con, adOpenStatic, adLockReadOnly, adCmdText
        If Not RS.BOF Then
            RS.Find "gledger='" + Trim(COMBOGENLEDGER.text) + "'"
            If RS.EOF Then
                COMBOGENLEDGER.SetFocus
            End If
        Else
            COMBOGENLEDGER.SetFocus
        End If
        RS.Close
    End If
End Sub

Private Sub Combosubledger_GotFocus()
 
    
Screen.MousePointer = vbHourglass
    
If PopUpValue3 <> "" Then
If b1 = 1 Then
  Combosubledger.text = PopUpValue3
 Else
  Combosubledger.text = PopUpValue2
End If

End If

If PopUpValue3 <> "" Then
    Call Commandshow_Click
    sendkeys "{Down}"
    sendkeys "{tab}"
End If

  PopUpValue3 = ""
  popupvalue5 = ""


Screen.MousePointer = vbDefault
    
End Sub

Private Sub Combosubledger_KeyDown(KeyCode As Integer, Shift As Integer)


If KeyCode = 113 Then
    b1 = 1
    searchType = "party"
    value = "select distinct(Party),Code,subledger from SLEDGER where gledger='SUNDRY DEBTORS' and " & stringyear & " order by party"
    'popuplistModel10 value, CON
    popuplist_client value, CCON
    set_focus = True
End If





If KeyCode = 116 Then
vs.SetFocus
For J = 1 To vs.rows - 1
   sendkeys "{down}"
   vs.Row = J
Next
End If


End Sub

Private Sub Combosubledger_KeyPress(KeyAscii As Integer)
If VB.Screen.ActiveForm.ActiveControl.SelLength > 0 Then
   sendkeys "{tab}"
   Exit Sub
End If
If KeyAscii = 13 Then
      
End If
End Sub

Private Sub Combosubledger_LostFocus()
On Error Resume Next

If Trim(Combosubledger.text) <> "" Then
    If Trim(COMBOGENLEDGER.text) <> "" Then
        If RS.State = 1 Then
            RS.Close
        End If
        RS.Open "select * from sledger where " & stringyear & " and gledger='" + Trim(COMBOGENLEDGER.text) + "' and subledger='" + Trim(Combosubledger.text) + "'", con, adOpenStatic, adLockReadOnly, adCmdText
        If RS.BOF Then
            Combosubledger.SetFocus
        End If
        RS.Close
    Else
        Combosubledger.text = ""
    End If
End If
End Sub

 Sub ALPHAB()
    If RS.State = 1 Then RS.Close
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    con.Execute ("delete  from treport where " & stringyear & "")
    
    If DateDiff("d", Trim(date1.text), Trim(date2.text)) <= 0 Then
        MsgBox "invalid date"
        Exit Sub
    End If
    Dim rs1 As New ADODB.Recordset
    Dim Balance As Double
    Dim OPBALANCE As Double
    Dim SDamount As Double
    Dim SCamount As Double
    Dim RsT As New ADODB.Recordset
    Dim viewsubledger As Boolean
    viewsubledger = False
    Balance = 0
    OPBALANCE = 0
    OPENINGSUBLEDGERS
    DoEvents
    If Trim(Alpha.text) <> "" And Alpha.Visible = True Then
      ' vouchers creditors
      con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid )  SELECT VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.DESCRIPTION, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.CBND , " & UId & " From VOUCHERS Where   " & stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.text) + "' and  Subledger like '" + Trim(Alpha.text) + "%'  AND   VOUCHERS.DebitorCredit='C' and VoucherDate >= cdate('" + Trim(date1.text) + "')   and VoucherDate <=cdate('" + Trim(date2.text) + "')    ORDER BY VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber,vsno"
      ' vouchers debtors
      con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid )  SELECT VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.DESCRIPTION, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.CBND , " & UId & " From VOUCHERS Where   " & stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.text) + "' and  Subledger like '" + Trim(Alpha.text) + "%' AND  VOUCHERS.DebitorCredit='D' and  VoucherDate >= cdate('" + Trim(date1.text) + "')   and VoucherDate <=cdate('" + Trim(date2.text) + "')      ORDER BY VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber,vsno"
      ' invoice
      con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid )  SELECT INVOICEA.GENLEDGER, INVOICEA.SUBLEDGER, INVOICEA.INVOICEDATE, 'I' AS Expr1, INVOICEA.INVOICENO, 'Sales Invoice' , INVOICEA.NETAMOUNT, 'D' , '' AS Expr3, " & UId & " FROM INVOICEA  where   " & stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.text) + "' and  Subledger like '" + Trim(Alpha.text) + "%'  and InvoiceDate >= cdate('" + Trim(date1.text) + "')   and InvoiceDate <=cdate('" + Trim(date2.text) + "') "
      ' cash credit
      con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid )  SELECT CASHA.GENLEDGER, CASHA.SUBLEDGER, CASHA.INVOICEDATE, 'S' AS Expr1, CASHA.INVOICENO, 'Sales C/M' , CASHA.NETAMOUNT, 'D' , '' AS Expr3, " & UId & "  FROM CASHA  where  " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.text) + "' and  Subledger like '" + Trim(Alpha.text) + "%'   and InvoiceDate >= cdate('" + Trim(date1.text) + "')   and InvoiceDate <=cdate('" + Trim(date2.text) + "')  "
      ' cash debit
      con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid )   SELECT CASHA.GENLEDGER, CASHA.SUBLEDGER, CASHA.INVOICEDATE, 'S' AS Expr1, CASHA.INVOICENO, 'Sales C/M' , CASHA.BAA, 'C' , '' AS Expr3, " & UId & "  FROM CASHA  where " & stringyear & " and CASHA.BAA<>0 and  genledger='" + Trim(COMBOGENLEDGER.text) + "' and  Subledger like '" + Trim(Alpha.text) + "%'   and InvoiceDate >= cdate('" + Trim(date1.text) + "')   and InvoiceDate <=cdate('" + Trim(date2.text) + "') "
       ' credit a
      con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ) SELECT CREDITA.GENLEDGER, CREDITA.SUBLEDGER, CREDITA.INVOICEDATE, 'C' AS Expr1, CREDITA.INVOICENO, 'Credit Note' , CREDITA.NETAMOUNT, 'C' , '' AS Expr3, " & UId & " FROM CREDITA " & stringyear & " and  where   genledger='" + Trim(COMBOGENLEDGER.text) + "' and  Subledger like '" + Trim(Alpha.text) + "%' and InvoiceDate >= cdate('" + Trim(date1.text) + "')   and InvoiceDate <=cdate('" + Trim(date2.text) + "')   "
       ' dnfadr
      con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid )  SELECT DNFA.PGLD, DNFA.PSLD, DNFA.DND, 'D' , DNFA.DNN, DNFA.N, DNFA.Na, DNFA.DC, '', " & UId & "   From DNFA  where " & stringyear & " and   Pgld ='" + Trim(COMBOGENLEDGER.text) + "' and  Psld like '" + Trim(Alpha.text) + "%'  and dnd >= cdate('" + Trim(date1.text) + "')   and dnd <=cdate('" + Trim(date2.text) + "')  ORDER BY DNFA.PGLD, DNFA.PSLD, DNFA.DND, DNFA.DNN "
       'cnf1cr
      con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid )  SELECT CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND, 'C' , CNF1A.CNN, CNF1A.N, CNF1A.NA, CNF1A.DC, '', " & UId & "  From CNF1A where " & stringyear & " and Pgld='" + Trim(COMBOGENLEDGER.text) + "' and  Psld like '" + Trim(Alpha.text) + "%'  and cnd >= cdate('" + Trim(date1.text) + "')   and cnd <=cdate('" + Trim(date2.text) + "')   ORDER BY CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND,CNF1A.CNN"
 
      con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid )  SELECT DNFB.GLD, DNFB.SLD, DNFB.DND, 'D' , DNFB.DNN, 'DNB' , DNFB.A, DNFB.DC, '' , " & UId & " From DNFB  where " & stringyear & " and gld='" + Trim(COMBOGENLEDGER.text) + "' and  sld like '" + Trim(Alpha.text) + "%'   and dnd >= cdate('" + Trim(date1.text) + "')   and dnd <=cdate('" + Trim(date2.text) + "')    ORDER BY DNFB.GLD, DNFB.SLD, DNFB.DND, DNFB.DNN"

      con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid )   SELECT CNF1B.GLD, CNF1B.SLD, CNF1B.CND, 'C' , CNF1B.CNN, 'CNB' , CNF1B.A, CNF1B.DC, '', " & UId & " From CNF1B where " & stringyear & " and gld='" + Trim(COMBOGENLEDGER.text) + "' and  sld like '" + Trim(Alpha.text) + "%'   and cnd >= cdate('" + Trim(date1.text) + "')   and cnd <=cdate('" + Trim(date2.text) + "')   ORDER BY    CNF1B.GLD, CNF1B.SLD, CNF1B.CND, CNF1B.CNN"
   End If
   
   
    If Trim(Alpha.text) = "" And Alpha.Visible = True Then
                
                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ) SELECT VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.DESCRIPTION, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.CBND , " & UId & " From VOUCHERS Where  " & stringyear & " and  genledger ='" + Trim(COMBOGENLEDGER.text) + "'  AND   VOUCHERS.DebitorCredit='C' and  VoucherDate >= cdate('" + Trim(date1.text) + "')   and VoucherDate <=cdate('" + Trim(date2.text) + "')   ORDER BY VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber,vsno"
                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid )  SELECT VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.DESCRIPTION, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.CBND , " & UId & " From VOUCHERS Where  " & stringyear & " and  genledger ='" + Trim(COMBOGENLEDGER.text) + "' AND  VOUCHERS.DebitorCredit='D' and genledger='" + Trim(COMBOGENLEDGER.text) + "' and VoucherDate >= cdate('" + Trim(date1.text) + "')   and VoucherDate <=cdate('" + Trim(date2.text) + "')    ORDER BY VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber,vsno"
                
                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno ,userid) SELECT INVOICEA.GENLEDGER, INVOICEA.SUBLEDGER, INVOICEA.INVOICEDATE, 'I' AS Expr1, INVOICEA.INVOICENO, 'Sales Invoice' , INVOICEA.NETAMOUNT, 'D' , '' AS Expr3  , " & UId & " FROM INVOICEA  where  " & stringyear & " and  genledger ='" + Trim(COMBOGENLEDGER.text) + "'  and InvoiceDate >= cdate('" + Trim(date1.text) + "')   and InvoiceDate <=cdate('" + Trim(date2.text) + "') "
    
                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno ,userid) SELECT CASHA.GENLEDGER, CASHA.SUBLEDGER, CASHA.INVOICEDATE, 'S' AS Expr1, CASHA.INVOICENO, 'Sales C/M' , CASHA.NETAMOUNT, 'D' , '' AS Expr3 , " & UId & " FROM CASHA   where  " & stringyear & " and  genledger='" + Trim(COMBOGENLEDGER.text) + "'    and InvoiceDate >= cdate('" + Trim(date1.text) + "')   and InvoiceDate <=cdate('" + Trim(date2.text) + "')   "
                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid )  SELECT CASHA.GENLEDGER, CASHA.SUBLEDGER, CASHA.INVOICEDATE, 'S' AS Expr1, CASHA.INVOICENO, 'Sales C/M' , CASHA.BAA, 'C' , '' AS Expr3, " & UId & "  FROM CASHA  where " & stringyear & " and  genledger='" + Trim(COMBOGENLEDGER.text) + "'   and InvoiceDate >= cdate('" + Trim(date1.text) + "')   and InvoiceDate <=cdate('" + Trim(date2.text) + "')  AND CASHA.BAA <>0  "
                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ) SELECT CREDITA.GENLEDGER, CREDITA.SUBLEDGER, CREDITA.INVOICEDATE, 'C' AS Expr1, CREDITA.INVOICENO, 'Credit Note' , CREDITA.NETAMOUNT, 'C' , '' AS Expr3, " & UId & "  FROM CREDITA  where " & stringyear & " and   genledger='" + Trim(COMBOGENLEDGER.text) + "'    and InvoiceDate >= cdate('" + Trim(date1.text) + "') and InvoiceDate <=cdate('" + Trim(date2.text) + "') "

                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ) SELECT DNFA.PGLD, DNFA.PSLD, DNFA.DND, 'D' , DNFA.DNN, DNFA.N, DNFA.Na, DNFA.DC, ''  , " & UId & " From DNFA  where  " & stringyear & " and  Pgld ='" + Trim(COMBOGENLEDGER.text) + "'and dnd >= cdate('" + Trim(date1.text) + "')   and dnd <=cdate('" + Trim(date2.text) + "')  ORDER BY DNFA.PGLD, DNFA.PSLD, DNFA.DND, DNFA.DNN "
 
                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ) SELECT CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND, 'C' , CNF1A.CNN, CNF1A.N, CNF1A.NA, CNF1A.DC, '' , " & UId & " From CNF1A where " & stringyear & " and Pgld='" + Trim(COMBOGENLEDGER.text) + "'  and cnd >= cdate('" + Trim(date1.text) + "')   and cnd <=cdate('" + Trim(date2.text) + "')  ORDER BY CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND,CNF1A.CNN"
 
                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ) SELECT DNFB.GLD, DNFB.SLD, DNFB.DND, 'D' , DNFB.DNN, 'DNB' , DNFB.A, DNFB.DC, '', " & UId & "  From DNFB  where " & stringyear & " and gld='" + Trim(COMBOGENLEDGER.text) + "'   and dnd >= cdate('" + Trim(date1.text) + "')   and dnd <=cdate('" + Trim(date2.text) + "')    ORDER BY DNFB.GLD, DNFB.SLD, DNFB.DND, DNFB.DNN"
                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ) SELECT CNF1B.GLD, CNF1B.SLD, CNF1B.CND, 'C' , CNF1B.CNN, 'CNB' , CNF1B.A, CNF1B.DC, '' , " & UId & " From CNF1B where " & stringyear & " and gld='" + Trim(COMBOGENLEDGER.text) + "' and  cnd >= cdate('" + Trim(date1.text) + "')   and cnd <=cdate('" + Trim(date2.text) + "') ORDER BY    CNF1B.GLD, CNF1B.SLD, CNF1B.CND, CNF1B.CNN"
End If
If Trim(Alpha.text) = "" And Alpha.Visible = False And Combosubledger.text <> "" Then
                
                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid)  SELECT VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.DESCRIPTION, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.CBND, " & UId & ",fyear,setupid From VOUCHERS Where   " & stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.text) + "' and  Subledger = '" & Combosubledger.text & "'  AND   VOUCHERS.DebitorCredit='C' and  convert(smalldatetime,VoucherDate,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,VoucherDate,103) <=convert(smalldatetime,'" + Trim(date2.text) + "',103)   ORDER BY VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber,vsno"
                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid)   SELECT GenLedger, SubLedger, VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.DESCRIPTION, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.CBND, " & UId & ",fyear,setupid From VOUCHERS Where  " & stringyear & " and  genledger ='" + Trim(COMBOGENLEDGER.text) + "' and  Subledger = '" & Combosubledger.text & "' AND  VOUCHERS.DebitorCredit='D' and genledger='" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,VoucherDate,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,VoucherDate,103) <=convert(smalldatetime,'" + Trim(date2.text) + "',103)    ORDER BY VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber,vsno"
                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid)  SELECT INVOICEA.GENLEDGER, INVOICEA.SUBLEDGER, INVOICEA.INVOICEDATE, 'I' AS Expr1, INVOICEA.INVOICENO, 'Sales Invoice' , INVOICEA.NETAMOUNT, 'D' , '' AS Expr3 , " & UId & ",fyear,setupid FROM INVOICEA  where  " & stringyear & " and  genledger ='" + Trim(COMBOGENLEDGER.text) + "' and  Subledger = '" & Combosubledger.text & "'  and convert(smalldatetime,invoiceDate,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,invoiceDate,103) <=convert(smalldatetime,'" + Trim(date2.text) + "',103) "
                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid)  SELECT CASHA.GENLEDGER, CASHA.SUBLEDGER, CASHA.INVOICEDATE, 'S' AS Expr1, CASHA.INVOICENO, 'Sales C/M' , CASHA.NETAMOUNT, 'D' , '' AS Expr3, " & UId & ",fyear,setupid  FROM CASHA  where  " & stringyear & " and  genledger='" + Trim(COMBOGENLEDGER.text) + "' and  Subledger = '" & Combosubledger.text & "' and convert(smalldatetime,invoiceDate,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,invoiceDate,103) <=convert(smalldatetime,'" + Trim(date2.text) + "',103)  "
                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid)  SELECT CASHA.GENLEDGER, CASHA.SUBLEDGER, CASHA.INVOICEDATE, 'S' AS Expr1, CASHA.INVOICENO, 'Sales C/M' , CASHA.BAA, 'C' , '' AS Expr3, " & UId & ",fyear,setupid  FROM CASHA  where " & stringyear & " and  genledger='" + Trim(COMBOGENLEDGER.text) + "' and  Subledger = '" & Combosubledger.text & "'   and convert(smalldatetime,invoiceDate,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,invoiceDate,103) <=convert(smalldatetime,'" + Trim(date2.text) + "',103)  AND CASHA.BAA <>0  "
                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid)    SELECT CREDITA.GENLEDGER, CREDITA.SUBLEDGER, CREDITA.INVOICEDATE, 'C' AS Expr1, CREDITA.INVOICENO, 'Credit Note' , CREDITA.NETAMOUNT, 'C' , '' AS Expr3 , " & UId & ",fyear,setupid FROM CREDITA  where  " & stringyear & " and  genledger='" + Trim(COMBOGENLEDGER.text) + "' and  Subledger = '" & Combosubledger.text & "' and convert(smalldatetime,invoiceDate,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,invoicedate,103) <=convert(smalldatetime,'" + Trim(date2.text) + "',103) "
                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid)   SELECT DNFA.PGLD, DNFA.PSLD, DNFA.DND, 'D' , DNFA.DNN, DNFA.N, DNFA.Na, DNFA.DC, '' , " & UId & ",fyear,setupid   From DNFA  where  " & stringyear & " and  Pgld ='" + Trim(COMBOGENLEDGER.text) + "' and  Psld = '" & Combosubledger.text & "' and convert(smalldatetime,dnd,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,dnd,103) <=convert(smalldatetime,'" + Trim(date2.text) + "',103)  ORDER BY DNFA.PGLD, DNFA.PSLD, DNFA.DND, DNFA.DNN "
                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno, userid,fyear,setupid)   SELECT CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND, 'C' , CNF1A.CNN, CNF1A.N, CNF1A.NA, CNF1A.DC, '' , " & UId & ",fyear,setupid From CNF1A where " & stringyear & " and Pgld='" + Trim(COMBOGENLEDGER.text) + "' and  Psld = '" & Combosubledger.text & "' and  convert(smalldatetime,cnd,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,cnd,103) <=convert(smalldatetime,'" + Trim(date2.text) + "',103)  ORDER BY CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND,CNF1A.CNN"
                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno, userid,fyear,setupid)   SELECT DNFB.GLD, DNFB.SLD, DNFB.DND, 'D' , DNFB.DNN, 'DNB' , DNFB.A, DNFB.DC, ''     , " & UId & ",fyear,setupid From DNFB  where " & stringyear & " and gld='" + Trim(COMBOGENLEDGER.text) + "' and  sld = '" & Combosubledger.text & "'   and convert(smalldatetime,dnd,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,dnd,103) <=convert(smalldatetime,'" + Trim(date2.text) + "',103)   ORDER BY DNFB.GLD, DNFB.SLD, DNFB.DND, DNFB.DNN"
                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno ,userid,fyear,setupid)   SELECT CNF1B.GLD, CNF1B.SLD, CNF1B.CND, 'C' , CNF1B.CNN, 'CNB' , CNF1B.A, CNF1B.DC, '' , " & UId & ",fyear,setupid From CNF1B where " & stringyear & " and gld='" + Trim(COMBOGENLEDGER.text) + "' and  sld = '" & Combosubledger.text & "' and convert(smalldatetime,cnd,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,cnd,103) <=convert(smalldatetime,'" + Trim(date2.text) + "',103) ORDER BY    CNF1B.GLD, CNF1B.SLD, CNF1B.CND, CNF1B.CNN"
End If
con.Execute "insert into Treport ( Genledger,Subledger,openingbalance,userid,fyear,setupid) SELECT '" + Trim(COMBOGENLEDGER.text) + "'as genled, SUBLEDGER, (sum(YEAROPENING) + SUM (OPAMOUNTDEBIT) - SUM(OPAMOUNTCREDIT)) AS OPCR ,  " & UId & " as UserId,fyear,setupid from subledgertrail where " & stringyear & " GROUP BY SUBLEDGER,fyear,setupid;"


'====================================================================================
Dim dr, cr As Double
Dim rs_inv As New ADODB.Recordset


dr = 0
cr = 0

vs.Clear
vs.rows = 2
J = 1

lblCr.Visible = False
lblDr.Visible = False

lblCr.Caption = 0
lblDr.Caption = 0
lblClosingBal.Caption = 0

lblDrCR.Caption = ""
lbl_crdr.Caption = ""


If rs_inv.State = 1 Then rs_inv.Close
rs_inv.Open "select AdviceStatus,invoiceno from invoicea where " & stringyear & "", con, adOpenKeyset, adLockReadOnly


If RS.State = 1 Then RS.Close
RS.Open "select vno,vdate,vtype,narration,ad,dorc from treport where " & stringyear & "  order by vdate,sno", con, adOpenDynamic, adLockOptimistic
While RS.EOF = False

If Not IsNull(RS!vdate) Then

vs.TextMatrix(J, 0) = RS!vno

If RS!vtype = "I" Then
   rs_inv.MoveFirst
   If rs_inv.EOF = False Then
      rs_inv.Find "invoiceno=" & RS!vno & ""
      If rs_inv.EOF = False Then
      
      If RS!vno = 1373 Then
       
     '  MsgBox "a"
      
      End If
      
      If LCase(rs_inv(0)) = "pending" Then
         vs.TextMatrix(J, 3) = RS!narration
         vs.Cell(flexcpBackColor, J, 3) = vbGreen
      Else
         vs.TextMatrix(J, 3) = RS!narration
         vs.Cell(flexcpBackColor, J, 3) = &HF7F9AC
      End If
      
      End If
   End If
   
End If

vs.TextMatrix(J, 1) = RS!vdate


If RS!vtype = "S" Then
   vs.TextMatrix(J, 2) = "C/M"
   vs.TextMatrix(J, 3) = "Cash Memo"
ElseIf RS!vtype = "I" Then
   vs.TextMatrix(J, 2) = "I"
   vs.TextMatrix(J, 3) = "Invoice Sales"
ElseIf RS!vtype = "C" Then
   vs.TextMatrix(J, 2) = "CI"
   vs.TextMatrix(J, 3) = "Credit Note Item"
Else
  vs.TextMatrix(J, 2) = RS!vtype
  vs.TextMatrix(J, 3) = RS!narration
End If


If RS!dorc = "D" Then
vs.TextMatrix(J, 4) = RS!ad
dr = dr + RS!ad
Else
vs.TextMatrix(J, 5) = RS!ad
cr = cr + RS!ad
End If

vs.rows = vs.rows + 1
J = J + 1

End If


RS.MoveNext
Wend


vs.Cols = 7

vs.FormatString = "<VNo|VDate|VType|Narration|>Dr|>Cr|Balance|Dr/Cr"


vs.ColWidth(0) = 1000
vs.ColWidth(1) = 1200
vs.ColWidth(2) = 800
vs.ColWidth(3) = 4000
vs.ColWidth(4) = 1000
vs.ColWidth(5) = 1000
vs.ColWidth(6) = 900
vs.ColWidth(7) = 500


vs.Cell(flexcpFontBold, 0, 0) = True
vs.Cell(flexcpFontBold, 0, 1) = True
vs.Cell(flexcpFontBold, 0, 2) = True
vs.Cell(flexcpFontBold, 0, 3) = True
vs.Cell(flexcpFontBold, 0, 4) = True
vs.Cell(flexcpFontBold, 0, 5) = True
vs.Cell(flexcpFontBold, 0, 6) = True
vs.Cell(flexcpFontBold, 0, 7) = True




If (dr > 0 Or cr > 0) Then

    If RS.State = 1 Then RS.Close
    RS.Open "select OpeningBalance from treport where " & stringyear & " and OpeningBalance<>0", con, adOpenDynamic, adLockOptimistic
    If RS.EOF = False Then
       If RS(0) < 0 Then
          lblCr.Caption = Abs(RS(0)) & " "
          
          lblOp.Caption = Format(Abs(RS(0)), "0.00")
          lbl_crdr.Caption = "Cr"
          
          lblCr.Visible = True
          lblDr.Visible = False
        Else
          
          lblOp.Caption = Format(Abs(RS(0)), "0.00")
          lbl_crdr.Caption = "Dr"
          
          lblDr.Caption = Abs(RS(0)) & " "
          lblCr.Visible = False
          lblDr.Visible = True
    
       End If
    End If

Else

    If RS.State = 1 Then RS.Close
    RS.Open "select OpeningBalance from treport where " & stringyear & " and OpeningBalance<>0", con, adOpenDynamic, adLockOptimistic
    If RS.EOF = False Then
       If RS(0) < 0 Then
          lblCr.Caption = Abs(RS(0)) & " "
          
          lblOp.Caption = Format(Abs(RS(0)), "0.00")
          lbl_crdr.Caption = "Cr"
          
          lblClosingBal.Caption = lblOp.Caption
          lblDrCR.Caption = "Cr"
          
          
          lblCr.Visible = True
          lblDr.Visible = False
        Else
          
          lblOp.Caption = Format(Abs(RS(0)), "0.00")
          lbl_crdr.Caption = "Dr"
          
          lblClosingBal.Caption = lblOp.Caption
          lblDrCR.Caption = "Dr"
          
          
          lblDr.Caption = Abs(RS(0)) & " "
          lblCr.Visible = False
          lblDr.Visible = True
    
       End If
       Exit Sub
    End If

End If





lblDrTotal.Caption = (Val(lblDr) + dr) & ""
lblCrTotal.Caption = (Val(lblCr) + cr) & ""


Dim bal As Double
dr = 0
cr = 0

'==============================================
For I = 1 To vs.rows - 2
   If I = 1 Then
      dr = Val(vs.TextMatrix(I, 4)) + Val(lblDr)
      cr = Val(vs.TextMatrix(I, 5)) + Val(lblCr)
      
      vs.TextMatrix(I, 6) = Round(Abs(dr - cr), 2)
      
      
      If (dr - cr) < 0 Then
         vs.TextMatrix(I, 7) = "Cr"
      Else
         vs.TextMatrix(I, 7) = "Dr"
      End If
      
      
      
      
      
   Else
      
      dr = Val(vs.TextMatrix(I, 4))
      cr = Val(vs.TextMatrix(I, 5))
      
      
      If vs.TextMatrix(I - 1, 7) = "Cr" Then
      bal = (dr - (cr + Val(vs.TextMatrix(I - 1, 6))))
      Else
      bal = ((dr + Val(vs.TextMatrix(I - 1, 6))) - cr)
      End If
      
      vs.TextMatrix(I, 6) = Round(Abs(bal), 2)
      
            
      If (bal) < 0 Then
         vs.TextMatrix(I, 7) = "Cr"
      Else
         vs.TextMatrix(I, 7) = "Dr"
      End If

      
      
   End If
   
   vs.Row = I + 1
Next
'----------------------------------------------

If vs.TextMatrix(vs.rows - 2, 7) <> "" Then
If vs.rows > 2 Then
    lblDrCR.Caption = vs.TextMatrix(vs.rows - 2, 7)
    lblClosingBal.Caption = Format(vs.TextMatrix(vs.rows - 2, 6), "0.00")
Else
    lblDrCR.Caption = ""
    lblClosingBal.Caption = ""

End If
End If


End Sub

Private Sub CommandReturn_Click()
    Unload Me
End Sub
Private Sub Commandshow_Click()

Commandshow.Enabled = False
'********sub for alpha wise and Partywise according to new fast mathed
 lblOp = "0"
 
 DoEvents
 con.Execute "Delete  from subledgertrail where " & stringyear & ""
 DoEvents
 DoEvents
 DoEvents
 DoEvents
 DoEvents
 DoEvents
 ALPHAB
 
 'CON.CommitTrans
 Commandshow.Enabled = True
 
End Sub
Private Sub date1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    date2.SetFocus
End If

End Sub

Private Sub date1_LostFocus()
    If Trim(date1.text) <> "" Then
        If Not checkdate(Trim(date1.text), date1) Then
            date1.SetFocus
            End If
    End If
End Sub

Private Sub date2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    sendkeys "{TAB}"
End If

End Sub

Private Sub date2_LostFocus()
    If Trim(date2.text) <> "" Then
        If Not checkdate(Trim(date2.text), date2) Then
            date2.SetFocus
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then
  If MsgBox("Want To Exit ..", vbQuestion + vbYesNo) = vbYes Then
     Unload Me
   End If
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 '   SendKeys "{TAB}"
End If



End Sub
Private Sub Form_Load()

Me.top = 100
Me.Left = 100

Me.Width = 11000
Me.Height = 10040
    
Me.Caption = "Subledger Ledger"


con.Execute "delete  from treport where " & stringyear & ""
con.Execute "Delete  from subledgertrail where " & stringyear & ""


On Error GoTo ac1

Do While Trim(UCase(VB.Screen.ActiveForm.Name)) <> (Trim(UCase("MainMenu") Or Trim(UCase("frmbilllist"))))
    If Trim(UCase(VB.Screen.ActiveForm.Name)) <> Trim(UCase("MainMenu")) Then
        Unload VB.Screen.ActiveForm
    End If
Loop

ac1:


Me.top = 0
Me.Left = 0

Set RS = New ADODB.Recordset
    RS.Open "select * from gledger where " & stringyear & " and slf=1", con, adOpenStatic, adLockReadOnly, adCmdText
    If Not RS.BOF Then
        Do While Not RS.EOF
            COMBOGENLEDGER.AddItem Trim(RS!gledger)
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.Close
    RS.Open "select * from setup1 where " & stringyear & "", con, adOpenStatic, adLockReadOnly
    CNSetup
    date1.text = RS!yarfrom
    date2.text = RS!yarto
    RS.Close
    
End Sub
Sub xx()

End Sub

Sub OPENINGSUBLEDGERS()

''''If Trim(Alpha.Text) <> "" Then
''''        'CON.Execute "Insert into subledgertrail  SELECT SUBLEDGER , YEAROPENING, 0 AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId FROM SLEDGER where  gledger='" + Trim(COMBOGENLEDGER.Text) + "' AND subledger like '" + Trim(alpha.Text) + "%'", p, adCmdText
''''        'subledger opening start
''''        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER ,YEAROPENING, 0 AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId " _
''''        & " FROM SLEDGER  where " & stringyear & " and  gledger='" + Trim(COMBOGENLEDGER.Text) + "'  AND SLEDGER.subledger like '" + Trim(Alpha.Text) + "%'"
''''    ' from invoice a
''''        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER , 0 as YEAROPENING,  sum(INVOICEA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId " _
''''        & " FROM (SLEDGER LEFT JOIN INVOICEA ON (SLEDGER.SUBLEDGER = INVOICEA.SUBLEDGER) )" _
''''        & " where  " & stringyear & " and gledger='" + Trim(COMBOGENLEDGER.Text) + "'and INVOICEDATE < cdate('" + Trim(date1.Text) + "') AND SLEDGER.subledger like '" + Trim(Alpha.Text) + "%' " _
''''        & " GROUP BY SLEDGER.SUBLEDGER "
''''   ' from casha
''''        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, Sum(CASHA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT,  " & UId & " as UserId " _
''''        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER) ) " _
''''        & " where " & stringyear & " and  gledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER <>'CASH PARTY'  and INVOICEDATE< cdate('" + Trim(date1.Text) + "') AND SLEDGER.subledger like '" + Trim(Alpha.Text) + "%'" _
''''        & " GROUP BY SLEDGER.SUBLEDGER " _
''''        & " HAVING  Sum(CASHA.NETAMOUNT) <> 0; "
''''
''''        con.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING , 0  AS OPAMOUNTDEBIT, Sum(CASHA.BAA) AS OPAMOUNTCREDIT ," & UId & " as UserId  " _
''''        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER))" _
''''        & " where " & stringyear & " and  gledger ='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER <>'CASH PARTY'   and INVOICEDATE < cdate('" + Trim(date1.Text) + "')  AND SLEDGER.subledger like '" + Trim(Alpha.Text) + "%' " _
''''        & " GROUP BY SLEDGER.SUBLEDGER " _
''''        & " HAVING  Sum(CASHA.baa) <> 0; "
''''
''''    ' from credita
''''        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER,    0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,SUM(CREDITA.NETAMOUNT)  AS OPAMOUNTCREDIT  , " & UId & " as UserId " _
''''        & " FROM (SLEDGER LEFT JOIN CREDITA ON (SLEDGER.SUBLEDGER = CREDITA.SUBLEDGER)) " _
''''        & " where " & stringyear & " and gledger='" + Trim(COMBOGENLEDGER.Text) + "' and INVOICEDATE < cdate('" + Trim(date1.Text) + "') AND SLEDGER.subledger like '" + Trim(Alpha.Text) + "%' " _
''''        & " GROUP BY SLEDGER.SUBLEDGER " _
''''        & " HAVING  Sum(CREDITA.NETAMOUNT) <> 0; "
''''
''''
''''   ' from vouchers
''''        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING,Sum(VOUCHERS.Amount) AS OPAMOUNTDEBIT ,   0 AS OPAMOUNTCREDIT ,  " & UId & " as UserId" _
''''        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) " _
''''        & " WHERE " & stringyear & " and DEBITORCREDIT = 'D' and gledger ='" + Trim(COMBOGENLEDGER.Text) + "' and VOUCHERDATE < cdate('" + Trim(date1.Text) + "')  AND SLEDGER.subledger like '" + Trim(Alpha.Text) + "%'" _
''''        & " GROUP BY SLEDGER.SUBLEDGER " _
''''        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0;"
''''
''''
''''        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, 0 AS OpAmountdebit,  Sum(VOUCHERS.Amount) AS OPAMOUNTCREDIT , " & UId & " as UserId " _
''''        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) " _
''''        & " WHERE " & stringyear & " and DEBITORCREDIT = 'C' and gledger ='" + Trim(COMBOGENLEDGER.Text) + "' and VOUCHERDATE < cdate('" + Trim(date1.Text) + "') AND SLEDGER.subledger like '" + Trim(Alpha.Text) + "%'" _
''''        & " GROUP BY SLEDGER.SUBLEDGER " _
''''        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0; "
''''         ''''ok
''''  'from cnf1a
''''        con.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, Sum(CNF1A.NA) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT  , " & UId & " as UserId " _
''''        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) " _
''''        & " WHERE " & stringyear & " and (((CNF1A.DC)='D')) and pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and cnd < cdate('" + Trim(date1.Text) + "') AND SLEDGER.subledger like '" + Trim(Alpha.Text) + "%'" _
''''        & " GROUP BY SLEDGER.SUBLEDGER " _
''''        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
''''
''''                con.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(CNF1A.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
''''        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) " _
''''        & " WHERE " & stringyear & " and (((CNF1A.DC)='C')) and pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and cnd < cdate('" + Trim(date1.Text) + "') AND SLEDGER.subledger like '" + Trim(Alpha.Text) + "%'" _
''''        & " GROUP BY SLEDGER.SUBLEDGER " _
''''        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
''''
''''
''''  ' from dnfa
''''        con.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFA.NA)  AS OPAMOUNTDEBIT,  0  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
''''        & " FROM SLEDGER LEFT JOIN DNFA ON (SLEDGER.gledger = DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD) " _
''''        & " WHERE " & stringyear & " and ((( DNFA.DC) = 'D' )) and  pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and dnd < cdate('" + Trim(date1.Text) + "') AND SLEDGER.subledger like '" + Trim(Alpha.Text) + "%'" _
''''        & " GROUP BY SLEDGER.SUBLEDGER " _
''''        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
''''
''''
''''        con.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(DNFA.NA)  AS OPAMOUNTCREDIT  , " & UId & " as UserId" _
''''        & " FROM SLEDGER LEFT JOIN DNFA  ON (SLEDGER.gledger =DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD)" _
''''        & " WHERE " & stringyear & " and (((DNFA.DC)='C')) and   pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and dnd < cdate('" + Trim(date1.Text) + "')AND SLEDGER.subledger like '" + Trim(Alpha.Text) + "%'" _
''''        & " GROUP BY SLEDGER.SUBLEDGER " _
''''        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
''''   ' from cnf1b
''''        con.Execute "Insert into subledgertrail     SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(CNF1B.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId " _
''''        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) " _
''''        & " WHERE " & stringyear & " and (((CNF1B.DC)='D')) and gld  = '" + Trim(COMBOGENLEDGER.Text) + "' and cnd < cdate('" + Trim(date1.Text) + "') AND SLEDGER.subledger like '" + Trim(Alpha.Text) + "%'" _
''''        & " GROUP BY SLEDGER.SUBLEDGER " _
''''        & " HAVING  Sum(CNF1B.A) <> 0; ", p, adCmdText
''''
''''
''''        con.Execute "Insert into subledgertrail    SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(CNF1B.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
''''        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) " _
''''        & " WHERE " & stringyear & " and (((CNF1B.DC)= 'C')) and gld  = '" + Trim(COMBOGENLEDGER.Text) + "' and cnd < cdate('" + Trim(date1.Text) + "') AND SLEDGER.SUBLEDGER like '" + Trim(Alpha.Text) + "%' " _
''''        & " GROUP BY SLEDGER.SUBLEDGER " _
''''        & " HAVING  Sum(CNF1B.A) <> 0;", p, adCmdText
''''
''''   ' dnfb
''''        con.Execute "Insert into subledgertrail     SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFB.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId " _
''''        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER = DNFB.SLD) " _
''''        & " WHERE " & stringyear & " and (((DNFB.DC)='D')) and   gld = '" + Trim(COMBOGENLEDGER.Text) + "'And dnd < cdate('" + Trim(date1.Text) + "')  AND SLEDGER.SUBLEDGER like '" + Trim(Alpha.Text) + "%'" _
''''        & " GROUP BY SLEDGER.SUBLEDGER " _
''''        & " HAVING  Sum(DNFB.A) <> 0; ", p, adCmdText
''''
''''        con.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(DNFB.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId  " _
''''        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER =DNFB.SLD)" _
''''        & " WHERE " & stringyear & " and (((DNFB.DC)= 'C')) and  gld = '" + Trim(COMBOGENLEDGER.Text) + "' and dnd < cdate('" + Trim(date1.Text) + "') AND SLEDGER.SUBLEDGER like '" + Trim(Alpha.Text) + "%'" _
''''        & " GROUP BY SLEDGER.SUBLEDGER " _
''''        & " HAVING  Sum(DNFB.A) <> 0 ", p, adCmdText
''''
''''
''''  End If
          
          


'''
'''If Trim(Alpha.Text) = "" And Combosubledger.Text = "" Then
'''
'''        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER , YEAROPENING,  0 AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT  , " & UId & " as UserId " _
'''        & " FROM SLEDGER where  " & stringyear & " and  gledger='" + Trim(COMBOGENLEDGER.Text) + "'", p, adCmdText
'''
'''        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER , 0 as YEAROPENING,  sum(INVOICEA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT  , " & UId & " as UserId " _
'''        & " FROM (SLEDGER LEFT JOIN INVOICEA ON (SLEDGER.SUBLEDGER = INVOICEA.SUBLEDGER))  " _
'''        & " where " & stringyear & " and  genledger='" + Trim(COMBOGENLEDGER.Text) + "'and INVOICEDATE < cdate('" + Trim(date1.Text) + "') " _
'''        & " GROUP BY SLEDGER.SUBLEDGER; ", p, adCmdText
'''
'''        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, Sum(CASHA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId " _
'''        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER)) " _
'''        & " where " & stringyear & " and gledger='" + Trim(COMBOGENLEDGER.Text) + "' and  SLEDGER.SUBLEDGER <>'CASH PARTY'    and INVOICEDATE< cdate('" + Trim(date1.Text) + "') " _
'''        & " GROUP BY SLEDGER.SUBLEDGER " _
'''        & " HAVING  Sum(CASHA.NETAMOUNT) <> 0; ", p, adCmdText
'''
'''        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING , 0  AS OPAMOUNTDEBIT, Sum(CASHA.BAA) AS OPAMOUNTCREDIT  , " & UId & " as UserId" _
'''        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER))" _
'''        & " where " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER <>'CASH PARTY' and INVOICEDATE < cdate('" + Trim(date1.Text) + "') " _
'''        & " GROUP BY SLEDGER.SUBLEDGER " _
'''        & " HAVING  Sum(CASHA.baa) <> 0; ", p, adCmdText
'''
'''
'''        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,    0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,SUM(CREDITA.NETAMOUNT)  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
'''        & " FROM (SLEDGER LEFT JOIN CREDITA ON (SLEDGER.SUBLEDGER = CREDITA.SUBLEDGER)) " _
'''        & " where " & stringyear & " and  genledger='" + Trim(COMBOGENLEDGER.Text) + "' and INVOICEDATE < cdate('" + Trim(date1.Text) + "') " _
'''        & " GROUP BY SLEDGER.SUBLEDGER " _
'''        & " HAVING  Sum(CREDITA.NETAMOUNT) <> 0; ", p, adCmdText
'''
'''
'''
'''        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING,Sum(VOUCHERS.Amount) AS OPAMOUNTDEBIT ,   0 AS OPAMOUNTCREDIT  , " & UId & " as UserId" _
'''        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) " _
'''        & " WHERE " & stringyear & " and DEBITORCREDIT = 'D' and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and VOUCHERDATE < cdate('" + Trim(date1.Text) + "') " _
'''        & " GROUP BY SLEDGER.SUBLEDGER " _
'''        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0;", p, adCmdText
'''
'''
'''        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, 0 AS OpAmountdebit,  Sum(VOUCHERS.Amount) AS OPAMOUNTCREDIT , " & UId & " as UserId " _
'''        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) " _
'''        & " WHERE " & stringyear & " and DEBITORCREDIT = 'C' and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and VOUCHERDATE < cdate('" + Trim(date1.Text) + "') " _
'''        & " GROUP BY SLEDGER.SUBLEDGER " _
'''        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0; ", p, adCmdText
'''         ''''ok
'''
'''        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, Sum(CNF1A.NA) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT  , " & UId & " as UserId" _
'''        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) " _
'''        & " WHERE " & stringyear & " and (((CNF1A.DC)='D')) and pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and cnd < cdate('" + Trim(date1.Text) + "') " _
'''        & " GROUP BY SLEDGER.SUBLEDGER " _
'''        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
'''
'''        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(CNF1A.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
'''        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) " _
'''        & " WHERE " & stringyear & " and (((CNF1A.DC)='C')) and pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and cnd < cdate('" + Trim(date1.Text) + "') " _
'''        & " GROUP BY SLEDGER.SUBLEDGER " _
'''        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
'''
'''
'''
'''        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFA.NA)  AS OPAMOUNTDEBIT,  0  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
'''        & " FROM SLEDGER LEFT JOIN DNFA ON (SLEDGER.gledger = DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD) " _
'''        & " WHERE " & stringyear & " and ((( DNFA.DC) = 'D' )) and  pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and dnd < cdate('" + Trim(date1.Text) + "') " _
'''        & " GROUP BY SLEDGER.SUBLEDGER " _
'''        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
'''
'''
'''        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(DNFA.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
'''        & " FROM SLEDGER LEFT JOIN DNFA  ON (SLEDGER.gledger =DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD)" _
'''        & " WHERE " & stringyear & " and (((DNFA.DC)='C')) and   pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and dnd < cdate('" + Trim(date1.Text) + "') " _
'''        & " GROUP BY SLEDGER.SUBLEDGER " _
'''        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
'''
'''        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(CNF1B.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId " _
'''        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) " _
'''        & " WHERE " & stringyear & " and (((CNF1B.DC)='D')) and gld  = '" + Trim(COMBOGENLEDGER.Text) + "' and cnd < cdate('" + Trim(date1.Text) + "')  " _
'''        & " GROUP BY SLEDGER.SUBLEDGER " _
'''        & " HAVING  Sum(CNF1B.A) <> 0; ", p, adCmdText
'''
'''
'''        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(CNF1B.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
'''        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) " _
'''        & " WHERE " & stringyear & " and (((CNF1B.DC)= 'C')) and gld  = '" + Trim(COMBOGENLEDGER.Text) + "' and cnd < cdate('" + Trim(date1.Text) + "')  " _
'''        & " GROUP BY SLEDGER.SUBLEDGER " _
'''        & " HAVING  Sum(CNF1B.A) <> 0;", p, adCmdText
'''
'''
'''        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFB.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId " _
'''        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER = DNFB.SLD) " _
'''        & " WHERE " & stringyear & " and (((DNFB.DC)='D')) and   gld = '" + Trim(COMBOGENLEDGER.Text) + "'And dnd < cdate('" + Trim(date1.Text) + "') " _
'''        & " GROUP BY SLEDGER.SUBLEDGER " _
'''        & " HAVING  Sum(DNFB.A) <> 0; ", p, adCmdText
'''
'''        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(DNFB.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
'''        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER =DNFB.SLD)" _
'''        & " WHERE " & stringyear & " and (((DNFB.DC)= 'C')) and  gld = '" + Trim(COMBOGENLEDGER.Text) + "' and dnd < cdate('" + Trim(date1.Text) + "') " _
'''        & " GROUP BY SLEDGER.SUBLEDGER " _
'''        & " HAVING  Sum(DNFB.A) <> 0 ", p, adCmdText
'''  End If
'''
  
If Trim(Alpha.text) = "" And Combosubledger.text <> "" Then

        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER , SLEDGER.YEAROPENING, 0 AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT  , " & UId & " as UserId," & setupid & ",'" & session & "'  FROM SLEDGER  where  " & stringyear & " and gledger='" + Trim(COMBOGENLEDGER.text) + "' and SLEDGER.SUBLEDGER = '" & Combosubledger.text & "'", p, adCmdText
        
        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER , 0 as YEAROPENING,  sum(INVOICEA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & "," & session & " " _
        & " FROM (SLEDGER LEFT JOIN INVOICEA ON (SLEDGER.SUBLEDGER = INVOICEA.SUBLEDGER) )" _
        & " where  INVOICEA.fyear='" & session & "' and  INVOICEA.setupid='" & setupid & "'  and gledger='" + Trim(COMBOGENLEDGER.text) + "' and SLEDGER.SUBLEDGER=  '" & Combosubledger.text & "' and convert(smalldatetime,INVOICEDATE,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.text) + "%' " _
        & " GROUP BY SLEDGER.SUBLEDGER "
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, Sum(CASHA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId ," & setupid & "," & session & " " _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER)) " _
        & " where cashA.fyear='" & session & "' and  cashA.setupid='" & setupid & "' and  gledger='" + Trim(COMBOGENLEDGER.text) + "' and SLEDGER.SUBLEDGER=  '" & Combosubledger.text & "'  and convert(smalldatetime,INVOICEDATE,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.NETAMOUNT) <> 0; ", p, adCmdText
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING , 0  AS OPAMOUNTDEBIT, Sum(CASHA.BAA) AS OPAMOUNTCREDIT  , " & UId & " as UserId," & setupid & "," & session & "" _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER))" _
        & " where cashA.fyear='" & session & "' and  cashA.setupid='" & setupid & "' and  genledger='" + Trim(COMBOGENLEDGER.text) + "' and SLEDGER.SUBLEDGER ='" & Combosubledger.text & "' and convert(smalldatetime,INVOICEDATE,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.baa) <> 0; ", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,    0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,SUM(CREDITA.NETAMOUNT)  AS OPAMOUNTCREDIT , " & UId & " as UserId ," & setupid & "," & session & " " _
        & " FROM (SLEDGER LEFT JOIN CREDITA ON (SLEDGER.SUBLEDGER = CREDITA.SUBLEDGER)) " _
        & " where CREDITA.fyear='" & session & "' and  CREDITA.setupid='" & setupid & "' and genledger='" + Trim(COMBOGENLEDGER.text) + "' and SLEDGER.SUBLEDGER='" & Combosubledger.text & "' and convert(smalldatetime,INVOICEDATE,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CREDITA.NETAMOUNT) <> 0; ", p, adCmdText
        
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING,Sum(VOUCHERS.Amount) AS OPAMOUNTDEBIT ,   0 AS OPAMOUNTCREDIT  , " & UId & " as UserId," & setupid & "," & session & " " _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) " _
        & " WHERE VOUCHERS.fyear='" & session & "' and  VOUCHERS.setupid='" & setupid & "' and DEBITORCREDIT = 'D' and genledger='" + Trim(COMBOGENLEDGER.text) + "' and SLEDGER.SUBLEDGER= '" & Combosubledger.text & "' and convert(smalldatetime,voucherDATE,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0;", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, 0 AS OpAmountdebit,  Sum(VOUCHERS.Amount) AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & "," & session & "  " _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) " _
        & " WHERE  VOUCHERS.fyear='" & session & "' and  VOUCHERS.setupid='" & setupid & "' and DEBITORCREDIT = 'C' and genledger='" + Trim(COMBOGENLEDGER.text) + "' and SLEDGER.SUBLEDGER= '" & Combosubledger.text & "' and convert(smalldatetime,voucherdate,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0; ", p, adCmdText
         ''''ok
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, Sum(CNF1A.NA) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT  , " & UId & " as UserId," & setupid & "," & session & " " _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) " _
        & " WHERE CNF1A.fyear='" & session & "' and  CNF1A.setupid='" & setupid & "' and (((CNF1A.DC)='D')) and pgld = '" + Trim(COMBOGENLEDGER.text) + "'and CNF1A.PSLD = '" & Combosubledger.text & "'  and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(CNF1A.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId ," & setupid & "," & session & "" _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) " _
        & " WHERE CNF1A.fyear='" & session & "' and  CNF1A.setupid='" & setupid & "' and (((CNF1A.DC)='C')) and pgld = '" + Trim(COMBOGENLEDGER.text) + "' and CNF1A.PSLD = '" & Combosubledger.text & "'  and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
        
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFA.NA)  AS OPAMOUNTDEBIT,  0  AS OPAMOUNTCREDIT , " & UId & " as UserId ," & setupid & "," & session & "" _
        & " FROM SLEDGER LEFT JOIN DNFA ON (SLEDGER.gledger = DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD) " _
        & " WHERE DNFA.fyear='" & session & "' and  DNFA.setupid='" & setupid & "' and ((( DNFA.DC) = 'D' )) and  pgld = '" + Trim(COMBOGENLEDGER.text) + "' and   DNFA.PSLD = '" & Combosubledger.text & "' and convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(DNFA.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId ," & setupid & "," & session & "" _
        & " FROM SLEDGER LEFT JOIN DNFA  ON (SLEDGER.gledger =DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD)" _
        & " WHERE DNFA.fyear='" & session & "' and  DNFA.setupid='" & setupid & "' and (((DNFA.DC)='C')) and   pgld = '" + Trim(COMBOGENLEDGER.text) + "'  and  DNFA.PSLD = '" & Combosubledger.text & "'   and convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(CNF1B.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId ," & setupid & "," & session & "" _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) " _
        & " WHERE CNF1B.fyear='" & session & "' and  CNF1B.setupid='" & setupid & "' and (((CNF1B.DC)='D')) and gld  = '" + Trim(COMBOGENLEDGER.text) + "' and  CNF1B.SLD = '" & Combosubledger.text & "'    and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0; ", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(CNF1B.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId ," & setupid & "," & session & "" _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) " _
        & " WHERE CNF1B.fyear='" & session & "' and  CNF1B.setupid='" & setupid & "' and (((CNF1B.DC)= 'C')) and gld  = '" + Trim(COMBOGENLEDGER.text) + "'  and  CNF1B.SLD = '" & Combosubledger.text & "' and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103)" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0;", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFB.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId ," & setupid & "," & session & "" _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER = DNFB.SLD) " _
        & " WHERE DNFB.fyear='" & session & "' and  DNFB.setupid='" & setupid & "' and (((DNFB.DC)='D')) and   gld = '" + Trim(COMBOGENLEDGER.text) + "'  and  DNFB.SLD = '" & Combosubledger.text & "'  And convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103)" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0; ", p, adCmdText
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(DNFB.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & "," & session & " " _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER =DNFB.SLD)" _
        & " WHERE DNFB.fyear='" & session & "' and  DNFB.setupid='" & setupid & "' and (((DNFB.DC)= 'C')) and  gld = '" + Trim(COMBOGENLEDGER.text) + "'  and  DNFB.SLD = '" & Combosubledger.text & "' and convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103)" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0 ", p, adCmdText
  End If
  
  
End Sub
Sub multi_Mail(party As String)

'======Opening Data

        con.Execute "Delete  from subledgertrail where " & stringyear & " and subledger='" & party & "'"
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER , SLEDGER.YEAROPENING, 0 AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT  , " & UId & " as UserId," & setupid & ",'" & session & "'  FROM SLEDGER  where  " & stringyear & " and gledger='" + Trim(COMBOGENLEDGER.text) + "' and SLEDGER.SUBLEDGER = '" & party & "'", p, adCmdText
        
        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER , 0 as YEAROPENING,  sum(INVOICEA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & "," & session & " " _
        & " FROM (SLEDGER LEFT JOIN INVOICEA ON (SLEDGER.SUBLEDGER = INVOICEA.SUBLEDGER) )" _
        & " where  INVOICEA.fyear='" & session & "' and  INVOICEA.setupid='" & setupid & "'  and gledger='" + Trim(COMBOGENLEDGER.text) + "' and SLEDGER.SUBLEDGER=  '" & party & "' and convert(smalldatetime,INVOICEDATE,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.text) + "%' " _
        & " GROUP BY SLEDGER.SUBLEDGER "
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, Sum(CASHA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId ," & setupid & "," & session & " " _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER)) " _
        & " where cashA.fyear='" & session & "' and  cashA.setupid='" & setupid & "' and  gledger='" + Trim(COMBOGENLEDGER.text) + "' and SLEDGER.SUBLEDGER=  '" & party & "'  and convert(smalldatetime,INVOICEDATE,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.NETAMOUNT) <> 0; ", p, adCmdText
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING , 0  AS OPAMOUNTDEBIT, Sum(CASHA.BAA) AS OPAMOUNTCREDIT  , " & UId & " as UserId," & setupid & "," & session & "" _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER))" _
        & " where cashA.fyear='" & session & "' and  cashA.setupid='" & setupid & "' and  genledger='" + Trim(COMBOGENLEDGER.text) + "' and SLEDGER.SUBLEDGER ='" & party & "' and convert(smalldatetime,INVOICEDATE,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.baa) <> 0; ", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,    0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,SUM(CREDITA.NETAMOUNT)  AS OPAMOUNTCREDIT , " & UId & " as UserId ," & setupid & "," & session & " " _
        & " FROM (SLEDGER LEFT JOIN CREDITA ON (SLEDGER.SUBLEDGER = CREDITA.SUBLEDGER)) " _
        & " where CREDITA.fyear='" & session & "' and  CREDITA.setupid='" & setupid & "' and genledger='" + Trim(COMBOGENLEDGER.text) + "' and SLEDGER.SUBLEDGER='" & party & "' and convert(smalldatetime,INVOICEDATE,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CREDITA.NETAMOUNT) <> 0; ", p, adCmdText
        
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING,Sum(VOUCHERS.Amount) AS OPAMOUNTDEBIT ,   0 AS OPAMOUNTCREDIT  , " & UId & " as UserId," & setupid & "," & session & " " _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) " _
        & " WHERE VOUCHERS.fyear='" & session & "' and  VOUCHERS.setupid='" & setupid & "' and DEBITORCREDIT = 'D' and genledger='" + Trim(COMBOGENLEDGER.text) + "' and SLEDGER.SUBLEDGER= '" & party & "' and convert(smalldatetime,voucherDATE,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0;", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, 0 AS OpAmountdebit,  Sum(VOUCHERS.Amount) AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & "," & session & "  " _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) " _
        & " WHERE  VOUCHERS.fyear='" & session & "' and  VOUCHERS.setupid='" & setupid & "' and DEBITORCREDIT = 'C' and genledger='" + Trim(COMBOGENLEDGER.text) + "' and SLEDGER.SUBLEDGER= '" & party & "' and convert(smalldatetime,voucherdate,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0; ", p, adCmdText
         ''''ok
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, Sum(CNF1A.NA) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT  , " & UId & " as UserId," & setupid & "," & session & " " _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) " _
        & " WHERE CNF1A.fyear='" & session & "' and  CNF1A.setupid='" & setupid & "' and (((CNF1A.DC)='D')) and pgld = '" + Trim(COMBOGENLEDGER.text) + "'and CNF1A.PSLD = '" & party & "'  and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(CNF1A.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId ," & setupid & "," & session & "" _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) " _
        & " WHERE CNF1A.fyear='" & session & "' and  CNF1A.setupid='" & setupid & "' and (((CNF1A.DC)='C')) and pgld = '" + Trim(COMBOGENLEDGER.text) + "' and CNF1A.PSLD = '" & party & "'  and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
        
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFA.NA)  AS OPAMOUNTDEBIT,  0  AS OPAMOUNTCREDIT , " & UId & " as UserId ," & setupid & "," & session & "" _
        & " FROM SLEDGER LEFT JOIN DNFA ON (SLEDGER.gledger = DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD) " _
        & " WHERE DNFA.fyear='" & session & "' and  DNFA.setupid='" & setupid & "' and ((( DNFA.DC) = 'D' )) and  pgld = '" + Trim(COMBOGENLEDGER.text) + "' and   DNFA.PSLD = '" & party & "' and convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(DNFA.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId ," & setupid & "," & session & "" _
        & " FROM SLEDGER LEFT JOIN DNFA  ON (SLEDGER.gledger =DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD)" _
        & " WHERE DNFA.fyear='" & session & "' and  DNFA.setupid='" & setupid & "' and (((DNFA.DC)='C')) and   pgld = '" + Trim(COMBOGENLEDGER.text) + "'  and  DNFA.PSLD = '" & party & "'   and convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(CNF1B.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId ," & setupid & "," & session & "" _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) " _
        & " WHERE CNF1B.fyear='" & session & "' and  CNF1B.setupid='" & setupid & "' and (((CNF1B.DC)='D')) and gld  = '" + Trim(COMBOGENLEDGER.text) + "' and  CNF1B.SLD = '" & party & "'    and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0; ", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(CNF1B.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId ," & setupid & "," & session & "" _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) " _
        & " WHERE CNF1B.fyear='" & session & "' and  CNF1B.setupid='" & setupid & "' and (((CNF1B.DC)= 'C')) and gld  = '" + Trim(COMBOGENLEDGER.text) + "'  and  CNF1B.SLD = '" & party & "' and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103)" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0;", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFB.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId ," & setupid & "," & session & "" _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER = DNFB.SLD) " _
        & " WHERE DNFB.fyear='" & session & "' and  DNFB.setupid='" & setupid & "' and (((DNFB.DC)='D')) and   gld = '" + Trim(COMBOGENLEDGER.text) + "'  and  DNFB.SLD = '" & party & "'  And convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103)" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0; ", p, adCmdText
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(DNFB.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & "," & session & " " _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER =DNFB.SLD)" _
        & " WHERE DNFB.fyear='" & session & "' and  DNFB.setupid='" & setupid & "' and (((DNFB.DC)= 'C')) and  gld = '" + Trim(COMBOGENLEDGER.text) + "'  and  DNFB.SLD = '" & party & "' and convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103)" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0 ", p, adCmdText

'======End Opening Data

      If RS.State = 1 Then RS.Close
      DoEvents
      DoEvents
      DoEvents
      DoEvents
      DoEvents
      DoEvents
      con.Execute ("delete  from treport where " & stringyear & "")
      
      





End Sub


Private Sub Form_Unload(cancel As Integer)
popupvalue5 = ""
End Sub

Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)

Screen.MousePointer = vbHourglass
If KeyCode = 116 Then
   Combosubledger.SetFocus
   Screen.MousePointer = vbDefault
End If


If KeyCode = 13 Then


If vs.TextMatrix(vs.RowSel, 2) = "I" Then
         inviceNo = vs.TextMatrix(vs.RowSel, 0)
         invoice.Show
ElseIf vs.TextMatrix(vs.RowSel, 2) = "CI" Then
         inviceNo = vs.TextMatrix(vs.RowSel, 0)
         Critnote.Show
ElseIf vs.TextMatrix(vs.RowSel, 2) = "CN" Then
         inviceNo = vs.TextMatrix(vs.RowSel, 0)
         Creditnotefile.Show
ElseIf vs.TextMatrix(vs.RowSel, 2) = "D" Then
   
         inviceNo = vs.TextMatrix(vs.RowSel, 0)
         Debitnotefile.Show
   
ElseIf vs.TextMatrix(vs.RowSel, 2) = "C/M" Then
         inviceNo = vs.TextMatrix(vs.RowSel, 0)
         countersale.Show
End If


End If


Screen.MousePointer = vbDefault

End Sub
