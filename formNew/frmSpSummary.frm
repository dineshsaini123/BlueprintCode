VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSpSummary 
   Caption         =   "Rep. Wise List"
   ClientHeight    =   9552
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   14856
   Icon            =   "frmSpSummary.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9552
   ScaleWidth      =   14856
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtbkName 
      Height          =   315
      Left            =   1035
      TabIndex        =   5
      Top             =   1260
      Width           =   3855
   End
   Begin VB.OptionButton Option2_Qty 
      Caption         =   "Specimen Qty"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2985
      TabIndex        =   4
      Top             =   645
      Width           =   1905
   End
   Begin VB.OptionButton Option1_amount 
      Caption         =   "Specimen Amt."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1020
      TabIndex        =   12
      Top             =   675
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.CommandButton cmdAdd_1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&View"
      Height          =   615
      Left            =   12060
      Picture         =   "frmSpSummary.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   1170
   End
   Begin VB.CommandButton cmdPrint_7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Print"
      Enabled         =   0   'False
      Height          =   615
      Left            =   13320
      Picture         =   "frmSpSummary.frx":0BF0
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      Width           =   1170
   End
   Begin VB.CommandButton cmdExit_12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "E&xit"
      Height          =   615
      Left            =   12060
      Picture         =   "frmSpSummary.frx":17D4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   900
      Width           =   2430
   End
   Begin VB.ComboBox cborep 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   330
      ItemData        =   "frmSpSummary.frx":23B8
      Left            =   1020
      List            =   "frmSpSummary.frx":23E9
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   150
      Width           =   3915
   End
   Begin MSComCtl2.DTPicker txtFrom 
      Height          =   315
      Left            =   5520
      TabIndex        =   2
      Top             =   150
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   550
      _Version        =   393216
      Format          =   155582465
      CurrentDate     =   42409
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   7635
      Left            =   60
      TabIndex        =   7
      Top             =   1740
      Width           =   14475
      _cx             =   25532
      _cy             =   13467
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   12582847
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
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1000
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmSpSummary.frx":25AD
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
         Left            =   14580
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   4020
         Width           =   195
      End
   End
   Begin MSComCtl2.DTPicker txtTo 
      Height          =   315
      Left            =   7320
      TabIndex        =   3
      Top             =   150
      Width           =   1410
      _ExtentX        =   2498
      _ExtentY        =   550
      _Version        =   393216
      Format          =   155582465
      CurrentDate     =   42409
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Group Name"
      Height          =   315
      Left            =   45
      TabIndex        =   13
      Top             =   1305
      Width           =   1065
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rep. Name"
      Height          =   315
      Left            =   60
      TabIndex        =   11
      Top             =   150
      Width           =   1020
   End
   Begin VB.Label lblAson 
      Caption         =   "To"
      Height          =   255
      Left            =   6960
      TabIndex        =   10
      Top             =   150
      Width           =   315
   End
End
Attribute VB_Name = "frmSpSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_1_Click()

Dim str_date As String
Dim sale, saleret

Screen.MousePointer = vbHourglass

vs.rows = 1


str_date = "(invoicedate>=convert(smalldatetime,'" & txtFrom.value & "',103) and invoicedate<=convert(smalldatetime,'" & txtto.value & "',103))"

If txtbkName.Text <> "" Then
   str_date = str_date & " and GROUPCODE ='" & txtbkName.Text & "'"
End If
 
 
vs.Cols = 5
vs.TextMatrix(0, 0) = "SNo."
vs.TextMatrix(0, 1) = "Representative"
vs.TextMatrix(0, 2) = "Specimen"
vs.TextMatrix(0, 3) = "Specimen Return"
vs.TextMatrix(0, 4) = "Net Specimen"
 
 
 con.Execute "delete from tmpRepSp where len(repname)>0"
 
 If Option1_amount.value = True Then
   
   If Me.cboRep.Text = "ALL" Then
      'con.Execute "insert into tmpRepSp(repname,NetAmt) SELECT Representative,sum(NetSale) FROM RepWisteNetSpcimen where " & str_date & " group by Representative"
      con.Execute "insert into tmpRepSp(repname,NetAmt) SELECT agentname,sum(NETAMOUNT) FROM invoiceSPBQry where " & str_date & " group by agentname"
      'con.Execute "insert into tmpRepSp(repname,NetAmt_Ret) SELECT Representative,sum(NetSale) FROM RepWisteNetSpcimen_ret where " & str_date & " group by Representative"
      con.Execute "insert into tmpRepSp(repname,NetAmt_Ret) SELECT agentname,sum(NETAMOUNT) FROM invoiceSPRETBQry where " & str_date & " group by agentname"
   Else
      'con.Execute "insert into tmpRepSp(repname,NetAmt) SELECT Representative,sum(NetSale) FROM RepWisteNetSpcimen where Representative='" & cborep.Text & "' and " & str_date & " group by Representative"
      'con.Execute "insert into tmpRepSp(repname,NetAmt_Ret) SELECT Representative,sum(NetSale) FROM RepWisteNetSpcimen_ret where Representative='" & cborep.Text & "' and " & str_date & " group by Representative"
      
      con.Execute "insert into tmpRepSp(repname,NetAmt) SELECT agentname,sum(NETAMOUNT) FROM invoiceSPBQry where agentname='" & cboRep.Text & "' and " & str_date & " group by agentname"
      con.Execute "insert into tmpRepSp(repname,NetAmt_Ret) SELECT agentname,sum(NETAMOUNT) FROM invoiceSPRETBQry where agentname='" & cboRep.Text & "' and " & str_date & " group by agentname"

   End If
   
   str11 = "SELECT RepName, sum(round(NetAmt,2)) as sale ,SUM(round(NetAmt_ret,2)) as saleret from tmpRepSp group by RepName"
   
 Else
 
   If Me.cboRep.Text = "ALL" Then
      'con.Execute "insert into tmpRepSp(repname,Qty) SELECT AgentName,sum(Qty) FROM SpecimenRegister where " & str_date & " group by AgentName"
      'con.Execute "insert into tmpRepSp(repname,Qty_Ret) SELECT AgentName,sum(Qty) FROM SpecimenReturnRegister where  " & str_date & " group by AgentName"
       con.Execute "insert into tmpRepSp(repname,Qty) SELECT agentname,sum(QUANTITY) FROM invoiceSPBQry where " & str_date & " group by agentname"
       con.Execute "insert into tmpRepSp(repname,Qty_Ret) SELECT agentname,sum(QUANTITY) FROM invoiceSPRETBQry where " & str_date & " group by agentname"
   Else
      'con.Execute "insert into tmpRepSp(repname,Qty) SELECT AgentName,sum(Qty) FROM SpecimenRegister where AgentName='" & cborep.Text & "' and " & str_date & " group by AgentName"
      'con.Execute "insert into tmpRepSp(repname,Qty_Ret) SELECT AgentName,sum(Qty) FROM SpecimenReturnRegister where AgentName='" & cborep.Text & "' and " & str_date & " group by AgentName"
       con.Execute "insert into tmpRepSp(repname,Qty) SELECT agentname,sum(QUANTITY) FROM invoiceSPBQry where agentname='" & cboRep.Text & "' and " & str_date & " group by agentname"
       con.Execute "insert into tmpRepSp(repname,Qty_Ret) SELECT agentname,sum(QUANTITY) FROM invoiceSPRETBQry where agentname='" & cboRep.Text & "' and " & str_date & " group by agentname"

   End If
   
   str11 = "SELECT RepName, sum(Qty) as sale ,SUM(Qty_ret) as saleret from tmpRepSp group by RepName"
    
 End If
 
 
 sale = 0
 saleret = 0
 rows_ = 1
 If RS.State = 1 Then RS.close
 RS.Open str11, con, adOpenKeyset, adLockOptimistic
 k1 = 1
 For J = 1 To RS.RecordCount
    sale = 0
    saleret = 0
    If Not IsNull(RS(1)) Then
       sale = Round(RS(1), 0)
    Else
       sale = 0
    End If
    If Not IsNull(RS(2)) Then
       saleret = Round(RS(2), 0)
    Else
       saleret = 0
    End If
    If (sale > 0 Or saleret > 0) Then
    
        vs.rows = vs.rows + 1
        vs.TextMatrix(rows_, 0) = J
        vs.TextMatrix(rows_, 1) = RS(0)
        vs.TextMatrix(rows_, 2) = sale
        vs.TextMatrix(rows_, 3) = saleret
        vs.TextMatrix(rows_, 4) = (sale - saleret)
        rows_ = rows_ + 1
    End If
    
    RS.MoveNext
 Next
 
 

 
txtTotal = 0

For I = 1 To vs.rows - 1
If vs.TextMatrix(I, 0) <> "" Then
   txtTotal = txtTotal + Val(vs.TextMatrix(I, 4))
End If
Next


vs.ColWidth(0) = 1100
vs.ColWidth(1) = 5200
vs.ColWidth(2) = 2500
vs.ColWidth(3) = 2500
vs.ColWidth(4) = 2500


 
cmdPrint_7.Enabled = True
Screen.MousePointer = vbDefault
    
    

End Sub

Private Sub cmdExit_12_Click()
Unload Me
End Sub

Private Sub cmdPrint_7_Click()
   
   DSNNew
   
   con.Execute "delete from TmpBook where head='" & UId & "'"
   con.Execute "delete from TmpBook where head is null"
   
   For I = 1 To vs.rows - 1
   If vs.TextMatrix(I, 0) <> "" Then
     con.Execute "insert into TmpBook(BName,BalanceQty,head,issueQty,OrderNo) values('" & vs.TextMatrix(I, 1) & "','" & vs.TextMatrix(I, 2) & "','" & UId & "','" & vs.TextMatrix(I, 3) & "','" & vs.TextMatrix(I, 4) & "')"
   End If
   Next
   
   If MsgBox("Want To Print ?", vbInformation + vbYesNo) = vbYes Then
    
    MainMenu.cr1.Reset
    MainMenu.cr1.ReportFileName = rptPath & "/NetSales_sp.rpt"
    MainMenu.cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    MainMenu.cr1.ReplaceSelectionFormula "{TmpBook.head}='" & UId & "'"
    MainMenu.cr1.WindowShowPrintSetupBtn = True
    
    MainMenu.cr1.WindowShowExportBtn = True
    MainMenu.cr1.Formulas(0) = "fdate='" & Me.txtFrom.value & "'"
    MainMenu.cr1.Formulas(1) = "tdate='" & Me.txtto.value & "'"
    MainMenu.cr1.WindowState = crptMaximized
    MainMenu.cr1.WindowShowRefreshBtn = True
    MainMenu.cr1.Action = 1
    
   End If

End Sub

Private Sub Form_Load()
 txtFrom.value = from_date
 txtto.value = to_date
 
   If RS.State = 1 Then RS.close
   RS.Open "select Rep as Representative from SalesRepQry order by Rep", CON_blue
    cboRep.Clear
    If Not RS.EOF Then
       Do While Not RS.EOF
          If IsNull(RS(0)) = False Then
            Me.cboRep.AddItem RS(0)
          End If
          If Not RS.EOF Then RS.MoveNext
        Loop
    End If
    
    Me.cboRep.AddItem "ALL"
    Me.cboRep.Text = "ALL"
    
    
End Sub
Private Sub txtbkName_GotFocus()
If PopUpValue1 <> "" Then
   txtbkName.Text = PopUpValue1
   PopUpValue1 = ""
End If
End Sub
Private Sub txtbkName_KeyDown(KeyCode As Integer, Shift As Integer)
   
If KeyCode = 113 Then

   searchType = "party"
   value = "select distinct iif(GROUPCODE_sub is null,GROUPCODE,GROUPCODE_sub) as GROUPCODE from BOOKS where (freeItem='n') order by GROUPCODE"
   popuplist_client value, con
   set_focus = True

End If


End Sub
