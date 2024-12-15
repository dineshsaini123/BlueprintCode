VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmChangeModule 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   3105
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3960
      _cx             =   6985
      _cy             =   5477
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   11162880
      BackColorFixed  =   7917545
      ForeColorFixed  =   16777215
      BackColorSel    =   16777153
      ForeColorSel    =   -2147483647
      BackColorBkg    =   -2147483639
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   16777215
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   6
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   500
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmChangeModule.frx":0000
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
      ForeColorFrozen =   16777215
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "frmChangeModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub permissionUserWise()
    
Dim D As Integer
Dim o As Object
    
With MainMenu
    
.menuGenralLedgerMaster.Visible = False
.menusubleadgermaster.Visible = False
.mnuDistict.Visible = False
.mnuCity.Visible = False
.mnuAgentMaster.Visible = False
.mnuTransport.Visible = False
.mnuBooksMast.Visible = False
.mnuBookgp.Visible = False
.mnuReportBookGp.Visible = False
.mnuDiscount.Visible = False
.mnuSchoolMaster.Visible = False
.mnuTeacherDet.Visible = False
.menujournalvoucher.Visible = False
.menusalesinvoice.Visible = False
.mnuCreditNItem.Visible = False
.menucountersale.Visible = False
.menucreditnote.Visible = False
.menudebitnote.Visible = False
.mnuIssuedBind.Visible = False
.mnuBoromBinder.Visible = False
.mnuBookRecFBind.Visible = False
.mnuStockTranstoGo.Visible = False
.mnuInvBasil.Visible = False
.mnuInvoiceBasilRet.Visible = False
.mnuBookIssueSp.Visible = False
.mnuSpRet.Visible = False
.mnuBookIssuetoSh.Visible = False
.mnuDonation.Visible = False
.mnuBookStock.Visible = False
.menucashbook.Visible = False
.mnudistrictwisesalesreturn.Visible = False
.menugeneralledgeraccounts.Visible = False
.menusubledgeraccounts.Visible = False
.alphaSubLedgerAccountsmnu.Visible = False
.menugenledgertrialbalance.Visible = False
.menusubledgertrialbalance.Visible = False
.menugenledgeropentrialbalance.Visible = False
.mnugpsale.Visible = False
.menudistrictwisesales.Visible = False
.menuGroupwisesales.Visible = False
.menubankadvice1.Visible = False
.mnubankadviceReconciliation.Visible = False
.mnubiltyreturnregister.Visible = False
.mnuPartyList.Visible = False
.mnuDispachedreg.Visible = False
.mnuBankReg.Visible = False
.mnuCashReg.Visible = False
.mnuPartyProf.Visible = False
.mnuSL_invoice.Visible = False
.mnuBR.Visible = False
.conven_mnuBookwiseAgnLedger.Visible = False
.conven_mnuTotalBookQty.Visible = False
.conven_mnuTotalBookAmt.Visible = False
.conven_mnuCollegeList.Visible = False
.conven_mnuAgentwiseIssue.Visible = False
.conven_mnuDispatchReg.Visible = False

'---End Part--------------------------------
.mnuInvoiceEnd.Visible = False
.menucreditnoteandpartmaster.Visible = False
.mnuCounterEndPart.Visible = False
.mnuInvoiceEnd_basil.Visible = False
.mnuInvoiceEnd_basil_ret.Visible = False
.mnuBKIssueEndP.Visible = False
.mnuBKReturnEndP1.Visible = False

.mnuBookRec_IssueCat.Visible = False
.mnuPaper_size.Visible = False
.mnuPaper_gsm.Visible = False
.mnuPaper_Maker.Visible = False
.mnuPBookMast.Visible = False
.mnuGodown.Visible = False

.mnuSchoolMaster.Visible = False
.mnuTeacherDet.Visible = False
.mnuLedger.Visible = False

'''''paper---------------
.mnuPaper_NegativePrint.Visible = False
.mnuPaper_TitlePrint.Visible = False
.mnuPaper_subjectprint.Visible = False
.mnuLedger.Visible = False
.mnuLedger_conven.Visible = False
.mnuOrder.Visible = False
.mnuPaerRec.Visible = False
.mnuStockTrans.Visible = False
.mnuPaperSize.Visible = False
.mnuBooksdet.Visible = False

'========================
'Stock System

.mnuBookRec_IssueCat.Visible = False
.mnuIssuedBind.Visible = False
.mnuBoromBinder.Visible = False
.mnuBookRecFBind.Visible = False
.mnuStockTranstoGo.Visible = False
.mnuLedger_Basil.Visible = False

End With
   
       
If RS.State = 1 Then RS.close
RS.Open "select taskname,tasktype,permission from UsrePermission where (module='" & module_ & "' and username= '" & main.username & "') order by tasktype", CON, adOpenKeyset, adLockReadOnly
While RS.EOF = False
For Each o In MainMenu
If o.Name = RS!taskType Then
   If RS!Permission = "y" Then
      o.Visible = True
   Else
      o.Visible = False
   End If
End If
Next
RS.MoveNext
Wend

    
End Sub
Private Sub Form_Load()

vs.Rows = 1
If RS.State = 1 Then RS.close
RS.Open "select * from UsrePermission where (module='Module' and username='" & username & "') order by taskname", CON, adOpenKeyset, adLockReadOnly
For J = 0 To RS.RecordCount - 1
vs.TextMatrix(J, 0) = RS!taskname
vs.Rows = vs.Rows + 1
RS.MoveNext
Next
  
vs.TextMatrix(J, 0) = "Exit"







Me.Top = 1500
Me.Left = 1500

End Sub
Private Sub vs_Click()
   If vs.TextMatrix(vs.RowSel, 0) = "Exit" Then
      Unload Me
   End If
End Sub
Private Sub vs_DblClick()

If vs.TextMatrix(vs.RowSel, 0) = "Exit" Then
   Unload Me
End If

module_ = vs.TextMatrix(vs.RowSel, 0)
permissionUserWise

Unload frmChangeModule

'MainMenu.Caption = "Publication Software System : " & module_
BackColorFrom MainMenu

End Sub

Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
   
   If vs.TextMatrix(vs.RowSel, 0) = "Exit" Then
      Unload Me
   End If
   
   module_ = vs.TextMatrix(vs.RowSel, 0)
   permissionUserWise
   
   Unload frmChangeModule
   
   'MainMenu.Caption = "Publication Software System : " & module_
   BackColorFrom MainMenu

   
   
ElseIf KeyCode = 27 Then
      Unload Me
End If

End Sub
