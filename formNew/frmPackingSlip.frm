VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPackingSlip 
   Caption         =   "Packing Slip"
   ClientHeight    =   8772
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   9756
   Icon            =   "frmPackingSlip.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8772
   ScaleWidth      =   9756
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1_print 
      Caption         =   "Direct Print"
      Height          =   255
      Left            =   2220
      TabIndex        =   31
      Top             =   7020
      Width           =   1575
   End
   Begin VB.TextBox txtOrderNo1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4320
      TabIndex        =   9
      Top             =   720
      Width           =   1005
   End
   Begin VB.TextBox txtState 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5460
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   1935
      Width           =   4170
   End
   Begin VB.TextBox txtAddress3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5460
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   1620
      Width           =   4170
   End
   Begin VB.TextBox txtBillNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   23
      Top             =   1080
      Width           =   1185
   End
   Begin VB.TextBox txtBundle 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      MaxLength       =   40
      TabIndex        =   21
      Top             =   1440
      Width           =   3105
   End
   Begin VB.TextBox txtOrderNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   14
      Top             =   720
      Width           =   1185
   End
   Begin VB.TextBox txtParty 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5460
      TabIndex        =   13
      Top             =   690
      Width           =   4170
   End
   Begin VB.TextBox txtNarration 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      MaxLength       =   40
      TabIndex        =   12
      Top             =   1800
      Width           =   3105
   End
   Begin VB.TextBox txtPartyAdd1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5460
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1005
      Width           =   4170
   End
   Begin VB.TextBox txtPartyAdd2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5460
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1320
      Width           =   4170
   End
   Begin VB.Frame buttonFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   120
      TabIndex        =   1
      Top             =   7440
      Width           =   9255
      Begin VB.CommandButton cmdPrintDelChallan 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Print Delivery &Challan"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   6420
         Picture         =   "frmPackingSlip.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   120
         Width           =   1245
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   5100
         Picture         =   "frmPackingSlip.frx":0BF0
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   1245
      End
      Begin VB.CommandButton cmdAdd_1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Re&fresh"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   45
         Picture         =   "frmPackingSlip.frx":17D4
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1170
      End
      Begin VB.CommandButton cmdSave_2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   1230
         Picture         =   "frmPackingSlip.frx":23B8
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   135
         Width           =   1230
      End
      Begin VB.CommandButton cmdDelete_3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Delete"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   2520
         Picture         =   "frmPackingSlip.frx":2F9C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1230
      End
      Begin VB.CommandButton cmdEdit_4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   3810
         Picture         =   "frmPackingSlip.frx":3B80
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   135
         Width           =   1230
      End
      Begin VB.CommandButton cmdExit_12 
         BackColor       =   &H00FFFFC0&
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   7740
         Picture         =   "frmPackingSlip.frx":3F8D
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1350
      End
   End
   Begin VB.TextBox txtTotQty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5940
      TabIndex        =   0
      Text            =   "0"
      Top             =   6960
      Width           =   1140
   End
   Begin Crystal.CrystalReport CR 
      Left            =   8640
      Top             =   6360
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker txtOrderDate 
      Height          =   315
      Left            =   2880
      TabIndex        =   8
      Top             =   720
      Width           =   1395
      _ExtentX        =   2455
      _ExtentY        =   550
      _Version        =   393216
      Format          =   181731329
      CurrentDate     =   39500
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   4590
      Left            =   0
      TabIndex        =   15
      Top             =   2340
      Width           =   8760
      _cx             =   15452
      _cy             =   8096
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   16761024
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16764622
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
      Rows            =   200
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   250
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
      BackStyle       =   0  'Transparent
      Caption         =   "Order No"
      Height          =   240
      Index           =   6
      Left            =   4380
      TabIndex        =   29
      Top             =   420
      Width           =   945
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Press F4 Key To Delete Item"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   60
      TabIndex        =   28
      Top             =   7020
      Width           =   2955
   End
   Begin VB.Label lblPacking 
      BackStyle       =   0  'Transparent
      Caption         =   "Packing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   6
      Left            =   60
      TabIndex        =   25
      Top             =   0
      Width           =   4725
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill/Challan No"
      Height          =   300
      Index           =   5
      Left            =   60
      TabIndex        =   24
      Top             =   1080
      Width           =   1185
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bundle(s)"
      Height          =   300
      Index           =   3
      Left            =   60
      TabIndex        =   22
      Top             =   1440
      Width           =   1005
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Slip No :"
      Height          =   270
      Index           =   0
      Left            =   60
      TabIndex        =   20
      Top             =   780
      Width           =   870
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Party Name"
      Height          =   240
      Index           =   2
      Left            =   5460
      TabIndex        =   19
      Top             =   420
      Width           =   1125
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date "
      Height          =   270
      Index           =   1
      Left            =   2460
      TabIndex        =   18
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Carton No."
      Height          =   300
      Index           =   4
      Left            =   60
      TabIndex        =   17
      Top             =   1800
      Width           =   1005
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "F2 For Search"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7080
      TabIndex        =   16
      Top             =   420
      Width           =   2505
   End
End
Attribute VB_Name = "frmPackingSlip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit As Boolean
Dim party As String


Private Sub cboOrderBy_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
 If cboOrderBy <> "" Then
   cboTrans.SetFocus
 End If
End If
End Sub

Private Sub cboTrans_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If cboTrans <> "" Then
   txtTransAdd.SetFocus
End If
End If
End Sub
Private Sub cmdAdd_1_Click()
   
   maxOrder
   Edit = False
   refreshFld
   cmdSave_2.Enabled = True
   cmdDelete_3.Enabled = False
   cmdEdit_4.Enabled = False
    txtOrderNo.SetFocus
    

End Sub
Sub refreshFld()
   
   'txtParty = ""
   'txtPartyAdd1 = ""
   'txtPartyAdd2 = ""
   'txtPartyAdd3 = ""
   'txtPartyAdd4 = ""
   
   cboTrans = ""
   'txtBillNo = ""
   'txtBundle = ""
   txtNarration = ""
   txtTransAdd = ""
   txtNarration = ""
   txtTotQty = 0
   'txtOrderNo1 = ""
   
   vs.Clear
   setVSWidth

End Sub
Sub maxOrder()
   
   If RS.State = 1 Then RS.close
   RS.Open "select  max(slipno) from PackinkSlipA", con
   If IsNull(RS(0)) Then
       txtOrderNo = 1
     Else
       txtOrderNo = RS(0) + 1
   End If
   
End Sub

Private Sub cmdDelete_3_Click()

  
If MsgBox("Are you Sure ", vbQuestion + vbYesNo) = vbYes Then
  con.Execute "delete from PackinkSlipa where SlipNo=" & txtOrderNo & ""
  con.Execute "delete from PackinkSlipb where SlipNo=" & txtOrderNo & ""
End If

End Sub

Private Sub cmdEdit_4_Click()
Edit = True
cmdSave_2.Enabled = True
cmdDelete_3.Enabled = True
cmdEdit_4.Enabled = False
cmdSave_2.SetFocus
End Sub

Private Sub cmdExit_12_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()

cr.Reset
cr.ReportFileName = rptPath & "/PackinkSlip.rpt"
cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
cr.ReplaceSelectionFormula "{PackinkSlipA.slipno}=" & txtOrderNo & ""
cr.WindowShowPrintSetupBtn = True
cr.WindowShowRefreshBtn = True
cr.WindowMaxButton = True
cr.WindowState = crptMaximized
If Check1_print.value = 0 Then
cr.Action = 1
Else
cr.Destination = crptToPrinter
cr.Action = 0
End If

End Sub

Private Sub cmdPrintDelChallan_Click()
cr.Reset
cr.ReportFileName = rptPath & "/DeliverChallan.rpt"
cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass


If txtBillNo <> "" Then
cr.ReplaceSelectionFormula "{PackinkSlipA.billno}=" & txtBillNo & " and {PackinkSlipA.Category}='" & packing_ & "'"
End If

cr.WindowShowPrintSetupBtn = True
cr.WindowShowRefreshBtn = True
cr.WindowMaxButton = True
cr.WindowState = crptMaximized
cr.Action = 1
End Sub
Function checkTitle_inOrder(ord As String, bcode As String) As Boolean
   
On Error GoTo abc
   
If rs1.State = 1 Then rs1.close
rs1.Open "select top 1 INVOICENO  from OrderBookList where INVOICENO=" & ord & " and bookcode='" & bcode & "'", con
If rs1.EOF = True Then
   checkTitle_inOrder = True
Else
   checkTitle_inOrder = False
End If
Exit Function

abc:

MsgBox "" & err.DESCRIPTION
   
      
End Function
Private Sub cmdSave_2_Click()



On Error GoTo save:

'===========================================================
If Edit = True Then
   con.Execute "delete from PackinkSlipA where slipno=" & Val(txtOrderNo) & ""
   con.Execute "delete from PackinkSlipB where slipno=" & Val(txtOrderNo) & ""
End If

Dim pQty, OrQty
pQty = 0
OrQty = 0

For I = 1 To vs.rows - 1
 If vs.TextMatrix(I, 1) <> "" Then
 If checkTitle_inOrder(txtOrderNo1.text, vs.TextMatrix(I, 1)) = True Then
      MsgBox "This Title is not found Related Order .... ", vbCritical
      Exit Sub
 End If
 
 pQty = 0
 OrQty = 0
 OrQty1 = 0
 
 Set rs1 = con.Execute("exec check_PackingQty '" & Trim(txtOrderNo1.text) & "','" & vs.TextMatrix(I, 1) & "'")
  If rs1.EOF = False Then
     pQty = rs1!qty
     OrQty1 = rs1!qty
 End If
 
 
 
 Set rs1 = con.Execute("exec check_Data '" & Trim(txtOrderNo1.text) & "','" & vs.TextMatrix(I, 1) & "'")
  If rs1.EOF = False Then
     If packing_ = "inv" Then
       OrQty = rs1!qty
     Else
       OrQty = rs1!Spqty
     End If
 End If
 
 
 
If Val(vs.TextMatrix(I, 3)) > 0 Then
 pQty = pQty + Val(vs.TextMatrix(I, 3))
End If
 
 If OrQty < pQty Then
      MsgBox "Qty. Exceed Related Order .... " & vs.TextMatrix(I, 1) & " Order Qty : " & OrQty & vbCrLf & " Already Made Packing Qty : " & OrQty1, vbCritical
      Exit Sub
 End If
 
 
 End If
 Next
 
 ''===========================================================
      

If Edit = False Then
   maxOrder
      
      
 
      
      
   If MsgBox("want to save ?", vbQuestion + vbYesNo) = vbNo Then
      Exit Sub
   End If
   
   If txtParty = "" Then
      MsgBox "Select Paty Name...", vbCritical
      txtParty.SetFocus
      Exit Sub
   End If
   
   If txtBillNo = "" Then
      MsgBox "Enter Bill/Challan No...", vbCritical
      txtBillNo.SetFocus
      Exit Sub
   End If
   
   CheckexistBill
   
   con.Execute "insert into PackinkSlipA(slipno,Dates,party,address1,address2,marka,bundles,billno,category,address3,states,orderno) " & _
   " values(" & Val(txtOrderNo) & ",'" & Format(txtOrderDate, "MM/dd/yyyy") & "','" & Trim(txtParty) & "','" & Trim(txtPartyAdd1) & "','" & Trim(txtPartyAdd2) & "','" & txtBundle.text & "','" & Trim(txtNarration.text) & "'," & Val(txtBillNo.text) & ",'" & packing_ & "','" & txtAddress3.text & "','" & txtstate & "'," & Val(txtOrderNo1) & ")"
   
   For I = 1 To vs.rows - 1
   If vs.TextMatrix(I, 1) <> "" Then
     
     If packing_ = "inv" Then
        con.Execute "insert into PackinkSlipB(slipno,Dates,Bookcode,Bookname,qty,spqty) values(" & Val(txtOrderNo) & ",'" & Format(txtOrderDate, "MM/dd/yyyy") & "','" & vs.TextMatrix(I, 1) & "','" & vs.TextMatrix(I, 2) & "','" & vs.TextMatrix(I, 3) & "','" & vs.TextMatrix(I, 5) & "')"
     Else
        con.Execute "insert into PackinkSlipB(slipno,Dates,Bookcode,Bookname,qty) values(" & Val(txtOrderNo) & ",'" & Format(txtOrderDate, "MM/dd/yyyy") & "','" & vs.TextMatrix(I, 1) & "','" & vs.TextMatrix(I, 2) & "','" & vs.TextMatrix(I, 3) & "')"
     End If
     
 
   End If
   Next
   
   
Else


   If MsgBox("want to Modify ?", vbQuestion + vbYesNo) = vbNo Then
      Exit Sub
   End If


   con.Execute "delete from PackinkSlipA where slipno=" & Val(txtOrderNo) & ""
   con.Execute "delete from PackinkSlipB where slipno=" & Val(txtOrderNo) & ""
      
   
   con.Execute "insert into PackinkSlipA(slipno,Dates,party,address1,address2,marka,bundles,billno,category,address3,states,orderno) " & _
   " values(" & Val(txtOrderNo) & ",'" & Format(txtOrderDate, "MM/dd/yyyy") & "','" & Trim(txtParty) & "','" & Trim(txtPartyAdd1) & "','" & Trim(txtPartyAdd2) & "','" & txtBundle.text & "','" & Trim(txtNarration.text) & "'," & Val(txtBillNo.text) & ",'" & packing_ & "','" & txtAddress3.text & "','" & txtstate & "'," & Val(txtOrderNo1) & ")"

   
   
   For I = 1 To vs.rows - 1
   If vs.TextMatrix(I, 1) <> "" Then
     
     If packing_ = "inv" Then
        con.Execute "insert into PackinkSlipB(slipno,Dates,Bookcode,Bookname,qty,spqty) values(" & Val(txtOrderNo) & ",'" & Format(txtOrderDate, "MM/dd/yyyy") & "','" & vs.TextMatrix(I, 1) & "','" & vs.TextMatrix(I, 2) & "','" & vs.TextMatrix(I, 3) & "','" & vs.TextMatrix(I, 5) & "')"
     Else
        con.Execute "insert into PackinkSlipB(slipno,Dates,Bookcode,Bookname,qty) values(" & Val(txtOrderNo) & ",'" & Format(txtOrderDate, "MM/dd/yyyy") & "','" & vs.TextMatrix(I, 1) & "','" & vs.TextMatrix(I, 2) & "','" & vs.TextMatrix(I, 3) & "')"
     End If
     
        'RS.AddNew
        'RS!slipno = Val(txtOrderNo)
        'RS!Dates = txtOrderDate
        'RS!Bookcode = vs.TextMatrix(I, 1)
        'RS!Bookname = vs.TextMatrix(I, 2)
        'RS!qty = vs.TextMatrix(I, 3)
        'RS.update
   End If
   Next


End If

   
   






cmdSave_2.Enabled = False
cmdDelete_3.Enabled = False
cmdEdit_4.Enabled = True
cmdEdit_4.SetFocus

Exit Sub

save:

MsgBox "" & err.DESCRIPTION

End Sub
Sub searchData()

   'If packing_ = "inv" Then
   '   popuplist10 "select SlipNo,Dates,Party from PackinkSlipA where Category='inv' group by SlipNo,Dates,Party", CON
   'Else
   '   popuplist10 "select SlipNo,Dates,Party from PackinkSlipA where Category='invsp' group by SlipNo,Dates,Party", CON
   'End If


   If RS.State = 1 Then RS.close
   If packing_ = "inv" Then
      RS.Open "select * from PackinkSlipA where Category='inv' and slipNo=" & txtOrderNo & "", con, adOpenDynamic, adLockOptimistic
   Else
      RS.Open "select * from PackinkSlipA where Category='invsp' and slipNo=" & txtOrderNo & "", con, adOpenDynamic, adLockOptimistic
   End If
   If RS.EOF = True Then
      Exit Sub
   End If
   
   If RS.EOF = False Then
      txtOrderDate = RS!dates
      txtParty = RS!party
      txtPartyAdd1 = RS!address1 & ""
      txtPartyAdd2 = RS!address2 & ""
      txtBundle.text = RS!marka & ""
      txtBillNo.text = RS!billno & ""
      txtNarration.text = RS!bundles & ""
      txtAddress3.text = RS!address3 & ""
      txtstate = RS!states & ""
      
      txtOrderNo1.text = RS!orderNo & ""

      
   End If
   
            
   If RS.State = 1 Then RS.close
   RS.Open "select * from PackinkSlipB where slipNo=" & txtOrderNo & " order by printorder", con, adOpenDynamic, adLockOptimistic
   For I = 1 To RS.RecordCount
   If RS.EOF = False Then
        vs.TextMatrix(I, 0) = I
        vs.TextMatrix(I, 1) = RS!Bookcode & ""
        vs.TextMatrix(I, 2) = RS!Bookname & ""
        vs.TextMatrix(I, 3) = RS!qty & ""
        vs.TextMatrix(I, 4) = RS!PRINTORDER
        If packing_ = "inv" Then
           vs.TextMatrix(I, 5) = RS!Spqty & ""
        End If
   End If
   RS.MoveNext
   Next
   
   
   
   
   cmdDelete_3.Enabled = False
   cmdSave_2.Enabled = False
   
Total

End Sub

Private Sub Form_Activate()
SetButton cmdEdit_4, cmdEdit_4, cmdSave_2, cmdDelete_3
End Sub
Private Sub Form_Load()


If packing_ = "inv" Then
   lblPacking(6) = "PACKING SLIP FOR INVOICE :- "
Else
   lblPacking(6) = "PACKING SLIP FOR SPECIMEN :-"
End If


bcode = ""
BName = ""

Me.top = 900
Me.Left = 300
setVSWidth

If RS.State = 1 Then RS.close
'RS.Open "select * from books where " & stringyear & " order by bookcode", CCON, adOpenDynamic, adLockReadOnly, adCmdText
RS.Open "select * from books where " & stringyear & " order by bookcode", CCON, adOpenDynamic, adLockReadOnly, adCmdText
If Not RS.BOF Then
    Do While Not RS.EOF

       If bcode = "" Then
        bcode = RS("bookcode")
       Else
        bcode = bcode & "|" & RS("bookcode")
       End If
        If Not RS.EOF Then
            RS.MoveNext
        End If
    Loop
End If
RS.close


If RS.State = 1 Then RS.close
RS.Open "select * from books where " & stringyear & " order by bookname", CCON, adOpenDynamic, adLockReadOnly, adCmdText
If Not RS.BOF Then
    Do While Not RS.EOF
        
       If bcode = "" Then
        BName = RS("bookname")
       Else
        BName = BName & "|" & RS("bookname")
       End If
        
        If Not RS.EOF Then
            RS.MoveNext
        End If
    Loop
End If
RS.close


'vs.ColComboList(1) = bcode
vs.ColComboList(2) = BName



maxOrder

txtOrderDate = Format(Date, "dd/MM/yyyy")

    
End Sub
Sub setVSWidth()
     
     
If packing_ = "inv" Then

    vs.Cols = 6
    vs.FormatString = "S.N|B.CODE|BOOK NAME|QTY||Sp.Qty"
    vs.ColWidth(0) = 500
    vs.ColWidth(1) = 1200
    vs.ColWidth(2) = 4200
    vs.ColWidth(3) = 1000
    vs.ColWidth(4) = 0
    vs.ColWidth(5) = 1100
    
Else
    vs.Cols = 5
    vs.FormatString = "S.N|B.CODE|BOOK NAME|QTY|"
    vs.ColWidth(0) = 500
    vs.ColWidth(1) = 1200
    vs.ColWidth(2) = 4200
    vs.ColWidth(3) = 1000
    vs.ColWidth(4) = 0


End If
    
    
    
    
End Sub



Private Sub txtBankAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtBookingStn.SetFocus
End Sub
Private Sub txtBankName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtBankAdd.SetFocus
End Sub
Private Sub txtBookingStn_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
vs.SetFocus
vs.Col = 1
End If
End Sub

Private Sub txtBillNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   
   txtBundle.SetFocus
End If
End Sub
Sub CheckexistBill()

If txtBillNo.text = "" Then Exit Sub

If packing_ = "inv" Then

 If RS.State = 1 Then RS.close
 RS.Open "select top 1 INVOICENO from invoicea where INVOICENO=" & txtBillNo.text & "", con, adOpenKeyset, adLockReadOnly
 If RS.EOF = False Then
    MsgBox "This Invoice No is Already exist....", vbCritical
    txtBillNo.SetFocus
    Exit Sub
 End If
 
Else

 If RS.State = 1 Then RS.close
 RS.Open "select top 1 INVOICENO from invoicea_sp where INVOICENO=" & txtBillNo.text & "", con, adOpenKeyset, adLockReadOnly
 If RS.EOF = False Then
    MsgBox "This Invoice No is Already exist....", vbCritical
    txtBillNo.SetFocus
    Exit Sub
 End If

 
End If

End Sub
Private Sub txtBillNo_LostFocus()
CheckexistBill
End Sub
Private Sub txtBundle_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtNarration.SetFocus
End Sub
Private Sub txtNarration_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   vs.SetFocus
   vs.Col = 1
   
   Do While vs.Row > 1
      sendkeys "{home}"
      vs.Row = vs.Row - 1
   Loop

End If


End Sub
Private Sub txtOrderDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then txtOrderNo1.SetFocus
End Sub

Private Sub txtOrderNo_GotFocus()
If PopUpValue1 <> "" Then

refreshFld
If party = "f2" Then
   txtOrderNo = PopUpValue1
Else
   txtOrderNo = PopUpValue3
End If

searchData
setVSWidth
'SetButton cmdEdit_4, cmdEdit_4, cmdSave_2, cmdDelete_3
PopUpValue1 = ""
PopUpValue2 = ""
PopUpValue3 = ""

End If
End Sub

Private Sub txtOrderNo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then
   
   searchType = "inv"
   party = "f2"
   If packing_ = "inv" Then
      'popuplist10 "select SlipNo,Dates,Party from PackinkSlipA where Category='inv' group by SlipNo,Dates,Party", con
      popuplistFast "select SlipNo,Dates,Party from PackinkSlipA where Category='inv' group by SlipNo,Dates,Party", con, , , "PINV"
   Else
      popuplistFast "select SlipNo,Dates,Party from PackinkSlipA where Category='invsp' group by SlipNo,Dates,Party", con, , , "PSP"
      'popuplist10 "select SlipNo,Dates,Party from PackinkSlipA where Category='invsp' group by SlipNo,Dates,Party", con
   End If
   
ElseIf KeyCode = 112 Then
   
   party = "f1"
   searchType = "inv"
   If packing_ = "inv" Then
      popuplist10 "select BillNo,Party,SlipNo,Dates from PackinkSlipA where Category='inv' group by BillNo,Party,SlipNo,Dates", con
   Else
      popuplist10 "select BillNo,Party,SlipNo,Dates from PackinkSlipA where Category='invsp' group by BillNo,Party,SlipNo,Dates", con
   End If

    
End If

End Sub

Private Sub txtOrderNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       refreshFld
       searchData
       setVSWidth
       
       txtOrderDate.SetFocus
    End If

End Sub
Function checkPendingOrder() As Boolean

Dim ss_11 As Long

If RS.State = 1 Then RS.close
RS.Open "select INVOICENO,BOOKCODE,QUANTITY,BookName,pending from OrderBookList where INVOICENO='" & txtOrderNo1 & "' and pending='n'", con, adOpenForwardOnly, adLockReadOnly
For I = 1 To RS.RecordCount

If rs1.State = 1 Then rs1.close
rs1.Open "select sum(QUANTITY) from invoiceBQry where (OrderNo=" & RS!invoiceNo & " and bookcode='" & RS!Bookcode & "')", con
If IsNull(rs1(0)) Then
   ss_11 = RS!QUANTITY
Else
   ss_11 = RS!QUANTITY - rs1(0)
End If
   
If ss_11 > 0 Then
   checkPendingOrder = True
   Exit Function
End If

RS.MoveNext
Next


End Function
Private Sub txtOrderNo1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then

    If packing_ = "inv" Then
        
    
    
        If txtOrderNo1.text <> "" Then
           If RS.State = 1 Then RS.close
           RS.Open "select SUBLEDGER,Address1,Address2,party_dist,party_state,Shipto," & _
           "Shipto_Add1,Shipto_Add2,Shipto_district,Shipto_States from ORDERA where INVOICENO=" & txtOrderNo1.text & " and sale_sp='sale'", con
           If RS.EOF = False Then
              
              If RS!Shipto = "" Then
                txtParty = RS!subledger
                txtPartyAdd1 = RS!address1
                txtPartyAdd2 = RS!address2
                txtAddress3 = RS!party_dist & "(" & RS!party_state & ")"
              Else
                txtParty = RS!Shipto
                txtPartyAdd1 = RS!Shipto_Add1
                txtPartyAdd2 = RS!Shipto_Add2
                txtAddress3 = RS!Shipto_district & "(" & RS!Shipto_States & ")"
        
           End If
        End If
        End If
        txtParty.SetFocus
        
    Else
    
        If txtOrderNo1.text <> "" Then
           If RS.State = 1 Then RS.close
           RS.Open "select SUBLEDGER,Address1,Address2,party_dist,party_state,Shipto," & _
           "Shipto_Add1,Shipto_Add2,Shipto_district,Shipto_States from ORDERA where INVOICENO=" & txtOrderNo1.text & " and sale_sp='sp'", con
           If RS.EOF = False Then
            
            If RS!Shipto = "" Then
              txtParty = RS!subledger
              txtPartyAdd1 = RS!address1
              txtPartyAdd2 = RS!address2
              txtAddress3 = RS!party_dist & "(" & RS!party_state & ")"
            Else
                txtParty = RS!Shipto
                txtPartyAdd1 = RS!Shipto_Add1
                txtPartyAdd2 = RS!Shipto_Add2
                txtAddress3 = RS!Shipto_district & "(" & RS!Shipto_States & ")"
           End If
        End If
        End If
        txtParty.SetFocus


End If
End If

End Sub

Private Sub txtParty_GotFocus()
  
If PopUpValue1 <> "" Then
  
If packing_ = "inv" Then

   txtParty = PopUpValue1
   txtPartyAdd1 = PopUpValue2
   txtPartyAdd2 = PopUpValue3
   If UCase(popupvalue4) = UCase(popupvalue5) Then
      txtAddress3 = popupvalue4
   Else
      txtAddress3 = popupvalue4 & "(" & popupvalue5 & ")"
   End If
   
Else

   txtParty = PopUpValue1
   txtPartyAdd1 = PopUpValue2
   txtPartyAdd2 = PopUpValue3
   If UCase(popupvalue4) = UCase(popupvalue5) Then
      txtAddress3 = popupvalue4
   Else
      txtAddress3 = popupvalue4 & "(" & popupvalue5 & ")"
   End If
   

End If
   
   
   PopUpValue1 = ""
   PopUpValue2 = ""
   PopUpValue3 = ""
   popupvalue4 = ""
   popupvalue5 = ""
     
'End If

End If
  
End Sub

Private Sub txtParty_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then



If packing_ = "inv" Then
    searchType = "party"
    
    popuplist10 "select DESCFORINVOICE as PartyName,Address1,Address2,Address3 as City,distcode  from Sledger order by DESCFORINVOICE", con
Else

    searchType = "party"
    popuplist10 "select Rep as Representative,Add1,Add2,City,District from SalesRepQry order by Rep", CON_blue


End If


End If

''If KeyCode = 113 Then
''    tblNo = 50
''    frmSearchItem.Show
''End If


If KeyCode = 13 Then
    If txtParty <> "" Then
     txtBillNo.SetFocus
    End If
End If

End Sub

Private Sub txtTransAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtNarration.SetFocus
End Sub

Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)
  
  If KeyCode = 115 Then
  If MsgBox("Want to delete ?", vbQuestion + vbYesNo) = vbYes Then
    vs.RemoveItem (vs.RowSel)
    Total
  End If
  End If

End Sub

Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then
      
      If vs.Col = 1 Then
          If RS.State = 1 Then RS.close
          RS.Open "select Bookcode,bookname,rate from BOOKS where bookcode='" & vs.TextMatrix(vs.RowSel, 1) & "'", con
          If RS.EOF = False Then
              vs.TextMatrix(vs.RowSel, 2) = RS!Bookname
              vs.TextMatrix(vs.RowSel, 1) = UCase(vs.TextMatrix(vs.RowSel, 1))
              vs.TextMatrix(vs.RowSel, 0) = vs.Row
              
              
'              If checkTitle_inOrder(txtOrderNo1.Text, vs.TextMatrix(vs.RowSel, 1)) = True Then
'                 MsgBox "This Title is not found Related Order .... ", vbCritical
'                 Exit Sub
'              End If
              
              sendkeys "{right}"
              sendkeys "{right}"
          End If
           
      ElseIf vs.Col = 2 Then
          
          If RS.State = 1 Then RS.close
          RS.Open "select Bookcode,bookname,rate from BOOKS where bookname='" & vs.TextMatrix(vs.RowSel, 2) & "'", con
          If RS.EOF = False Then
              vs.TextMatrix(vs.RowSel, 1) = UCase(RS!Bookcode)
              vs.TextMatrix(vs.RowSel, 0) = vs.Row
              sendkeys "{right}"
          End If
      ElseIf vs.Col = 3 Then
           sendkeys "{down}"
           sendkeys "{home}"
           Total
      ElseIf vs.Col = 5 Then
           sendkeys "{down}"
           sendkeys "{home}"
           Total
     
      End If
      
      End If
End Sub
Sub Total()
    txtTotQty = 0
    
    For I = 1 To vs.rows - 1
       If vs.TextMatrix(I, 2) <> "" Then
          txtTotQty = Val(txtTotQty) + Val(vs.TextMatrix(I, 3))
       End If
    Next
    
    
End Sub
Private Sub vs_SelChange()
   'If vs.Col > 3 Then
   '   vs.Editable = flexEDNone
   'Else
      vs.Editable = flexEDKbdMouse
   'End If
End Sub


