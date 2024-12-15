VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBookSt 
   Caption         =   "Book Status"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11148
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   11148
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport cr 
      Left            =   360
      Top             =   7515
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
      Height          =   480
      Left            =   9315
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   855
      Width           =   1320
   End
   Begin VB.TextBox txtClosingBk 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5880
      TabIndex        =   13
      Top             =   7620
      Width           =   975
   End
   Begin VB.TextBox txtStockOut 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3660
      TabIndex        =   11
      Top             =   7620
      Width           =   975
   End
   Begin VB.TextBox txtSTin 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1800
      TabIndex        =   9
      Top             =   7620
      Width           =   975
   End
   Begin VB.TextBox txtGodown 
      Height          =   315
      Left            =   3960
      TabIndex        =   8
      Top             =   510
      Width           =   1335
   End
   Begin VB.TextBox txtbkname 
      Height          =   315
      Left            =   3960
      TabIndex        =   4
      Top             =   120
      Width           =   3795
   End
   Begin VB.TextBox txtBkCode 
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      Top             =   90
      Width           =   1335
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   3675
      Left            =   90
      TabIndex        =   0
      Top             =   1485
      Width           =   9915
      _cx             =   17489
      _cy             =   6482
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
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
      ForeColorFixed  =   8388608
      BackColorSel    =   16777153
      ForeColorSel    =   -2147483647
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   12
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
   Begin MSComCtl2.DTPicker dateAson 
      Height          =   330
      Left            =   1080
      TabIndex        =   5
      Top             =   510
      Width           =   1350
      _ExtentX        =   2392
      _ExtentY        =   572
      _Version        =   393216
      Format          =   156499969
      CurrentDate     =   39795
   End
   Begin VSFlex7Ctl.VSFlexGrid vs_billwise 
      Height          =   2100
      Left            =   45
      TabIndex        =   15
      Top             =   5220
      Visible         =   0   'False
      Width           =   9915
      _cx             =   17489
      _cy             =   3704
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   11162880
      BackColorFixed  =   13758456
      ForeColorFixed  =   8388608
      BackColorSel    =   16777153
      ForeColorSel    =   -2147483647
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   11
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   280
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
   Begin MSComCtl2.DTPicker txtFrom 
      Height          =   330
      Left            =   9315
      TabIndex        =   17
      Top             =   90
      Width           =   1350
      _ExtentX        =   2392
      _ExtentY        =   572
      _Version        =   393216
      Format          =   156499969
      CurrentDate     =   39795
   End
   Begin MSComCtl2.DTPicker txtToDate 
      Height          =   330
      Left            =   9315
      TabIndex        =   19
      Top             =   450
      Width           =   1350
      _ExtentX        =   2392
      _ExtentY        =   572
      _Version        =   393216
      Format          =   156499969
      CurrentDate     =   39795
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   1365
      Left            =   8235
      Top             =   45
      Width           =   2580
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   8370
      TabIndex        =   20
      Top             =   495
      Width           =   690
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   8370
      TabIndex        =   18
      Top             =   90
      Width           =   690
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Closing Book"
      Height          =   255
      Left            =   4800
      TabIndex        =   14
      Top             =   7620
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Stock out"
      Height          =   255
      Left            =   2880
      TabIndex        =   12
      Top             =   7620
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Stock in"
      Height          =   255
      Left            =   1140
      TabIndex        =   10
      Top             =   7620
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Godown :"
      Height          =   315
      Left            =   2940
      TabIndex        =   7
      Top             =   570
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "As On "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   120
      TabIndex        =   6
      Top             =   510
      Width           =   690
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Book Name :"
      Height          =   315
      Left            =   2940
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Book Code :"
      Height          =   315
      Left            =   90
      TabIndex        =   1
      Top             =   135
      Width           =   1095
   End
End
Attribute VB_Name = "frmBookSt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str1 As String
Dim str_go As String
Dim str_go_issue As String
Private Sub Text1_Change()

End Sub
Private Sub cmdPrint_Click()

Dim qtyrec, qtyissue, op As Double
con.Execute "delete from tmpBookWiseStock"

qtyrec = 0
qtyissue = 0
op = 0

str_date = "(dates<convert(smalldatetime,'" & txtFrom.value & "',103) and bookcode='" & txtBkCode.Text & "')"
If RS.State = 1 Then RS.close

If frmBookStock.Check1_crm.value = 1 Then
   
   If frmBookStock.lblStokType.Caption = "Stock Reflect With Order" Then
      RS.Open "select QtyRec,QtyIssue from bookwiseStockQry_WithOrder where " & str_date, con
   Else
      RS.Open "select QtyRec,QtyIssue from bookwiseStockQry where " & str_date, con
   End If

Else
   RS.Open "select QtyRec,QtyIssue from bookwiseStockQry_withoutcrm where godown='" & frmBookStock.cboBinder_Godown & "' and " & str_date, con
End If

While RS.EOF = False

If Not IsNull(RS!qtyrec) Then
   qtyrec = qtyrec + RS!qtyrec
End If

If Not IsNull(RS!qtyissue) Then
   qtyissue = qtyissue + RS!qtyissue
End If

RS.MoveNext
Wend

If qtyissue > 0 Then
   qtyissue = qtyissue * -1
End If


op = qtyrec + qtyissue


str_date = "((dates>=convert(smalldatetime,'" & txtFrom.value & "',103) and dates<=convert(smalldatetime,'" & txtToDate.value & "',103)) and bookcode='" & txtBkCode.Text & "')"

con.Execute "insert tmpBookWiseStock(invoiceno,dates,QtyRec,QtyIssue,narr,bookcode,op) " & _
" values('" & 0 & "','" & Format(txtFrom.value, "MM/dd/yyyy") & "','" & 0 & "','" & 0 & "','Opening','" & txtBkCode.Text & "','" & op & "') "


If RS.State = 1 Then RS.close
If frmBookStock.Check1_crm.value = 1 Then
   If frmBookStock.lblStokType.Caption = "Stock Reflect With Order" Then
      RS.Open "select invoiceno,dates,QtyRec,QtyIssue,narr,bookcode from bookwiseStockQry_WithOrder where " & str_date, con
   Else
      RS.Open "select invoiceno,dates,QtyRec,QtyIssue,narr,bookcode from bookwiseStockQry where " & str_date, con
   End If
Else
   RS.Open "select invoiceno,dates,QtyRec,QtyIssue,narr,bookcode from bookwiseStockQry_withoutcrm where godown='" & frmBookStock.cboBinder_Godown & "' and " & str_date, con
End If

While RS.EOF = False

If Not IsNull(RS!qtyrec) Then
   qtyrec = RS!qtyrec
Else
   qtyrec = 0
End If

If Not IsNull(RS!qtyissue) Then
   qtyissue = RS!qtyissue
Else
   qtyissue = 0
End If

If (qtyissue > 0 Or qtyrec > 0) Then
con.Execute "insert tmpBookWiseStock(invoiceno,dates,QtyRec,QtyIssue,narr,bookcode,op) " & _
" values('" & RS!invoiceNo & "','" & Format(RS!dates, "MM/dd/yyyy") & "','" & qtyrec & "','" & qtyissue & "','" & RS!NARR & "','" & RS!Bookcode & "',0) "
End If

RS.MoveNext
Wend



If MsgBox("Want to View ?", vbQuestion + vbYesNo) = vbYes Then
    cr.Reset
    cr.ReportFileName = rptPath & "/bookwiseStock.rpt"
    cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    cr.WindowShowPrintSetupBtn = True
    cr.WindowState = crptMaximized
    cr.Action = 1
End If



End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Unload Me
End If
End Sub

Private Sub Form_Load()

  On Error GoTo aa12

  Me.txtBkCode = PopUpValue1
  dateAson.value = frmBookStock.dateAson.value
  txtbkName.Text = PopUpValue2
  txtGodown.Text = frmBookStock.cboBinder_Godown.Text
  
  Dim rs_kit As New ADODB.Recordset
  Dim rs_fill As New ADODB.Recordset
  
  Dim q_in, q_out As Integer
  
  PopUpValue1 = ""
  PopUpValue2 = ""
  
  Me.Top = 1000
  Me.Left = 500
  
  txtFrom.value = from_date
  txtToDate.value = to_date
  'Noida Stock=============================================================
  q_in = 0
  q_out = 0
  
  If (frmBookStock.Check1_crm.value = 1 Or PopUpValue6 = 2) Then
   
   If rs_fill.State = 1 Then rs_fill.close
   rs_fill.Open "select BookCode,sum(Qty) As Qty,sum(QtySP) As QtySp,Status,status_ from tmpNSStock where bookcode='" & Me.txtBkCode & "' and convert(smalldatetime,Dates,103)<= convert(smalldatetime,'" & dateAson.value & "',103)  group by BookCode,Status,status_", con
   Set vs.DataSource = rs_fill
   
   For I = 1 To vs.rows - 1
        
        If vs.TextMatrix(I, 4) = "IN" Then
            q_in = q_in + IIf(vs.TextMatrix(I, 2) = "", 0, vs.TextMatrix(I, 2))
            q_in = q_in + IIf(vs.TextMatrix(I, 3) = "", 0, vs.TextMatrix(I, 3))
        Else
            q_out = q_out + IIf(vs.TextMatrix(I, 2) = "", 0, vs.TextMatrix(I, 2))
            q_out = q_out + IIf(vs.TextMatrix(I, 3) = "", 0, vs.TextMatrix(I, 3))
          End If
       
   Next
   
   txtStockOut = q_out
   txtSTin = q_in
   txtClosingBk = (q_in + q_out)
   
   
   
   
   
   Exit Sub
   End If
   
   'Noida stock End=========================================================
   
   str1 = "Godown='" & txtGodown.Text & "'"
   str_go = "Godown_In='" & txtGodown.Text & "'"
   str_go_issue = "Godown_Out='" & txtGodown.Text & "'"
  
  
  vs.FormatString = "Transaction|>Qty|Stock Out/IN|Remarks.."
  
  txtSTin = 0
  txtStockOut = 0
  Dim qty As Long
  qty = 0
  
  '1 sale
  
   If RS.State = 1 Then RS.close
   RS.Open "select sum(Qty) from SaleReturnRegister where convert(smalldatetime,invoiceDate,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
   " BookCode='" & txtBkCode & "' and " & _
   " " & str1 & " and " & stringyear & "", con, adOpenKeyset
   If Not IsNull(RS(0)) Then qty = RS(0)
   
   If RS.State = 1 Then RS.close
   RS.Open "select sum(Qty) from SaleReturnRegister_Free where convert(smalldatetime,invoiceDate,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
   " BookCode='" & txtBkCode & "' and " & _
   " " & str1 & " and " & stringyear & "", con, adOpenKeyset
   If Not IsNull(RS(0)) Then qty = qty + RS(0)

   If qty > 0 Then
      vs.TextMatrix(1, 0) = "Sale Return"
      vs.TextMatrix(1, 1) = qty
      txtSTin = qty
      vs.TextMatrix(1, 2) = "STOCK IN"
   Else
      vs.TextMatrix(1, 0) = "Sale Return"
      vs.TextMatrix(1, 1) = 0
      vs.TextMatrix(1, 2) = "STOCK IN"
   End If
   
   '2
   qty = 0
   
    If RS.State = 1 Then RS.close
    RS.Open "select sum(Qty) from SpecimenReturnRegister where convert(smalldatetime,invoiceDate,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
    " BookCode='" & txtBkCode & "' and " & _
    " " & str1 & " and " & stringyear & "", con, adOpenKeyset
    If Not IsNull(RS(0)) Then qty = RS(0)
    
    
    If RS.State = 1 Then RS.close
    RS.Open "select sum(Qty) from SpecimenReturnRegister_Free where convert(smalldatetime,invoiceDate,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
    " BookCode='" & txtBkCode & "' and " & _
    " " & str1 & " and " & stringyear & "", con, adOpenKeyset
    If Not IsNull(RS(0)) Then qty = qty + RS(0)
    
    
    If qty > 0 Then
      vs.TextMatrix(2, 0) = "Specimen Return"
      vs.TextMatrix(2, 1) = qty
      txtSTin = Val(txtSTin) + qty
      vs.TextMatrix(2, 2) = "STOCK IN"
   Else
      vs.TextMatrix(2, 0) = "Specimen Return"
      vs.TextMatrix(2, 1) = 0
      vs.TextMatrix(2, 2) = "STOCK IN"
   End If

   
    '3
   
    If RS.State = 1 Then RS.close
    RS.Open "select sum(Qty) from BinderReceiveRegister where convert(smalldatetime,invoiceDate,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
    " Book_Code='" & txtBkCode & "' and " & _
    " " & str1 & " and " & stringyear & "", con, adOpenKeyset
    If Not IsNull(RS(0)) Then
      vs.TextMatrix(3, 0) = "Binder Receive"
      vs.TextMatrix(3, 1) = RS(0)
      txtSTin = Val(txtSTin) + RS(0)
      vs.TextMatrix(3, 2) = "STOCK IN"
    Else
      vs.TextMatrix(3, 0) = "Binder Receive"
      vs.TextMatrix(3, 1) = 0
      vs.TextMatrix(3, 2) = "STOCK IN"
    End If
  
   
   '4
    
    qty = 0
   
    If RS.State = 1 Then RS.close
    RS.Open "select sum(Qty) from BookStock where convert(smalldatetime,Dates,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
    " BookCode='" & txtBkCode & "' and " & _
    " " & str_go & " and " & stringyear & "", con, adOpenKeyset
    If Not IsNull(RS(0)) Then
       qty = RS(0)
    End If
    
    
    
    If rs_kit.State = 1 Then rs_kit.close
    rs_kit.Open "select KITCODE from BOOKS_KIT  where BOOKCODE='" & txtBkCode & "'", con
    'If rs_kit.EOF = False Then
    While rs_kit.EOF = False
        If RS.State = 1 Then RS.close
        RS.Open "select sum(Qty) from BookStock where convert(smalldatetime,Dates,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
        " BOOKCODE='" & rs_kit.Fields("kitCODE").value & "' and " & _
        " " & str_go & " and " & stringyear & "", con, adOpenKeyset
        If Not IsNull(RS(0)) Then qty = qty + RS(0)
    'End If
    rs_kit.MoveNext
    Wend
    
    
    If qty > 0 Then
      vs.TextMatrix(4, 0) = "Stock Transfar"
      vs.TextMatrix(4, 1) = qty
      txtSTin = Val(txtSTin) + qty
      vs.TextMatrix(4, 2) = "STOCK IN"
    Else
      vs.TextMatrix(4, 0) = "Stock Transfar"
      vs.TextMatrix(4, 1) = 0
      vs.TextMatrix(4, 2) = "STOCK IN"
    End If
  
  '5
  qty = 0
   
   If RS.State = 1 Then RS.close
   RS.Open "select sum(Balance) from BookOpening where BOOKCODE='" & txtBkCode & "' and Godown='" & txtGodown & "'", con, adOpenKeyset
   
   If Not IsNull(RS(0)) Then
      vs.TextMatrix(5, 0) = "Opening Book"
      vs.TextMatrix(5, 1) = RS(0)
      txtSTin = Val(txtSTin) + RS(0)
      vs.TextMatrix(5, 2) = "STOCK IN"
    Else
      vs.TextMatrix(5, 0) = "Opening Book"
      vs.TextMatrix(5, 1) = 0
      vs.TextMatrix(5, 2) = "STOCK IN"
    End If
   
   '-------- end stock in-------------------------------
   
   Dim godown_issue As Long
   
    If RS.State = 1 Then RS.close
    RS.Open "select sum(Qty) from SaleRegister where convert(smalldatetime,invoiceDate,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
    " BookCode='" & txtBkCode & "' and " & _
    " " & str1 & " and " & stringyear & "", con, adOpenKeyset
    If Not IsNull(RS(0)) Then
       godown_issue = RS(0)
       txtStockOut = godown_issue
    Else
      godown_issue = 0
    End If
    
  
    
    If RS.State = 1 Then RS.close
    RS.Open "select sum(QUANTITY) from SaleRegister_Free where convert(smalldatetime,invoiceDate,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
    " BookCode='" & txtBkCode & "' and " & _
    " " & str1 & " and " & stringyear & "", con, adOpenKeyset
    If Not IsNull(RS(0)) Then
       godown_issue = godown_issue + RS(0)
       txtStockOut = Val(txtStockOut) + RS(0)
    End If
   
   '6
    If godown_issue > 0 Then
      vs.TextMatrix(6, 0) = "sale"
      vs.TextMatrix(6, 1) = godown_issue
      vs.TextMatrix(6, 2) = "STOCK OUT"
    Else
      vs.TextMatrix(6, 0) = "sale"
      vs.TextMatrix(6, 1) = 0
      vs.TextMatrix(6, 2) = "STOCK OUT"
    End If
    

   
   '7
   
    godown_issue = 0
    
    If RS.State = 1 Then RS.close
    RS.Open "select sum(Qty) from SpecimenRegister where convert(smalldatetime,invoicedate,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
    " BookCode='" & txtBkCode & "' and " & _
    " " & str1 & " and " & stringyear & "", con, adOpenKeyset
    If Not IsNull(RS(0)) Then
       godown_issue = godown_issue + RS(0)
       txtStockOut = Val(txtStockOut) + RS(0)
    End If
    
    
    ''------------Kit Qry (SpecimenRegister)-----------------------------------------
    If RS.State = 1 Then RS.close
    RS.Open "select sum(QUANTITY) from SPRegister_Free where convert(smalldatetime,invoicedate,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
    " BookCode='" & txtBkCode & "' and " & _
    " " & str1 & " and " & stringyear & "", con, adOpenKeyset
    If Not IsNull(RS(0)) Then
       godown_issue = godown_issue + RS(0)
       txtStockOut = Val(txtStockOut) + RS(0)
    End If
   
    If godown_issue > 0 Then
      vs.TextMatrix(7, 0) = "Specimen"
      vs.TextMatrix(7, 1) = godown_issue
      vs.TextMatrix(7, 2) = "STOCK OUT"
    Else
      vs.TextMatrix(7, 0) = "Specimen"
      vs.TextMatrix(7, 1) = 0
      vs.TextMatrix(7, 2) = "STOCK OUT"
    End If
   
   
   '8
   godown_issue = 0
   
   If RS.State = 1 Then RS.close
   RS.Open "select netbook from BinderIssueRegister where convert(smalldatetime,invoiceDate,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
   " Book_Code='" & txtBkCode & "' and " & _
   " " & str1 & " and " & stringyear & "", con, adOpenKeyset
   If RS.EOF = False Then
   While RS.EOF = False
      godown_issue = godown_issue + RS(0)
      txtStockOut = Val(txtStockOut) + RS(0)
      
      RS.MoveNext
   Wend
   End If

    If godown_issue > 0 Then
      vs.TextMatrix(8, 0) = "Binder Issue"
      vs.TextMatrix(8, 1) = godown_issue
      vs.TextMatrix(8, 2) = "STOCK OUT"
    Else
      vs.TextMatrix(8, 0) = "Binder Issue"
      vs.TextMatrix(8, 1) = 0
      vs.TextMatrix(8, 2) = "STOCK OUT"
    End If
   


''''Stock Transfer
''''Issue Qty
'9
'

godown_issue = 0

If RS.State = 1 Then RS.close
RS.Open "select sum(Qty) from BookStock where convert(smalldatetime,Dates,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
" BOOKCODE='" & txtBkCode & "' and " & _
" " & str_go_issue & " and " & stringyear & "", con, adOpenKeyset
If Not IsNull(RS(0)) Then
   godown_issue = godown_issue + RS(0)
   txtStockOut = Val(txtStockOut) + RS(0)
End If




    


'
'''------------Kit Qry (Stock Trans)-----------------------------------------


If RS.State = 1 Then RS.close
RS.Open "select sum(Qty) from BookStock_free where convert(smalldatetime,Dates,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
" BOOKCODE='" & txtBkCode & "' and " & _
" godown='" & txtGodown & "' and " & stringyear & "", con, adOpenKeyset
If Not IsNull(RS(0)) Then
   godown_issue = godown_issue + RS(0)
   txtStockOut = Val(txtStockOut) + RS(0)
End If
   
   
If godown_issue > 0 Then
      vs.TextMatrix(9, 0) = "Book Stock"
      vs.TextMatrix(9, 1) = godown_issue
      vs.TextMatrix(9, 2) = "STOCK OUT"
Else
      vs.TextMatrix(9, 0) = "Book Stock"
      vs.TextMatrix(9, 1) = 0
      vs.TextMatrix(9, 2) = "STOCK OUT"
End If

 
 
   '====================================================
    godown_issue = 0
    If RS.State = 1 Then RS.close
    RS.Open "select sum(Qty) from CashSaleRegister where convert(smalldatetime,invoiceDate,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
    " BookCode='" & txtBkCode & "' and " & _
    " " & str1 & " and " & stringyear & "", con, adOpenKeyset
    If Not IsNull(RS(0)) Then
       godown_issue = RS(0)
       
       txtStockOut = Val(txtStockOut) + RS(0)
    Else
      godown_issue = 0
    End If
   '6
    If godown_issue > 0 Then
      vs.TextMatrix(10, 0) = "Cash sale"
      vs.TextMatrix(10, 1) = godown_issue
      vs.TextMatrix(10, 2) = "STOCK OUT"
    Else
      vs.TextMatrix(10, 0) = "Cash sale"
      vs.TextMatrix(10, 1) = 0
      vs.TextMatrix(10, 2) = "STOCK OUT"
    End If

   
   '====================================================
   

   
   vs.ColWidth(0) = 2000
   vs.ColWidth(1) = 1000
   vs.ColWidth(2) = 2000
   vs.ColWidth(3) = 3000

   txtClosingBk = (Val(txtSTin) - Val(txtStockOut))
  
  
Exit Sub

aa12:

MsgBox "" & err.DESCRIPTION
  
End Sub
Private Sub vs_Click()

Dim rsf As New ADODB.Recordset

Screen.MousePointer = vbHourglass
vs_billwise.Visible = True

If vs.TextMatrix(vs.RowSel, 0) = "Sale Return" Then
   str_ = "SELECT INVOICENO,sum(QUANTITY) as Qty FROM SaleRetBQry_forgross where BOOKCODE='" & txtBkCode & "'  group by INVOICENO order by INVOICENO"
ElseIf vs.TextMatrix(vs.RowSel, 0) = "sale" Then
   str_ = "SELECT INVOICENO,sum(QUANTITY) as Qty FROM invoiceBQry where BOOKCODE='" & txtBkCode & "'  group by INVOICENO order by INVOICENO"
ElseIf vs.TextMatrix(vs.RowSel, 0) = "Specimen" Then
   str_ = "SELECT INVOICENO,sum(QUANTITY) as Qty FROM invoiceSPBQry where BOOKCODE='" & txtBkCode & "'  group by INVOICENO order by INVOICENO"
ElseIf vs.TextMatrix(vs.RowSel, 0) = "Specimen Return" Then
   str_ = "SELECT INVOICENO,sum(QUANTITY) as Qty FROM invoiceSPRETBQry where BOOKCODE='" & txtBkCode & "'  group by INVOICENO order by INVOICENO"
ElseIf vs.TextMatrix(vs.RowSel, 0) = "Cash sale" Then
   str_ = "SELECT INVOICENO,sum(QUANTITY) as Qty FROM cashBQry where BOOKCODE='" & txtBkCode & "'  group by INVOICENO order by INVOICENO"

ElseIf vs.TextMatrix(vs.RowSel, 0) = "Binder Receive" Then
   str_ = "SELECT INVOICENO,sum(NetBook) as Qty FROM BinderReceiveBQry where BOOK_CODE='" & txtBkCode & "'  group by INVOICENO order by INVOICENO"

ElseIf vs.TextMatrix(vs.RowSel, 0) = "Binder Issue" Then
   str_ = "SELECT INVOICENO,sum(NetBook) as Qty FROM BinderIssueBQry where BOOK_CODE='" & txtBkCode & "'  group by INVOICENO order by INVOICENO"


End If


If Not IsEmpty(str_) Then
    If rsf.State = 1 Then rsf.close
    rsf.Open str_, con, adOpenDynamic, adLockOptimistic
    Set vs_billwise.DataSource = rsf
Else
    vs_billwise.Clear
End If


Screen.MousePointer = vbDefault

End Sub

