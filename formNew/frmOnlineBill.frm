VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmOnlineBill 
   Caption         =   "Bill Print"
   ClientHeight    =   5796
   ClientLeft      =   60
   ClientTop       =   396
   ClientWidth     =   10800
   Icon            =   "frmOnlineBill.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   5796
   ScaleWidth      =   10800
   Begin VB.TextBox txtThrough 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   0
      Top             =   675
      Width           =   1725
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
      Height          =   660
      Left            =   3915
      Picture         =   "frmOnlineBill.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4725
      Width           =   1395
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
      Height          =   660
      Left            =   5445
      Picture         =   "frmOnlineBill.frx":0BF0
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4725
      Width           =   1410
   End
   Begin VB.TextBox txtNet 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   9180
      TabIndex        =   14
      Top             =   5085
      Width           =   1095
   End
   Begin VB.TextBox txtAddDisRate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8685
      TabIndex        =   2
      Top             =   4365
      Width           =   465
   End
   Begin VB.TextBox txtAdddisAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   9180
      TabIndex        =   12
      Top             =   4365
      Width           =   1095
   End
   Begin VB.TextBox txtdisAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9180
      TabIndex        =   3
      Top             =   4725
      Width           =   1095
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   8550
      TabIndex        =   9
      Top             =   4005
      Width           =   1725
   End
   Begin VB.TextBox txtOrderNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Top             =   135
      Width           =   1725
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   2325
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   10770
      _cx             =   18997
      _cy             =   4101
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   16761024
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16777215
      ForeColorSel    =   0
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
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   300
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   5
      RowHeightMin    =   400
      RowHeightMax    =   400
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmOnlineBill.frx":17D4
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
   Begin VB.Label Label7 
      BackColor       =   &H00FFFF00&
      Height          =   60
      Left            =   0
      TabIndex        =   16
      Top             =   495
      Width           =   10815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice No"
      Height          =   285
      Left            =   270
      TabIndex        =   15
      Top             =   720
      Width           =   1275
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Net Amount"
      Height          =   285
      Left            =   7380
      TabIndex        =   13
      Top             =   5130
      Width           =   1275
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Add. Discount (%)"
      Height          =   285
      Left            =   7380
      TabIndex        =   11
      Top             =   4410
      Width           =   1320
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Add/Less"
      Height          =   285
      Left            =   7380
      TabIndex        =   10
      Top             =   4770
      Width           =   960
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   285
      Left            =   7380
      TabIndex        =   8
      Top             =   4050
      Width           =   960
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Order No"
      Height          =   285
      Left            =   270
      TabIndex        =   6
      Top             =   135
      Width           =   960
   End
End
Attribute VB_Name = "frmOnlineBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExit_12_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()

Dim discount, AddDiscount, discountAmt, AddDiscountAmt As Double
DSNNew

'discount = IIf(txtDisRate.Text = "", 0, txtDisRate.Text)
AddDiscount = IIf(txtAddDisRate = "", 0, txtAddDisRate)

discountAmt = IIf(txtdisAmt = "", 0, txtdisAmt)
AddDiscountAmt = IIf(txtAdddisAmt = "", 0, txtAdddisAmt)

con.Execute "update ordera set onlineAmt=" & txtAmount.Text & ",discountAmt=" & discountAmt & "" & _
",Adddiscount=" & AddDiscount & ",AdddiscountAmt=" & AddDiscountAmt & ",OnlineNetAmt=" & Val(txtNet.Text) & ",through_='" & txtThrough.Text & "' where invoiceno=" & txtOrderNo & ""


For I = 1 To vs.rows - 1
       If vs.TextMatrix(I, 1) <> "" Then
         con.Execute "update orderb set onlineAmt=" & vs.TextMatrix(I, 7) & ",onlineDis=" & vs.TextMatrix(I, 5) & ",onlineDisAmt=" & vs.TextMatrix(I, 6) & "  where invoiceno=" & txtOrderNo & " and bookcode='" & vs.TextMatrix(I, 0) & "'"
       End If
Next


If MsgBox("Want to print ? ", vbQuestion + vbYesNo) = vbYes Then
 

frmINVOrder.cr.Reset
frmINVOrder.cr.ReportFileName = rptPath & "/onlineBill.rpt"
frmINVOrder.cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
frmINVOrder.cr.ReplaceSelectionFormula "{ordera.invoiceno}=" & txtOrderNo & ""
frmINVOrder.cr.WindowShowPrintSetupBtn = True
frmINVOrder.cr.WindowShowRefreshBtn = True
frmINVOrder.cr.WindowMaxButton = True
frmINVOrder.cr.WindowState = crptMaximized
frmINVOrder.cr.Action = 1

End If

End Sub
Sub setVSWidth()
      vs.FormatString = "Code|BookName|Quantity|Rate|Amount|Dis.(%)|Dis.Amt|Net Amount"
      
      vs.ColWidth(0) = 800
      vs.ColWidth(1) = 3000
      vs.ColWidth(2) = 1000
      vs.ColWidth(3) = 1000
      vs.ColWidth(4) = 1000
      vs.ColWidth(5) = 1200
      vs.ColWidth(6) = 1200
      
      
End Sub
Private Sub Form_Load()
    
    Me.Top = 2000
    Me.Left = 500
    Me.Width = 10900
    Me.Height = 6000
    
    txtOrderNo = frmINVOrder.txtOrderNo.Text
    txtAmount.Text = Format(frmINVOrder.txtOnlineAmt.Text, "0.00")
    
    '========================================================
    vs.rows = 1
    If RS.State = 1 Then RS.close
    RS.Open " SELECT   ORDERB.INVOICENO, ORDERB.BOOKCODE, BOOKS.BOOKNAME, ORDERB.QUANTITY, ORDERB.RATE, orderb.onlineDis,orderb.onlineDisAmt,orderb.onlineAmt " & _
    " FROM    ORDERB INNER JOIN BOOKS ON ORDERB.BOOKCODE = BOOKS.BOOKCODE  where ORDERB.INVOICENO=" & txtOrderNo & " order by   orderb.sno", con
    For I = 1 To RS.RecordCount
    
    vs.rows = vs.rows + 1
    
    vs.TextMatrix(I, 0) = RS!Bookcode
    vs.TextMatrix(I, 1) = RS!Bookname
    vs.TextMatrix(I, 2) = RS!QUANTITY
    vs.TextMatrix(I, 3) = RS!rate
    vs.TextMatrix(I, 4) = (RS!QUANTITY * RS!rate)
    
    If IsNull(RS!onlineDis) Then
       vs.TextMatrix(I, 5) = 0
    Else
       vs.TextMatrix(I, 5) = RS!onlineDis
    End If
    
    If IsNull(RS!onlineDisAmt) Then
       vs.TextMatrix(I, 6) = 0
    Else
       vs.TextMatrix(I, 6) = RS!onlineDisAmt
    End If
    
    If IsNull(RS!onlineAmt) Then
       vs.TextMatrix(I, 7) = 0
    Else
       vs.TextMatrix(I, 7) = RS!onlineAmt
    End If
    
    
    'If Val(vs.TextMatrix(I, 5)) > 0 Then
    '   vs.TextMatrix(I, 7) = Val(vs.TextMatrix(I, 4)) - Val(vs.TextMatrix(I, 5))
    'Else
    '   vs.TextMatrix(I, 7) = Val(vs.TextMatrix(I, 4))
    'End If
    
    RS.MoveNext
    
    Next
    
    Total
    
    
    setVSWidth
    vs.rows = vs.rows + 1
    '========================================================
    If rs1.State = 1 Then rs1.close
    rs1.Open "select  onlineAmt,discount,discountAmt,Adddiscount,AdddiscountAmt,OnlineNetAmt,through_ from ordera where invoiceno=" & txtOrderNo & "", con
    If rs1.EOF = False Then
       'txtAmount.Text = rs1!onlineAmt & ""
       
       txtdisAmt.Text = rs1!discountAmt & ""
       
       txtAddDisRate.Text = rs1!AddDiscount & ""
       txtAdddisAmt.Text = rs1!AddDiscountAmt & ""
       
       txtNet.Text = rs1!OnlineNetAmt & ""
       txtThrough = rs1!through_ & ""
    
    End If

    
End Sub
Sub Total()
txtAmount.Text = 0

For I = 1 To vs.rows - 1

If vs.TextMatrix(I, 7) <> "" Then
   txtAmount = Val(txtAmount) + Val(vs.TextMatrix(I, 7))
End If

Next

End Sub
Sub calc()
     
    Dim addAmt As Double
    addAmt = 0


    addAmt = Val(txtAmount.Text)
    
    If Val(txtAddDisRate) > 0 Then
       txtAdddisAmt.Text = addAmt * Val(txtAddDisRate.Text) / 100
    Else
       txtAdddisAmt.Text = 0
    End If
    
    
    txtAdddisAmt.Text = Round(txtAdddisAmt.Text, 2)
    
        
    txtNet.Text = (Val(txtAmount.Text) - Val(txtAdddisAmt.Text) + Val(txtdisAmt.Text))
    txtNet.Text = Format(Round(txtNet.Text, 2), "0.00")
     
End Sub

Private Sub txtAddDisRate_Change()
calc
End Sub

Private Sub txtDisRate_Change()
calc
End Sub

Private Sub txtdisAmt_Change()
calc
End Sub

Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)
     
     Dim discount As Double
     discount = 0
     
     If KeyCode = 13 Then
     
     If vs.Col = 5 Then
     
        If (Val(vs.TextMatrix(vs.RowSel, 5)) > 0) Then
            discount = Round(Val(vs.TextMatrix(vs.RowSel, 4)) * Val(vs.TextMatrix(vs.RowSel, 5)) / 100, 2)
            vs.TextMatrix(vs.RowSel, 6) = discount
            vs.TextMatrix(vs.RowSel, 7) = (Val(vs.TextMatrix(vs.RowSel, 4)) - Val(vs.TextMatrix(vs.RowSel, 6)))
        Else
             vs.TextMatrix(vs.RowSel, 7) = Val(vs.TextMatrix(vs.RowSel, 4))
        End If
        SendKeys "{down}"
         Total
         
     End If
     
     End If
     
End Sub
