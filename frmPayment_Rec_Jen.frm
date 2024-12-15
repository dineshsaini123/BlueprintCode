VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmPayment_Rec_Jen 
   Caption         =   "Voucher Details"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14085
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8640
   ScaleWidth      =   14085
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   435
      Left            =   8040
      TabIndex        =   10
      Top             =   60
      Width           =   1035
   End
   Begin VB.TextBox txtDr 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   10860
      TabIndex        =   9
      Top             =   8100
      Width           =   1515
   End
   Begin VB.TextBox txtCr 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   12480
      TabIndex        =   8
      Top             =   8100
      Width           =   1395
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   435
      Left            =   6960
      TabIndex        =   7
      Top             =   60
      Width           =   1035
   End
   Begin VB.ComboBox vtype 
      Height          =   315
      ItemData        =   "frmPayment_Rec_Jen.frx":0000
      Left            =   5400
      List            =   "frmPayment_Rec_Jen.frx":0010
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   120
      Width           =   1485
   End
   Begin MSComCtl2.DTPicker toDate 
      Height          =   315
      Left            =   2520
      TabIndex        =   2
      Top             =   180
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      _Version        =   393216
      Format          =   71827457
      CurrentDate     =   40617
   End
   Begin MSComCtl2.DTPicker fromDate 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   180
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Format          =   71827457
      CurrentDate     =   40617
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   7155
      Left            =   300
      TabIndex        =   0
      Top             =   840
      Width           =   13875
      _cx             =   24474
      _cy             =   12621
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      BackColorBkg    =   -2147483636
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
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
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
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   540
      Width           =   13935
   End
   Begin VB.Label Label2 
      Caption         =   "Voucher Type :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4020
      TabIndex        =   4
      Top             =   180
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1800
      TabIndex        =   3
      Top             =   240
      Width           =   915
   End
End
Attribute VB_Name = "frmPayment_Rec_Jen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub view()
   
   Dim i As Integer
   Dim dr, CR As Double
   
   Dim vnuber, vdate As String
   Dim cat As Integer
   
   
   vs.Cols = 10
   
   dr = 0
   CR = 0
   cat = 1
   
   
   
   
   
   vs.FormatString = "SN|VoucherDate|VNo|Genledger|Subledger|DESCRIPTION|Amount(Dr.)|Amount(Cr.)|"
   
   i = 1
   vnuber = ""
   vdate = ""
  
   
   If rs.State = 1 Then rs.Close
   
   If Len(vtype.Text) > 2 Then
        rs.Open "select * from vouchers WHERE" & _
        " (convert(smalldatetime,voucherdate,103)>=convert(smalldatetime,'" + Trim(fromdate.Value) + "',103)   and convert(smalldatetime,voucherdate,103)<=convert(smalldatetime,'" + Trim(todate.Value) + "',103)) and " & stringyear & " order by voucherDate,VoucherNumber,vsno", CON
   Else
        rs.Open "select * from vouchers WHERE VoucherTYPE='" & vtype.Text & "' AND " & _
        " (convert(smalldatetime,voucherdate,103)>=convert(smalldatetime,'" + Trim(fromdate.Value) + "',103)   and convert(smalldatetime,voucherdate,103)<=convert(smalldatetime,'" + Trim(todate.Value) + "',103)) and " & stringyear & " order by voucherDate,VoucherNumber,vsno", CON
   End If
   
   
   While rs.EOF = False
  
   vs.Rows = rs.RecordCount + 1
   
   vs.TextMatrix(i, 0) = i
   
   
   vs.TextMatrix(i, 1) = rs!VoucherDate
   vs.TextMatrix(i, 2) = rs!VOUCHERNUMBER
   vs.TextMatrix(i, 3) = rs!Genledger
   vs.TextMatrix(i, 4) = rs!subledger
   vs.TextMatrix(i, 5) = rs!DESCRIPTION
   
   If (vnuber <> "" Or vdate <> "") Then
            
      If (vnuber = rs!VOUCHERNUMBER And vdate = rs!VoucherDate) Then
      Else
         cat = cat + 1
      End If
      
   End If
   
   vnuber = rs!VOUCHERNUMBER
   vdate = rs!VoucherDate
   
   
   If rs!DEBITORCREDIT = "D" Then
      dr = dr + rs!amount
      vs.TextMatrix(i, 6) = rs!amount
      vs.TextMatrix(i, 7) = 0
   Else
      CR = CR + rs!amount
      vs.TextMatrix(i, 6) = 0
      vs.TextMatrix(i, 7) = rs!amount
   End If
    
   vs.TextMatrix(i, 8) = cat
   vs.TextMatrix(i, 9) = rs!VoucherType
    
    
    
   i = i + 1
   rs.MoveNext
   Wend
   
   vs.ColWidth(0) = 500
   vs.ColWidth(1) = 1000
   vs.ColWidth(2) = 500
   vs.ColWidth(3) = 2900
   vs.ColWidth(4) = 3000
   vs.ColWidth(5) = 3100
   
   vs.ColWidth(6) = 1200
   vs.ColWidth(7) = 1200
   
   
   txtDr.Text = dr
   txtCr.Text = CR
   
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()
  
   
  Screen.MousePointer = vbHourglass
  
  CON.Execute "delete from vdetails where len(sn)>0"
  
  If rs.State = 1 Then rs.Close
  rs.Open "select * from vdetails", CON
  For i = 1 To vs.Rows - 1
     CON.Execute "insert into vdetails(SN,Vdate,VNo,Gledger,Sledger,Narr,Amt1,Amt2,cat,vtype) " & _
     "values('" & vs.TextMatrix(i, 0) & "','" & vs.TextMatrix(i, 1) & "','" & vs.TextMatrix(i, 2) & "','" & vs.TextMatrix(i, 3) & "'," & _
     "'" & vs.TextMatrix(i, 4) & "','" & vs.TextMatrix(i, 5) & "','" & vs.TextMatrix(i, 6) & "','" & vs.TextMatrix(i, 7) & "','" & vs.TextMatrix(i, 8) & "','" & vs.TextMatrix(i, 9) & "')"
  Next
  
  
  
 If MsgBox("Want To Print ?", vbInformation + vbYesNo) = vbYes Then
  
  
    MainMenu.cr1.Reset
    MainMenu.cr1.Connect = constr
    MainMenu.cr1.ReportFileName = App.Path & "\REPORTS\VoucherDet.rpt"
    MainMenu.cr1.Formulas(0) = "fromdate='" & fromdate.Value & "'"
    MainMenu.cr1.Formulas(1) = "todate='" & todate.Value & "'"
    
    If vtype.Text = "J" Then
    MainMenu.cr1.Formulas(2) = "rtype='" & "Journal Voucher Details" & "'"
    ElseIf vtype.Text = "P" Then
    MainMenu.cr1.Formulas(2) = "rtype='" & "Payment Voucher Details" & "'"
    ElseIf vtype.Text = "R" Then
    MainMenu.cr1.Formulas(2) = "rtype='" & "Receipt Voucher Details" & "'"
    Else
    MainMenu.cr1.Formulas(2) = "rtype='" & "Day Book Register" & "'"
    End If
    
    MainMenu.cr1.WindowState = crptMaximized
    MainMenu.cr1.WindowShowPrintSetupBtn = True
    MainMenu.cr1.Action = 1

  
  End If
  
  Screen.MousePointer = vbDefault
  
End Sub

Private Sub vtype_Click()
view
End Sub
