VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmTaxableRG1 
   BackColor       =   &H00C0E0FF&
   ClientHeight    =   8775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16995
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8775
   ScaleWidth      =   16995
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport cr 
      Left            =   9600
      Top             =   780
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "&View"
      Height          =   495
      Left            =   3900
      TabIndex        =   4
      Top             =   660
      Width           =   1215
   End
   Begin VB.CommandButton Command1_print 
      Caption         =   "&Print"
      Height          =   495
      Left            =   5220
      TabIndex        =   1
      Top             =   660
      Width           =   1215
   End
   Begin VB.CommandButton Command2_exit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   6540
      TabIndex        =   0
      Top             =   660
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dt2 
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      Top             =   660
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Format          =   19660801
      CurrentDate     =   41274
   End
   Begin MSComCtl2.DTPicker dt1 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   660
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      _Version        =   393216
      Format          =   19660801
      CurrentDate     =   41274
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   8235
      Left            =   120
      TabIndex        =   5
      Top             =   1260
      Width           =   18135
      _cx             =   31988
      _cy             =   14526
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   16710321
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
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   2
      FixedCols       =   0
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
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "Taxable RG -1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   60
      Width           =   18015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   600
      Width           =   555
   End
End
Attribute VB_Name = "frmTaxableRG1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim billNoForm_to As String
Dim rss As New ADODB.Recordset

Sub vsINI()
   
   
 vs.Cols = 20
 
 vs.MergeCells = flexMergeFixedOnly
 vs.WordWrap = True
 
  With vs
   
   .RowHeight(1) = 800
   
   vs.MergeRow(0) = True
   vs.MergeCol(0) = True
   vs.MergeCol(1) = True
   vs.MergeCol(2) = True
   
   .MergeCol(9) = True
   .MergeCol(10) = True
   .MergeCol(11) = True
   .MergeCol(12) = True
   
   .MergeCol(13) = True
   .MergeCol(14) = True
   .MergeCol(15) = True
   
    .MergeCol(17) = True
   .MergeCol(18) = True
   .MergeCol(19) = True
   
  
   
   
  
    
   For K = 3 To 8
      vs.ColAlignment(K) = flexAlignCenterCenter
   Next
  
  
     For K = 13 To 16
      vs.ColAlignment(K) = flexAlignCenterCenter
   Next

   
   .MergeCells = flexMergeFree
   .MergeRow(0) = True
   
   .MergeCells = flexMergeFree
   .MergeCol(0) = True
   .MergeCol(1) = True
   .MergeCol(2) = True
   .MergeCol(3) = True
   
   
   vs.TextMatrix(0, 0) = "Date"
   vs.TextMatrix(0, 1) = "Opening Balance"
   vs.TextMatrix(0, 2) = "Quantity Manufactured"
   vs.TextMatrix(0, 3) = "Total (2+3)"
   
   vs.TextMatrix(0, 4) = "For Home Use (Domestic Sales)"
   vs.TextMatrix(0, 5) = "For Home Use (Domestic Sales)"
   
   vs.TextMatrix(0, 6) = "For Export Under Claim for Rebate of Duty"
   vs.TextMatrix(0, 7) = "For Export Under Claim for Rebate of Duty"
   vs.TextMatrix(0, 8) = "For Export Under Claim for Rebate of Duty"
   
   .TextMatrix(0, 9) = "For Other Factories or Warehouse Under Bond"
   
   
   .TextMatrix(0, 10) = "For other Purpose"
   .TextMatrix(0, 11) = "For other Purpose"
   .TextMatrix(0, 12) = "For other Purpose"


   .TextMatrix(0, 13) = "Duty Payable & Paid"
   .TextMatrix(0, 14) = "Duty Payable & Paid"
   .TextMatrix(0, 15) = "Duty Payable & Paid"

  
   .TextMatrix(0, 16) = "Total Duty"
  
   .TextMatrix(0, 17) = "closing balance"
   .TextMatrix(0, 18) = "closing balance"
   
   .TextMatrix(0, 19) = "Remark(Bill No.)"
   
   
   
   
   
   vs.TextMatrix(1, 0) = "Date"
   vs.TextMatrix(1, 1) = "Opening Balance"
   vs.TextMatrix(1, 2) = "Quantity Manufactured"
   vs.TextMatrix(1, 3) = "Total (2+3)"
   
   vs.TextMatrix(1, 4) = "Quantity"
   vs.TextMatrix(1, 5) = "Value"
   
   .TextMatrix(1, 6) = "Quantity"
   .TextMatrix(1, 7) = "Value"
   .TextMatrix(1, 8) = "For Export Under Bond"
   
   
   .TextMatrix(1, 9) = "For Other Factories or Warehouse Under Bond"
   
   
   .TextMatrix(1, 10) = "Purpose"
   .TextMatrix(1, 11) = "Quantity"
   .TextMatrix(1, 12) = "Rate"
   
   
   
   .TextMatrix(1, 13) = "BED @2%"
   .TextMatrix(1, 14) = "Edu. Cess @0.02%"
   .TextMatrix(1, 15) = "S.H.Edu. Cess @0.01%"
   
   .TextMatrix(1, 16) = "Total Duty"
   
   
   .TextMatrix(1, 17) = "In Finishing Room"
   .TextMatrix(1, 18) = "In Bonded Store Room"
   
   .TextMatrix(1, 19) = "Remark(Bill No.)"

   
  
 
 End With
  
  
 For i = 1 To 19
    vs.ColWidth(i) = 1500
 Next




   
   
End Sub
Function billDesc(dt As Date) As String
  
  Dim str1, str2 As String
  
  str1 = ""
  str2 = ""
  
  If rss.State = 1 Then rss.Close
  rss.Open "select INVOICEno from salesQry_withEducess where " & stringyear & " and " & _
  " convert(smalldatetime,INVOICEDATE,103)=convert(smalldatetime,'" & dt & "',103)" & _
  " ", CON, adOpenKeyset, adLockReadOnly
  While rss.EOF = False
     
     If str1 = "" Then
        str1 = rss(0) & " to "
     End If
     
     str2 = rss(0)
     
     rss.MoveNext
  Wend

 If str1 = "" Then
    str1 = rss(0)
 Else
    str1 = str1 & "" & str2
 End If


 billDesc = str1

End Function
Private Sub cmdshow_Click()


Dim rs1 As New ADODB.Recordset

Dim Rec_book, sale_book As Double

Dim balNextDay As Double


balNextDay = 0

Rec_book = 0
sale_book = 0

vs.Clear
vsINI



vs.MergeCells = flexMergeFixedOnly
K1 = 2



If rs.State = 1 Then rs.Close
rs.Open "select INVOICEDATE,sum(GAMOUNT),sum(Qty),sum(aexp6am),sum(aexp7am) from salesQry_withEducess where " & stringyear & " and " & _
" convert(smalldatetime,INVOICEDATE,103)>=convert(smalldatetime,'" & dt1.Value & "',103) and " & _
" convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & dt2.Value & "',103)" & _
" group by INVOICEDATE", CON, adOpenKeyset, adLockReadOnly


For i = 1 To rs.RecordCount

  
  
  Rec_book = 0
  sale_book = 0
  
  
  'Fatch Received Data
  
  If rs1.State = 1 Then rs1.Close
  rs1.Open "select sum(QUANTITY) from ProductReceipt where " & stringyear & " and " & _
  " convert(smalldatetime,RecDate,103)<convert(smalldatetime,'" & rs(0) & "',103)" & _
  " ", CON, adOpenKeyset, adLockReadOnly
  If Not IsNull(rs1(0)) Then
     Rec_book = rs1(0)
  End If


  If rs1.State = 1 Then rs1.Close
  rs1.Open "select sum(QTY) from salesQry_withEducess where " & stringyear & " and " & _
  " convert(smalldatetime,INVOICEDATE,103)<convert(smalldatetime,'" & rs(0) & "',103)" & _
  " ", CON, adOpenKeyset, adLockReadOnly
  If Not IsNull(rs1(0)) Then
     sale_book = rs1(0)
  End If



  vs.TextMatrix(K1, 0) = rs!InvoiceDate
  
  
  'If K1 >= 2 Then
     vs.TextMatrix(K1, 1) = ((Rec_book - sale_book) + balNextDay)
  'Else
  '   vs.TextMatrix(K1, 1) = (Rec_book - sale_book)
  'End If
  
  
  If rs1.State = 1 Then rs1.Close
  rs1.Open "select sum(QUANTITY) from ProductReceipt where " & stringyear & " and " & _
  " convert(smalldatetime,recdate,103)=convert(smalldatetime,'" & rs(0) & "',103) " & _
  " ", CON, adOpenKeyset, adLockReadOnly
  If Not IsNull(rs1(0)) Then
     vs.TextMatrix(K1, 2) = rs1(0)
  End If
  
  
  vs.TextMatrix(K1, 3) = (Val(vs.TextMatrix(K1, 1)) + Val(vs.TextMatrix(K1, 2)))
  
  vs.TextMatrix(K1, 4) = rs(2)
  
  vs.TextMatrix(K1, 5) = rs(1)
  
  vs.TextMatrix(K1, 14) = rs(3)                  'aexp6am
  vs.TextMatrix(K1, 15) = rs(4)
  
  vs.TextMatrix(K1, 16) = (rs(3) + rs(4))
  
  
  
  'GAMOUNT
  
  
  
  
  If K1 >= 2 Then
     balNextDay = (Val(vs.TextMatrix(K1, 3)) - Val(vs.TextMatrix(K1, 4)))
     vs.TextMatrix(K1, 18) = balNextDay
  End If
 
  vs.TextMatrix(K1, 19) = billDesc(rs(0))
  
  
 
  
  
  
  K1 = K1 + 1
  rs.MoveNext
  
Next


 For i = 0 To 18
    vs.ColWidth(i) = 1500
 Next

vs.ColWidth(11) = 1600
'vs.ColWidth(21) = 1700



End Sub
Private Sub Command1_print_Click()




CON.Execute "delete from taxable_RG1"


Dim bed, Ecess, SEcess, TotalDuty As Double

bed = 0: Ecess = 0: SEcess = 0: TotalDuty = 0



If rs.State = 1 Then rs.Close
rs.Open "select * from taxable_RG1", CON, adOpenDynamic, adLockOptimistic

For i = 2 To vs.Rows - 1

If vs.TextMatrix(i, 0) <> "" Then

rs.addNew
rs!InvDate = vs.TextMatrix(i, 0)

rs!opBal = vs.TextMatrix(i, 1)
rs!Qty_mfg = vs.TextMatrix(i, 2)
rs!Total = vs.TextMatrix(i, 3)
rs!Sale_Qty = vs.TextMatrix(i, 4)
rs!Sale_Amount = vs.TextMatrix(i, 5)
rs!Ex_Qty = vs.TextMatrix(i, 6)
rs!Ex_Value = vs.TextMatrix(i, 7)
rs!UnderBond = vs.TextMatrix(i, 8)
rs!Warehouse_ubond = vs.TextMatrix(i, 9)
rs!Perpose = vs.TextMatrix(i, 10)
rs!P_qty = vs.TextMatrix(i, 11)
rs!P_rate = vs.TextMatrix(i, 12)

rs!bed = vs.TextMatrix(i, 13)
bed = bed + Val(vs.TextMatrix(i, 13))



rs!Ecess = vs.TextMatrix(i, 14)
Ecess = Ecess + Val(vs.TextMatrix(i, 14))

rs!SEcess = vs.TextMatrix(i, 15)
SEcess = SEcess + Val(vs.TextMatrix(i, 15))

rs!TotalDuty = vs.TextMatrix(i, 16)
TotalDuty = TotalDuty + Val(vs.TextMatrix(i, 16))

rs!Froom = vs.TextMatrix(i, 17)
rs!StoreRoom = vs.TextMatrix(i, 18)
rs!billno = vs.TextMatrix(i, 19)

      
rs.Update

End If




Next
      
      
rs.addNew
rs!Perpose = "Total : "

rs!bed = Round(bed, 0)
rs!Ecess = Round(Ecess, 0)
rs!SEcess = Round(SEcess, 0)
rs!TotalDuty = Round(TotalDuty, 0)
rs.Update

DoEvents
DoEvents
DoEvents
DoEvents
    
cr.Reset
cr.Connect = constr
cr.ReportFileName = App.Path & "\REPORTS\taxableRg1.rpt"

cr.Formulas(0) = "fromdate='" & dt1.Value & "'"
cr.Formulas(1) = "todate='" & dt2.Value & "'"

cr.WindowState = crptMaximized
cr.WindowShowPrintSetupBtn = True
cr.WindowShowRefreshBtn = True
cr.Action = 1





End Sub
Private Sub Command2_exit_Click()
Unload Me
End Sub

Private Sub Form_Load()
vsINI

dt1.Value = Date
dt2.Value = Date
 

End Sub

