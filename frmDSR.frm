VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmDSR 
   BackColor       =   &H00C0E0FF&
   Caption         =   "D.S.R.  Register"
   ClientHeight    =   9570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18330
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9570
   ScaleWidth      =   18330
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport cr 
      Left            =   16080
      Top             =   660
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command2_exit 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6540
      TabIndex        =   6
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1_print 
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5220
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dt2 
      Height          =   315
      Left            =   2160
      TabIndex        =   3
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Format          =   19726337
      CurrentDate     =   41274
   End
   Begin MSComCtl2.DTPicker dt1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      _Version        =   393216
      Format          =   19726337
      CurrentDate     =   41274
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "&View"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3900
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   1140
      Width           =   18135
      _cx             =   31988
      _cy             =   13150
      _ConvInfo       =   1
      Appearance      =   0
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
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "D.S.R. Register"
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
      Left            =   60
      TabIndex        =   7
      Top             =   0
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
      TabIndex        =   4
      Top             =   600
      Width           =   555
   End
End
Attribute VB_Name = "frmDSR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub vsINI()
   
   
   vs.Cols = 22
 
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
   
   .MergeCol(20) = True
   .MergeCol(21) = True
   
   
   
   .MergeCells = flexMergeFree
   .WordWrap = True
    
    
  ' vs.MergeCells = flexMergeFixedOnly
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

   
   
   vs.TextMatrix(0, 0) = "S.N."
   vs.TextMatrix(0, 1) = "Invoice Date"
   vs.TextMatrix(0, 2) = "Invoice No"
   vs.TextMatrix(0, 3) = "Sales"
   vs.TextMatrix(0, 4) = "Sales"
   vs.TextMatrix(0, 5) = "Sales"
   vs.TextMatrix(0, 6) = "Sales"
   vs.TextMatrix(0, 7) = "Sales"
   vs.TextMatrix(0, 8) = "Sales"
   .TextMatrix(0, 9) = "Total Sale Vaue (Excisable)"
   .TextMatrix(0, 10) = "Total Quantity (Excisable)"
   .TextMatrix(0, 11) = "Total Sale Vaue (Non-Excisable)"
   .TextMatrix(0, 12) = "Total Quantity(Non - Excisable)"

   .TextMatrix(0, 13) = "Centra Exice Duty"
   .TextMatrix(0, 14) = "Centra Exice Duty"
   .TextMatrix(0, 15) = "Centra Exice Duty"
   .TextMatrix(0, 16) = "Centra Exice Duty"
 
 
   .TextMatrix(0, 17) = "VAT"
   .TextMatrix(0, 18) = "CST"
   .TextMatrix(0, 19) = "Postate Exp."
   .TextMatrix(0, 20) = "Short & Excess"
   .TextMatrix(0, 21) = "Total Invoice Value"

   
   
   
   .TextMatrix(1, 0) = "S.N."
   .TextMatrix(1, 1) = "Invoice Date"
   .TextMatrix(1, 2) = "Invoice No"
   .TextMatrix(1, 3) = "Sales -Tax Invoice"
   .TextMatrix(1, 4) = "Sales -Sale Invoice"
   .TextMatrix(1, 5) = "Sale-Exempte (Non Taxable)"
   .TextMatrix(1, 6) = "Sales -Agst. Form-'C'"
   .TextMatrix(1, 7) = "Sales -Without Form-'C'"
   .TextMatrix(1, 8) = "Sales -Agst. Form-'H'"
   .TextMatrix(1, 9) = "Total Sale Vaue (Excisable)"
   .TextMatrix(1, 10) = "Total Quantity (Excisable)"
   .TextMatrix(1, 11) = "Total Sale Vaue (Non-Excisable)"
   .TextMatrix(1, 12) = "Total Quantity(Non - Excisable)"
   
   .TextMatrix(1, 13) = "CENVAT"
   .TextMatrix(1, 14) = "Edu. Cess"
   .TextMatrix(1, 15) = "S.H. Edu. Cess"
   .TextMatrix(1, 16) = "Total Duty"

   
   .TextMatrix(1, 17) = "VAT"
   .TextMatrix(1, 18) = "CST"
   .TextMatrix(1, 19) = "Postate Exp."
   .TextMatrix(1, 20) = "Short & Excess"
   .TextMatrix(1, 21) = "Total Invoice Value"

 
  
 
 End With
  
  
 For i = 1 To 20
    vs.ColWidth(i) = 1500
 Next

vs.ColWidth(11) = 1600
vs.ColWidth(21) = 1800


 
   
   
End Sub
Private Sub cmdShow_Click()




vs.Clear
vsINI


vs.MergeCells = flexMergeFixedOnly

K1 = 2

If rs.State = 1 Then rs.Close
rs.Open "select * from salesQry_withEducess where " & stringyear & " and " & _
" convert(smalldatetime,INVOICEDATE,103)>=convert(smalldatetime,'" & dt1.Value & "',103) and " & _
" convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & dt2.Value & "',103)" & _
" ", CON, adOpenKeyset, adLockReadOnly


For i = 1 To rs.RecordCount

  vs.TextMatrix(K1, 0) = i
  vs.TextMatrix(K1, 1) = rs!InvoiceDate
  vs.TextMatrix(K1, 2) = rs!INVOICENO
  
  
  If rs!with_withoutFormc = "1" Then
     vs.TextMatrix(K1, 6) = rs!GAmount
  ElseIf rs!with_withoutFormc = "2" Then
     vs.TextMatrix(K1, 7) = rs!GAmount
  ElseIf rs!with_withoutFormc = "3" Then
     vs.TextMatrix(K1, 5) = rs!GAmount
  ElseIf rs!with_withoutFormc = "4" Then
     vs.TextMatrix(K1, 8) = rs!GAmount
  ElseIf rs!with_withoutFormc = "5" Then
     vs.TextMatrix(K1, 3) = rs!GAmount
  ElseIf rs!with_withoutFormc = "6" Then
     vs.TextMatrix(K1, 4) = rs!GAmount
  End If
  
  
  
  
  
  
  vs.TextMatrix(K1, 9) = (Val(vs.TextMatrix(K1, 4)) + Val(vs.TextMatrix(K1, 3)) + Val(vs.TextMatrix(K1, 6)) + Val(vs.TextMatrix(K1, 7)) + Val(vs.TextMatrix(K1, 8)))
  
  If vs.TextMatrix(K1, 9) > 0 Then
     vs.TextMatrix(K1, 10) = rs!qty
  ElseIf vs.TextMatrix(K1, 5) > 0 Then
     vs.TextMatrix(K1, 11) = Val(vs.TextMatrix(K1, 5))
     vs.TextMatrix(K1, 12) = rs!qty
  End If
  
  
  If rs!aexp5am > 0 Then
     vs.TextMatrix(K1, 13) = rs!aexp5am
  End If
  
  If rs!aexp6am > 0 Then
     vs.TextMatrix(K1, 14) = rs!aexp6am
  End If
  
  If rs!aexp7am > 0 Then
     vs.TextMatrix(K1, 15) = rs!aexp7am
  End If
  
  vs.TextMatrix(K1, 16) = (Val(vs.TextMatrix(K1, 13)) + Val(vs.TextMatrix(K1, 14)) + Val(vs.TextMatrix(K1, 15)))
  
  
  '------------------------------------
  If InStr(rs!aexp2, "VAT") > 0 Then
    If rs!aexp2am > 0 Then vs.TextMatrix(K1, 17) = rs!aexp2am
     
  End If
  
  If InStr(rs!aexp2, "CST") > 0 Then
     vs.TextMatrix(K1, 18) = rs!aexp2am
  End If
  
  If rs!aexp3am > 0 Then vs.TextMatrix(K1, 19) = rs!aexp3am
  
  If rs!netamount > 0 Then vs.TextMatrix(K1, 21) = rs!netamount
  
  'aexp6am
  
  If rs!aexp6am > 0 Then vs.TextMatrix(K1, 20) = rs!aexp4am
  
  
  K1 = K1 + 1
  
  rs.MoveNext
  
Next


 For i = 1 To 20
    vs.ColWidth(i) = 1500
 Next

vs.ColWidth(11) = 1600
vs.ColWidth(21) = 1700



End Sub

Private Sub Command1_print_Click()


CON.Execute "delete from tmpdsr"



Dim Saleinvoice, Sale_Exempte, Agn_FormC, Agn_WFormC, Agn_FormH, Value_Excise As Double
Dim Qty_Excise, Value_NonExcise, Qty_NonExcise, CenVat, Ecess, SEcess As Double
Dim TotalDuty, VAT, cst, Postage, Short, InvValue As Double

If rs.State = 1 Then rs.Close
rs.Open "select * from tmpDSR", CON, adOpenDynamic, adLockOptimistic
For i = 2 To vs.Rows - 1
If vs.TextMatrix(i, 0) <> "" Then
rs.addNew
rs!sno = vs.TextMatrix(i, 0)
rs!InvDate = vs.TextMatrix(i, 1)
rs!InvNo = vs.TextMatrix(i, 2)

rs!Saletaxinvoice = vs.TextMatrix(i, 3)
Saletaxinvoice = Saletaxinvoice + Val(vs.TextMatrix(i, 3))

rs!Saleinvoice = vs.TextMatrix(i, 4)
Saleinvoice = Saleinvoice + Val(vs.TextMatrix(i, 4))

rs!Sale_Exempte = vs.TextMatrix(i, 5)
Sale_Exempte = Sale_Exempte + Val(vs.TextMatrix(i, 5))


rs!Agn_FormC = vs.TextMatrix(i, 6)
Agn_FormC = Agn_FormC + Val(vs.TextMatrix(i, 6))

rs!Agn_WFormC = vs.TextMatrix(i, 7)
Agn_WFormC = Agn_WFormC + Val(vs.TextMatrix(i, 7))


rs!Agn_FormH = vs.TextMatrix(i, 8)
Agn_FormH = Agn_FormH + Val(vs.TextMatrix(i, 8))

rs!Value_Excise = vs.TextMatrix(i, 9)
Value_Excise = Value_Excise + Val(vs.TextMatrix(i, 9))

rs!Qty_Excise = vs.TextMatrix(i, 10)
Qty_Excise = Qty_Excise + Val(vs.TextMatrix(i, 10))

rs!Value_NonExcise = vs.TextMatrix(i, 11)
Value_NonExcise = Value_NonExcise + Val(vs.TextMatrix(i, 11))

rs!Qty_NonExcise = vs.TextMatrix(i, 12)
Qty_NonExcise = Qty_NonExcise + Val(vs.TextMatrix(i, 12))

rs!CenVat = vs.TextMatrix(i, 13)
CenVat = CenVat + Val(vs.TextMatrix(i, 13))

rs!Ecess = vs.TextMatrix(i, 14)
Ecess = Ecess + Val(vs.TextMatrix(i, 14))

rs!SEcess = vs.TextMatrix(i, 15)
SEcess = SEcess + Val(vs.TextMatrix(i, 15))

rs!TotalDuty = vs.TextMatrix(i, 16)
TotalDuty = TotalDuty + Val(vs.TextMatrix(i, 16))

rs!VAT = vs.TextMatrix(i, 17)
VAT = VAT + Val(vs.TextMatrix(i, 17))

rs!cst = vs.TextMatrix(i, 18)
cst = cst = Val(vs.TextMatrix(i, 18))


rs!Postage = vs.TextMatrix(i, 19)
Postage = Postage + Val(vs.TextMatrix(i, 19))

rs!Short = vs.TextMatrix(i, 20)
Short = Short + Val(vs.TextMatrix(i, 20))

rs!InvValue = vs.TextMatrix(i, 21)
InvValue = InvValue + Val(vs.TextMatrix(i, 21))

      
rs.Update

End If

Next



rs.addNew
rs!InvNo = "Total"
rs!Saletaxinvoice = Round(Saletaxinvoice, 0)
rs!Saleinvoice = Round(Saleinvoice, 0)
rs!Sale_Exempte = Round(Sale_Exempte, 0)
rs!Agn_FormC = Round(Agn_FormC, 0)
rs!Agn_WFormC = Round(Agn_WFormC, 0)
rs!Agn_FormH = Round(Agn_FormH, 0)
rs!Value_Excise = Round(Value_Excise, 0)
rs!Qty_Excise = Round(Qty_Excise, 0)
rs!Value_NonExcise = Round(Value_NonExcise, 0)
rs!Qty_NonExcise = Round(Qty_NonExcise, 0)
rs!CenVat = Round(CenVat, 0)
rs!Ecess = Round(Ecess, 0)
rs!SEcess = Round(SEcess, 0)
rs!TotalDuty = Round(TotalDuty, 0)
rs!VAT = Round(VAT, 0)
rs!cst = Round(cst, 0)
rs!Postage = Round(Postage, 0)
rs!Short = Round(Short, 0)
rs!InvValue = Round(InvValue, 0)
rs.Update
      

DoEvents
DoEvents
DoEvents
DoEvents
    
cr.Reset
cr.Connect = constr
cr.ReportFileName = App.Path & "\REPORTS\dsr.rpt"

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
