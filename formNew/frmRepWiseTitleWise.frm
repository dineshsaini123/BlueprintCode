VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRepWiseTitleWise 
   Caption         =   "Rep.WiseTitle Wise Net Qty."
   ClientHeight    =   2064
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   9756
   Icon            =   "frmRepWiseTitleWise.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2064
   ScaleWidth      =   9756
   Begin VB.ComboBox cboAgent 
      Height          =   288
      ItemData        =   "frmRepWiseTitleWise.frx":000C
      Left            =   1404
      List            =   "frmRepWiseTitleWise.frx":001C
      TabIndex        =   16
      Top             =   1620
      Width           =   2988
   End
   Begin VB.CommandButton Command1_headwise 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&HEADWISE Rep.WiseSales"
      Height          =   645
      Left            =   4545
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   945
      Width           =   2985
   End
   Begin VB.ComboBox cbogp 
      Height          =   288
      ItemData        =   "frmRepWiseTitleWise.frx":002E
      Left            =   1395
      List            =   "frmRepWiseTitleWise.frx":003E
      TabIndex        =   2
      Top             =   1260
      Width           =   1365
   End
   Begin VB.CommandButton Command1_excel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Export To Excel"
      Height          =   645
      Left            =   11790
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   135
      Width           =   1095
   End
   Begin VB.CommandButton CommandPrint 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Print (Rep.Wise And Title Wise)"
      Height          =   645
      Left            =   5940
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   285
      Width           =   1590
   End
   Begin VB.CommandButton CommandReturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Return"
      Height          =   645
      Left            =   7560
      Picture         =   "frmRepWiseTitleWise.frx":0050
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   285
      Width           =   1275
   End
   Begin VB.CommandButton cmdFilterData 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Filter Data"
      Height          =   645
      Left            =   4545
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   285
      Width           =   1365
   End
   Begin Crystal.CrystalReport cr 
      Left            =   11295
      Top             =   420
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker txtFromSale 
      Height          =   315
      Left            =   1380
      TabIndex        =   0
      Top             =   285
      Width           =   1335
      _ExtentX        =   2350
      _ExtentY        =   550
      _Version        =   393216
      Format          =   473825281
      CurrentDate     =   42409
   End
   Begin MSComCtl2.DTPicker txttoSale 
      Height          =   315
      Left            =   3015
      TabIndex        =   1
      Top             =   285
      Width           =   1410
      _ExtentX        =   2498
      _ExtentY        =   550
      _Version        =   393216
      Format          =   473825281
      CurrentDate     =   42409
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   120
      Left            =   5076
      TabIndex        =   7
      Top             =   2088
      Visible         =   0   'False
      Width           =   3456
      _cx             =   6085
      _cy             =   212
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
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
      ForeColorSel    =   12582912
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
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   500
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmRepWiseTitleWise.frx":0C34
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
      WordWrap        =   -1  'True
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
   Begin MSComCtl2.DTPicker txtspfdate 
      Height          =   315
      Left            =   1380
      TabIndex        =   11
      Top             =   675
      Width           =   1335
      _ExtentX        =   2350
      _ExtentY        =   572
      _Version        =   393216
      Format          =   473825281
      CurrentDate     =   42409
   End
   Begin MSComCtl2.DTPicker txtsptdate 
      Height          =   315
      Left            =   3015
      TabIndex        =   12
      Top             =   675
      Width           =   1410
      _ExtentX        =   2498
      _ExtentY        =   572
      _Version        =   393216
      Format          =   473825281
      CurrentDate     =   42409
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Agent Name :"
      Height          =   252
      Left            =   108
      TabIndex        =   17
      Top             =   1692
      Width           =   1356
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Specimen Range :"
      Height          =   255
      Left            =   90
      TabIndex        =   14
      Top             =   675
      Width           =   1350
   End
   Begin VB.Label Label3 
      Caption         =   "To"
      Height          =   255
      Left            =   2745
      TabIndex        =   13
      Top             =   675
      Width           =   315
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Group Code :"
      Height          =   255
      Left            =   90
      TabIndex        =   10
      Top             =   1260
      Width           =   1350
   End
   Begin VB.Label lblAson 
      Caption         =   "To"
      Height          =   255
      Left            =   2745
      TabIndex        =   9
      Top             =   285
      Width           =   315
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Range :"
      Height          =   255
      Left            =   90
      TabIndex        =   8
      Top             =   285
      Width           =   1350
   End
End
Attribute VB_Name = "frmRepWiseTitleWise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFilterData_Click()

Screen.MousePointer = vbHourglass

'con.Execute "exec rep_and_TitelWiseSp '" & txtFromSale.value & "','" & txttoSale.value & "'"
con.Execute "exec rep_and_TitelWiseSp '" & txtFromSale.value & "','" & txttoSale.value & "','" & txtspfdate.value & "','" & txtsptdate.value & "'"

'con.Execute "delete from tmpRepWiseSaleReturn where GROUPCODE<>'BP'"

DoEvents
DoEvents
DoEvents
DoEvents
DoEvents

Screen.MousePointer = vbDefault



End Sub
Private Sub Command1_excel_Click()

Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim I As Integer
Dim xl As Excel.Application
Dim str_ As String




If xl Is Nothing Then
   Set xl = CreateObject("Excel.Application")
End If

xl.Visible = True
Set xlBook = xl.Workbooks.Add
Set xlSheet = xlBook.Worksheets.Add

Dim c, r As Long
Dim Q1, q2, J As Double



row_ = 1
col_ = 1

xl.Columns("A:H").ColumnWidth = 12
J = 2


For I = 0 To vs.rows - 1
    For J = 0 To vs.Cols - 1
      
        xlSheet.Cells(row_, col_).value = vs.TextMatrix(I, J)
       
        col_ = col_ + 1
    Next
    row_ = row_ + 1
    col_ = 1
Next

End Sub

Private Sub Command1_headwise_Click()
Screen.MousePointer = vbHourglass
DSNNew

    MainMenu.cr1.Reset
    MainMenu.cr1.ReportFileName = rptPath & "/HeadWiseRepWise.rpt"
    MainMenu.cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    
    If cbogp.Text <> "" Then
        MainMenu.cr1.ReplaceSelectionFormula "{HEADWISE_RepwiseSales.GROUPCODE}='" & cbogp.Text & "'"
        
    End If
     MainMenu.cr1.Formulas(0) = "fdate='" & txtFromSale.value & "'"
    MainMenu.cr1.Formulas(1) = "tdate='" & txttoSale.value & "'"
   
    If cbogp.Text = "" Then
       MainMenu.cr1.Formulas(2) = "code_='" & "All" & "'"
    Else
       MainMenu.cr1.Formulas(2) = "code_='" & cbogp.Text & "'"
    End If
    
    MainMenu.cr1.WindowShowPrintSetupBtn = True
    MainMenu.cr1.WindowShowExportBtn = True
    MainMenu.cr1.WindowShowRefreshBtn = True
    MainMenu.cr1.WindowState = crptMaximized
    MainMenu.cr1.Action = 1

Screen.MousePointer = vbDefault

End Sub

Private Sub CommandPrint_Click()

Screen.MousePointer = vbHourglass
DSNNew

MainMenu.cr1.Reset
MainMenu.cr1.ReportFileName = rptPath & "/RepWiseSaleReturn.rpt"
MainMenu.cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass

If cbogp.Text = "" Then
   
   If cboAgent.Text = "" Then
     'MainMenu.cr1.ReplaceSelectionFormula "({RepWiseSaleReturnQry.invoicedate}>=datevalue('" & Format(txtFromSale.value, "MM/dd/yyyy") & "') and {RepWiseSaleReturnQry.invoicedate}<=datevalue('" & Format(txttoSale.value, "MM/dd/yyyy") & "'))"
   Else
       MainMenu.cr1.ReplaceSelectionFormula "({RepWiseSaleReturnQry.agentname}='" & cboAgent.Text & "')"
   End If
Else

 If cboAgent.Text = "" Then
    MainMenu.cr1.ReplaceSelectionFormula "({RepWiseSaleReturnQry.groupcode}='" & cbogp.Text & "')"
 Else
    MainMenu.cr1.ReplaceSelectionFormula "({RepWiseSaleReturnQry.groupcode}='" & cbogp.Text & "' and {RepWiseSaleReturnQry.agentname}='" & cboAgent.Text & "')"
 End If
 
End If

MainMenu.cr1.WindowShowPrintSetupBtn = True
MainMenu.cr1.WindowShowExportBtn = True
MainMenu.cr1.WindowShowRefreshBtn = True
MainMenu.cr1.WindowState = crptMaximized
MainMenu.cr1.Action = 1




Screen.MousePointer = vbDefault



End Sub
Private Sub CommandReturn_Click()
 Unload Me
End Sub
Private Sub Form_Load()

Me.Width = 9030
Me.Height = 3000


txtFromSale.value = Format(Date, "dd/MM/yyyy")
txttoSale.value = Format(Date, "dd/MM/yyyy")

txtspfdate.value = Format(Date, "dd/MM/yyyy")
txtsptdate.value = Format(Date, "dd/MM/yyyy")
Set RS = New ADODB.Recordset
RS.Open "select Rep as Representative,Email from SalesRepQry where (email is not null and len(email)>1) order by Rep", CON_blue
  cboAgent.Clear
  If Not RS.EOF Then
     Do While Not RS.EOF
        If IsNull(RS(0)) = False Then
          Me.cboAgent.AddItem RS(0)
        End If
        If Not RS.EOF Then RS.MoveNext
      Loop
  End If
    

BackColorFrom Me
End Sub


