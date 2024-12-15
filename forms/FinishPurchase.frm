VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form FinishPurchase 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Purchase "
   ClientHeight    =   9060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14715
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9060
   ScaleWidth      =   14715
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtbillno 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "0"
      Top             =   600
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrint1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Print"
      Enabled         =   0   'False
      Height          =   765
      Left            =   7635
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   7500
      Width           =   1035
   End
   Begin VB.ComboBox CBOCREDIT 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5130
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   6615
      Visible         =   0   'False
      Width           =   1890
   End
   Begin MSDataListLib.DataCombo Cmbmedi 
      Height          =   2505
      Left            =   180
      TabIndex        =   26
      Top             =   2295
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4419
      _Version        =   393216
      Appearance      =   0
      Style           =   1
      ListField       =   ""
      BoundColumn     =   ""
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Height          =   7875
      Left            =   11160
      TabIndex        =   17
      Top             =   90
      Width           =   3180
      Begin VB.ComboBox cbogp 
         Height          =   315
         Left            =   225
         TabIndex        =   2
         Top             =   1845
         Width           =   2850
      End
      Begin VB.CommandButton cmdsearch 
         BackColor       =   &H80000013&
         Caption         =   "&Search"
         Height          =   375
         Left            =   630
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1125
         Width           =   1935
      End
      Begin VB.ListBox listno 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5310
         Left            =   225
         TabIndex        =   18
         Top             =   2250
         Width           =   2865
      End
      Begin MSComCtl2.DTPicker todate 
         Height          =   315
         Left            =   1290
         TabIndex        =   19
         Top             =   645
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         Format          =   19595265
         CurrentDate     =   38923
      End
      Begin MSComCtl2.DTPicker fromdate 
         Height          =   315
         Left            =   1290
         TabIndex        =   22
         Top             =   225
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         Format          =   19595265
         CurrentDate     =   38923
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "From Date :"
         Height          =   225
         Left            =   495
         TabIndex        =   21
         Top             =   285
         Width           =   915
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "To Date :"
         Height          =   255
         Left            =   495
         TabIndex        =   20
         Top             =   645
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdprint 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Exit"
      Height          =   765
      Left            =   8670
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7500
      Width           =   1035
   End
   Begin VB.TextBox txtrem 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1665
      TabIndex        =   5
      Top             =   1620
      Width           =   6435
   End
   Begin VB.TextBox txtparty 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1665
      TabIndex        =   4
      Top             =   1230
      Width           =   4245
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8370
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "0"
      Top             =   6600
      Width           =   1245
   End
   Begin VB.CommandButton cmdmodify 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Modify"
      Enabled         =   0   'False
      Height          =   765
      Left            =   5550
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7500
      Width           =   1065
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Save"
      Height          =   765
      Left            =   4515
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7500
      Width           =   1035
   End
   Begin VB.CommandButton cmdref 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Refresh"
      Height          =   765
      Left            =   3450
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7500
      Width           =   1065
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Delete Bill"
      Enabled         =   0   'False
      Height          =   765
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7500
      Width           =   1035
   End
   Begin MSComCtl2.DTPicker dtpdate1 
      Height          =   270
      Left            =   6735
      TabIndex        =   1
      Top             =   540
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   19595267
      CurrentDate     =   38338
   End
   Begin VSFlex7LCtl.VSFlexGrid fg 
      Height          =   4275
      Left            =   135
      TabIndex        =   24
      Top             =   2025
      Width           =   9645
      _cx             =   17013
      _cy             =   7541
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   0
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
      BackColor       =   16777215
      ForeColor       =   16711680
      BackColorFixed  =   12632256
      ForeColorFixed  =   0
      BackColorSel    =   16769505
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
      GridColor       =   16711680
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   8388608
      SheetBorder     =   8388608
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   8
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   2
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
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   510
         Left            =   6210
         TabIndex        =   25
         Top             =   6375
         Width           =   2655
      End
   End
   Begin MSComCtl2.DTPicker chdate 
      Height          =   270
      Left            =   6750
      TabIndex        =   3
      Top             =   900
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   19595267
      CurrentDate     =   38338
   End
   Begin Crystal.CrystalReport cr 
      Left            =   270
      Top             =   7065
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label8 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Challan Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5535
      TabIndex        =   28
      Top             =   900
      Width           =   1410
   End
   Begin VB.Label Label5 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Challan No."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   345
      TabIndex        =   27
      Top             =   915
      Width           =   1185
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   345
      TabIndex        =   15
      Top             =   1590
      Width           =   1275
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier/Firm"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   345
      TabIndex        =   14
      Top             =   1230
      Width           =   1440
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Total :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   7335
      TabIndex        =   13
      Top             =   6660
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   345
      TabIndex        =   12
      Top             =   585
      Width           =   825
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5535
      TabIndex        =   11
      Top             =   555
      Width           =   1095
   End
End
Attribute VB_Name = "FinishPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mrs As New ADODB.Recordset
Dim editflag As Boolean
Private Sub cboParty_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys "{tab}"
   End If
End Sub
Private Sub chdate_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then txtparty.SetFocus
End Sub
Private Sub Cmbmedi_KeyDown(KeyCode As Integer, Shift As Integer)
  
  If KeyCode = 27 Then
  If Cmbmedi.Visible = True Then
     Cmbmedi.Visible = False
  End If
  End If
  
End Sub
Private Sub Cmbmedi_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
     
    If fg.Col = 0 Then
       Cmbmedi.Visible = False
       fg.TextMatrix(fg.RowSel, 0) = Cmbmedi.Text
       'fg.TextMatrix(fg.RowSel, 1) = Cmbmedi.BoundText
       fg.SetFocus
       SendKeys "{right}"
       
       If rs.State = 1 Then rs.Close
       rs.Open "select * from ItemCreation where CourseName='" & Cmbmedi.Text & "'", CON
       If rs.EOF = False Then
          fg.TextMatrix(fg.RowSel, 3) = rs.Fields("unit").Value & ""
          fg.TextMatrix(fg.RowSel, 4) = rs.Fields("Price").Value
       End If
       
       
    Else
       Cmbmedi.Visible = False
       fg.TextMatrix(fg.RowSel, 1) = Cmbmedi.Text
       'fg.TextMatrix(fg.RowSel, 0) = Cmbmedi.BoundText
       fg.SetFocus
       SendKeys "{right}"
       'SendKeys "{right}"
       
       If Cmbmedi.BoundText = "" Then Exit Sub
       
       If rs.State = 1 Then rs.Close
       rs.Open "select unit,price from ItemCreation where ItemCode=" & Cmbmedi.BoundText & "", CON
       If rs.EOF = False Then
          fg.TextMatrix(fg.RowSel, 3) = rs.Fields("unit").Value & ""
          fg.TextMatrix(fg.RowSel, 4) = rs.Fields("Price").Value
       End If
    End If
      
      
    
    End If
End Sub
Private Sub cmdCal_Click()
     '=============================
     
     
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdModify_Click()

If MsgBox("Are you sure to Modify ?", vbQuestion + vbYesNo) = vbYes Then
   deleteFinishPurchase
   save
   
 
End If

End Sub
Sub max()
'    Set rs = New ADODB.Recordset
'    If rs.State = 1 Then rs.Close
'    rs.Open "select max(billno) from finishpurchase", con
'    If IsNull(rs.Fields(0).Value) Then
'       txtbillno.Text = 1
'       Else
'       txtbillno.Text = rs.Fields(0).Value + 1
'    End If
End Sub


Private Sub cmdPrint_Click()
Unload Me
End Sub

Private Sub cmdprint1_Click()
CR.Reset
CR.ReportFileName = App.Path & "\purchaseSlip.RPT"
CR.Connect = "FILEDSN=hotel;pwd=java;"
CR.ReplaceSelectionFormula "{FinishPurchase.billno}='" & txtbillno.Text & "'"
CR.WindowShowCloseBtn = True
CR.WindowShowPrintBtn = True
CR.WindowControlBox = True
CR.WindowShowPrintSetupBtn = True
CR.WindowShowProgressCtls = True
CR.WindowState = crptMaximized
CR.Action = 1

End Sub

Private Sub cmdRef_Click()
     fg.Clear
     
     SeWidth
     'max
     fg.Rows = 2
     cmdSave.Enabled = True
     
     
     
     txtparty.Text = ""
     txtrem.Text = ""
     dtpdate1.Value = Date
     txtTotal.Text = 0
     'CBOCREDIT.ListIndex = 0
     txtbillno.Text = ""
     'txtch.Text = ""
     txtbillno.SetFocus
     
End Sub

Private Sub cmdSave_Click()
  If MsgBox("Do U Want Save ?", vbQuestion + vbYesNo) = vbYes Then
      save
  End If
 
End Sub
Sub updateMaster()
     
   For i = 1 To fg.Rows - 1
      If fg.TextMatrix(i, 4) <> "" Then
      CON.Execute "update ItemCreation set price=" & fg.TextMatrix(i, 4) & " where itemname='" & fg.TextMatrix(i, 1) & "'"
      End If
   Next
     
End Sub

Sub save()
   If txtbillno.Text = "" Then
      MsgBox "Please Enter No !!", vbExclamation
      Exit Sub
   End If
   Set rs = New ADODB.Recordset
   If rs.State = 1 Then rs.Close
   rs.Open "select * from finishpurchase where billno='" & txtbillno.Text & "'", CON, adOpenDynamic, adLockOptimistic
   If rs.EOF = True Then
   
   For i = 1 To fg.Rows - 1
   
   If fg.TextMatrix(i, 1) <> "" Then
           
      rs.addNew
      rs!billno = Trim(txtbillno.Text)
      rs!Supplier = Trim(txtparty.Text)
      rs!dates = dtpdate1.Value
      rs!Remarks = txtrem.Text
      rs!challan = txtch.Text
      rs!challan_date = chdate.Value
      rs!totalAmt = Trim(txtTotal.Text)
      
      rs!gp = Trim(fg.TextMatrix(i, 0))
      rs!itemname = Trim(fg.TextMatrix(i, 1))
      rs!unit = Trim(fg.TextMatrix(i, 3))
      rs!qty = fg.TextMatrix(i, 2)
      rs!price = fg.TextMatrix(i, 4)
      rs!amt = Val(fg.TextMatrix(i, 5))
      rs!credit = CBOCREDIT.Text
      
      rs.Update
                 
      updateMaster
                 
    End If
  Next
   
   
   Else
   
      MsgBox "This  No Already Exist", vbInformation
   
   End If

   Call cmdRef_Click

End Sub
Sub search()
 
  Dim rs1 As New ADODB.Recordset
   
   
  Dim billno As Long
  
  If listno.Text <> "" Then
     txtbillno.Text = listno.Text
  End If

   
   
   
   Set rs = New ADODB.Recordset
   If rs.State = 1 Then rs.Close
   rs.Open "select * from finishpurchase where billno='" & txtbillno.Text & "'", CON, adOpenDynamic, adLockOptimistic
   If rs.EOF = False Then
      
      fg.Rows = rs.RecordCount + 1
      
      cmdmodify.Enabled = True
      Command2.Enabled = True
      txtparty.Text = rs!Supplier
       
         
      dtpdate1.Value = rs!dates
      txtTotal.Text = Format(rs!totalAmt, "0.00")
      txtrem.Text = rs!Remarks
      txtch.Text = rs!challan & ""
      chdate.Value = IIf(Not IsNull(rs!challan_date), rs!challan_date, Date)
      'CBOCREDIT.Text = rs!credit
      
      For i = 1 To rs.RecordCount
      fg.TextMatrix(i, 0) = rs!gp
      fg.TextMatrix(i, 1) = rs!itemname
      fg.TextMatrix(i, 3) = rs!unit
      fg.TextMatrix(i, 2) = Format(rs!qty, "0.000")
      fg.TextMatrix(i, 4) = Format(rs!price, "0.00")
      fg.TextMatrix(i, 5) = Format(rs!amt, "0.00")
      'fg.Rows = fg.Rows + 1
          
      rs.MoveNext
      Next

   End If
        
   'fg.Rows = fg.Rows + 2
      
      
      
   
End Sub



    


Private Sub cmdSearch_Click()
    Me.listno.Clear
    
    Set rs = Nothing
    If cbogp.Text = "" Then
       Set rs = CON.Execute("Select BillNo from finishpurchase where convert(smalldatetime,Dates,103)>= convert(smalldatetime,'" & fromdate.Value & "',103) and convert(smalldatetime,Dates,103) <= convert(smalldatetime,'" & todate.Value & "',103) Group By BillNo Order by BillNo")
    Else
       Set rs = CON.Execute("Select BillNo from finishpurchase where gp='" & cbogp.Text & "' and convert(smalldatetime,Dates,103)>= convert(smalldatetime,'" & fromdate.Value & "',103) and convert(smalldatetime,Dates,103) <= convert(smalldatetime,'" & todate.Value & "',103) Group By BillNo Order by BillNo")
    End If
    
    Do While Not rs.EOF = True
            Me.listno.AddItem (rs("BillNo"))
    rs.MoveNext
    Loop

End Sub

Private Sub Command2_Click()

If MsgBox("Are you sure to delete", vbQuestion + vbYesNo) = vbYes Then
   deleteFinishPurchase
   
   Call cmdRef_Click

End If

End Sub
Sub deleteFinishPurchase()
  Set mrs = Nothing
  Set mrs = CON.Execute("Delete from finishpurchase where billno='" & txtbillno.Text & "'")
  
  
  cmdmodify.Enabled = False
  Command2.Enabled = False

End Sub
Private Sub CommandButton4_Click()
Unload Me
End Sub
Private Sub dtpdate_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      cboParty.SetFocus
   End If
End Sub
Sub Total()
    txtTotal.Text = 0
    For i = 1 To fg.Rows - 1
       If fg.TextMatrix(i, 2) <> "" And fg.TextMatrix(i, 4) <> "" Then
          
          fg.TextMatrix(i, 5) = (Val(fg.TextMatrix(i, 2)) * Val(fg.TextMatrix(i, 4)))
          txtTotal.Text = (Val(txtTotal.Text) + Val(fg.TextMatrix(i, 5)))
          
       End If
    Next
    
    txtTotal.Text = Format(txtTotal.Text, "0.00")
    
End Sub

Private Sub Command3_Click()
''   Skinner1.Enabled = Not Skinner1.Enabled
''   Set Skinner1.Forms = Forms
''    If Skinner1.Enabled Then
''        Command3.Caption = "Disable &skin"
''    Else
''        Command3.Caption = "Enable &skin"
''    End If
End Sub

Private Sub date1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      txtTime2.SetFocus
   End If
End Sub
Private Sub datePLA_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then txtPLArs.SetFocus
   
End Sub
Private Sub daterg_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtregRs.SetFocus
End Sub

Private Sub dtpdate1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then txtch.SetFocus
End Sub

Private Sub fg_EnterCell()
If Me.fg.Col = 2 Or Me.fg.Col = 4 Then
Me.fg.Editable = flexEDKbd
Else
Me.fg.Editable = flexEDNone
End If
End Sub

Private Sub fg_KeyDown(KeyCode As Integer, Shift As Integer)




 If KeyCode = 13 Then

           
     If fg.Col = 1 Then
           
           Cmbmedi.Text = ""
           
           cellposi
           Dim filldata As New ADODB.Recordset
            filldata.Open "select ItemCode,ItemName from ItemCreation where CourseName='" & fg.TextMatrix(fg.RowSel, 0) & "' order by ItemName", CON
            
            Set Cmbmedi.RowSource = filldata
            Cmbmedi.ListField = "ItemName"
            Cmbmedi.BoundColumn = "ItemCode"
            Cmbmedi.ReFill
            
            Cmbmedi.Visible = True
            Cmbmedi.SetFocus
        End If
  End If


 



If KeyCode = 13 Then
If Me.fg.Col = 5 Then
Me.fg.Rows = Me.fg.Rows + 1
Me.fg.Row = Me.fg.Row + 1
Me.fg.Col = 0
Else
End If




End If


If KeyCode = 46 Then
   fg.RemoveItem (fg.RowSel)
   Total
End If



   If KeyCode = 13 Then
     If fg.Col = 0 Then
       
       cellposi
       fillcmb
       fg.Editable = flexEDNone
       Cmbmedi.Visible = True
       Cmbmedi.SetFocus
     ElseIf fg.Col = 1 Then
            
       fg.Editable = flexEDNone
       SendKeys "{right}"
       'SendKeys "{right}"
     
     ElseIf fg.Col = 2 Then
       fg.Editable = flexEDNone
       SendKeys "{right}"
       SendKeys "{right}"
       
       
     ElseIf fg.Col = 3 Then
       fg.Editable = flexEDNone
       
     End If
     
     
     
   End If

   
  





End Sub

Private Sub fg_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)

On Error Resume Next

Dim d1
Dim D2


If KeyCode = 13 Then

If Me.fg.Row = Me.fg.Rows - 1 And Me.fg.Col = 4 Then

   
  
  
  
    
    
    Me.fg.Rows = Me.fg.Rows + 1
    Me.fg.Row = Me.fg.Row + 1
    Me.fg.Col = 0
      
    
   If fg.TextMatrix(fg.RowSel - 1, 4) <> "" Then
     fg.TextMatrix(fg.RowSel - 1, 5) = (CDbl(fg.TextMatrix(fg.RowSel - 1, 4)) * CDbl(fg.TextMatrix(fg.RowSel - 1, 2)))
    End If

    
    
    
    Total
   
Else
   
   On Error Resume Next
    
    
    
   If fg.TextMatrix(fg.RowSel - 1, 4) <> "" Then
     fg.TextMatrix(fg.RowSel - 1, 5) = (CDbl(fg.TextMatrix(fg.RowSel - 1, 4)) * CDbl(fg.TextMatrix(fg.RowSel - 1, 2)))
    End If

    
    
    
    Total

    
    Me.fg.Col = Me.fg.Col + 1

End If






End If


End Sub

Private Sub fg_KeyUp(KeyCode As Integer, Shift As Integer)
  If fg.Col = 3 Then
  
  SendKeys "{right}"
   ' If fg.TextMatrix(fg.RowSel, 4) <> "" Then
    ' fg.TextMatrix(fg.RowSel, 5) = (CDbl(fg.TextMatrix(fg.RowSel, 4)) * CDbl(fg.TextMatrix(fg.RowSel, 3)))
  'End If

 'End If
 
 
 ElseIf fg.Col = 5 Then
 
   Total
 
 End If

End Sub

Private Sub fg_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    If Col = 4 Then


    End If
End Sub



Sub cellposi()
  Cmbmedi.Width = fg.CellWidth
  Cmbmedi.TOP = fg.TOP + fg.CellTop
  Cmbmedi.Left = fg.Left + fg.CellLeft - 45
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then
    If Val(txtTotal.Text) > 0 Then
     Call cmdSave_Click
     
    End If
    Unload Me
End If

End Sub

Private Sub Form_Load()
''Main


chdate.Value = Date
Me.fg.ColComboList(0) = lst

SeWidth

'max

dtpdate1.Value = Date


fillcmb

'Call frmBackColor(Me)

fromdate.Value = Date
todate.Value = Date
'CBOCREDIT.ListIndex = 0


If mrs.State = 1 Then mrs.Close
mrs.Open "select distinct(CourseName) from ItemCreation", CON
While mrs.EOF = False
cbogp.AddItem mrs(0)
mrs.MoveNext
Wend

End Sub
Sub fillcmb()

Dim filldata As New ADODB.Recordset


filldata.Open "select distinct(CourseName) from ItemCreation order by CourseName", CON

Set Cmbmedi.RowSource = filldata

Cmbmedi.ListField = "CourseName"
Cmbmedi.BoundColumn = "CourseName"
Cmbmedi.ReFill

End Sub


Sub SeWidth()
    
    fg.Cols = 6
    fg.FormatString = "Group|Item Name|Quantity|Unit|Rate|Amount"
    fg.ColWidth(0) = 2200
    fg.ColWidth(1) = 2800
    fg.ColWidth(2) = 1200
    fg.ColWidth(3) = 1200
    fg.ColWidth(4) = 1200
    fg.ColWidth(5) = 2000
    
    
    
    
End Sub

Private Sub FromDate_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      todate.SetFocus
   End If
End Sub

Private Sub List1_Click()
'''''''''''''''''''''''

'''                On Error Resume Next
'''                Set mrs = Nothing
'''                Set mrs = con.Execute("Select * from RawPurchaseMain where billno=" & Me.List1.Text & "")
'''                If mrs.EOF = True Then
'''                   MsgBox "finishpurchase No. dose't exist.", vbInformation
'''                   cmdSave.Enabled = True
'''                Else
'''                   dtpdate.Value = mrs.Fields("Dates").Value
'''                   cboParty.Text = mrs.Fields("PartyName").Value
'''                   txtbillno.Text = mrs.Fields("billno").Value
'''                   txtTotal.Text = mrs.Fields("amt").Value
'''                   cmdSave.Enabled = False
'''                End If
'''
'''
'''                Set mrs = Nothing
'''                Set mrs = con.Execute("Select * from RawPurchase where billno=" & Me.List1.Text & "")
'''
'''                Me.fg.Rows = 1
'''                i = 1
'''                            Do While Not mrs.EOF = True
'''
'''                                Me.fg.Rows = Me.fg.Rows + 1
'''                                Me.fg.TextMatrix(i, 0) = mrs("Itemcode")
'''                                Me.fg.TextMatrix(i, 1) = mrs("Itemname")
'''                                Me.fg.TextMatrix(i, 2) = mrs("Unit")
'''                                Me.fg.TextMatrix(i, 3) = mrs("Qty")
'''
'''
'''                                mrs.MoveNext
'''                                i = i + 1
'''                            Loop
'''
'''                          cmdmodify.Enabled = True
'''                          Command2.Enabled = True
'''
'''                          Total
'''
'''                       ' Call DeletePermissin(Command2)
'''                       ' Call SavePermissin(cmdSave)
'''                        'Call ModifyPermissin(cmdmodify)
'''
'''
'''''''''''''''''''''''''''
End Sub

Private Sub Option1_Click()
'Me.txtchequeno.Text = ""
'Me.txtchequeno.Enabled = False
'Me.txtchequeno.Visible = False
'Me.Label5.Enabled = False
'Me.Label5.Visible = False
'
End Sub

Private Sub Option1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.fg.SetFocus
Me.fg.Col = 0
End If
End Sub

Private Sub Option2_Click()
'Me.Label5.Enabled = True
'Me.Label5.Visible = True
'Me.txtchequeno.Enabled = True
'Me.txtchequeno.Visible = True
'Me.txtchequeno.SetFocus
End Sub

Private Sub Option2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.fg.SetFocus
Me.fg.Col = 0
End If
End Sub

Private Sub Option3_Click()
'Me.txtchequeno.Text = ""
'Me.txtchequeno.Enabled = False
'Me.txtchequeno.Visible = False
'Me.Label5.Enabled = False
'Me.Label5.Visible = False
End Sub

Private Sub Option3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.fg.SetFocus
Me.fg.Col = 0
End If
End Sub

Private Sub txtchequeno_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
'If KeyCode = 13 Then
'Me.fg.SetFocus
'Me.fg.Col = 0
'ElseIf KeyCode = vbKeyUp Then
'Me.Option1.SetFocus
'End If
End Sub

Private Sub SearchVs_KeyDown(KeyCode As Integer, Shift As Integer)
  
'''  If KeyCode = 38 Then
'''    If SearchVs.Row = 0 Then
'''     txtSearch.SetFocus
'''    End If
'''  ElseIf KeyCode = 13 Then
'''
'''     fg.SetFocus
'''
'''
'''    Set mrs = Nothing
'''    Set mrs = con.Execute("Select * from ItemCreation where ItemCode='" & Me.SearchVs.TextMatrix(Me.SearchVs.RowSel, 0) & "'")
'''    If Not mrs.EOF = True Then
'''       Me.fg.TextMatrix(fg.RowSel, 0) = mrs("ItemCode")
'''       Me.fg.TextMatrix(fg.RowSel, 1) = mrs("ItemName")
'''       Me.fg.TextMatrix(fg.RowSel, 2) = mrs("Unit")
'''       SendKeys "{right}"
'''    End If
'''
'''
'''
'''
'''     SearchFrame.Visible = False
'''  End If

End Sub

Private Sub listno_Click()
  search
End Sub

Private Sub todate_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then cmdSearch.SetFocus
End Sub

Private Sub txtSearch_Change()

Set mrs = Nothing
Set mrs = CON.Execute("Select ItemCode from ItemCreation where ItemName like '" & txtSearch.Text & "%' order by ItemName")
If mrs.EOF = False Then
   Set SearchVs.DataSource = mrs
End If

End Sub

Private Sub txtSearch_GotFocus()
  txtSearch.BackColor = &HFFC0C0
End Sub
Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 40 Then
     SearchVs.SetFocus
  End If
End Sub

Private Sub txtSearch_LostFocus()
  txtSearch.BackColor = &HFFFFFF
End Sub
Sub searchdate()

If rs.State = 1 Then rs.Close
rs.Open "Select subledger from sledger where subledger='" & PopUpValue1 & "'", CON, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
   
   txtparty.Text = rs!subledger
     
   PopUpValue1 = ""
   
End If

End Sub

Private Sub Text2_Change()

End Sub

Private Sub Text3_Change()

End Sub

Private Sub txtdutyinword2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtTotalcases.SetFocus
End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtbillno_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
      search
   End If
End Sub

Private Sub Label29_Click()

End Sub

Private Sub txtExiceDuty_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
       txtPLANo.SetFocus
   End If
End Sub

Private Sub txtgrno_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      txtPono.SetFocus
   End If
End Sub

Private Sub txtModeTr_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      txtgrno.SetFocus
   End If
End Sub

Private Sub txtch_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = 13 Then chdate.SetFocus
End Sub

Private Sub txtparty_GotFocus()
   searchdate
End Sub

Private Sub txtparty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode >= 65 And KeyCode <= 122 Then
       txtparty.Text = ""
       popuplist10 "Select subledger from SLedger order by subledger", CON
    End If
    
    If KeyCode = 13 Then
       txtrem.SetFocus
    End If
    
End Sub
Private Sub txtPLANo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      datePLA.SetFocus
   End If
End Sub

Private Sub txtPLArs_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtRG23.SetFocus
End Sub

Private Sub txtPono_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      fg.SetFocus
   End If
End Sub

Private Sub txttime_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      date1.SetFocus
   End If
End Sub

Private Sub txtrem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then fg.SetFocus
End Sub

