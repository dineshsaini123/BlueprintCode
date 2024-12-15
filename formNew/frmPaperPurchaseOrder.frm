VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPaperPurchaseOrder 
   Caption         =   "Purchase Order (Paper)..."
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   18384
   Icon            =   "frmPaperPurchaseOrder.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   18384
   Begin VB.Frame panel 
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   8985
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   18312
      Begin VB.TextBox txtRate 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1512
         TabIndex        =   30
         Top             =   1512
         Width           =   5040
      End
      Begin VB.TextBox txtOrderNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1515
         TabIndex        =   0
         Top             =   675
         Width           =   1068
      End
      Begin VB.TextBox txtParty 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8220
         TabIndex        =   4
         Top             =   810
         Width           =   5010
      End
      Begin VB.TextBox txtRemarks 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8244
         TabIndex        =   32
         Top             =   1500
         Width           =   5040
      End
      Begin VB.Frame Frame3 
         Caption         =   "Output Weight"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   210
         TabIndex        =   16
         Top             =   8970
         Visible         =   0   'False
         Width           =   465
         Begin VB.TextBox txtRawAndCasting 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3690
            TabIndex        =   17
            Text            =   "0"
            Top             =   1035
            Width           =   1320
         End
         Begin VB.Label Label1 
            Caption         =   "Raw Issue Weight"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   3
            Left            =   135
            TabIndex        =   20
            Top             =   570
            Width           =   1635
         End
         Begin VB.Label Label1 
            Caption         =   "Receiving  from casting"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   5
            Left            =   150
            TabIndex        =   19
            Top             =   885
            Width           =   2325
         End
         Begin VB.Shape Shape1 
            Height          =   615
            Left            =   0
            Top             =   1590
            Width           =   3135
         End
         Begin VB.Label Label9 
            Caption         =   "Receiving Semi Finish &&  Finish"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   150
            TabIndex        =   18
            Top             =   1515
            Width           =   3060
         End
         Begin VB.Shape Shape2 
            Height          =   585
            Left            =   75
            Top             =   1365
            Width           =   3150
         End
      End
      Begin VB.Frame buttonFrame 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   360
         TabIndex        =   15
         Top             =   7980
         Width           =   7560
         Begin VB.CommandButton cmdAdd_1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Add"
            Height          =   615
            Left            =   75
            Picture         =   "frmPaperPurchaseOrder.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   75
            Width           =   1230
         End
         Begin VB.CommandButton cmdSave_2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Save"
            Height          =   615
            Left            =   1305
            Picture         =   "frmPaperPurchaseOrder.frx":0BF0
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   75
            Width           =   1230
         End
         Begin VB.CommandButton cmdDelete_3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Delete"
            Height          =   615
            Left            =   2535
            Picture         =   "frmPaperPurchaseOrder.frx":17D4
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   75
            Width           =   1230
         End
         Begin VB.CommandButton cmdEdit_4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Edit"
            Height          =   615
            Left            =   3765
            Picture         =   "frmPaperPurchaseOrder.frx":23B8
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   75
            Width           =   1230
         End
         Begin VB.CommandButton cmdPrint_7 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Print"
            Height          =   615
            Left            =   5010
            Picture         =   "frmPaperPurchaseOrder.frx":27FA
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   75
            Width           =   1230
         End
         Begin VB.CommandButton cmdExit_12 
            BackColor       =   &H00FFFFFF&
            Caption         =   "E&xit"
            Height          =   615
            Left            =   6255
            Picture         =   "frmPaperPurchaseOrder.frx":33DE
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   75
            Width           =   1230
         End
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   11088
         TabIndex        =   14
         Top             =   7632
         Width           =   1392
      End
      Begin VB.TextBox txtLoose 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6540
         TabIndex        =   13
         Top             =   8415
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdMaster 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   13272
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   765
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.TextBox txtTotal1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6660
         TabIndex        =   11
         Text            =   "0"
         Top             =   8385
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.ComboBox txtFirmName 
         Height          =   288
         ItemData        =   "frmPaperPurchaseOrder.frx":3FC2
         Left            =   8220
         List            =   "frmPaperPurchaseOrder.frx":3FC4
         TabIndex        =   3
         Top             =   450
         Width           =   5010
      End
      Begin VB.TextBox txtIndent 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1515
         TabIndex        =   2
         Top             =   1125
         Width           =   5040
      End
      Begin Crystal.CrystalReport CR 
         Left            =   9075
         Top             =   8175
         _ExtentX        =   593
         _ExtentY        =   593
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin MSComCtl2.DTPicker txtDates 
         Height          =   315
         Left            =   3285
         TabIndex        =   1
         Top             =   630
         Width           =   1395
         _ExtentX        =   2455
         _ExtentY        =   550
         _Version        =   393216
         Format          =   504823809
         CurrentDate     =   39500
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   0
         Top             =   9165
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   572
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "filedsn=saru;"
         OLEDBString     =   "filedsn=saru;"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "ItemMaster"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VSFlex7Ctl.VSFlexGrid vs 
         Height          =   5580
         Left            =   360
         TabIndex        =   33
         Top             =   1980
         Width           =   17616
         _cx             =   31073
         _cy             =   9842
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
         BackColorFixed  =   12640511
         ForeColorFixed  =   8388608
         BackColorSel    =   16777153
         ForeColorSel    =   8404992
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
         AllowUserResizing=   3
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   100
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPaperPurchaseOrder.frx":3FC6
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rates settled :"
         Height          =   300
         Index           =   9
         Left            =   324
         TabIndex        =   31
         Top             =   1512
         Width           =   1092
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H0078CFE9&
         BorderWidth     =   4
         Height          =   870
         Left            =   315
         Top             =   7920
         Width           =   7665
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Order No :"
         Height          =   270
         Index           =   0
         Left            =   300
         TabIndex        =   28
         Top             =   690
         Width           =   1080
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   300
         Index           =   2
         Left            =   7224
         TabIndex        =   27
         Top             =   852
         Width           =   588
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date "
         Height          =   270
         Index           =   1
         Left            =   2790
         TabIndex        =   26
         Top             =   675
         Width           =   510
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks :"
         Height          =   300
         Index           =   4
         Left            =   7188
         TabIndex        =   25
         Top             =   1500
         Width           =   1092
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   10332
         TabIndex        =   24
         Top             =   7668
         Width           =   708
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "F2 For Search Order"
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
         Left            =   1500
         TabIndex        =   23
         Top             =   315
         Width           =   2505
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Firm Name :"
         Height          =   276
         Index           =   7
         Left            =   7224
         TabIndex        =   22
         Top             =   492
         Width           =   996
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Indent No :"
         Height          =   300
         Index           =   8
         Left            =   315
         TabIndex        =   21
         Top             =   1125
         Width           =   1005
      End
   End
End
Attribute VB_Name = "frmPaperPurchaseOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit As Boolean
Sub setWidth()


vs.Cols = 8

vs.FormatString = "SN.|Paper Mill|Quality|G.S.M.|Reels/Sheet|Size in Cms.|Quantity in Wt.|Delivery"
vs.ColWidth(0) = 700
vs.ColWidth(1) = 3400
vs.ColWidth(2) = 1500
vs.ColWidth(3) = 1200
vs.ColWidth(4) = 1500
vs.ColWidth(5) = 1800
vs.ColWidth(6) = 2000
vs.ColWidth(7) = 5000

End Sub
Sub Total()

txtTotal = 0

For I = 1 To vs.rows - 1

If vs.TextMatrix(I, 6) <> "" Then

txtTotal = Val(txtTotal) + Val(vs.TextMatrix(I, 6))


End If

Next


txtTotal = Round(txtTotal, 0)


End Sub
Private Sub cmdAdd_1_Click()

refresh_
cmdEdit_4.Enabled = False
cmdDelete_3.Enabled = False
cmdSave_2.Enabled = True


End Sub
Sub refresh_()

Edit = False
txtIndent.text = ""
txtRemarks.text = ""
txtParty.text = ""
'txtDespatched.text = ""
txtOrderNo.SetFocus

vs.Clear
setWidth

txtOrderNo.text = MaxSNo("Paper_PurchaseOrder", "OrderNo")

End Sub

Private Sub cmdDelete_3_Click()
If MsgBox("Want to delete ?", vbQuestion + vbYesNo) = vbYes Then

con.Execute "delete from Paper_PurchaseOrder where orderno='" & txtOrderNo & "'"
con.Execute "delete from Paper_PurchaseOrderdet where orderno='" & txtOrderNo & "'"

refresh_
txtOrderNo.SetFocus

End If

End Sub

Private Sub cmdEdit_4_Click()
Edit = True
cmdEdit_4.Enabled = False
cmdDelete_3.Enabled = True
cmdSave_2.Enabled = True
cmdSave_2.SetFocus
End Sub

Private Sub cmdExit_12_Click()
Unload Me
End Sub

Private Sub cmdMaster_Click()
HeadTbl = "Paperdealer"
frmMasters.Show 1
End Sub

Private Sub cmdPrint_7_Click()


DSNNew

If MsgBox("Want To Print ?", vbInformation + vbYesNo) = vbYes Then
    cr.Reset
    
    If Trim(txtFirmName.text) = "BLUEPRINT EDUCATION" Then
     cr.ReportFileName = rptPath & "/PaperPOrder.rpt"
    ElseIf Trim(txtFirmName.text) = "CHITRA PRAKASHAN (I) PVT LTD" Then
     cr.ReportFileName = rptPath & "/PaperPOrderCH.rpt"
    ElseIf Trim(txtFirmName.text) = "RAJLUXMI PUBLICATIONS" Then
     cr.ReportFileName = rptPath & "/PaperPOrderRL.rpt"
    End If
    
    cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    cr.ReplaceSelectionFormula "{Paper_PurchaseOrder.orderno}=" & txtOrderNo.text & ""
    'CR.Formulas(0) = "address='" & Text1.Text & "'"
    cr.WindowShowPrintSetupBtn = True
    cr.WindowState = crptMaximized
    cr.Action = 1
End If

End Sub

Private Sub cmdSave_2_Click()

If Edit = False Then

    If RS.State = 1 Then RS.close
    RS.Open "select * from Paper_PurchaseOrder", con, adOpenDynamic, adLockOptimistic
    RS.AddNew
    RS!orderNo = txtOrderNo.text
    RS!ORDERDATE = txtDates.value
    RS!IndentNo = txtIndent.text
    RS!subledger = txtParty.text
    RS!firmname = txtFirmName.text
    RS!OrderTo = txtOrderNo.text
    RS!rates = Trim(txtRate.text)
    RS!remarks = txtRemarks.text
    RS.update
    
    If RS.State = 1 Then RS.close
    RS.Open "select * from Paper_PurchaseOrderdet", con, adOpenDynamic, adLockOptimistic
    
    For I = 1 To vs.rows - 1
      If (vs.TextMatrix(I, 1) <> "") Then
        RS.AddNew
        RS!orderNo = txtOrderNo.text
        RS!PaperMill = vs.TextMatrix(I, 1)
        RS!quality = vs.TextMatrix(I, 2)
        RS!GSM = vs.TextMatrix(I, 3)
        RS!Reels_Sheet = vs.TextMatrix(I, 4)
        RS!Size_cm = vs.TextMatrix(I, 5)
        RS!qty = vs.TextMatrix(I, 6)
        RS!Delivery = vs.TextMatrix(I, 7)
        
      If rs1.State = 1 Then rs1.close
        rs1.Open "select Address,GSTIN,ContactName,ContactNo from Godownmaster where Godwn='" & vs.TextMatrix(I, 7) & "'", con
        If rs1.EOF = False Then
           RS!Address = rs1!Address & ""
           RS!GSTIN = "GSTIN : " & rs1!GSTIN
           RS!ContactName = "Contact Person : " & rs1!ContactName & " Mob: " & rs1!ContactNo
           RS!ContactNo = rs1!ContactNo
        End If
        
        
        
        
        
        RS.update
      End If
    Next

    cmdSave_2.Enabled = False
Else

    If RS.State = 1 Then RS.close
    RS.Open "select * from Paper_PurchaseOrder where OrderNo='" & txtOrderNo.text & "'", con, adOpenDynamic, adLockOptimistic
    If RS.EOF = False Then
        RS!orderNo = txtOrderNo.text
        RS!ORDERDATE = txtDates.value
        RS!IndentNo = txtIndent.text
        RS!subledger = txtParty.text
        RS!firmname = txtFirmName.text
        RS!OrderTo = txtOrderNo.text
        RS!rates = Trim(txtRate.text)
        RS!remarks = txtRemarks.text
        
        
        
        RS.update
    End If
    
    
    con.Execute "delete from Paper_PurchaseOrderdet where orderno='" & txtOrderNo & "'"
    
    For I = 1 To vs.rows - 1
      If (vs.TextMatrix(I, 1) <> "") Then
        
        If RS.State = 1 Then RS.close
        RS.Open "select * from Paper_PurchaseOrderdet where (orderNo='" & txtOrderNo & "' and sn='" & vs.TextMatrix(I, 0) & "')", con, adOpenDynamic, adLockOptimistic
        If RS.EOF = True Then
           RS.AddNew
        End If
        
        RS!orderNo = txtOrderNo.text
        RS!PaperMill = vs.TextMatrix(I, 1)
        RS!quality = vs.TextMatrix(I, 2)
        RS!GSM = vs.TextMatrix(I, 3)
        
        RS!Reels_Sheet = vs.TextMatrix(I, 4)
        RS!Size_cm = vs.TextMatrix(I, 5)
        RS!qty = vs.TextMatrix(I, 6)
        RS!Delivery = vs.TextMatrix(I, 7)

        
 
            
      If rs1.State = 1 Then rs1.close
        rs1.Open "select Address,GSTIN,ContactName,ContactNo from Godownmaster where Godwn='" & vs.TextMatrix(I, 7) & "'", con
        If rs1.EOF = False Then
           RS!Address = rs1!Address & ""
           RS!GSTIN = "GSTIN : " & rs1!GSTIN
           RS!ContactName = "Contact Person : " & rs1!ContactName & " Mob: " & rs1!ContactNo
           RS!ContactNo = rs1!ContactNo
        End If
            
        RS.update
        
     
     End If
     
    Next




End If


cmdSave_2.Enabled = True
MsgBox "Data Save ...", vbInformation, "Saved.."


End Sub

Private Sub Form_Load()
Me.Left = 50
Me.top = 10
Edit = False
'Me.Width = 14200
'Me.Height = 8500

Me.Width = 18400
Me.Height = 9610

setWidth
 
txtDates.value = Date
txtOrderNo.text = MaxSNo("Paper_PurchaseOrder", "OrderNo")

Dim s As String


BackColorFrom Me

txtFirmName.Clear
If RS.State = 1 Then RS.close
RS.Open "select FirmName,Add1,Add2 from FirmMaster order by firmname", con, adOpenStatic, adLockReadOnly
While RS.EOF = False
 txtFirmName.AddItem RS(0)
 RS.MoveNext
Wend

txtFirmName.ListIndex = 0

st_ = ""
If rs1.State = 1 Then rs1.close
rs1.Open "select Godwn as [Binder Name],Address from Godownmaster where " & stringyear & " and (Binder_Printer='b' or Binder_Printer='pb') order by Godwn", con, adOpenStatic, adLockReadOnly
While rs1.EOF = False
  If st_ = "" Then
     st_ = rs1(0)
  Else
     st_ = st_ & "|" & rs1(0)
  End If
  rs1.MoveNext
Wend

vs.ColComboList(7) = st_


End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then sendkeys ("{tab}")
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then sendkeys ("{tab}")
End Sub

Private Sub txtDates_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then sendkeys ("{tab}")
End Sub


Private Sub txtDespatched_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  txtRemarks.SetFocus
End If
End Sub

Private Sub txtFirmName_Click()
   createIndent
End Sub
Private Sub txtFirmName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
     txtParty.SetFocus
  End If
End Sub
Sub createIndent()
    If txtFirmName.text = "BLUEPRINT EDUCATION " Then
       txtIndent.text = "BP/" & session & "/" & txtOrderNo
    ElseIf txtFirmName.text = "CHITRA PRAKASHAN (I) PVT LTD" Then
       txtIndent.text = "CP/" & session & "/" & txtOrderNo
    ElseIf txtFirmName.text = "RAJLUXMI PUBLICATIONS" Then
       txtIndent.text = "RL/" & session & "/" & txtOrderNo
    End If
End Sub
Private Sub txtFirmName_LostFocus()
createIndent
End Sub

Private Sub txtIndent_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     txtFirmName.SetFocus
  End If
End Sub

Private Sub txtOrderNo_GotFocus()
If PopUpValue1 <> "" Then
   txtOrderNo.text = PopUpValue1
   
   searchData
   
   PopUpValue1 = ""
   PopUpValue2 = ""
   PopUpValue3 = ""
   popupvalue4 = ""
End If
End Sub
Sub searchData()


If RS.State = 1 Then RS.close
RS.Open "select * from Paper_PurchaseOrder where orderNo='" & txtOrderNo & "'", con, adOpenDynamic, adLockOptimistic
If RS.EOF = False Then

    cmdDelete_3.Enabled = False
    cmdSave_2.Enabled = False
    cmdEdit_4.Enabled = True


    txtOrderNo.text = RS!orderNo
    txtDates.value = RS!ORDERDATE
    txtIndent.text = RS!IndentNo
    txtParty.text = RS!subledger
    txtFirmName.text = RS!firmname
    txtOrderNo.text = RS!OrderTo
    'txtDespatched.text = RS!DespatchedBy
    txtRate.text = RS!rates & ""
    txtRemarks.text = RS!remarks
    
End If
    
vs.Clear
setWidth
vs.rows = 50

If RS.State = 1 Then RS.close
RS.Open "select * from Paper_PurchaseOrderdet where orderno='" & txtOrderNo & "' order by sn", con, adOpenDynamic, adLockOptimistic
For I = 1 To RS.RecordCount
    
    vs.TextMatrix(I, 0) = I
    vs.TextMatrix(I, 1) = RS!PaperMill
    vs.TextMatrix(I, 2) = RS!quality
    vs.TextMatrix(I, 3) = RS!GSM
    vs.TextMatrix(I, 4) = RS!Reels_Sheet & ""
    vs.TextMatrix(I, 5) = RS!Size_cm
    vs.TextMatrix(I, 6) = RS!qty
    vs.TextMatrix(I, 7) = RS!Delivery
    
    
    RS.MoveNext
Next

Total
   
End Sub
Private Sub txtOrderNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then sendkeys ("{tab}")

If KeyCode = 113 Then
   value = "select OrderNo,OrderDATE,SUBLEDGER as Party from Paper_PurchaseOrder order by orderno"
   popuplist1 value, con
End If

End Sub

Private Sub txtParty_GotFocus()

If PopUpValue1 <> "" Then
   txtParty.text = PopUpValue1
   PopUpValue1 = ""
End If

End Sub
Private Sub txtParty_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then
        searchType = "party"
        value = "select distinct(Party),Code,subledger from SLEDGER where gledger='SUNDRY CREDITORS' and " & stringyear & "  order by party"
        popuplist_client value, CCON
        set_focus = True
End If


If KeyCode = 13 Then
  txtRate.SetFocus
End If
End Sub

Private Sub txtRate_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
   txtRemarks.SetFocus
End If

End Sub

Private Sub txtRemarks_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then sendkeys ("{tab}")
End Sub
Private Sub vs_GotFocus()
If PopUpValue1 <> "" Then
   vs.TextMatrix(vs.RowSel, 0) = vs.Row
   vs.TextMatrix(vs.RowSel, 1) = PopUpValue1
   vs.TextMatrix(vs.RowSel, 2) = PopUpValue2
   vs.TextMatrix(vs.RowSel, 3) = popupvalue5
   vs.TextMatrix(vs.RowSel, 4) = PopUpValue3
   vs.TextMatrix(vs.RowSel, 5) = popupvalue4
   
   vs.Col = 6
   PopUpValue1 = ""
   PopUpValue2 = ""
   PopUpValue3 = ""
   popupvalue4 = ""
   popupvalue5 = ""
   
End If
End Sub
Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)
If vs.Col = 1 Then
If (KeyCode = 113 Or KeyCode = 13) Then
   searchType = "paper"
   value = "Select papermaker_name as [Paper Name],Eco,Size as [Reels/Sheet],SizeValue1 + ' X ' +SizeValue2 as PSize, GSM,papermaker_Id as Code from papermakemaster where papermaker_id <> '' and  " & stringyear
   popuplist10 value, con
End If

End If


If KeyCode = 115 Then
   If MsgBox("Want to delete ?", vbQuestion + vbYesNo) = vbYes Then
      con.Execute "delete from Paper_PurchaseOrderDet where orderno='" & txtOrderNo.text & "' and papermill='" & vs.TextMatrix(vs.RowSel, 1) & "' and delivery='" & vs.TextMatrix(vs.RowSel, 7) & "'"
      vs.RemoveItem (vs.RowSel)
      Total
   End If
End If


End Sub
Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If vs.Col = 4 Then
       If KeyCode = 13 Then
          If vs.TextMatrix(vs.RowSel, 4) <> "" Then
          sendkeys "{right}"
          End If
       End If
    ElseIf vs.Col = 5 Then
       If KeyCode = 13 Then
          If vs.TextMatrix(vs.RowSel, 5) <> "" Then
          sendkeys "{right}"
          End If
       End If
    ElseIf vs.Col = 6 Then
       If KeyCode = 13 Then
          If vs.TextMatrix(vs.RowSel, 6) <> "" Then
          sendkeys "{right}"
          End If
       End If
    
    ElseIf vs.Col = 7 Then
       If KeyCode = 13 Then
          If vs.TextMatrix(vs.RowSel, 6) <> "" Then
             sendkeys "{home}"
             sendkeys "{down}"
             Total
          End If
       End If
    End If
End If
End Sub

Private Sub vs_SelChange()

  If vs.Col <= 2 Then
     vs.Editable = flexEDNone
  Else
     vs.Editable = flexEDKbdMouse
     
  End If

End Sub
