VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBinderRecChallan 
   Caption         =   "Binder Book  Receive"
   ClientHeight    =   7824
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   12936
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.4
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7824
   ScaleWidth      =   12936
   Begin VB.Frame panel 
      Caption         =   "Book Receive From Binder"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   11.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   7752
      Left            =   60
      TabIndex        =   13
      Top             =   180
      Width           =   12825
      Begin VB.ComboBox Bookcode 
         Height          =   2184
         ItemData        =   "frmBinderRecChallan.frx":0000
         Left            =   120
         List            =   "frmBinderRecChallan.frx":0002
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   42
         Top             =   2220
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.CommandButton cmdAddRemarks 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Add Remarks"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   11400
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   1680
         Width           =   1035
      End
      Begin VB.OptionButton Option2_old 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Old Book"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3825
         TabIndex        =   39
         Top             =   900
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.OptionButton Option1_new 
         BackColor       =   &H00FFFFFF&
         Caption         =   "New Book"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3825
         TabIndex        =   38
         Top             =   525
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtTotal1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7560
         TabIndex        =   29
         Text            =   "0"
         Top             =   5970
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Frame buttonFrame 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   780
         Left            =   1935
         TabIndex        =   28
         Top             =   6645
         Width           =   10452
         Begin VB.CommandButton Command1_Binder 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Binder Ledger"
            Height          =   705
            Left            =   7884
            Picture         =   "frmBinderRecChallan.frx":0004
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   36
            Width           =   1275
         End
         Begin VB.CommandButton cmdExit_12 
            BackColor       =   &H00FFFFFF&
            Caption         =   "E&xit"
            Height          =   705
            Left            =   9192
            Picture         =   "frmBinderRecChallan.frx":0BE8
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   45
            Width           =   1200
         End
         Begin VB.CommandButton cmdPrint_7 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Print"
            Height          =   705
            Left            =   6612
            Picture         =   "frmBinderRecChallan.frx":17CC
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   45
            Width           =   1236
         End
         Begin VB.CommandButton cmdUndo_5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Undo"
            Height          =   705
            Left            =   5292
            Picture         =   "frmBinderRecChallan.frx":23B0
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   45
            Width           =   1275
         End
         Begin VB.CommandButton cmdEdit_4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Edit"
            Height          =   705
            Left            =   4005
            Picture         =   "frmBinderRecChallan.frx":2BF4
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   45
            Width           =   1236
         End
         Begin VB.CommandButton cmdDelete_3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Delete"
            Height          =   705
            Left            =   2685
            Picture         =   "frmBinderRecChallan.frx":3036
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   45
            Width           =   1275
         End
         Begin VB.CommandButton cmdSave_2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Save"
            Height          =   705
            Left            =   1365
            Picture         =   "frmBinderRecChallan.frx":3C1A
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   45
            Width           =   1275
         End
         Begin VB.CommandButton cmdAdd_1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Add"
            Height          =   705
            Left            =   45
            Picture         =   "frmBinderRecChallan.frx":47FE
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   45
            Width           =   1275
         End
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3330
         TabIndex        =   27
         Top             =   5970
         Width           =   1455
      End
      Begin VB.TextBox txtLoose 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6270
         TabIndex        =   26
         Top             =   5970
         Width           =   1215
      End
      Begin VB.TextBox txtHeating 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3105
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   495
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.TextBox txtId 
         Height          =   285
         Left            =   1725
         TabIndex        =   0
         Top             =   495
         Width           =   1410
      End
      Begin VB.TextBox txtParty 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7305
         TabIndex        =   3
         Top             =   360
         Width           =   4470
      End
      Begin VB.TextBox txtRemarks 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1725
         TabIndex        =   6
         Top             =   1500
         Width           =   4065
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7305
         TabIndex        =   15
         Top             =   690
         Width           =   4470
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7305
         TabIndex        =   14
         Top             =   990
         Width           =   4470
      End
      Begin VB.ComboBox cbogodown 
         Height          =   315
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1125
         Width           =   1395
      End
      Begin VB.TextBox txtTopay 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7320
         TabIndex        =   4
         Top             =   1305
         Width           =   1275
      End
      Begin VB.ComboBox cboFirm 
         Height          =   315
         Left            =   7320
         TabIndex        =   5
         Top             =   1620
         Width           =   3705
      End
      Begin MSComCtl2.DTPicker Dates 
         Height          =   315
         Left            =   1725
         TabIndex        =   1
         Top             =   780
         Width           =   1395
         _ExtentX        =   2455
         _ExtentY        =   550
         _Version        =   393216
         Format          =   138543105
         CurrentDate     =   39500
      End
      Begin VSFlex7Ctl.VSFlexGrid vs 
         Height          =   3804
         Left            =   120
         TabIndex        =   7
         Top             =   2052
         Width           =   12336
         _cx             =   21759
         _cy             =   6710
         _ConvInfo       =   1
         Appearance      =   2
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
         BackColorFixed  =   7917545
         ForeColorFixed  =   -2147483630
         BackColorSel    =   12907771
         ForeColorSel    =   16711680
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   8388608
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
         Rows            =   150
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   800
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
         ExplorerBar     =   7
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
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "F1 For Search Books"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   120
         TabIndex        =   41
         Top             =   5880
         Width           =   1545
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H0078CFE9&
         BorderWidth     =   4
         Height          =   876
         Left            =   1896
         Top             =   6600
         Width           =   10560
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total :"
         Height          =   195
         Index           =   6
         Left            =   2250
         TabIndex        =   30
         Top             =   5970
         Width           =   660
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "F2 For Search Binder"
         Height          =   285
         Left            =   7290
         TabIndex        =   25
         Top             =   135
         Width           =   2505
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "F2 For Search Challan"
         Height          =   285
         Left            =   1710
         TabIndex        =   24
         Top             =   270
         Width           =   2505
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Challan No :"
         Height          =   270
         Index           =   0
         Left            =   210
         TabIndex        =   23
         Top             =   510
         Width           =   1530
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Binder Name "
         Height          =   300
         Index           =   2
         Left            =   6060
         TabIndex        =   21
         Top             =   390
         Width           =   1245
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date "
         Height          =   270
         Index           =   1
         Left            =   210
         TabIndex        =   20
         Top             =   795
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks "
         Height          =   300
         Index           =   4
         Left            =   210
         TabIndex        =   19
         Top             =   1500
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Godown Name"
         Height          =   270
         Index           =   7
         Left            =   210
         TabIndex        =   18
         Top             =   1125
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "To Pay"
         Height          =   300
         Index           =   8
         Left            =   6030
         TabIndex        =   17
         Top             =   1350
         Width           =   1140
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Firm Name :"
         Height          =   300
         Index           =   9
         Left            =   6060
         TabIndex        =   16
         Top             =   1665
         Width           =   1140
      End
   End
   Begin Crystal.CrystalReport CR 
      Left            =   12240
      Top             =   7260
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
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
      Left            =   75
      TabIndex        =   8
      Top             =   8805
      Visible         =   0   'False
      Width           =   465
      Begin VB.TextBox txtRawAndCasting 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3690
         TabIndex        =   9
         Text            =   "0"
         Top             =   1035
         Width           =   1320
      End
      Begin VB.Shape Shape2 
         Height          =   585
         Left            =   75
         Top             =   1365
         Width           =   3150
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
         TabIndex        =   12
         Top             =   1515
         Width           =   3060
      End
      Begin VB.Shape Shape1 
         Height          =   615
         Left            =   0
         Top             =   1590
         Width           =   3135
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
         TabIndex        =   11
         Top             =   885
         Width           =   2325
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
         TabIndex        =   10
         Top             =   570
         Width           =   1635
      End
   End
End
Attribute VB_Name = "frmBinderRecChallan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim itemgp As String
Dim rates As Double
Dim I As Integer
Dim Status As String
Dim Item_Name As String
Dim unit As String
Dim qty As Integer
Dim iitem1 As String
Dim Edit As Boolean
Dim rs2 As New ADODB.Recordset
Dim StockFlag As String

Private Sub cmdMain_Click()
Unload Me

End Sub
Sub cellposi()
 'VsFrame.Width = 3165
 'VsFrame.Top = vs.Top + ((vs.CellTop)) - 1400
 'VsFrame.Left = (vs.Left) - 200
End Sub
Sub Total()
txtTotal.text = 0
txtLoose.text = 0

For J = 1 To vs.rows - 1
If vs.TextMatrix(J, 0) <> "" Then
txtTotal.text = (Val(txtTotal.text) + Val(vs.TextMatrix(J, 1)))
txtLoose.text = (Val(txtLoose.text) + Val(vs.TextMatrix(J, 3)))
End If
Next

End Sub

Sub cellposiVs()
 Vs1Frame.Width = 2500
 Vs1Frame.top = vs1.top + ((vs1.CellTop))
 Vs1Frame.Left = (vs1.Left) + 550
End Sub

Sub AddItemInGrid()
    
'    Dim rs_1 As New ADODB.Recordset
'
'    rs_1.Open "select * from books order by Bookname", con, adOpenDynamic, adLockOptimistic
'
'    Set cboItem.RowSource = rs_1
'    cboItem.ListField = "Bookcode"
'    cboItem.BoundColumn = "Bookname"
'    cboItem.ReFill
    
End Sub
Private Sub cboFinishItem_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtQty.SetFocus
End Sub
Private Sub Bookcode_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 27 Then
    Bookcode.Visible = False
    vs.SetFocus
 End If
End Sub

Private Sub Bookcode_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      vs.TextMatrix(vs.RowSel, 0) = Mid(Bookcode.text, InStr(Bookcode.text, ":") + 1)
      Bookcode.Visible = False
      vs.SetFocus
      vs.Col = 0
   End If
End Sub

Private Sub cboFirm_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Len(cboFirm.text) > 0 Then
txtRemarks.SetFocus
End If
End If
End Sub

Private Sub cbogodown_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then txtTopay.SetFocus
End Sub

Private Sub cbogodown_LostFocus()
If cboGodown = "" Then
   MsgBox "Select Godown Name ..", vbCritical
   cboGodown.SetFocus
   Exit Sub
End If
End Sub

Private Sub cboItem_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 13 Then
        cellposi
        If cboItem.text = "" Then
           VsFrame.Visible = False
           cmdSave_2.SetFocus
        Exit Sub
        End If
        vs.TextMatrix(vs.RowSel, 0) = cboItem.text
        vs.TextMatrix(vs.RowSel, 6) = cboItem.BoundText
        
        
        vs.SetFocus
        
     ElseIf KeyCode = 27 Then
       
          VsFrame.Visible = False
        
     End If
End Sub
Sub saveInMaster()
         On Error Resume Next
      
         If RS.State = 1 Then RS.close
         RS.Open "select * from ItemMaster where " & stringyear & " and ItemName='" & iitem1 & "'", con, adOpenDynamic, adLockOptimistic
         If RS.EOF = True Then
            RS.AddNew
            RS.Fields("ItemGp").value = frmAddMaster.cbogp.text
            RS.Fields("ItemName").value = iitem1
            RS.Fields("Unit").value = "Kg"
            RS.update
         Else
            MsgBox "This Item Already Exist !!", vbCritical
            Exit Sub
         End If
         frmAddMaster.Visible = False
  
End Sub
Private Sub cboitemvs1_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 13 Then
        cellposiVs
        vs1.TextMatrix(vs1.RowSel, 0) = cboitemvs1.text
        
        If RS.State = 1 Then RS.close
        RS.Open "select * from ItemMaster where " & stringyear & " and ItemName='" & cboitemvs1.text & "'", con, adOpenKeyset, adLockReadOnly
        If RS.EOF = True Then
            If MsgBox("Want To Add In Master ?", vbQuestion + vbYesNo) = vbYes Then
                frmAddMaster.Show 1
                iitem1 = cboitemvs1.text
                saveInMaster
                cboitemvs1.text = ""
                Vs1Frame.Visible = False
                vs1.SetFocus
             End If
        End If
        vs1.SetFocus
     ElseIf KeyCode = 27 Then
        Vs1Frame.Visible = False
     End If
End Sub
Private Sub cboItemVs2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
        
        
        'cellposiVs2
        vs3.TextMatrix(vs3.RowSel, 0) = cboItemVs2.text
        Set RS = New ADODB.Recordset
        If RS.State = 1 Then RS.close
        RS.Open "select * from ItemMaster where " & stringyear & " and ItemName='" & cboItemVs2.text & "'", con
        If RS.EOF = True Then
            If MsgBox("Want To Add In Master ?", vbQuestion + vbYesNo) = vbYes Then
                frmAddMaster.Show 1
                iitem1 = cboItemVs2.text
                saveInMaster
                
                cboItemVs2.text = ""
             End If
        End If
        vs3.SetFocus
        
ElseIf KeyCode = 27 Then
         FrameVs2.Visible = False
End If
End Sub

Private Sub cboItemVs3_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 13 Then
        
        vs2.TextMatrix(vs2.RowSel, 0) = cboItemVs3.text
        Set RS = New ADODB.Recordset
        If RS.State = 1 Then RS.close
        RS.Open "select * from ItemMaster where " & stringyear & " and ItemName='" & cboItemVs3.text & "'", con
        If RS.EOF = True Then
            If MsgBox("Want To Add In Master ?", vbQuestion + vbYesNo) = vbYes Then
                Vs3Frame.Visible = False
                iitem1 = cboItemVs3.text
                frmAddMaster.Show 1
                saveInMaster
                cboItemVs3.text = ""
                vs2.SetFocus
             End If
        End If
        Vs3Frame.Visible = False
         vs2.SetFocus
     ElseIf KeyCode = 27 Then
        Vs3Frame.Visible = False
     End If

End Sub

Private Sub cmdAdd_Click()
 If RS.State = 1 Then RS.close
 RS.Open "select HeatingNo from IssueMaster where " & stringyear & " and HeatingDate >=datevalue('" & fromdate.value & "') and HeatingDate <=datevalue('" & todate.value & "') order by HeatingNo", con
 ListHeatingNo.Clear
 If RS.EOF = False Then
    While RS.EOF = False
       ListHeatingNo.AddItem RS(0)
       RS.MoveNext
    Wend
 End If
End Sub
Private Sub cmdDelete_Click()
   
   If MsgBox("Want To Delete ?", vbQuestion + vbYesNo) = vbYes Then
      
      Call cmdref_Click
      
   End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFatch_Click()
AddSemifinish

End Sub

Private Sub cmdFind_Click()
 Frame1.Visible = True
 fromdate.SetFocus
End Sub

Private Sub cmdModify_Click()
   If MsgBox("Want To Modify ?", vbQuestion + vbYesNo) = vbYes Then
      
      UpdateIssue
      
      Call cmdref_Click
   End If
End Sub
Sub UpdateIssue()

Dim rss As New ADODB.Recordset
Dim search As New ADODB.Recordset
    
If search.State = 1 Then search.close
search.Open "select ItemName,qty from Invoice where " & stringyear & " and HeatNo='" & txtHeating.text & "'", con
If search.EOF = False Then
While search.EOF = False

    If rss.State = 1 Then rss.close
    rss.Open "select * from IssueRawMetrial where " & stringyear & " and HeatingNo=" & txtHeating.text & " and ItemName='" & search.Fields(0).value & "'", con, adOpenDynamic, adLockOptimistic
    If rss.EOF = False Then
       rss.Fields("Issue").value = (CDbl(rss.Fields("Issue").value) + CDbl(search.Fields("qty").value))
       rss.update
    End If
    
    search.MoveNext
    
Wend
  
End If
  
End Sub

Private Sub cmdref_Click()
      txtHeating.text = ""
      txtParty.text = ""
      
      txtRemarks.text = ""
      
      
      txtTotal1.text = 0
      txtTotal2.text = 0
      txtTotal3.text = 0
      txtTotal4.text = 0
      
      txtSize.text = ""
      txtGrade.text = ""
      txtRawAndCasting.text = 0
      
      vs.Clear
      vs1.Clear
      vs2.Clear
      vs3.Clear
      
      setWidth
      txtHeating.SetFocus
      cmdDelete.Enabled = False
      cmdModify.Enabled = False
      cmdSave.Enabled = True
      
      Record = ""
      
End Sub


Private Sub Command4_Click()
   Unload Me
End Sub
Private Sub cmdSave_Click()
    
    
    
    
    If RS.State = 1 Then RS.close
    RS.Open "select * from IssueMaster where " & stringyear & " and HeatingNo=" & txtHeating.text & "", con
    If RS.EOF = False Then
       MsgBox "Heating No. Already Exist !!", vbInformation
       Exit Sub
    End If
    
    If txtHeating.text = "" Then
       MsgBox "Please Enter Heating No !!", vbCritical
       txtHeating.SetFocus
       Exit Sub
    End If
    
    
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.close
    RS.Open "select * from IssueMaster where " & stringyear & " and HeatingNo=" & txtHeating.text & "", con
    If RS.EOF = True Then
       If MsgBox("Want To Save ?", vbQuestion + vbYesNo) = vbYes Then
        '  SaveData
       End If
    Else
          MsgBox "Dublicate Heating No !!", vbCritical
    End If
End Sub
Sub ItemGpSearch(Str As String)
    
    If rs1.State = 1 Then rs1.close
    rs1.Open "select ItemGp,Rate from ItemMaster where " & stringyear & " and ItemName='" & Str & "'", con
    If rs1.EOF = False Then
       itemgp = rs1.Fields(0).value
       rates = rs1.Fields(1).value
    End If
    
End Sub
Sub UpdateStock()
    Dim rr As New ADODB.Recordset
    Dim rs_u As New ADODB.Recordset
    Dim openning As Double
    
 
    
    
 '================ Issue For Casting
 
 
 If StockFlag = "1" Then
    
    If rs_u.State = 1 Then rs_u.close
    rs_u.Open "select * from Stock where " & stringyear & " and ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", con, adOpenDynamic, adLockOptimistic
    If rs_u.EOF = True Then
        rs_u.AddNew
        rs_u!itemname = Item_Name
        ItemGpSearch Item_Name
        rs_u!itemgp = itemgp
        rs_u!unit = unit
        rs_u!rate = rates
        rs_u!qty = (-1 * qty)
        rs_u.update
     Else
        rs_u!qty = rs_u!qty - qty
        rs_u.update
    End If
    
 End If
 
 '====================================
 
 
 '================ Receive For Casting
 
 If StockFlag = "2" Then
    
   If itemgp <> "Losses" Then
    
        If rs_u.State = 1 Then rs_u.close
        rs_u.Open "select * from Stock where " & stringyear & " and ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", con, adOpenDynamic, adLockOptimistic
        If rs_u.EOF = True Then
            rs_u.AddNew
            rs_u!itemname = Item_Name
            ItemGpSearch Item_Name
            rs_u!itemgp = itemgp
            rs_u!unit = unit
            rs_u!rate = rates
            rs_u!qty = qty
            rs_u.update
         Else
            rs_u!qty = rs_u!qty + qty
            rs_u.update
        End If
    
    End If
    
 End If
 
 '====================================
 
 
 '================ Receive For Finish
 
 If StockFlag = "3" Then
    
   If itemgp <> "Losses" Then
    
        If rs_u.State = 1 Then rs_u.close
        rs_u.Open "select * from Stock where " & stringyear & " and ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", con, adOpenDynamic, adLockOptimistic
        If rs_u.EOF = True Then
            rs_u.AddNew
            rs_u!itemname = Item_Name
            ItemGpSearch Item_Name
            rs_u!itemgp = itemgp
            rs_u!unit = unit
            rs_u!rate = rates
            rs_u!qty = qty
            rs_u.update
         Else
            rs_u!qty = rs_u!qty + qty
            rs_u.update
        End If
    
    End If
    
 End If
 
 '====================================
 
 
  '================ Issue For Finish
 
 If StockFlag = "4" Then
    
   If itemgp <> "Losses" Then
    
        If rs_u.State = 1 Then rs_u.close
        rs_u.Open "select * from Stock where " & stringyear & " and ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", con, adOpenDynamic, adLockOptimistic
        If rs_u.EOF = True Then
            rs_u.AddNew
            rs_u!itemname = Item_Name
            ItemGpSearch Item_Name
            rs_u!itemgp = itemgp
            rs_u!unit = unit
            rs_u!rate = rates
            rs_u!qty = (-1 * qty)
            rs_u.update
         Else
            rs_u!qty = rs_u!qty - qty
            rs_u.update
        End If
    
    End If
    
 End If
 
 '====================================
    
    
    
    
End Sub
 
 
     
  

Private Sub cmdAdd_1_Click()
   
    Edit = False
    txtHeating.text = ""
    dates.value = Date
    txtParty.text = ""
    txtRemarks.text = ""
    Text1.text = ""
    Text2.text = ""
    txtTotal.text = ""
    txtLoose.text = ""
   txtTopay = ""
   cboFirm.ListIndex = -1
   
   'RefData Me
   vs.Clear
   setWidth
   cmdDelete_3.Enabled = False
   cmdEdit_4.Enabled = True
   'cmdPrint_7.Enabled = False
   cmdSave_2.Enabled = True
   cmdUndo_5.Enabled = True
   cmdAdd_1.Enabled = True
   cmdExit_12.Enabled = True
   'txtHeating.SetFocus
   txtHeating.text = MaxSNo("BinderBkReceive", "INVOICENO")
   txtId.text = MaxSNo("BinderBkReceive", "INVOICENO")
   
   cboGodown.ListIndex = -1
   
   'formButtonValidation cmdDelete_3, cmdEdit_4
   
   'Frame1.Enabled = True
   vs.Enabled = True
   'txtHeating.SetFocus
   txtId.SetFocus
   
End Sub

Private Sub cmdAddRemarks_Click()
HeadTbl = "bookrecfrombinder"
frmMasters.Show 1
End Sub

Private Sub cmdDelete_3_Click()

If rs1.State = 1 Then rs1.close
rs1.Open "select * from BinderBkReceive where " & stringyear & " and INVOICENO=" & txtHeating.text & "", con
If rs1.EOF = False Then
   If rs2.State = 1 Then rs2.close
   rs2.Open "select * from BinderBkReceive where " & stringyear & " and INVOICENO=" & txtHeating.text & "", con
      If rs2!bAuthorized = True Then
          MsgBox "You can'nt change, Bill Already Locked !!", vbExclamation, "Alert"
          Exit Sub
      End If
End If


If MsgBox("Want to delete ?", vbQuestion + vbYesNo) = vbYes Then
    con.Execute "delete from BookReceiveDet where " & stringyear & " and INVOICENO=" & txtHeating.text & ""
    con.Execute "delete from BinderBkReceive where " & stringyear & " and INVOICENO=" & txtHeating.text & ""
    Call cmdAdd_1_Click
End If
End Sub

Private Sub cmdEdit_4_Click()
   
If rs1.State = 1 Then rs1.close
rs1.Open "select * from BinderBkReceive where " & stringyear & " and INVOICENO=" & txtHeating.text & "", con
If rs1.EOF = False Then
   If rs2.State = 1 Then rs2.close
   rs2.Open "select * from BinderBkReceive where " & stringyear & " and INVOICENO=" & txtHeating.text & "", con
      If rs2!bAuthorized = True Then
          MsgBox "You can'nt change, Bill Already Locked !!", vbExclamation, "Alert"
          Exit Sub
      End If
End If
   
   
   cmdDelete_3.Enabled = True
   cmdEdit_4.Enabled = False
   cmdPrint_7.Enabled = True
   cmdSave_2.Enabled = True
   cmdUndo_5.Enabled = True
   cmdAdd_1.Enabled = False
   cmdExit_12.Enabled = True
   'Frame1.Enabled = True
   Edit = True
   txtId.SetFocus
   vs.Enabled = True
End Sub

Private Sub cmdExit_12_Click()
Unload Me
End Sub

Private Sub cmdPrint_7_Click()


DSNNew

CR.Reset
CR.ReportFileName = rptPath & "/CHALLAN_bkrec.rpt"
CR.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
CR.ReplaceSelectionFormula "{BinderBkReceive.invoiceno}=" & txtHeating.text & ""
CR.WindowShowPrintSetupBtn = True
CR.WindowState = crptMaximized
CR.Action = 1

End Sub

Private Sub cmdSave_2_Click()

On Error GoTo aa1

Dim newoldbk As String

If Option1_new.value = True Then
newoldbk = "NEW"
Else
newoldbk = "OLD"
End If


If txtParty.text = "" Then
MsgBox "Please Enter Binder Name !!", vbInformation
Exit Sub
End If

If Edit = False Then
    txtHeating.text = MaxSNo("BinderBkReceive", "INVOICENO")
    txtId.text = txtHeating.text
End If


If MsgBox("Want to Save ?", vbYesNo + vbQuestion) = vbYes Then

  

If RS.State = 1 Then RS.close
RS.Open "select * from BinderBkReceive where " & stringyear & " and INVOICENO=" & txtId.text & "", con, adOpenDynamic, adLockOptimistic
If RS.EOF = True Then
RS.AddNew
RS.Fields("INVOICENO").value = txtHeating.text
RS.Fields("INVOICEDATE").value = dates.value
RS.Fields("SUBLEDGER").value = txtParty.text
RS.Fields("GENLEDGER").value = "Sundry Debtors"
RS.Fields("Remarks").value = txtRemarks.text
RS.Fields("add1").value = Text1.text
RS.Fields("add2").value = Text2.text
RS.Fields("godown").value = cboGodown
RS.Fields("topay").value = Val(txtTopay)
RS.Fields("firmname").value = cboFirm.text
RS.Fields("NewOldBook").value = newoldbk
RS.Fields("fyear").value = main.session
RS.Fields("setupid").value = main.setupid

RS.update
Else

RS.Fields("godown").value = cboGodown
RS.Fields("INVOICEDATE").value = dates.value
RS.Fields("SUBLEDGER").value = txtParty.text
RS.Fields("GENLEDGER").value = "Sundry Debtors"
RS.Fields("Remarks").value = txtRemarks.text
RS.Fields("add1").value = Text1.text
RS.Fields("add2").value = Text2.text
RS.Fields("NetBook").value = Val(txtTotal1.text)
RS.Fields("topay").value = Val(txtTopay)
RS.Fields("firmname").value = cboFirm.text
RS.Fields("NewOldBook").value = newoldbk
RS.Fields("fyear").value = main.session
RS.Fields("setupid").value = main.setupid


RS.update
cmdSave_2.Enabled = False
cmdPrint_7.SetFocus
End If



If RS.State = 1 Then RS.close
RS.Open "select * from BookReceiveDet where " & stringyear & " and INVOICENO=" & txtHeating.text & "", con, adOpenDynamic, adLockOptimistic
If RS.EOF = True Then


For I = 1 To vs.rows - 1

If vs.TextMatrix(I, 0) <> "" Then

RS.AddNew
RS.Fields("INVOICENO").value = txtHeating.text
RS.Fields("INVOICEDATE").value = dates.value
RS.Fields("SUBLEDGER").value = txtParty.text
RS.Fields("GENLEDGER").value = "Sundry Debtors"
RS.Fields("BOOKCODE").value = vs.TextMatrix(I, 6)
RS.Fields("TBook").value = IIf(vs.TextMatrix(I, 1) = "", 0, vs.TextMatrix(I, 1))
RS.Fields("LoosBook").value = vs.TextMatrix(I, 2)
RS.Fields("TotalBook").value = Val(vs.TextMatrix(I, 3))
RS.Fields("NetBook").value = vs.TextMatrix(I, 4)
RS.Fields("Remarks").value = vs.TextMatrix(I, 5)
RS.Fields("Book_Code").value = vs.TextMatrix(I, 0)
RS.Fields("fyear").value = main.session
RS.Fields("setupid").value = main.setupid

RS.update



DoEvents
DoEvents
DoEvents


If rs1.State = 1 Then rs1.close
rs1.Open "select GROUPCODE,SerName from BOOKS where BOOKCODE='" & vs.TextMatrix(I, 0) & "'", con
If rs1.EOF = False Then
   con.Execute "update BookReceiveDet set [gp]='" & rs1(0) & "',sername='" & rs1(1) & "' where BOOK_CODE='" & vs.TextMatrix(I, 0) & "' and INVOICENO=" & txtHeating.text & ""
End If



cmdSave_2.Enabled = False
cmdPrint_7.SetFocus
End If

Next

Else
con.Execute "delete from BookReceiveDet where " & stringyear & " and  INVOICENO=" & txtHeating.text & ""

For I = 1 To vs.rows - 1

If vs.TextMatrix(I, 0) <> "" Then

RS.AddNew
RS.Fields("INVOICENO").value = txtHeating.text
RS.Fields("INVOICEDATE").value = dates.value
RS.Fields("SUBLEDGER").value = txtParty.text
RS.Fields("GENLEDGER").value = "Sundry Debtors"
RS.Fields("BOOKCODE").value = vs.TextMatrix(I, 6)
RS.Fields("TBook").value = vs.TextMatrix(I, 1)
RS.Fields("LoosBook").value = vs.TextMatrix(I, 2)
RS.Fields("TotalBook").value = Val(vs.TextMatrix(I, 3))
RS.Fields("NetBook").value = vs.TextMatrix(I, 4)
RS.Fields("remarks").value = vs.TextMatrix(I, 5)
RS.Fields("Book_Code").value = vs.TextMatrix(I, 0)
RS.Fields("fyear").value = main.session
RS.Fields("setupid").value = main.setupid

RS.update

DoEvents
DoEvents
DoEvents


If rs1.State = 1 Then rs1.close
rs1.Open "select GROUPCODE,SerName from BOOKS where BOOKCODE='" & vs.TextMatrix(I, 0) & "'", con
If rs1.EOF = False Then
   con.Execute "update BookReceiveDet set [gp]='" & rs1(0) & "',sername='" & rs1(1) & "' where BOOK_CODE='" & vs.TextMatrix(I, 0) & "' and INVOICENO=" & txtHeating.text & ""
End If



cmdSave_2.Enabled = False
End If

Next


End If
'Call cmdAdd_1_Click

End If


Exit Sub
aa1:
MsgBox err.DESCRIPTION


End Sub
Sub searchData()


If RS.State = 1 Then RS.close
RS.Open "select * from BinderBkReceive where INVOICENO=" & txtHeating.text & "", con, adOpenDynamic, adLockOptimistic
If RS.EOF = False Then

cmdSave_2.Enabled = False

If RS!newoldbook = "NEW" Then
   Option1_new.value = True
Else
   Option2_old.value = True
End If

txtParty.text = RS.Fields("SUBLEDGER").value
txtRemarks.text = RS.Fields("Remarks").value & ""
Text1.text = RS.Fields("add1").value & ""
Text2.text = RS.Fields("add2").value & ""
txtTopay = RS.Fields("topay").value

cboFirm.text = RS.Fields("firmname").value

If Not IsNull(RS.Fields("godown").value) Then
cboGodown = RS.Fields("godown").value & ""
Else
cboGodown.ListIndex = -1
End If

End If



If RS.State = 1 Then RS.close
RS.Open "select * from BookReceiveDet where " & stringyear & " and INVOICENO=" & txtHeating.text & "", con, adOpenDynamic, adLockOptimistic
For I = 1 To RS.RecordCount
If RS.EOF = False Then
vs.TextMatrix(I, 6) = RS.Fields("BOOKCODE").value
vs.TextMatrix(I, 1) = RS.Fields("TBook").value
vs.TextMatrix(I, 2) = RS.Fields("LoosBook").value
vs.TextMatrix(I, 3) = RS.Fields("TotalBook").value
vs.TextMatrix(I, 4) = RS.Fields("NetBook").value
vs.TextMatrix(I, 5) = RS.Fields("remarks").value & ""
vs.TextMatrix(I, 0) = RS.Fields("Book_Code").value & ""
RS.MoveNext
End If
Next

'formButtonValidation cmdDelete_3, cmdEdit_4

Total

End Sub
Private Sub cmdUndo_5_Click()
   cmdDelete_3.Enabled = False
   cmdEdit_4.Enabled = False
   cmdPrint_7.Enabled = True
   cmdSave_2.Enabled = False
   cmdUndo_5.Enabled = False
   cmdAdd_1.Enabled = True
   cmdExit_12.Enabled = True
End Sub



Private Sub Command1_Binder_Click()

    PopUpValue7 = txtParty.text
    frmBinderLedger.Show
    
End Sub

Private Sub dates_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtParty.SetFocus
End Sub
Private Sub Dates_LostFocus()

If Trim(dates.value) <> "" Then
    If Not checkdate(Trim(dates.value), dates) Then
       dates.SetFocus
    End If
End If

End Sub

Private Sub Form_Activate()
'txtHeating.SetFocus
s = 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then
     'If MsgBox("Want To Exit ?", vbQuestion + vbYesNo) = vbYes Then
        Unload Me
     'End If
 End If
 
 
 
End Sub
Sub TotalFinal()
   
   If txtTotal3.text = "" Then
      txtTotal3.text = 0
   End If
   
   If txtTotal2.text = "" Then
      txtTotal2.text = 0
   End If
   
   
    txtRawAndCasting.text = (CDbl(txtTotal2.text) + CDbl(txtTotal3.text))
    txtRawAndCasting.text = Format(txtRawAndCasting.text, "#,###.000")
End Sub
Private Sub Form_Load()
  
Me.top = 100
Me.Left = 100

Me.Width = 13500
Me.Height = 9500
  
  
 BackColorFrom Me
  
 setWidth
 AddItemInGrid
' AddItemInGrid1
' AddItemInGrid2
' AddItemInGrid3
 setWidth
 
 
 txtHeating.text = MaxSNo("BinderBkReceive", "INVOICENO")
 txtId = txtHeating.text
 'Dates.value = Date
 
 Dim s As String
 
 s = ""
 


 If RS.State = 1 Then RS.close
 RS.Open "select name from MasterTbl where Category='bookrecfrombinder' order by name", con
 While RS.EOF = False
 If s = "" Then
 s = RS(0)
 Else
 s = s & "|" & RS(0)
 End If
 RS.MoveNext
 Wend
 
 vs.ColComboList(5) = s


 If RS.State = 1 Then RS.close
 RS.Open "select Godwn from Godownmaster where Binder_Printer='g' and " & stringyear & " order by id", con, adOpenKeyset, adLockReadOnly
 While RS.EOF = False
       cboGodown.AddItem RS(0)
       RS.MoveNext
 Wend
 
 If RS.State = 1 Then RS.close
 RS.Open "select firmname from firmname order by firmname", con
 While RS.EOF = False
       cboFirm.AddItem RS(0)
       RS.MoveNext
 Wend
 
 
 
 
' If RS.State = 1 Then RS.close
' Set RS = con.Execute("exec bookSearch '" & session & "'," & main.setupid & ",'" & "" & "'")
' 'RS.Open "select BOOKCODE,BOOKNAME from BOOKS where " & stringyear & "  order by BOOKCODE", con
' While RS.EOF = False
'    Bookcode.AddItem RS!Bookname & ":" & RS!Bookcode
'    RS.MoveNext
' Wend
 
 dates.value = Date
 
 
 
 If inviceNo <> "" Then
    
    txtHeating.text = inviceNo
    txtId.text = inviceNo
    vs.Clear
    setWidth
    searchData
    inviceNo = ""

End If
 
 
 
 
End Sub
Sub setWidth()
vs.Cols = 7
vs.FormatString = "BookCode|^Gaddi|^Books in a Gaddi|^Loose Books|^Total Books|Remarks|BookName"
vs.ColWidth(0) = 1500
vs.ColWidth(1) = 1500
vs.ColWidth(2) = 1500
vs.ColWidth(3) = 1200
vs.ColWidth(4) = 1200
vs.ColWidth(5) = 2000
vs.ColWidth(6) = 2000
End Sub
Private Sub FromDate_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then todate.SetFocus
End Sub
Private Sub HeatingDate_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then txtParty.SetFocus
End Sub

Private Sub ListHeatingNo_Click()
  Call cmdref_Click
  searchData
  TotalFinal
  'Frame1.Visible = False
End Sub
Private Sub Todate_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then Call cmdAdd_Click
End Sub
Private Sub txtGrade_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtRemarks.SetFocus
End Sub

Private Sub Form_Resize()
panel.Left = (Me.ScaleWidth - panel.Width) / 2
panel.top = (Me.ScaleHeight - panel.Height) / 2

End Sub

Private Sub txtHeating_GotFocus()
If PopUpValue1 <> "" Then
txtHeating.text = PopUpValue1
vs.Clear
setWidth
searchData
dates.value = PopUpValue2

PopUpValue1 = ""
PopUpValue2 = ""


'formButtonValidation cmdDelete_3, cmdEdit_4

Frame1.Enabled = False
vs.Enabled = False

End If
End Sub

Private Sub txtHeating_KeyDown(KeyCode As Integer, Shift As Integer)

'If KeyCode = 113 Then
'popuplist2 "select INVOICENO as [ChallanNo],INVOICEDATE as [Date],SUBLEDGER as Binder from BinderBkReceive order by INVOICENO", CON
'End If
'
'If KeyCode = 13 Then
'   Dates.SetFocus
'End If

End Sub

Private Sub txtHeating_KeyPress(KeyAscii As Integer)
   
''   If KeyAscii = 13 Then
''
''  'If Dates.Enabled = True Then
''  ' Dates.SetFocus
''  'End If
''
''  End If
  

End Sub
Sub search()


If inviceNo <> "" Then

    txtHeating.text = inviceNo
    txtId.text = inviceNo
    
    vs.Clear
    setWidth
    searchData
    vs.Enabled = False


End If


End Sub

Private Sub txtId_GotFocus()
If PopUpValue1 <> "" Then
txtHeating.text = PopUpValue1

txtId.text = PopUpValue1

vs.Clear
setWidth
searchData
dates.value = PopUpValue2

PopUpValue1 = ""
PopUpValue2 = ""


'formButtonValidation cmdDelete_3, cmdEdit_4
'Frame1.Enabled = False

vs.Enabled = False

End If

End Sub

Private Sub txtId_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then

   sqlQry = "select INVOICENO as [ChallanNo],INVOICEDATE as [Date],SUBLEDGER as Binder from BinderBkReceive where INVOICENO"
   orderby = "order by INVOICENO"


popuplist10 "select INVOICENO as [ChallanNo],INVOICEDATE as [Date],SUBLEDGER as Binder from BinderBkReceive where " & stringyear & "  order by INVOICENO", con
End If

If KeyCode = 13 Then
If dates.Enabled = True Then
   dates.SetFocus
End If
End If
End Sub

Private Sub txtId_KeyPress(KeyAscii As Integer)
'If KeyCode = 13 Then txtId.SetFocus

End Sub

Private Sub txtParty_GotFocus()
If PopUpValue1 <> "" Then
txtParty.text = PopUpValue1
Text1.text = PopUpValue2
'Text2.Text = PopUpValue3
PopUpValue1 = ""
PopUpValue2 = ""
PopUpValue3 = ""
End If
End Sub

Private Sub txtParty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
   popuplistModel10 "select Godwn as [Binder Name],Address from Godownmaster where " & stringyear & " and (binder_printer='b' or binder_printer='pb') order by Godwn", con
End If

If KeyCode = 13 Then
cboGodown.SetFocus
End If

End Sub

Private Sub txtParty_LostFocus()
   Record = ""
End Sub
Private Sub txtQty_GotFocus()
     txtQty.SelLength = 10
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtRemarks.SetFocus
End Sub
Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
 
Dim w%
 
 If KeyAscii = 13 Then
  vs.SetFocus

w = vs.RowSel
For I = 1 To 100
sendkeys "{up}"
w = w - 1
If w = 1 Then
   Exit For
End If
  
Next

'======================================

For I = 1 To 100
If vs.TextMatrix(I, 0) <> "" Then
sendkeys "{down}"
Else
Exit For
End If
Next



sendkeys "{home}"

End If
End Sub
Private Sub txtSize_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtGrade.SetFocus
End Sub



Private Sub txtTopay_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cboFirm.SetFocus
End Sub

Private Sub vs_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'     If vs.Col = 0 Then
'        cellposi
'      If cboItem.text <> "" Then
'        vs.TextMatrix(vs.RowSel, 0) = cboItem.text
'      End If
'     End If
End Sub
Private Sub vs_GotFocus()
   If PopUpValue1 <> "" Then
      vs.TextMatrix(vs.RowSel, 0) = PopUpValue1
      PopUpValue1 = ""
   End If
End Sub
Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)

  
  
  If KeyCode = 115 Then
  If MsgBox("Want to delete ?", vbQuestion + vbYesNo) = vbYes Then
    vs.RemoveItem (vs.RowSel)
    Total
  End If
  End If
  
  If KeyCode = 112 Then

    ' searchType = "inv"
    ' sqlqry = "select BOOKCODE,BOOKNAME from BOOKS where BOOKCODE"
    ' orderby = "order by BOOKCODE"
    ' popuplist10 "select BOOKCODE,BOOKNAME from BOOKS where " & stringyear & "  order by BOOKCODE", CON
         '------------------------------------
     If vs.Col = 0 Then
        Bookcode.Visible = True: Bookcode.Enabled = True
        Bookcode.ZOrder
        Bookcode.text = vs.text
        Bookcode.top = vs.top + vs.CellTop
        Bookcode.Left = vs.CellLeft + leftAlign + 150
        'Bookcode.Width = vs.ColWidth(vs.Col)
        Bookcode.SetFocus
     End If
     '-----------------------------------
  End If
  
  
  If KeyCode = 13 Then
  
  
  

     
     If vs.Col = 0 Then
        If Val(txtTotal) > 0 Then
        Call cmdSave_2_Click
     End If
     
     Else
        cellposi
     End If

  End If
  
  
  
  
  
End Sub

Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)

On Error Resume Next

If KeyCode = 13 Then
        
          
 If vs.Col = 0 Then
 
  
 
 If RS.State = 1 Then RS.close
 Set RS = con.Execute("exec BookSearch_bycode '" & session & "'," & main.setupid & ",'" & Trim(vs.TextMatrix(vs.RowSel, 0)) & "'")
 If RS.EOF = False Then
       vs.TextMatrix(vs.RowSel, 0) = UCase(vs.TextMatrix(vs.RowSel, 0))
       vs.TextMatrix(vs.RowSel, 2) = RS!BooksInGaddi & ""
       vs.TextMatrix(vs.RowSel, 4) = (IIf(vs.TextMatrix(vs.RowSel, 1) = "", 0, vs.TextMatrix(vs.RowSel, 1)) * IIf(vs.TextMatrix(vs.RowSel, 2) = "", 0, vs.TextMatrix(vs.RowSel, 2)) + IIf(vs.TextMatrix(vs.RowSel, 3) = "", 0, vs.TextMatrix(vs.RowSel, 3)))
       vs.TextMatrix(vs.RowSel, 6) = RS!Bookname & ""

Else
   vs.SetFocus
   Exit Sub
End If

  If vs.TextMatrix(vs.RowSel, 0) <> "" Then
       sendkeys "{right}"
  End If
 
    
 End If

If vs.Col = 1 Then

If Val(vs.TextMatrix(vs.RowSel, 1)) = 0 Then Exit Sub

 sendkeys "{right}"
 sendkeys "{right}"

 End If
    
If vs.Col = 3 Then
sendkeys "{right}"
sendkeys "{right}"
vs.TextMatrix(vs.RowSel, 4) = (IIf(vs.TextMatrix(vs.RowSel, 1) = "", 0, Val(vs.TextMatrix(vs.RowSel, 1))) * IIf(vs.TextMatrix(vs.RowSel, 2) = "", 0, vs.TextMatrix(vs.RowSel, 2)) + IIf(vs.TextMatrix(vs.RowSel, 3) = "", 0, vs.TextMatrix(vs.RowSel, 3)))

End If


If vs.Col = 5 Then

If Len(vs.TextMatrix(vs.RowSel, 5)) = 0 Then Exit Sub

vs.TextMatrix(vs.RowSel, 4) = (IIf(vs.TextMatrix(vs.RowSel, 1) = "", 0, vs.TextMatrix(vs.RowSel, 1)) * IIf(vs.TextMatrix(vs.RowSel, 2) = "", 0, vs.TextMatrix(vs.RowSel, 2)) + IIf(vs.TextMatrix(vs.RowSel, 3) = "", 0, vs.TextMatrix(vs.RowSel, 3)))
cboItem.text = ""
sendkeys "{home}"
sendkeys "{down}"
Total
End If
    
       
Total

End If

End Sub

Private Sub vs_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)

'If KeyCode = 13 Then
'
'If vs.Col = 1 Then
'If Not (KeyCode >= 48 And KeyCode <= 57) Then
'MsgBox "Enter Only Numeric Value !!", vbInformation
'vs.TextMatrix(vs.RowSel, 1) = ""
'vs.SetFocus
'Exit Sub
'End If
'End If
'
'
'If vs.Col = 2 Then
'If Not (KeyCode >= 48 And KeyCode <= 57) Then
'MsgBox "Enter Only Numeric Value !!", vbInformation
'vs.TextMatrix(vs.RowSel, 2) = ""
'vs.SetFocus
'Exit Sub
'End If
'End If
'
'
'If vs.Col = 3 Then
'If Not (KeyCode >= 48 And KeyCode <= 57) Then
'MsgBox "Enter Only Numeric Value !!", vbInformation
'vs.TextMatrix(vs.RowSel, 3) = ""
'vs.SetFocus
'Exit Sub
'End If
'End If
'
'
'If vs.Col = 4 Then
'If Not (KeyCode >= 48 And KeyCode <= 57) Then
'MsgBox "Enter Only Numeric Value !!", vbInformation
'vs.TextMatrix(vs.RowSel, 4) = ""
'vs.SetFocus
'Exit Sub
'End If
'End If
'
'
'
'
'End If

End Sub

Private Sub vs_LeaveCell()
  Total
End Sub

Private Sub vs1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     If vs1.Col = 0 Then
        cellposiVs
        vs1.TextMatrix(vs1.RowSel, 0) = cboitemvs1.text
     End If

End Sub
Private Sub vs1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 46 Then
    vs1.RemoveItem (vs1.RowSel)
    'Total1
    TotalFinal
  End If
  
  If KeyCode = 13 Then
     If vs1.Col = 0 Then
        vs1.Editable = flexEDNone
        Vs1Frame.Visible = True
        cboitemvs1.Visible = True
        cboitemvs1.SetFocus
     Else
        vs1.Editable = flexEDKbdMouse
        cellposiVs
     End If
  End If
End Sub

Private Sub vs1_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
        
 If vs1.Col = 0 Then
    vs1.Editable = flexEDNone
    Vs1Frame.Visible = True
    cboitemvs1.SetFocus
          
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.close
    RS.Open "select * from ItemMaster where " & stringyear & " and ItemName='" & vs1.TextMatrix(vs1.RowSel, 0) & "'", con, adOpenKeyset, adLockReadOnly
    If RS.EOF = False Then
       vs1.TextMatrix(vs1.RowSel, 1) = RS.Fields("Unit").value
       sendkeys "{right}"
       sendkeys "{right}"
       Vs1Frame.Visible = False
       vs1.Editable = flexEDKbdMouse
       vs1.SetFocus
    Else
       vs1.TextMatrix(vs1.RowSel, 1) = "Kg"
       sendkeys "{right}"
       sendkeys "{right}"
       Vs1Frame.Visible = False
       vs1.Editable = flexEDKbdMouse
       vs1.SetFocus
    
    End If
    
 End If
    
 If vs1.Col = 2 Then
           
    sendkeys "{home}"
    sendkeys "{down}"
    
'    AddItemInGrid1
 End If
    
    

 'Total1
 TotalFinal

End If


End Sub
Sub AddSemifinish()
   Dim J As Integer
   
   J = 1
    
   vs3.Clear
   For I = 1 To vs1.rows - 1
    
   If vs1.TextMatrix(I, 0) <> "" Then
      If RS.State = 1 Then RS.close
      RS.Open "select * from ItemMaster where " & stringyear & " and ItemName='" & vs1.TextMatrix(I, 0) & "'", con, adOpenKeyset, adLockReadOnly
      If RS.Fields("itemgp").value = "Semi Finish (R/D)" Or RS.Fields("itemgp").value = "Semi Finish (Store)" Then
         vs3.TextMatrix(J, 0) = vs1.TextMatrix(I, 0)
         vs3.TextMatrix(J, 1) = vs1.TextMatrix(I, 1)
         vs3.TextMatrix(J, 2) = vs1.TextMatrix(I, 2)
         J = J + 1
      End If
   End If
        
   Next
    
    
End Sub
Private Sub vs1_LeaveCell()
   'Total1
End Sub

Private Sub vs2_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     If vs2.Col = 0 Then
        'cellposiVs3
        vs2.TextMatrix(vs2.RowSel, 0) = cboItemVs3.text
     End If

End Sub

Private Sub vs2_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 46 Then
    vs2.RemoveItem (vs2.RowSel)
    'Total2
    TotalFinal
 End If

  
  If KeyCode = 13 Then
     
     If vs2.Col = 0 Then
        vs2.Editable = flexEDNone
        Vs3Frame.Visible = True
        cboItemVs3.Visible = True
        cboItemVs3.SetFocus
     Else
        vs2.Editable = flexEDKbdMouse
        'cellposiVs3
     End If

  End If

End Sub

Private Sub vs2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
        
          
 If vs2.Col = 0 Then
 
      vs2.Editable = flexEDNone
      Vs3Frame.Visible = True
      
      
          
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.close
    RS.Open "select * from ItemMaster where " & stringyear & " and ItemName='" & vs2.TextMatrix(vs2.RowSel, 0) & "'", con
    If RS.EOF = False Then
       vs2.TextMatrix(vs2.RowSel, 1) = RS.Fields("Unit").value
       sendkeys "{right}"
       sendkeys "{right}"
       Vs3Frame.Visible = False
       vs2.Editable = flexEDKbdMouse
       vs2.SetFocus
    Else
       vs2.TextMatrix(vs2.RowSel, 1) = "Kg"
       sendkeys "{right}"
       sendkeys "{right}"
       Vs3Frame.Visible = False
       vs2.Editable = flexEDKbdMouse
       vs2.SetFocus
    
    End If
    
 End If
 
    
    If vs2.Col = 2 Then
           
           sendkeys "{home}"
           sendkeys "{down}"
           Vs3Frame.top = Vs3Frame.top + 170
    End If
    
       
   'Total2

End If

End Sub
Private Sub vs2_LeaveCell()
   'Total2
End Sub
Private Sub vs3_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     If vs3.Col = 0 Then
        'cellposiVs2
        'vs3.TextMatrix(vs3.RowSel, 0) = cboitemvscboItemVs2.Text
     End If
 
End Sub

Private Sub vs3_KeyDown(KeyCode As Integer, Shift As Integer)
    
  If KeyCode = 46 Then
    vs3.RemoveItem (vs3.RowSel)
    'Total4
  End If
  
  If KeyCode = 13 Then
     If vs3.Col = 0 Then
        
        vs3.Editable = flexEDNone
        FrameVs2.Visible = True
        cboItemVs2.Visible = True
        cboItemVs2.SetFocus
     Else
        
        vs3.Editable = flexEDKbdMouse
        
     End If
  End If

End Sub

Private Sub vs3_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
 
 If vs3.Col = 0 Then
    vs3.Editable = flexEDNone
    FrameVs2.Visible = True
    cboItemVs2.SetFocus
    
    
 
          
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.close
    RS.Open "select * from ItemMaster where " & stringyear & " and ItemName='" & vs3.TextMatrix(vs3.RowSel, 0) & "'", con
    If RS.EOF = False Then
       vs3.TextMatrix(vs3.RowSel, 1) = RS.Fields("Unit").value
       sendkeys "{right}"
       sendkeys "{right}"
       FrameVs2.Visible = False
       vs3.Editable = flexEDKbdMouse
       vs3.SetFocus
    Else
       vs3.TextMatrix(vs3.RowSel, 1) = "Kg"
       sendkeys "{right}"
       sendkeys "{right}"
       FrameVs2.Visible = False
       vs3.Editable = flexEDKbdMouse
       vs3.SetFocus
    
    End If
    
 End If
    
 If vs3.Col = 2 Then
    
   If RS.State = 1 Then RS.close
   RS.Open "select  OpeningStock from ItemMaster where " & stringyear & " and ItemName='" & vs3.TextMatrix(vs.RowSel, 0) & "'", con
   If RS.EOF = False Then
      If Val(RS.Fields(0).value) < Val(vs3.TextMatrix(vs3.RowSel, 2)) Then
         MsgBox "Stock Less !!", vbInformation
         
      End If
   End If
    
    
    sendkeys "{home}"
    sendkeys "{down}"
    
    FrameVs2.top = FrameVs2.top + 170
    'AddItemInGrid2
 End If
    
    

 'Total4
 
End If
 
End Sub
