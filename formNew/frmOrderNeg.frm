VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOrderNeg 
   ClientHeight    =   9240
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   17640
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9240
   ScaleWidth      =   17640
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
      Left            =   45
      TabIndex        =   1
      Top             =   90
      Width           =   17112
      Begin VB.ComboBox txtFirmName 
         Height          =   315
         ItemData        =   "frmOrderNeg.frx":0000
         Left            =   6570
         List            =   "frmOrderNeg.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   495
         Width           =   5100
      End
      Begin VB.TextBox txtHeating 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1830
         TabIndex        =   0
         Top             =   630
         Width           =   1740
      End
      Begin VB.TextBox txtParty 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6585
         TabIndex        =   21
         Top             =   930
         Width           =   5070
      End
      Begin VB.TextBox txtRemarks 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1830
         TabIndex        =   20
         Top             =   1305
         Width           =   4065
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
         TabIndex        =   14
         Top             =   8970
         Visible         =   0   'False
         Width           =   465
         Begin VB.TextBox txtRawAndCasting 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3690
            TabIndex        =   15
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
            TabIndex        =   18
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
            TabIndex        =   17
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
            TabIndex        =   16
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
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6585
         TabIndex        =   13
         Top             =   1260
         Width           =   5070
      End
      Begin VB.Frame buttonFrame 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   804
         Left            =   360
         TabIndex        =   5
         Top             =   7944
         Width           =   8775
         Begin VB.CommandButton cmdAdd_1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Add"
            Height          =   615
            Left            =   75
            Picture         =   "frmOrderNeg.frx":0004
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   90
            Width           =   1230
         End
         Begin VB.CommandButton cmdSave_2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Save"
            Height          =   648
            Left            =   1305
            Picture         =   "frmOrderNeg.frx":0BE8
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   75
            Width           =   1230
         End
         Begin VB.CommandButton cmdDelete_3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Delete"
            Height          =   648
            Left            =   2535
            Picture         =   "frmOrderNeg.frx":17CC
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   75
            Width           =   1230
         End
         Begin VB.CommandButton cmdEdit_4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Edit"
            Height          =   648
            Left            =   3765
            Picture         =   "frmOrderNeg.frx":23B0
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   75
            Width           =   1230
         End
         Begin VB.CommandButton cmdUndo_5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Undo"
            Height          =   648
            Left            =   4995
            Picture         =   "frmOrderNeg.frx":27F2
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   75
            Width           =   1230
         End
         Begin VB.CommandButton cmdPrint_7 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Print"
            Height          =   648
            Left            =   6225
            Picture         =   "frmOrderNeg.frx":2D7C
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   75
            Width           =   1230
         End
         Begin VB.CommandButton cmdExit_12 
            BackColor       =   &H00FFFFFF&
            Caption         =   "E&xit"
            Height          =   648
            Left            =   7470
            Picture         =   "frmOrderNeg.frx":3960
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   75
            Width           =   1230
         End
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   9240
         TabIndex        =   4
         Top             =   7440
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtLoose 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7755
         TabIndex        =   3
         Top             =   8415
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdMaster 
         Appearance      =   0  'Flat
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
         Left            =   11940
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   870
         Visible         =   0   'False
         Width           =   450
      End
      Begin Crystal.CrystalReport CR 
         Left            =   9075
         Top             =   8175
         _ExtentX        =   593
         _ExtentY        =   593
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin MSComCtl2.DTPicker Dates 
         Height          =   315
         Left            =   1830
         TabIndex        =   22
         Top             =   945
         Width           =   1395
         _ExtentX        =   2455
         _ExtentY        =   550
         _Version        =   393216
         Format          =   547880961
         CurrentDate     =   39500
      End
      Begin VSFlex7Ctl.VSFlexGrid vs 
         Height          =   5736
         Left            =   312
         TabIndex        =   23
         Top             =   1668
         Width           =   16452
         _cx             =   29019
         _cy             =   10118
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
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16777215
         ForeColorSel    =   16711680
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   8388608
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmOrderNeg.frx":4544
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
      Begin VB.TextBox txtTotal1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7785
         TabIndex        =   19
         Text            =   "0"
         Top             =   8385
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Firm Name :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   7
         Left            =   5670
         TabIndex        =   31
         Top             =   540
         Width           =   990
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H0078CFE9&
         BorderWidth     =   4
         Height          =   870
         Left            =   315
         Top             =   7920
         Width           =   8880
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Order No :"
         Height          =   270
         Index           =   0
         Left            =   345
         TabIndex        =   29
         Top             =   600
         Width           =   1530
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   5655
         TabIndex        =   28
         Top             =   960
         Width           =   585
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date "
         Height          =   270
         Index           =   1
         Left            =   330
         TabIndex        =   27
         Top             =   915
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks "
         Height          =   300
         Index           =   4
         Left            =   315
         TabIndex        =   26
         Top             =   1260
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total :"
         Height          =   240
         Index           =   6
         Left            =   8760
         TabIndex        =   25
         Top             =   7440
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "F2 For Search Challan"
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
         Left            =   1815
         TabIndex        =   24
         Top             =   405
         Width           =   2505
      End
   End
End
Attribute VB_Name = "frmOrderNeg"
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
Dim rs_book As New ADODB.Recordset
Dim iitem1 As String
Dim StockFlag As String
Dim rs1 As New ADODB.Recordset
Private Sub cmdMain_Click()
  Unload Me
End Sub

Sub cellposi()
 'VsFrame.Width = 3165
 VsFrame.top = vs.top + ((vs.CellTop)) - 1500
 VsFrame.Left = (vs.Left) - 190
End Sub
Sub Total()

txtTotal.text = 0
txtLoose.text = 0

For J = 1 To vs.rows - 1
If vs.TextMatrix(J, 0) <> "" Then
txtTotal.text = (Val(txtTotal.text) + Val(vs.TextMatrix(J, 6)))
'txtLoose.Text = (Val(txtLoose.Text) + Val(vs.TextMatrix(j, 3)))
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
'    'rs_1.Open "select * from BookMaster order by Book", con, adOpenDynamic, adLockOptimistic
'    Set rs_1 = con.Execute("exec searchList '" & "book_master" & "'")
'
'    Set cboItem.RowSource = rs_1
'    cboItem.ListField = "Book"
'    cboItem.BoundColumn = "BookNo"
'    cboItem.ReFill
    
End Sub
Private Sub cboFinishItem_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtQty.SetFocus
End Sub


Private Sub cbogodown_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then txtRemarks.SetFocus
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
        vs.TextMatrix(vs.RowSel, 1) = cboItem.BoundText
         
        If RS.State = 1 Then RS.close
        If cboItem.BoundText = "" Then Exit Sub
        RS.Open "select headdata1,headdata2,headdata3,dividevalue,((cast(headdata1 as int)+cast(headdata2 as int)+cast(headdata3 as int))/(dividevalue)) as noofform from bookmaster " & _
        "where " & stringyear & " and bookno='" & cboItem.BoundText & "' and categories='Negative'", con, adOpenKeyset, adLockReadOnly
        'cast(INVOICENO as int)
        If RS.EOF = False Then
           vs.TextMatrix(vs.RowSel, 2) = RS(0) & ""
           vs.TextMatrix(vs.RowSel, 3) = RS(1) & ""
           vs.TextMatrix(vs.RowSel, 4) = RS(2) & ""
           vs.TextMatrix(vs.RowSel, 5) = Round(RS(4), 2) & ""
        End If
        
        
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
        RS.Open "select * from ItemMaster where " & stringyear & " and ItemName='" & cboitemvs1.text & "'", con
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
'Total4

End Sub

Private Sub cmdFind_Click()
 Frame1.Visible = True
 fromdate.SetFocus
End Sub

Private Sub cmdModify_Click()
   If MsgBox("Want To Modify ?", vbQuestion + vbYesNo) = vbYes Then
      
        
      
      'SaveData
      
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
   
    txtHeating.text = ""
    dates.value = Date
    txtParty.text = ""
    txtRemarks.text = ""
    Text1.text = ""
    txtTotal.text = ""
    txtLoose.text = ""

   
   vs.Clear
   setWidth
   cmdDelete_3.Enabled = False
   cmdEdit_4.Enabled = True
   'cmdPrint_7.Enabled = False
   cmdSave_2.Enabled = True
   cmdUndo_5.Enabled = True
   cmdAdd_1.Enabled = True
   cmdExit_12.Enabled = True
   txtHeating.SetFocus
   
   
   'txtHeating.Text = MaxSNoNew("BillMaster", "bill_id", "Negative")
   txtHeating.text = MaxOrderNo(txtFirmName)
   
End Sub

Private Sub cmdDelete_3_Click()
If MsgBox("Want to delete ?", vbQuestion + vbYesNo) = vbYes Then
con.Execute "delete from BillMaster where categories='Negative' and bill_id='" & txtHeating.text & "' and " & stringyear
con.Execute "delete from Billtrans where categories='Negative' and bill_id='" & txtHeating.text & "' and " & stringyear
Call cmdAdd_1_Click
End If
End Sub

Private Sub cmdEdit_4_Click()
   cmdDelete_3.Enabled = True
   cmdEdit_4.Enabled = False
   cmdPrint_7.Enabled = True
   cmdSave_2.Enabled = True
   cmdUndo_5.Enabled = True
   cmdAdd_1.Enabled = False
   cmdExit_12.Enabled = True
End Sub

Private Sub cmdExit_12_Click()
Unload Me
End Sub
Sub updateRecord()

Dim fonttype As String

J = 1
k1 = 1

set1 = ""

If rs1.State = 1 Then rs1.close
rs1.Open "select set_name,BookCode from Billtrans where categories='Negative' and bill_id='" & txtHeating.text & "'" & _
" and " & stringyear & " order by set_name", con, adOpenKeyset, adLockReadOnly
While rs1.EOF = False
If Len(set1) > 0 Then
If (set1 = rs1(0)) Then
Else
   J = 1
   k1 = k1 + 1
End If
End If

If rs_book.State = 1 Then rs_book.close
rs_book.Open "select bookfont from BookMaster where BookNo='" & rs1(1) & "'" & _
" and " & stringyear, con, adOpenKeyset, adLockReadOnly
If rs_book.EOF = False Then
   fonttype = rs_book(0)
 Else
   fonttype = "e"
End If


Set RS = New ADODB.Recordset
RS.Open "select set_name from billtrans where categories='Negative' and BookCode= '" & rs1(1) & "' and " & _
"bill_id='" & txtHeating.text & "' order by AutoId", con, adOpenDynamic, adLockOptimistic
If RS.EOF = False Then
 con.Execute "update Billtrans set bookfont='" & fonttype & "',GPOrderPrinting=" & J & ",OrderPrinting=" & k1 & " where categories='Negative' and bill_id='" & txtHeating.text & "' and BookCode= '" & rs1(1) & "' and " & stringyear
 J = J + 1
End If
set1 = rs1(0)
rs1.MoveNext
Wend

End Sub

Private Sub cmdMaster_Click()

HeadTbl = "Binder"
frmMasters.Show 1

End Sub

Private Sub cmdPrint_7_Click()


DSNNew

If MsgBox("Want To Print ?", vbInformation + vbYesNo) = vbYes Then
    
    CR.Reset
    CR.ReportFileName = rptPath & "/NegativePrint.rpt"
    CR.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    CR.ReplaceSelectionFormula "{BillMaster.categories}='Negative' and {BillMaster.bill_id}='" & txtHeating.text & "'"
    'CR.Formulas(0) = "address='" & Text1.Text & "'"
    If txtFirmName.text = "BLUEPRINT EDUCATION " Then
       CR.Formulas(0) = "unit_='" & "(A division of Chitra Prakashan (I) Pvt.Ltd.)" & "'"
    Else
       CR.Formulas(0) = "unit_='" & "" & "'"
    End If
    CR.Formulas(1) = "address='" & Text1.text & "'"
    CR.WindowShowPrintSetupBtn = True
    CR.WindowState = crptMaximized
    CR.Action = 1
    
End If


End Sub
Private Sub cmdSave_2_Click()

Dim paperdetails As String
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset

k1 = 0

'On Error GoTo aa1



If txtParty.text = "" Then
   MsgBox "Please Enter Binder Name !!", vbInformation
   Exit Sub
End If


If rs1.State = 1 Then rs1.close
rs1.Open "select * from PaperMakeMaster where " & stringyear, con, adOpenKeyset, adLockReadOnly
If MsgBox("Want to Save ?", vbYesNo + vbQuestion) = vbYes Then

If RS.State = 1 Then RS.close
RS.Open "select * from BillMaster where categories='Negative' and bill_id='" & txtHeating.text & "' and " & stringyear, con, adOpenDynamic, adLockOptimistic
If RS.EOF = False Then
   con.Execute "delete from BillMaster where categories='Negative' and bill_id='" & txtHeating.text & "' and " & stringyear
   con.Execute "delete from billtrans where categories='Negative' and bill_id='" & txtHeating.text & "' and " & stringyear
End If


RS.AddNew
RS.Fields("firm_id").value = txtFirmName.text
RS.Fields("bill_id").value = txtHeating.text
RS.Fields("dat").value = dates.value
RS.Fields("PrinterName").value = txtParty.text
RS.Fields("Remarks").value = txtRemarks.text
RS.Fields("categories").value = "Negative"
RS!setupid = setupid
RS!fyear = session

RS.update

If rs2.State = 1 Then rs2.close
rs2.Open "select * from Billtrans where categories='Negative' and " & stringyear, con, adOpenDynamic, adLockOptimistic


For I = 1 To vs.rows - 1

If vs.TextMatrix(I, 0) <> "" Then

rs2.AddNew

rs2.Fields("categories").value = "Negative"

rs2.Fields("bill_id").value = txtHeating.text
rs2.Fields("BookName").value = vs.TextMatrix(I, 0)
rs2.Fields("BookCode").value = vs.TextMatrix(I, 1)

rs2.Fields("Inners").value = vs.TextMatrix(I, 2)

rs2.Fields("text").value = vs.TextMatrix(I, 3)

rs2.Fields("paper").value = vs.TextMatrix(I, 4)
rs2.Fields("NoOfForm").value = Round(vs.TextMatrix(I, 5), 2)

rs2.Fields("Types").value = vs.TextMatrix(I, 6)

rs2!setupid = setupid
rs2!fyear = session

rs2.Fields("BidningType").value = vs.TextMatrix(I, 7)
rs2.Fields("Neg_Remarks").value = vs.TextMatrix(I, 8)
rs2!Particulars = vs.TextMatrix(I, 9)


rs2.update


End If

Next


updateRecord

cmdSave_2.Enabled = False
cmdPrint_7.SetFocus



End If





Exit Sub
aa1:
MsgBox err.DESCRIPTION


End Sub
Sub searchData()

vs.rows = 50
I = 1

Dim rs1 As New ADODB.Recordset


Set RS = New ADODB.Recordset
RS.Open "select * from BillMaster where " & stringyear & " and bill_id='" & txtHeating.text & "' and categories='Negative'", con, adOpenDynamic, adLockOptimistic
If RS.EOF = False Then

cmdSave_2.Enabled = False
cmdDelete_3.Enabled = False
cmdEdit_4.Enabled = True


txtFirmName.text = RS.Fields("firm_id").value & ""

txtHeating.text = RS.Fields("bill_id").value
dates.value = RS.Fields("dat").value
txtParty.text = RS.Fields("PrinterName").value
txtRemarks.text = RS.Fields("Remarks").value




If rs1.State = 1 Then rs1.close
rs1.Open "select Address from Godownmaster where Godwn ='" & txtParty.text & "'", con
If rs1.EOF = False Then
   Text1.text = rs1(0)
Else
   Text1.text = ""
End If


End If


If RS.State = 1 Then RS.close
RS.Open "select * from Billtrans as b inner join billmaster as bm on (b.bill_id = bm.bill_id and b.categories=bm.categories) " & _
" where bm.fyear='" & session & "' and bm.setupid=" & setupid & " and b.bill_id='" & txtHeating.text & "' and bm.categories='Negative'", con, adOpenDynamic, adLockOptimistic


While RS.EOF = False




If RS!bookfont = "h" Then
   
   vs.Col = 0
   vs.CellFontName = hindi
   vs.CellFontSize = 14
   
   vs.TextMatrix(I, 0) = RS.Fields("BookName").value
Else
   vs.CellFontName = english
   vs.CellFontSize = 12
   vs.TextMatrix(I, 0) = RS.Fields("BookName").value

End If


vs.TextMatrix(I, 1) = RS.Fields("BookCode").value

vs.TextMatrix(I, 2) = RS.Fields("Inners").value & ""
vs.TextMatrix(I, 3) = RS.Fields("text").value & ""
vs.TextMatrix(I, 4) = RS.Fields("paper").value & ""
vs.TextMatrix(I, 5) = Round(RS.Fields("NoOfForm").value, 2) & ""
vs.TextMatrix(I, 6) = RS.Fields("Types").value & ""

vs.TextMatrix(I, 7) = RS.Fields("BidningType").value & ""

vs.TextMatrix(I, 8) = RS.Fields("Neg_Remarks").value & ""

vs.TextMatrix(I, 9) = RS!Particulars & ""

I = I + 1
RS.MoveNext

Wend

'total

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



Private Sub dates_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtParty.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

 
If KeyCode = 27 Then
     If MsgBox("Want To Exit ?", vbQuestion + vbYesNo) = vbYes Then
        Unload Me
     End If
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

 Me.Caption = "Order For Negative Printing"
 
 Me.Left = 50
 Me.top = 10

'Me.Width = 14200
'Me.Height = 8500

Me.Width = 16700
Me.Height = 9810
 
 setWidth
 AddItemInGrid
 
 
 setWidth
 
 dates.value = Date
 'txtHeating.Text = MaxSNo("BillMaster", "bill_id")
 'txtHeating.Text = MaxSNoNew("BillMaster", "bill_id", "Negative")
 
 
s100 = "Double Colour|Four Colour|Single Colour"
vs.ColComboList(2) = s100


txtFirmName.Clear
Set RS = New ADODB.Recordset
RS.Open "select FirmName,Add1,Add2 from FirmMaster order by firmname", con, adOpenStatic, adLockReadOnly
While RS.EOF = False
 txtFirmName.AddItem RS(0)
 RS.MoveNext
Wend

txtFirmName.ListIndex = 0
txtHeating.text = MaxOrderNo(txtFirmName)

Dim s As String
 
BackColorFrom Me
 
vs.ColComboList(6) = "Web|Sheet"
 
End Sub
Sub setWidth()
vs.Cols = 10

vs.FormatString = "Books Name|Books Code|Colour|Text/Inner|No of Page|No Of Forms|Type|BindingType|Remarks|Size"
vs.ColWidth(0) = 3300
vs.ColWidth(1) = 1200
vs.ColWidth(2) = 1700
vs.ColWidth(3) = 1000
vs.ColWidth(4) = 1400
vs.ColWidth(5) = 1200
vs.ColWidth(6) = 1000

vs.ColWidth(7) = 1200
vs.ColWidth(8) = 2000
vs.ColWidth(9) = 1000

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

Private Sub txtFirmName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
dates.SetFocus
End If

End Sub

Private Sub txtFirmName_LostFocus()
txtHeating.text = MaxOrderNo(txtFirmName)
End Sub

Private Sub txtHeating_GotFocus()
If PopUpValue1 <> "" Then
txtHeating.text = PopUpValue1
dates.value = PopUpValue2
vs.Clear
setWidth
searchData
PopUpValue1 = ""
PopUpValue2 = ""
End If
End Sub

Private Sub txtHeating_KeyDown(KeyCode As Integer, Shift As Integer)


If KeyCode = 113 Then
   value = "select bill_id,dat as OrderDate,PrinterName from BillMaster where firm_id='" & txtFirmName.text & "' and categories='Negative'  group by bill_id,dat,PrinterName  order by convert(int,bill_id) "
   popuplist1 value, con
End If


If KeyCode = 13 Then
searchData
End If



End Sub

Private Sub txtHeating_KeyPress(KeyAscii As Integer)
   
   If KeyAscii = 13 Then
        
   txtFirmName.SetFocus

        
  End If
  

End Sub
Private Sub txtParty_GotFocus()
If PopUpValue1 <> "" Then
txtParty.text = PopUpValue1
Text1.text = PopUpValue2
PopUpValue1 = ""
'PopUpValue2 = ""
'PopUpValue3 = ""
End If
End Sub
Private Sub txtParty_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then
   'value = "select Godwn as [Binder Name],Address from Godownmaster where " & stringyear & " and (Binder_Printer='p' or Binder_Printer='pb') order by Godwn"
   value = "select Godwn as [Binder Name],Address from Godownmaster where " & stringyear & " and len(Godwn)>3 order by Godwn"
   popuplist1 value, con
End If


End Sub

Private Sub txtParty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtRemarks.SetFocus
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
 If KeyAscii = 13 Then
   'SendKeys "{tab}"
   vs.SetFocus
   vs.Col = 1
 End If
End Sub
Private Sub txtSize_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtGrade.SetFocus
End Sub

Private Sub vs_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     If vs.Col = 0 Then
        cellposi
        vs.TextMatrix(vs.RowSel, 0) = cboItem.text
     End If
End Sub

Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)

  
  
  If KeyCode = 115 Then
  If MsgBox("Want to delete ?", vbQuestion + vbYesNo) = vbYes Then
    vs.RemoveItem (vs.RowSel)
    Total
  End If
  End If
  
  If KeyCode = 13 Then
     
     If vs.Col = 0 Then
        vs.Editable = flexEDNone
        VsFrame.Visible = True
        cboItem.SetFocus
     Else
        vs.Editable = flexEDKbdMouse
        cellposi
     End If

  End If
  
  
  
  
  
End Sub
Sub addType(bk As String)
s11 = ""

Dim k1 As Integer
k1 = 6



Dim fnext As Integer
Dim fld As Integer

kk2 = vs.RowSel

If vs.RowSel = 1 Then
   fnext = 8
Else
   fnext = kk2 + 8
End If

fld = 1

If rs1.State = 1 Then rs1.close
rs1.Open "select Head1,Head2,Head3,Head4,Head5,txtHead6 as Head6,txtHead7 as Head7,txtHead8 as Head8," & _
"color1,color2,color3,color4,color5,color6,color7,color8,HeadData1,HeadData2,HeadData3,HeadData4,HeadData5," & _
"txtheadData6 as HeadData6,txtheadData7 as HeadData7,txtheadData8 as HeadData8,Inn_Forms as Forms1,text_Forms as Forms2," & _
"Exam_Forms as Forms3,Supp_Forms as Forms4,Title_Forms as Forms5,txtTextSupp6 as Forms6,txtTextSupp7 as Forms7," & _
"txtTextSupp8 as Forms8 from BookMaster where bookno='" & bk & "'", con
If rs1.EOF = False Then
For Q1 = kk2 To fnext - 1


    If Not IsNull(rs1.Fields("head" & fld).value) Then
        If rs1.Fields("head" & fld).value <> "" Then
            If s11 = "" Then
               s11 = rs1.Fields("head" & fld).value
            Else
               s11 = s11 & "|" & rs1.Fields("head" & fld).value
            End If
        
        
        
            If Q1 > 1 Then
               vs.TextMatrix(Q1, 0) = vs.TextMatrix(vs.RowSel, 0)
               vs.TextMatrix(Q1, 1) = vs.TextMatrix(vs.RowSel, 1)
               
               vs.TextMatrix(Q1, 6) = vs.TextMatrix(vs.RowSel, 6)
               vs.TextMatrix(Q1, 7) = vs.TextMatrix(vs.RowSel, 7)
               vs.TextMatrix(Q1, 9) = vs.TextMatrix(vs.RowSel, 9)
            End If
            
            vs.TextMatrix(Q1, 2) = rs1.Fields("color" & fld)
            vs.TextMatrix(Q1, 3) = rs1.Fields("head" & fld).value
            vs.TextMatrix(Q1, 4) = rs1.Fields("headData" & fld).value
            vs.TextMatrix(Q1, 5) = rs1.Fields("forms" & fld).value & ""
        
        
        End If
        
     
        
    End If


     
fld = fld + 1

Next

End If

vs.ColComboList(3) = s11

  
   
End Sub

Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
        
          
 If vs.Col = 0 Then
 
      vs.Editable = flexEDNone
      VsFrame.Visible = True
      cboItem.SetFocus
          
  If vs.TextMatrix(vs.RowSel, 0) <> "" Then
       sendkeys "{right}"
       sendkeys "{right}"
       sendkeys "{right}"
       sendkeys "{right}"
       sendkeys "{right}"
       sendkeys "{right}"
  End If
       
       VsFrame.Visible = False
       vs.Editable = flexEDKbdMouse
       vs.SetFocus
 '  End If
    
 End If
 

If vs.Col = 1 Then

If vs.TextMatrix(vs.RowSel, 1) <> "" Then
    
    'addType Trim(vs.TextMatrix(vs.RowSel, 1))
    
    If RS.State = 1 Then RS.close
    
    RS.Open "select Book,bookfont,book_size,binding,websheet from bookmaster " & _
    "where " & stringyear & " and bookno='" & Trim(vs.TextMatrix(vs.RowSel, 1)) & "'", con, adOpenKeyset, adLockReadOnly
    
    
    If RS.EOF = True Then Exit Sub
    If RS.EOF = False Then
        
        
         If RS!bookfont = "h" Then
           
           vs.Col = 0
           vs.CellFontName = hindi
           vs.CellFontSize = 14
            vs.TextMatrix(vs.RowSel, 0) = RS!Book & ""
           vs.Col = 1
        Else
           vs.CellFontName = english
           vs.CellFontSize = 12
           vs.TextMatrix(vs.RowSel, 0) = RS!Book & ""
           vs.TextMatrix(vs.RowSel, 1) = UCase(vs.TextMatrix(vs.RowSel, 1))
           vs.TextMatrix(vs.RowSel, 7) = RS!Binding & ""
           vs.TextMatrix(vs.RowSel, 9) = RS!book_size & ""
           vs.TextMatrix(vs.RowSel, 6) = RS!websheet & ""
           
        End If
        
        If Trim(vs.TextMatrix(vs.RowSel, 2)) = "" Then
          addType Trim(vs.TextMatrix(vs.RowSel, 1))
        End If

        
        sendkeys "{right}"

        
          
    End If
End If

ElseIf vs.Col = 2 Then
     sendkeys "{right}"
ElseIf vs.Col = 3 Then


    If RS.State = 1 Then RS.close
    RS.Open "select head1,head2,head3,head4,head5,headData1,headData2," & _
    "headData3,headData4,headData5,text_DBy,Inn_DBy,Exam_DBy,Inn_forms,text_forms,Exam_forms from bookmaster " & _
    " where bookno='" & Trim(vs.TextMatrix(vs.RowSel, 1)) & "'", con, adOpenKeyset, adLockReadOnly
    If RS.EOF = False Then
       For k1 = 1 To 5
           If RS.Fields("head" & k1) = vs.TextMatrix(vs.RowSel, 3) Then
                 vs.TextMatrix(vs.RowSel, 4) = RS.Fields("headData" & k1)
              
              If LCase(RS.Fields("head" & k1).Name) = "head1" Then
              'If LCase(Mid(RS.Fields("head" & k1), 1, 3)) = "inn" Then
                 vs.TextMatrix(vs.RowSel, 5) = RS.Fields("Inn_forms").value
              'ElseIf LCase(Mid(RS.Fields("head" & k1), 1, 3)) = "tex" Then
              ElseIf LCase(RS.Fields("head" & k1).Name) = "head2" Then
                 vs.TextMatrix(vs.RowSel, 5) = RS.Fields("text_forms").value
              'ElseIf LCase(Mid(RS.Fields("head" & k1), 1, 3)) = "exa" Then
              ElseIf LCase(RS.Fields("head" & k1).Name) = "head3" Then
                 vs.TextMatrix(vs.RowSel, 5) = RS.Fields("Exam_forms").value
              End If
           End If
       Next
    End If
    

    sendkeys "{right}"
ElseIf vs.Col = 4 Then
    
    If RS.State = 1 Then RS.close
    RS.Open "select DivideValue  from bookmaster " & _
    " where bookno='" & Trim(vs.TextMatrix(vs.RowSel, 1)) & "'", con, adOpenKeyset, adLockReadOnly
    If RS.EOF = False Then
       If Not IsNull(RS(0)) Then
         vs.TextMatrix(vs.RowSel, 5) = Round((IIf(vs.TextMatrix(vs.RowSel, 4) = "", 0, vs.TextMatrix(vs.RowSel, 4)) / RS(0)), 2)
       End If
    End If
    
    sendkeys "{right}"
ElseIf vs.Col = 5 Then
    sendkeys "{right}"
ElseIf vs.Col = 6 Then
    If vs.TextMatrix(vs.RowSel, 6) <> "" Then
       sendkeys "{right}"
    End If
ElseIf vs.Col = 7 Then
    If vs.TextMatrix(vs.RowSel, 7) <> "" Then
       sendkeys "{right}"
    End If
ElseIf vs.Col = 8 Then
    If vs.TextMatrix(vs.RowSel, 8) <> "" Then
       sendkeys "{right}"
    End If
ElseIf vs.Col = 9 Then
    sendkeys "{right}"
    sendkeys "{home}"
    sendkeys "{down}"
    sendkeys "{right}"
    Total

'ElseIf vs.Col = 8 Then
End If


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
    RS.Open "select * from ItemMaster where " & stringyear & " and ItemName='" & vs1.TextMatrix(vs1.RowSel, 0) & "'", con
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
      RS.Open "select * from ItemMaster where " & stringyear & " and ItemName='" & vs1.TextMatrix(I, 0) & "'", con
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



