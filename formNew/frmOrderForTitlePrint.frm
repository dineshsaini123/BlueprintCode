VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOrderForTitlePrint 
   ClientHeight    =   9084
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   14400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9084
   ScaleWidth      =   14400
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
      Height          =   8730
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Width           =   14340
      Begin VB.ComboBox txtFirmName 
         Height          =   315
         ItemData        =   "frmOrderForTitlePrint.frx":0000
         Left            =   6600
         List            =   "frmOrderForTitlePrint.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   270
         Width           =   5100
      End
      Begin VB.TextBox txtTotal1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   690
         Left            =   360
         TabIndex        =   20
         Text            =   "0"
         Top             =   9225
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.TextBox txtHeating 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1890
         TabIndex        =   0
         Top             =   495
         Width           =   1380
      End
      Begin VB.TextBox txtParty 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6600
         TabIndex        =   22
         Top             =   645
         Width           =   5070
      End
      Begin VB.TextBox txtRemarks 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1890
         TabIndex        =   21
         Top             =   1305
         Width           =   4620
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
         Left            =   570
         TabIndex        =   15
         Top             =   9285
         Visible         =   0   'False
         Width           =   465
         Begin VB.TextBox txtRawAndCasting 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3690
            TabIndex        =   16
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
            TabIndex        =   19
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
            TabIndex        =   18
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
            TabIndex        =   17
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
         Left            =   6600
         TabIndex        =   14
         Top             =   975
         Width           =   5070
      End
      Begin VB.Frame buttonFrame 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   810
         Left            =   315
         TabIndex        =   5
         Top             =   7800
         Width           =   7665
         Begin VB.CommandButton cmdAdd_1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Add"
            Height          =   660
            Left            =   45
            Picture         =   "frmOrderForTitlePrint.frx":0004
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   60
            Width           =   1065
         End
         Begin VB.CommandButton cmdSave_2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Save"
            Height          =   660
            Left            =   1155
            Picture         =   "frmOrderForTitlePrint.frx":0BE8
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   60
            Width           =   1065
         End
         Begin VB.CommandButton cmdDelete_3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Delete"
            Height          =   660
            Left            =   2265
            Picture         =   "frmOrderForTitlePrint.frx":17CC
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   60
            Width           =   1065
         End
         Begin VB.CommandButton cmdEdit_4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Edit"
            Height          =   660
            Left            =   3375
            Picture         =   "frmOrderForTitlePrint.frx":23B0
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   60
            Width           =   1065
         End
         Begin VB.CommandButton cmdUndo_5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Undo"
            Height          =   660
            Left            =   4485
            Picture         =   "frmOrderForTitlePrint.frx":27F2
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   60
            Width           =   1005
         End
         Begin VB.CommandButton cmdPrint_7 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Print"
            Height          =   660
            Left            =   5535
            Picture         =   "frmOrderForTitlePrint.frx":3036
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   60
            Width           =   1005
         End
         Begin VB.CommandButton cmdExit_12 
            BackColor       =   &H00FFFFFF&
            Caption         =   "E&xit"
            Height          =   660
            Left            =   6585
            Picture         =   "frmOrderForTitlePrint.frx":3C1A
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   60
            Width           =   1005
         End
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8685
         TabIndex        =   4
         Top             =   7410
         Width           =   1065
      End
      Begin VB.TextBox txtLoose 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   405
         TabIndex        =   3
         Top             =   9270
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtBinder 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6600
         TabIndex        =   2
         Top             =   1305
         Visible         =   0   'False
         Width           =   5070
      End
      Begin Crystal.CrystalReport CR 
         Left            =   3240
         Top             =   9360
         _ExtentX        =   593
         _ExtentY        =   593
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin MSComCtl2.DTPicker Dates 
         Height          =   315
         Left            =   1890
         TabIndex        =   13
         Top             =   810
         Width           =   1395
         _ExtentX        =   2455
         _ExtentY        =   550
         _Version        =   393216
         Format          =   166002689
         CurrentDate     =   39500
      End
      Begin VSFlex7Ctl.VSFlexGrid vs 
         Height          =   5550
         Left            =   270
         TabIndex        =   23
         Top             =   1785
         Width           =   13965
         _cx             =   24633
         _cy             =   9790
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
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
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
         Left            =   5535
         TabIndex        =   31
         Top             =   270
         Width           =   990
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H0078CFE9&
         BorderWidth     =   4
         Height          =   900
         Left            =   270
         Top             =   7755
         Width           =   7770
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Order No :"
         Height          =   270
         Index           =   0
         Left            =   420
         TabIndex        =   29
         Top             =   495
         Width           =   1530
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Printer Name "
         Height          =   300
         Index           =   2
         Left            =   5520
         TabIndex        =   28
         Top             =   630
         Width           =   1245
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date "
         Height          =   270
         Index           =   1
         Left            =   405
         TabIndex        =   27
         Top             =   855
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks "
         Height          =   300
         Index           =   4
         Left            =   390
         TabIndex        =   26
         Top             =   1290
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total :"
         Height          =   240
         Index           =   6
         Left            =   8205
         TabIndex        =   25
         Top             =   7425
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
         Left            =   1890
         TabIndex        =   24
         Top             =   270
         Width           =   2505
      End
   End
End
Attribute VB_Name = "frmOrderForTitlePrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim itemgp As String
Dim rates As Double
Dim I As Integer
Dim Status As String

Dim BB10 As Boolean

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
 
'' VsFrame.Top = vs.Top + ((vs.CellTop)) - 1500
'' VsFrame.Left = (vs.Left) - 150
End Sub
Sub Total()
    
txtTotal.text = 0
txtLoose.text = 0
For J = 1 To vs.rows - 1
If vs.TextMatrix(J, 0) <> "" Then
  txtTotal.text = (Val(txtTotal.text) + Val(vs.TextMatrix(J, 4)))
End If
Next
    

End Sub

Sub cellposiVs()
 Vs1Frame.Width = 2500
 Vs1Frame.top = vs1.top + ((vs1.CellTop))
 Vs1Frame.Left = (vs1.Left) + 550
End Sub
Sub AddItemInGrid()
Dim rs_1 As ADODB.Recordset
ff1 = ""

Set rs_1 = New ADODB.Recordset
Set rs_1 = con.Execute("exec searchList " & "bookmaster" & "")

While rs_1.EOF = False

If ff1 = "" Then
   ff1 = rs_1(0)
Else
   ff1 = ff1 & "|" & rs_1(0)
End If

rs_1.MoveNext
Wend

vs.ColComboList(0) = ff1


'''rs_1.Open "select BOOK,BookNo from BookMaster where " & stringyear & " order by Book", con, adOpenStatic, adLockReadOnly
''Set cboItem.RowSource = rs_1
''cboItem.ListField = "Book"
''cboItem.BoundColumn = "BookNo"
''cboItem.ReFill
    
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
        
        vs.TextMatrix(vs.RowSel, 5) = vs.Row
        
        
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
        'cellposiVs3
        
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
        'cboItemVs3.Visible = False
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
    Dates.value = Date
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
   
   'txtHeating.Text = MaxSNoNew("BillMaster", "bill_id", "Title")
   txtHeating.text = MaxOrderNo(txtFirmName)
   
End Sub

Private Sub cmdDelete_3_Click()
If MsgBox("Want to delete ?", vbQuestion + vbYesNo) = vbYes Then
con.Execute "delete from BillMaster where categories='Title' and bill_id='" & txtHeating.text & "' and " & stringyear
con.Execute "delete from Billtrans where categories='Title' and bill_id='" & txtHeating.text & "' and " & stringyear
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
rs1.Open "select set_name,BookCode from Billtrans where categories='Title' and bill_id='" & txtHeating.text & "' and " & stringyear & " order by set_name", con, adOpenStatic, adLockReadOnly
While rs1.EOF = False
If Len(set1) > 0 Then
If (set1 = rs1(0)) Then
Else
   J = 1
   k1 = k1 + 1
End If
End If

If RS.State = 1 Then RS.close
RS.Open "select set_name from billtrans where categories='Title' and BookCode= '" & rs1(1) & "' and set_name='" & rs1(0) & "' and bill_id='" & txtHeating.text & "' and " & stringyear, con, adOpenKeyset, adLockReadOnly
If RS.EOF = False Then

If rs_book.State = 1 Then rs_book.close
rs_book.Open "select bookfont from BookMaster where BookNo='" & rs1(1) & "' and " & stringyear, con, adOpenKeyset, adLockReadOnly
If rs_book.EOF = False Then
   fonttype = rs_book(0)
 Else
   fonttype = "e"
End If
 
 con.Execute "update Billtrans set bookfont='" & fonttype & "',GPOrderPrinting=" & J & ",OrderPrinting=" & k1 & " where categories='Title' and set_name='" & rs1(0) & "' and bill_id='" & txtHeating.text & "' and BookCode= '" & rs1(1) & "' and " & stringyear
 
 J = J + 1
End If
set1 = rs1(0)
rs1.MoveNext
Wend

End Sub
Private Sub cmdPrint_7_Click()

DSNNew

If MsgBox("Want To Print ?", vbInformation + vbYesNo) = vbYes Then
    CR.Reset
    CR.ReportFileName = rptPath & "/OrderTitle.rpt"
    CR.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    CR.ReplaceSelectionFormula "{BillMaster.categories}='Title' and {BillMaster.bill_id}='" & txtHeating.text & "' and {BillMaster.setupid}=" & setupid & " and {BillMaster.fyear}='" & session & "'"
    CR.Formulas(0) = "address='" & Text1.text & "'"
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

On Error GoTo aa1


If txtParty.text = "" Then
   MsgBox "Please Enter Binder Name !!", vbInformation
   Exit Sub
End If



If rs1.State = 1 Then rs1.close
rs1.Open "select * from PaperMakeMaster", con



If MsgBox("Want to Save ?", vbYesNo + vbQuestion) = vbYes Then

If RS.State = 1 Then RS.close
RS.Open "select * from BillMaster where " & stringyear & " and bill_id='" & txtHeating.text & "' and categories='Title'", con, adOpenDynamic, adLockOptimistic
If RS.EOF = False Then
   con.Execute "delete from BillMaster where categories='Title' and bill_id='" & txtHeating.text & "' and " & stringyear
   con.Execute "delete from billtrans where categories='Title' and bill_id='" & txtHeating.text & "' and " & stringyear
End If


RS.AddNew
RS.Fields("firm_id").value = txtFirmName.text
RS.Fields("bill_id").value = txtHeating.text
RS.Fields("dat").value = Dates.value
RS.Fields("PrinterName").value = txtParty.text
RS.Fields("Remarks").value = txtRemarks.text
RS.Fields("categories").value = "Title"
RS!setupid = setupid
RS!fyear = session

RS.update

If rs2.State = 1 Then rs2.close
rs2.Open "select * from Billtrans", con, adOpenDynamic, adLockOptimistic


For I = 1 To vs.rows - 1

If vs.TextMatrix(I, 0) <> "" Then

    rs2.AddNew
    rs2.Fields("bill_id").value = txtHeating.text
    rs2.Fields("BookName").value = vs.TextMatrix(I, 0)
    rs2.Fields("BookCode").value = vs.TextMatrix(I, 1)
    rs2.Fields("PaperName").value = vs.TextMatrix(I, 2)
    rs2.Fields("Set_Name").value = vs.TextMatrix(I, 3)
    rs2.Fields("Qty").value = vs.TextMatrix(I, 4)
    rs2.Fields("OrderPrinting").value = Val(vs.TextMatrix(I, 5))
    rs2.Fields("GPOrderPrinting").value = Val(vs.TextMatrix(I, 6))
    rs2.Fields("binderName1").value = vs.TextMatrix(I, 7)
    rs2.Fields("remdetails").value = vs.TextMatrix(I, 8)
    rs2.Fields("categories").value = "Title"
    
    k1 = InStr(vs.TextMatrix(I, 2), ":")
    If k1 > 0 Then
    
        rs1.MoveFirst
        rs1.Find "papermaker_id='" & Mid(vs.TextMatrix(I, 2), k1 + 1, 2) & "'"
        If rs1.EOF = False Then
          paperdetails = rs1!papermaker_name & " " & rs1!SizeValue1 & " X " & rs1!SizeValue2 & " - " & rs1!GSM & " " & IIf(rs1!eco = "None", "", rs1!eco)
        End If
          rs2.Fields("PaperDetails").value = paperdetails
        End If
        
        rs2!setupid = setupid
        rs2!fyear = session
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

On Error Resume Next

vs.rows = 50
I = 1

Dim rs1 As New ADODB.Recordset


Set RS = New ADODB.Recordset
'If rs.State = 1 Then rs.close
RS.Open "select * from BillMaster where " & stringyear & " and bill_id='" & txtHeating.text & "' and categories='Title'", con, adOpenDynamic, adLockOptimistic
If RS.EOF = False Then

cmdSave_2.Enabled = False
cmdDelete_3.Enabled = False
cmdEdit_4.Enabled = True

txtFirmName.text = RS.Fields("firm_id").value & ""

txtHeating.text = RS.Fields("bill_id").value
Dates.value = RS.Fields("dat").value
txtParty.text = RS.Fields("PrinterName").value
txtRemarks.text = RS.Fields("Remarks").value

'txtBinder.Text = rs.Fields("binderName").value & ""

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
" where (bm.fyear='" & session & "' and bm.setupid=" & setupid & " and b.bill_id='" & txtHeating.text & "' and bm.categories='Title')", con, adOpenDynamic, adLockOptimistic
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

'vs.TextMatrix(i, 0) = rs.Fields("BookName").value
vs.TextMatrix(I, 1) = RS.Fields("BookCode").value
vs.TextMatrix(I, 2) = RS.Fields("PaperName").value
vs.TextMatrix(I, 3) = RS.Fields("Set_Name").value
vs.TextMatrix(I, 4) = RS.Fields("Qty").value
vs.TextMatrix(I, 5) = RS.Fields("OrderPrinting").value
vs.TextMatrix(I, 6) = RS.Fields("GPOrderPrinting").value
vs.TextMatrix(I, 7) = RS.Fields("binderName1").value & ""
vs.TextMatrix(I, 8) = RS.Fields("remdetails").value & ""

I = I + 1
RS.MoveNext

Wend

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



Private Sub dates_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtFirmName.SetFocus
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

Me.Caption = "Order For Title Printing ..."
 
 Me.Left = 0
 Me.top = 100
 
 'Me.Width = 14000
 'Me.Height = 8500
 
 Me.Width = 14400
 Me.Height = 9495

 Screen.MousePointer = vbHourglass
 
 
 AddItemInGrid
 
 
 
 setWidth
 
 Dates.value = Date
 'txtHeating.Text = MaxSNo("BillMaster", "bill_id")
 'txtHeating.Text = MaxSNoNew("BillMaster", "bill_id", "Title")
 
 
 Dim s As String
 
 
 s = ""
 Set RS = New ADODB.Recordset
 RS.Open "select * from remarks order by head", con
 While RS.EOF = False
 If s = "" Then
 s = RS(0)
 Else
 s = s & "|" & RS(0)
 End If
 RS.MoveNext
 Wend
 
 vs.ColComboList(3) = s
 
'----------------------------------
 
 s = ""
 If RS.State = 1 Then RS.close
 RS.Open "select * from PaperMakeMaster where " & stringyear & " order by papermaker_name", con, adOpenStatic, adLockReadOnly
 While RS.EOF = False
 If s = "" Then
 s = RS!papermaker_name & ":" & RS!papermaker_id
 Else
 s = s & "|" & RS!papermaker_name & ":" & RS!papermaker_id & "," & RS!ptype & ":" & RS!eco & ":" & RS!SizeValue1 & "X" & RS!SizeValue2 & ":" & RS!GSM
 End If
 RS.MoveNext
 Wend
 
 vs.ColComboList(2) = s
 
 
 '================================
 s = ""
 If RS.State = 1 Then RS.close
 
 RS.Open "select Godwn as [Binder Name],Address from Godownmaster where " & stringyear & " and (Binder_Printer='b' or Binder_Printer='pb')", con, adOpenStatic, adLockReadOnly
 
 While RS.EOF = False
 If s = "" Then
 s = RS(0)
 Else
 s = s & "|" & RS(0)
 End If
 RS.MoveNext
 Wend
 
 vs.ColComboList(7) = s
 BackColorFrom Me
 
 
txtFirmName.Clear
If RS.State = 1 Then RS.close
RS.Open "select FirmName,Add1,Add2 from FirmMaster order by firmname", con, adOpenStatic, adLockReadOnly
While RS.EOF = False
 txtFirmName.AddItem RS(0)
 RS.MoveNext
Wend

txtFirmName.ListIndex = 0
txtHeating.text = MaxOrderNo(txtFirmName)
 
Screen.MousePointer = vbDefault
 
End Sub
Sub setWidth()
vs.Cols = 9

vs.FormatString = "Books Name|Books Code|<Paper Name & Code|^SET|>Qty|||Binder Name|Remarks"
vs.ColWidth(0) = 3000
vs.ColWidth(1) = 1100
vs.ColWidth(2) = 3000
vs.ColWidth(3) = 1000
vs.ColWidth(4) = 1000
vs.ColWidth(5) = 0
vs.ColWidth(6) = 0
vs.ColWidth(7) = 2200
vs.ColWidth(8) = 1800

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

Private Sub txtBinder_GotFocus()
If PopUpValue1 <> "" Then
  txtBinder = PopUpValue1
  PopUpValue1 = ""
End If

End Sub

Private Sub txtBinder_KeyDown(KeyCode As Integer, Shift As Integer)


If KeyCode = 113 Then
   value = "select Binder_name from BinderMaster order by Binder_name"
   popuplist1 value, con
End If


End Sub
Private Sub txtBinder_KeyPress(KeyAscii As Integer)

'If KeyAscii = 13 Then
'   If KeyAscii = 13 Then txtRemarks.SetFocus
'End If

End Sub
Private Sub txtFirmName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtParty.SetFocus
End Sub

Private Sub txtFirmName_LostFocus()
txtHeating.text = MaxOrderNo(txtFirmName)
End Sub

Private Sub txtHeating_GotFocus()
If PopUpValue1 <> "" Then
txtHeating.text = PopUpValue1
Dates.value = PopUpValue2
vs.Clear
setWidth
searchData
PopUpValue1 = ""
PopUpValue2 = ""
End If
End Sub

Private Sub txtHeating_KeyDown(KeyCode As Integer, Shift As Integer)


If KeyCode = 113 Then
   value = "select distinct bill_id,dat as OrderDate,PrinterName from BillMaster where " & stringyear & " and categories='Title' order by bill_id"
   popuplist1 value, con
End If


If KeyCode = 13 Then
searchData
End If

End Sub

Private Sub txtHeating_KeyPress(KeyAscii As Integer)
   
   If KeyAscii = 13 Then
        
   Dates.SetFocus

        
  End If
  

End Sub
Private Sub txtParty_GotFocus()
If PopUpValue1 <> "" Then
txtParty.text = PopUpValue1
Text1.text = PopUpValue2
PopUpValue1 = ""
PopUpValue2 = ""
PopUpValue3 = ""
End If
End Sub
Private Sub txtParty_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then
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
'     If vs.Col = 0 Then
'        cellposi
'        vs.TextMatrix(vs.RowSel, 0) = cboItem.Text
'     End If
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
        'vs.Editable = flexEDNone
        'VsFrame.Visible = False
        'cboItem.SetFocus
     Else
        'vs.Editable = flexEDKbdMouse
        'cellposi
     End If

  End If
  
  
  
  
  
End Sub

Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
        
          
 If vs.Col = 0 Then
 
      vs.Editable = flexEDNone
      VsFrame.Visible = False
      'cboItem.SetFocus
      
      If rs1.State = 1 Then rs1.close
      rs1.Open "Select top 1  BookNo from BookMaster where BOOK='" & vs.TextMatrix(vs.RowSel, 0) & "'", con
      If rs1.EOF = False Then
      vs.TextMatrix(vs.RowSel, 1) = rs1(0)
      End If
      
      
          
  If vs.TextMatrix(vs.RowSel, 0) <> "" Then
       sendkeys "{right}"
  End If
       
       VsFrame.Visible = False
       vs.Editable = flexEDKbdMouse
       vs.SetFocus
 '  End If
    
 End If
 

If vs.Col = 1 Then
    
    
    

If vs.TextMatrix(vs.RowSel, 1) <> "" Then
    If RS.State = 1 Then RS.close
    RS.Open "select Book,Binder,bookfont from bookmaster " & _
    "where " & stringyear & " and bookno='" & vs.TextMatrix(vs.RowSel, 1) & "'", con, adOpenKeyset, adLockReadOnly
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
        
        End If

        
        
        vs.TextMatrix(vs.RowSel, 7) = RS!Binder & ""
        sendkeys "{right}"
    End If
    
    vs.Col = 1
End If




ElseIf vs.Col = 2 Then
     sendkeys "{right}"
ElseIf vs.Col = 3 Then
    
    
    BB10 = False
    
    sendkeys "{right}"
    vs.TextMatrix(vs.RowSel, 5) = vs.Row
    
    For J = 1 To vs.rows - 1
    If vs.TextMatrix(vs.RowSel, 3) = vs.TextMatrix(J, 3) Then
        
     If BB10 = False Then
        pname = vs.TextMatrix(J, 2)
        vs.TextMatrix(vs.RowSel, 2) = pname
        vs.TextMatrix(vs.RowSel, 4) = vs.TextMatrix(J, 4)
        BB10 = True
     End If
        
    Else
    
      If vs.TextMatrix(vs.RowSel, 2) = "" Then
       '  vs.Col = 2
      End If
        
    End If
    Next
    
    
    
 
ElseIf vs.Col = 4 Then
    sendkeys "{right}"
ElseIf vs.Col = 7 Then
    sendkeys "{right}"

ElseIf vs.Col = 8 Then
    sendkeys "{home}"
    sendkeys "{down}"
    sendkeys "{right}"
    Total
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
'  total
End Sub
Private Sub vs1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'     If vs1.Col = 0 Then
'        cellposiVs
'        vs1.TextMatrix(vs1.RowSel, 0) = cboitemvs1.Text
'     End If

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

Private Sub vs_SelChange()

'If vs.Col = 5 Then
'   vs.Editable = flexEDNone
'Else
'   vs.Editable = flexEDKbdMouse
'End If

End Sub


