VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBookIssue 
   ClientHeight    =   9048
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11376
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.4
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBookIssue.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9048
   ScaleWidth      =   11376
   Visible         =   0   'False
   Begin VB.Frame panel 
      Caption         =   "Book Receive/Issue"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   9108
      Left            =   135
      TabIndex        =   4
      Top             =   30
      Width           =   11160
      Begin VB.TextBox txtAmt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   9504
         TabIndex        =   35
         Text            =   "0"
         Top             =   7176
         Width           =   1296
      End
      Begin VB.Frame panel1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   900
         Left            =   360
         TabIndex        =   18
         Top             =   7752
         Width           =   7965
         Begin VB.CommandButton cmdRef 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Refresh"
            Height          =   675
            Left            =   90
            Picture         =   "frmBookIssue.frx":57EE
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   120
            Width           =   1245
         End
         Begin VB.CommandButton cmdExit 
            BackColor       =   &H00FFFFFF&
            Caption         =   "E&xit"
            Height          =   675
            Left            =   6615
            Picture         =   "frmBookIssue.frx":63D2
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   120
            Width           =   1245
         End
         Begin VB.CommandButton cmdDelete 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Delete"
            Height          =   675
            Left            =   4005
            Picture         =   "frmBookIssue.frx":6FB6
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   120
            Width           =   1245
         End
         Begin VB.CommandButton cmdSave 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Save"
            Height          =   675
            Left            =   1380
            Picture         =   "frmBookIssue.frx":7B9A
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   120
            Width           =   1245
         End
         Begin VB.CommandButton cmdEdit 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Edit"
            Height          =   675
            Left            =   2700
            Picture         =   "frmBookIssue.frx":877E
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   120
            Width           =   1245
         End
         Begin VB.CommandButton cmdPrint1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Print"
            Height          =   675
            Left            =   5340
            Picture         =   "frmBookIssue.frx":8BC0
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   120
            Width           =   1245
         End
      End
      Begin VB.TextBox txtTotal1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   8556
         TabIndex        =   17
         Text            =   "0"
         Top             =   7176
         Width           =   900
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2025
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1440
         Width           =   1545
      End
      Begin VB.ComboBox cboCategory 
         Height          =   315
         ItemData        =   "frmBookIssue.frx":97A4
         Left            =   2565
         List            =   "frmBookIssue.frx":97A6
         TabIndex        =   16
         Text            =   "cboCategory"
         Top             =   1830
         Width           =   3345
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00B8E4F1&
         Height          =   555
         Left            =   1965
         TabIndex        =   15
         Top             =   450
         Width           =   3840
         Begin VB.OptionButton OptionReceive 
            BackColor       =   &H00B8E4F1&
            Caption         =   "Receive Item"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   1890
            TabIndex        =   1
            Top             =   135
            Width           =   1770
         End
         Begin VB.OptionButton OptionIssue 
            BackColor       =   &H00B8E4F1&
            Caption         =   "Issue Item"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   135
            TabIndex        =   0
            Top             =   135
            Value           =   -1  'True
            Width           =   1770
         End
      End
      Begin VB.Frame Godown 
         BackColor       =   &H0078CFE9&
         Caption         =   "Godown Stock Transfer Option"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   2445
         Left            =   6930
         TabIndex        =   9
         Top             =   480
         Width           =   3696
         Begin VB.ComboBox cboGodown_in 
            Height          =   315
            Left            =   315
            Style           =   2  'Dropdown List
            TabIndex        =   12
            ToolTipText     =   "Select Godown Name"
            Top             =   1980
            Width           =   2985
         End
         Begin VB.ComboBox cboGodown_Out 
            Height          =   315
            ItemData        =   "frmBookIssue.frx":97A8
            Left            =   360
            List            =   "frmBookIssue.frx":97AA
            Style           =   2  'Dropdown List
            TabIndex        =   11
            ToolTipText     =   "Select Godown Name"
            Top             =   675
            Width           =   2940
         End
         Begin VB.CommandButton cmdGodown 
            Caption         =   "Add"
            Height          =   375
            Left            =   1350
            TabIndex        =   10
            Top             =   1080
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00B8E4F1&
            Caption         =   "Select Source Godown"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   360
            TabIndex        =   14
            Top             =   405
            Width           =   2940
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H00B8E4F1&
            Caption         =   "Select Destination Godown"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   315
            TabIndex        =   13
            Top             =   1710
            Width           =   2940
         End
      End
      Begin VB.Frame godown1 
         BackColor       =   &H0078CFE9&
         Caption         =   "Select Godown Name"
         ForeColor       =   &H00000000&
         Height          =   915
         Left            =   6930
         TabIndex        =   7
         Top             =   1875
         Width           =   3510
         Begin VB.ComboBox cbogodown1 
            Height          =   315
            Left            =   90
            Style           =   2  'Dropdown List
            TabIndex        =   8
            ToolTipText     =   "Select Godown Name"
            Top             =   315
            Width           =   2985
         End
      End
      Begin VB.ComboBox txtBinderName 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2025
         Style           =   1  'Simple Combo
         TabIndex        =   6
         Top             =   2190
         Width           =   4890
      End
      Begin VB.TextBox txtRemarks 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2025
         MaxLength       =   50
         TabIndex        =   5
         Top             =   2550
         Width           =   4905
      End
      Begin MSComCtl2.DTPicker recdate 
         Height          =   330
         Left            =   4350
         TabIndex        =   3
         Top             =   1425
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   572
         _Version        =   393216
         Format          =   135790593
         CurrentDate     =   39795
      End
      Begin VSFlex7Ctl.VSFlexGrid vs 
         Height          =   4080
         Left            =   300
         TabIndex        =   25
         Top             =   3000
         Width           =   10812
         _cx             =   19071
         _cy             =   7197
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
         BackColorSel    =   16762566
         ForeColorSel    =   16711680
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   12632256
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
         GridLines       =   8
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   500
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
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
         ExplorerBar     =   7
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
      Begin VB.TextBox txtBinder 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   5985
         MaxLength       =   50
         TabIndex        =   26
         Top             =   2250
         Width           =   825
      End
      Begin VB.Label Label4 
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
         Height          =   228
         Left            =   300
         TabIndex        =   34
         Top             =   7152
         Width           =   1548
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H0078CFE9&
         BorderWidth     =   4
         Height          =   1008
         Left            =   312
         Top             =   7692
         Width           =   8076
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total :"
         Height          =   192
         Index           =   6
         Left            =   7956
         TabIndex        =   33
         Top             =   7224
         Width           =   660
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "F2 For Search"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   1995
         TabIndex        =   32
         Top             =   1200
         Width           =   2220
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rec./Issue No :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   1
         Left            =   315
         TabIndex        =   31
         Top             =   1470
         Width           =   1920
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   2
         Left            =   3630
         TabIndex        =   30
         Top             =   1470
         Width           =   915
      End
      Begin VB.Label bindercode 
         BackStyle       =   0  'Transparent
         Caption         =   "Description :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   3
         Left            =   315
         TabIndex        =   29
         Top             =   2235
         Width           =   1860
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rec./Issue Category :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   3
         Left            =   315
         TabIndex        =   28
         Top             =   1875
         Width           =   2550
      End
      Begin VB.Label bindercode 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   0
         Left            =   315
         TabIndex        =   27
         Top             =   2550
         Width           =   1500
      End
   End
   Begin Crystal.CrystalReport CR 
      Left            =   10890
      Top             =   10350
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmBookIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim kk As Integer
Dim Edit As Boolean
Dim save_1 As Boolean
Dim rs_max As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Private Sub cboCategory_Click()
Dim rss As New ADODB.Recordset
If cboCategory.text = "Binder" Then
   txtBinderName.Enabled = True
ElseIf cboCategory.text = "Exchange" Then
   txtBinderName.Enabled = True
   bindercode(3).Enabled = True
Else
End If
End Sub
Private Sub cboCategory_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
If IssueBook = "StockTransfar" Then
   cboGodown_Out.SetFocus
Else
   cbogodown1.SetFocus
End If
End If
End Sub
Private Sub cboGodown_in_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If IssueBook = "StockTransfar" Then
   txtRemarks.SetFocus
End If
End If
End Sub
Private Sub cboGodown_Out_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If IssueBook = "StockTransfar" Then
   cboGodown_in.SetFocus
End If
End If

End Sub
Private Sub cbogodown1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
txtBinderName.Enabled = True
txtBinderName.SetFocus
End If
End Sub
Private Sub cboItem_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   vs.Col = 1
   vs.TextMatrix(vs.RowSel, 1) = cboItem.text
   cboItem.Visible = False
End If
End Sub
Private Sub cbogodown1_LostFocus()
If cbogodown1 = "" Then
MsgBox "Select Godown Name ...", vbInformation
cbogodown1.SetFocus
End If
End Sub
Private Sub cmdBookLedger_Click()
frmBookLedger.Height = ((Me.Height))
frmBookLedger.Width = (Me.Width)
frmBookLedger.top = 0
frmBookLedger.Left = 0

vsledger.Height = 6500
vsledger.Width = 11000
vsledger.top = 1000
frmBookLedger.Visible = True
vsledger.Cols = 6
vsledger.FormatString = "Book Code|Book Name|>Opening|Receive|>Issue|>Closing"
vsledger.ColWidth(0) = 1100
vsledger.ColWidth(1) = 4000
vsledger.ColWidth(2) = 1400
vsledger.ColWidth(3) = 1400
vsledger.ColWidth(4) = 1400
vsledger.ColWidth(5) = 1400


End Sub
Private Sub cmdDelete_Click()

If rs1.State = 1 Then rs1.close
rs1.Open "select * from BookStock Where " & stringyear & " and EntryNo = " & txtCode.text & ""
If rs1.EOF = False Then
   If rs2.State = 1 Then rs2.close
   rs2.Open "select * from BookStock Where " & stringyear & " and EntryNo = " & txtCode.text & ""
      If rs2!bAuthorized = True Then
          MsgBox "You can'nt change, Bill Already Locked !!", vbExclamation, "Alert"
          Exit Sub
      End If
End If



If MsgBox("Want to Delete ?", vbQuestion + vbYesNo) = vbYes Then
   con.Execute "delete from BookStock Where " & stringyear & " and EntryNo = " & txtCode.text & ""
   con.Execute "delete from BookStock_free Where " & stringyear & " and EntryNo = " & txtCode.text & ""
   
   cmdEdit.Enabled = False
   cmdDelete.Enabled = False
   cmdSave.Enabled = True

End If

setWidth
Call cmdref_Click

End Sub
Private Sub cmdedit_Click()

If rs1.State = 1 Then rs1.close
rs1.Open "select * from BookStock Where " & stringyear & " and EntryNo = " & txtCode.text & "", con
If rs1.EOF = False Then
   If rs2.State = 1 Then rs2.close
   rs2.Open "select * from BookStock Where " & stringyear & " and EntryNo = " & txtCode.text & "", con
      If rs2!bAuthorized = True Then
          MsgBox "You can'nt change, Bill Already Locked !!", vbExclamation, "Alert"
          Exit Sub
      End If
End If



Edit = True
vs.Editable = flexEDKbdMouse
cmdSave.Enabled = True
cmdDelete.Enabled = True
cmdEdit.Enabled = False
End Sub
Private Sub cmdexit_Click()
Unload Me
End Sub
Private Sub cmdExit1_Click()
frmBookLedger.Visible = False
End Sub
Private Sub cmdPrint1_Click()
If MsgBox(" Print ?", vbQuestion + vbYesNo) = vbYes Then
   
   DSNNew
   
   CR.Reset
   CR.ReportFileName = rptPath + "\printSlip.rpt"
   CR.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
   CR.ReplaceSelectionFormula "{bookstock.EntryNo}=" & txtCode.text & ""
   
   If OptionIssue.value = True Then
      CR.Formulas(0) = "head='" & "Issue Slip" & "'"
   Else
      CR.Formulas(0) = "head='" & "Receive Slip" & "'"
   End If
   
   CR.WindowShowCloseBtn = True
   CR.WindowShowPrintBtn = True
   CR.WindowControlBox = True
   CR.WindowShowPrintSetupBtn = True
   CR.WindowShowProgressCtls = True
   CR.WindowState = crptMaximized
   CR.Action = 1
End If
End Sub
Private Sub cmdref_Click()

save_1 = False
cmdSave.Enabled = True
vs.Editable = flexEDKbdMouse
setWidth
Edit = False
    
If RS.State = 1 Then RS.close
RS.Open "SELECT MAX(EntryNo) FROM BOOKSTOCK", con
If IsNull(RS(0)) Then
   txtCode.text = 1
Else
   txtCode.text = RS(0) + 1
End If


cbogodown1.ListIndex = -1
cboGodown_Out.ListIndex = -1
cboGodown_in.ListIndex = -1
txtBinder.text = ""
txtBinderName.ListIndex = -1
cboCategory.ListIndex = -1
txtTotal1.text = 0
txtRemarks.text = ""

If IssueBook = "Issue" Then
   cboCategory.Enabled = True
ElseIf IssueBook = "StockTransfar" Then
   cboCategory.text = "Transfer"
   cboCategory.Enabled = False
End If

txtBinderName.text = ""
cmdEdit.Enabled = False
cmdDelete.Enabled = False
cmdSave.Enabled = True
txtCode.SetFocus

End Sub
Private Sub cmdSave_Click()
   
   If IssueBook = "StockTransfar" Then
      
      
   If cboGodown_Out.text = "" Then
      MsgBox "Plz. Select Source godown ...", vbExclamation
      cboGodown_Out.SetFocus
      Exit Sub
   End If
   
   If cboGodown_in.text = "" Then
      MsgBox "Plz. Select Destination godown ...", vbExclamation
      cboGodown_in.SetFocus
      Exit Sub
   End If
      
    If cboGodown_in.text = cboGodown_Out.text Then
       MsgBox "Destination and source godown can'nt be same ...", vbCritical
       Exit Sub
    End If

Else

  
   If cbogodown1.text = "" Then
      MsgBox "Plz. Select godown ...", vbExclamation
      cbogodown1.SetFocus
      Exit Sub
   End If
   
   
   End If
   
   
   If cboCategory.text = "" Then
      MsgBox "Plz. Select Category ...", vbCritical
      cboCategory.SetFocus
      Exit Sub
   End If
   
   
   If cboCategory.text = "Binder" Then
        If txtBinderName.text = "" Then
           MsgBox "Select Binder Name ...", vbCritical
           txtBinderName.SetFocus
           Exit Sub
        End If
   End If
   
   
   If Val(txtTotal1.text) = 0 Then
      MsgBox "Plz. Enter Books ..", vbCritical
      txtTotal1.SetFocus
      Exit Sub
   End If
   
   
   
If MsgBox("Want to save ?", vbQuestion + vbYesNo) = vbYes Then
   save
   cmdSave.Enabled = False
   vs.Editable = flexEDNone
   cmdPrint1.SetFocus
End If
End Sub
Sub searchData()
On Error Resume Next

setWidth
Dim rs1 As New ADODB.Recordset

If RS.State = 1 Then RS.close
If IssueBook = "StockTransfar" Then

RS.Open "SELECT BookStock.BOOKCODE, BOOKS.BOOKNAME, BOOKS.RATE, BookStock.Qty,BookStock.EntryNo," & _
"BookStock.Dates,BookStock.Binder_Code,BookStock.Category,BookStock.Godown_In,BookStock.Godown_Out,BookStock.Issue_Receive,bookstock.remarks,BOOKS.discount,bookstock.amount FROM BOOKS " & _
"INNER JOIN BookStock ON BOOKS.BOOKCODE = BookStock.BOOKCODE Where BookStock.fyear='" & session & "' and BookStock.setupid=" & setupid & " and EntryNo = " & PopUpValue1 & " and GodownHead='StockTransfar' order by bookstock.auto", con, adOpenDynamic, adLockOptimistic

Else

RS.Open "SELECT BookStock.BOOKCODE, BOOKS.BOOKNAME, BOOKS.RATE, BookStock.Qty,BookStock.EntryNo," & _
"BookStock.Dates,BookStock.Binder_Code,BookStock.Category,BookStock.Godown_In,BookStock.Godown_Out,BookStock.Issue_Receive,bookstock.remarks,BOOKS.discount,bookstock.amount FROM BOOKS " & _
"INNER JOIN BookStock ON BOOKS.BOOKCODE = BookStock.BOOKCODE Where BookStock.fyear='" & session & "' and BookStock.setupid=" & setupid & " and EntryNo = " & PopUpValue1 & " and (GodownHead='Issue' or GodownHead='Receive') order by bookstock.auto", con, adOpenDynamic, adLockOptimistic

End If

If RS.EOF = False Then

vs.Editable = flexEDNone

Edit = True
cmdEdit.Enabled = True
cmdSave.Enabled = False

For I = 1 To vs.rows
If RS.EOF = False Then


If IssueBook = "StockTransfar" Then
    cboGodown_in.text = RS!Godown_in
    cboGodown_Out.text = RS!Godown_Out
    txtRemarks.text = RS.Fields("remarks").value & ""
Else
    cbogodown1.text = IIf(RS!Godown_Out = "-", RS!Godown_in, RS!Godown_Out)
End If

If RS!Issue_Receive = "Receive" Then
   OptionReceive.value = True
   OptionIssue.value = False
Else
   OptionReceive.value = False
   OptionIssue.value = True
End If

txtRemarks.text = RS.Fields("remarks").value & ""
txtCode.text = RS.Fields(4).value
recdate.value = RS.Fields(5).value
If RS.Fields("Binder_Code").value <> "" Then
  txtBinderName.text = RS.Fields("Binder_Code").value & ""
End If
cboCategory.text = RS!category



vs.TextMatrix(I, 0) = RS.Fields(0).value
vs.TextMatrix(I, 1) = RS.Fields(1).value
vs.TextMatrix(I, 2) = RS.Fields(2).value
vs.TextMatrix(I, 3) = RS.Fields(3).value

'If IsNull(RS!Rate) Then
'vs.TextMatrix(I, 2) = 0
'Else
'vs.TextMatrix(I, 2) = RS!Rate
'End If

If IsNull(RS!amount) Then
vs.TextMatrix(I, 4) = 0
Else
vs.TextMatrix(I, 4) = RS!amount
End If

If IsNull(RS!discount) Then
vs.TextMatrix(I, 5) = 0
Else
vs.TextMatrix(I, 5) = RS!discount
End If


RS.MoveNext
End If
Next

Total

End If

PopUpValue1 = ""

If RS.State = 1 Then RS.close
RS.Open "select * from SETUP where uname='" & UserName & "'", con
If RS.EOF = False Then
   If (RS!bedit = True) Then
      sendkeys "{tab}"
      cmdEdit.Enabled = True
   ElseIf (RS!bsave = True) Then
   
   If Edit = True Then
      cmdEdit.Enabled = False
   Else
      sendkeys "{tab}"
   End If
   
   End If
Else
   sendkeys "{tab}"
End If

End Sub
Sub save()

Dim kk As New ADODB.Recordset


If Edit = False Then

    If rs_max.State = 1 Then rs_max.close
    rs_max.Open "SELECT MAX(EntryNo) FROM BOOKSTOCK where " & stringyear, con, adOpenDynamic, adLockOptimistic
    If IsNull(rs_max(0)) Then
      txtCode.text = 1
    Else
      txtCode.text = rs_max(0) + 1
    End If

End If


If RS.State = 1 Then RS.close
If RS.State = 1 Then RS.close
RS.Open "select * from BookStock where " & stringyear & " and EntryNo=" & txtCode.text & " and GodownHead='" & IssueBook & "' order by auto", con, adOpenDynamic, adLockOptimistic
If RS.EOF = False Then

con.Execute "delete from BookStock where " & stringyear & " and EntryNo=" & txtCode.text & " and GodownHead='" & IssueBook & "'"
con.Execute "delete from BookStock_free where " & stringyear & " and EntryNo=" & txtCode.text & ""

For I = 1 To vs.rows - 1
If vs.TextMatrix(I, 0) <> "" Then
RS.AddNew


If IssueBook = "StockTransfar" Then
    RS!Godown_in = cboGodown_in.text
    RS!Godown_Out = cboGodown_Out.text
    RS!GodownHead = "StockTransfar"
Else
    
If OptionIssue = True Then
    RS!GodownHead = "Issue"
    RS!Godown_Out = cbogodown1.text
    RS!Godown_in = "-"
Else
    RS!GodownHead = "Receive"
    RS!Godown_in = cbogodown1.text
    RS!Godown_Out = "-"
End If

End If


RS.Fields("EntryNo").value = txtCode.text
RS.Fields("Dates").value = recdate.value
RS.Fields("BOOKCODE").value = vs.TextMatrix(I, 0)
RS.Fields("Qty").value = vs.TextMatrix(I, 3)
RS.Fields("Category").value = cboCategory.text
RS.Fields("Binder_Code").value = txtBinderName.text
If OptionIssue.value = True Then
   RS.Fields("Issue_Receive").value = "Issue"
Else
   RS.Fields("Issue_Receive").value = "Receive"
End If
RS.Fields("remarks").value = txtRemarks.text
RS.Fields("Fyear").value = session
RS.Fields("setupid").value = setupid
RS!BookDesc = ReturnBookDesc(vs.TextMatrix(I, 0))

RS!rate = vs.TextMatrix(I, 2)
RS!amount = vs.TextMatrix(I, 4)
RS!discount = vs.TextMatrix(I, 5)

RS.update

'===============================================================
If kk.State = 1 Then kk.close
kk.Open "select BOOKNAME,NoPrintDesc,Qty,bookcode,rate from KitQry where kitcode='" & vs.TextMatrix(I, 0) & "'", con
bkdesc = ""
While kk.EOF = False
 con.Execute "insert into BookStock_free(EntryNo,Dates,BOOKCODE,Qty,setupid,Fyear,Godown) " & _
 " values('" & RS!entryNo & "','" & Format(RS.Fields("Dates").value, "MM/dd/yyyy") & "','" & kk!Bookcode & "','" & Val(vs.TextMatrix(I, 3)) & "','" & setupid & "','" & session & "','" & RS!Godown_Out & "')"
 kk.MoveNext
Wend

'-----------------------------------------------------------------
'=================================================================

End If
Next

Else



con.Execute "delete from BookStock where " & stringyear & " and EntryNo=" & txtCode.text & ""
con.Execute "delete from BookStock_free where " & stringyear & " and EntryNo=" & txtCode.text & ""


For I = 1 To vs.rows - 1
If vs.TextMatrix(I, 0) <> "" Then
RS.AddNew

If IssueBook = "StockTransfar" Then
    RS!Godown_in = cboGodown_in.text
    RS!Godown_Out = cboGodown_Out.text
    RS!GodownHead = "StockTransfar"
Else
    
If OptionIssue = True Then
    RS!GodownHead = "Issue"
    RS!Godown_Out = cbogodown1.text
    RS!Godown_in = "-"
Else
    RS!GodownHead = "Receive"
    RS!Godown_in = cbogodown1.text
    RS!Godown_Out = "-"
End If
    
    
End If


RS.Fields("EntryNo").value = txtCode.text
RS.Fields("Dates").value = recdate.value
RS.Fields("BOOKCODE").value = vs.TextMatrix(I, 0)
RS.Fields("Qty").value = vs.TextMatrix(I, 3)
RS.Fields("Category").value = cboCategory.text
RS.Fields("Binder_Code").value = txtBinderName.text
If OptionIssue.value = True Then
   RS.Fields("Issue_Receive").value = "Issue"
Else
   RS.Fields("Issue_Receive").value = "Receive"
End If
RS.Fields("remarks").value = txtRemarks.text

RS.Fields("Fyear").value = session
RS.Fields("setupid").value = setupid
RS!BookDesc = ReturnBookDesc(vs.TextMatrix(I, 0))

RS!rate = vs.TextMatrix(I, 2)
RS!amount = vs.TextMatrix(I, 4)
RS!discount = vs.TextMatrix(I, 5)


RS.update

'===============================================================
If kk.State = 1 Then kk.close
kk.Open "select BOOKNAME,NoPrintDesc,Qty,bookcode,rate from KitQry where kitcode='" & vs.TextMatrix(I, 0) & "'", con
bkdesc = ""
While kk.EOF = False
 con.Execute "insert into BookStock_free(EntryNo,Dates,BOOKCODE,Qty,setupid,Fyear,Godown) " & _
 " values('" & RS!entryNo & "','" & Format(RS.Fields("Dates").value, "MM/dd/yyyy") & "','" & kk!Bookcode & "','" & (kk!qty * Val(vs.TextMatrix(I, 3))) & "','" & setupid & "','" & session & "','" & RS!Godown_Out & "')"
 kk.MoveNext
Wend

'-----------------------------------------------------------------
'=================================================================


End If
Next

cmdSave.Enabled = False

End If



cmdEdit.Enabled = False
cmdDelete.Enabled = False
cmdSave.Enabled = True


End Sub
Private Sub Form_Load()

Me.top = 100
Me.Left = 100
Me.Width = 11200
Me.Height = 9700

BackColorFrom Me, panel, panel1

Edit = False
cmdEdit.Enabled = False
cmdDelete.Enabled = False
cmdSave.Enabled = True

'txtBinderName.Clear
'If RS.State = 1 Then RS.close
'RS.Open "select SUBLEDGER as [Binder Name] from SLEDGER order by SUBLEDGER", con
'While RS.EOF = False
'   txtBinderName.AddItem RS(0)
'   RS.MoveNext
'Wend



If RS.State = 1 Then RS.close
RS.Open "select Godwn from Godownmaster where Binder_Printer='g' order by Godwn", con
While RS.EOF = False
    cboGodown_in.AddItem RS(0)
    cboGodown_Out.AddItem RS(0)
    cbogodown1.AddItem RS(0)
    RS.MoveNext
Wend

recdate.value = Date

setWidth


AddCategory

If IssueBook = "Issue" Or IssueBook = "Receice" Then
   Godown.Visible = False
   Frame1.Visible = True
   godown1.Visible = True
   Label1(1).Caption = "Issue/Rec. No"
   cboCategory.Enabled = True
   
   bindercode(3).Visible = True
   txtBinderName.Visible = True
   txtBinder.Visible = True

ElseIf IssueBook = "StockTransfar" Then
   Godown.Visible = True
   Frame1.Visible = False
   godown1.Visible = False
   Label1(1).Caption = "Transfer No."
   cboCategory.text = "Transfer"
   cboCategory.Enabled = False
   
   'bindercode(3).Visible = False
   'txtBinderName.Visible = False
   txtBinder.Visible = False
End If




If Book_Stock = "Book_Sp" Or Book_Stock = "Book_dem" Then
cmdBookLedger.Visible = False
Else
'cmdBookLedger.Visible = True
End If



'-----------------------------------------------------

If IssueBook = "Issue" Then
    If RS.State = 1 Then RS.close
    RS.Open "SELECT MAX(EntryNo) FROM BOOKSTOCK where " & stringyear & " and GodownHead='Issue'", con
    If IsNull(RS(0)) Then
       txtCode.text = 1
    Else
       txtCode.text = RS(0) + 1
    End If
Else
    If RS.State = 1 Then RS.close
    RS.Open "SELECT MAX(EntryNo) FROM BOOKSTOCK where " & stringyear & " and GodownHead='StockTransfar'", con
    If IsNull(RS(0)) Then
       txtCode.text = 1
    Else
       txtCode.text = RS(0) + 1
    End If

End If

'-----------------------------------------------------

If RS.State = 1 Then RS.close
RS.Open "SELECT MAX(EntryNo) FROM BOOKSTOCK", con
If IsNull(RS(0)) Then
   txtCode.text = 1
Else
   txtCode.text = RS(0) + 1
End If



cmdDelete.Enabled = False
cmdSave.Enabled = True

End Sub
Sub setWidth()
vs.Clear
vs.Cols = 6
vs.FormatString = "Code|Book Name||>Qty|>Amt(Rs.)|"

vs.ColWidth(0) = 1200
vs.ColWidth(1) = 6100
vs.ColWidth(2) = 0
vs.ColWidth(3) = 1800
vs.ColWidth(4) = 1000
vs.ColWidth(5) = 0

txtAmt.text = 0

sendkeys "{HOME}"

'formButtonValidation cmdDelete, cmdEdit

End Sub
Sub Total()
On Error Resume Next

txtTotal1.text = 0
txtAmt.text = 0


For I = 1 To vs.rows - 1
If vs.TextMatrix(I, 3) <> "" Then

txtTotal1.text = (Val(txtTotal1.text) + vs.TextMatrix(I, 3))
txtAmt.text = (Val(txtAmt.text) + vs.TextMatrix(I, 4))

End If
Next

End Sub
Sub AddCategory()
Dim rs_ca As New ADODB.Recordset
cboCategory.Clear


If IssueBook = "Issue" Then

If rs_ca.State = 1 Then rs_ca.close
If OptionIssue.value = True Then
rs_ca.Open "select name from Issue_ReceiveMaster where Category='Issue' order by Name", con
While rs_ca.EOF = False
   cboCategory.AddItem rs_ca(0)
   rs_ca.MoveNext
Wend
End If


ElseIf IssueBook = "Receive" Then

If rs_ca.State = 1 Then rs_ca.close
If OptionReceive.value = True Then
rs_ca.Open "select name from Issue_ReceiveMaster where Category='Receive' order by Name", con
While rs_ca.EOF = False
   cboCategory.AddItem rs_ca(0)
   rs_ca.MoveNext
Wend

End If


ElseIf IssueBook = "StockTransfar" Then

If rs_ca.State = 1 Then rs_ca.close
If OptionIssue.value = True Then
rs_ca.Open "select name from Issue_ReceiveMaster where Category='Transfer' order by Name", con
End If

While rs_ca.EOF = False
   cboCategory.AddItem rs_ca(0)
   rs_ca.MoveNext
Wend

End If



End Sub

Private Sub Form_Resize()
panel.Left = (Me.ScaleWidth - panel.Width) / 2
panel.top = (Me.ScaleHeight - panel.Height) / 2

End Sub

Private Sub OptionIssue_Click()

IssueBook = "Issue"
AddCategory
End Sub

Private Sub OptionIssue_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtCode.SetFocus
End Sub

Private Sub OptionReceive_Click()

IssueBook = "Receive"

AddCategory

End Sub

Private Sub OptionReceive_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtCode.SetFocus
End Sub

Private Sub recDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If cboCategory.Enabled = True Then
   cboCategory.SetFocus
Else
   cboGodown_Out.SetFocus
End If
End If
End Sub
Private Sub recdate_LostFocus()

If Trim(recdate.value) <> "" Then
    If Not checkdate(Trim(recdate.value), recdate) Then
       recdate.SetFocus
    End If
End If

End Sub

Private Sub txtBinderName_GotFocus()

If PopUpValue1 <> "" Then
''  txtCode.Text = PopUpValue2
txtBinderName.text = PopUpValue1
''  searchData
PopUpValue1 = ""
''  PopUpValue2 = ""
''  cmdDelete.Enabled = True
End If
''
End Sub

Private Sub txtBinderName_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
   txtRemarks.SetFocus
End If

If KeyCode = 113 Then
   
If cboCategory = "Binder" Then
    popuplist10 "select SUBLEDGER as [Binder Name],Address1 as Address,Address2 as City from SLEDGER order by SUBLEDGER,Code", con
ElseIf cboCategory = "Exchange" Then
    popuplist10 "select Party as [Customer],Code from SLEDGER where " & stringyear & " and gledger='SUNDRY DEBTORS' order by Party", con
End If

End If

End Sub

Private Sub txtCode_GotFocus()
If PopUpValue1 <> "" Then
   searchData
   cboCategory_Click
   
End If
End Sub

Private Sub txtcode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then

PopUpValue1 = txtCode
searchData
PopUpValue1 = ""







End If

If KeyCode = 113 Then
If IssueBook = "StockTransfar" Then
   popuplist10 "select EntryNo,Dates,Binder_Code as [Description],Remarks from bookstock where " & stringyear & " and GodownHead='StockTransfar'  group by EntryNo,Dates,Binder_Code,Remarks", con
ElseIf IssueBook = "Issue" Or IssueBook = "Receive" Then
   popuplist10 "select EntryNo,Dates,Binder_Code as [Description],Remarks from bookstock where (" & stringyear & " and GodownHead='Issue' or GodownHead='Receive')  group by EntryNo,Dates,Binder_Code,Remarks", con
End If
End If
End Sub
Private Sub txtRemarks_GotFocus()
  If PopUpValue1 <> "" Then
     txtRemarks.text = PopUpValue1
     PopUpValue1 = ""
  End If
End Sub
Private Sub txtRemarks_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 113 Then
 If cboCategory.text = "Exchange" Then
    popuplist10 "Select party as [SUBLEDGER],Code from SLEDGER order by party", con
 End If
 End If
 
 If KeyCode = 13 Then
    vs.SetFocus
 End If
 
End Sub

Private Sub vs_GotFocus()

If PopUpValue1 <> "" Then
If vs.Col = 0 Then
    If RS.State = 1 Then RS.close
    RS.Open "SELECT BOOKNAME,RATE,BOOKCODE FROM bOOKS WHERE " & stringyear & " and BOOKCODE='" & PopUpValue2 & "'"
    If RS.EOF = False Then
        vs.TextMatrix(vs.RowSel, 0) = RS!Bookcode
        vs.TextMatrix(vs.RowSel, 1) = RS!Bookname
        vs.TextMatrix(vs.RowSel, 2) = RS!rate
        PopUpValue1 = ""
    End If
ElseIf vs.Col = 1 Then
    If RS.State = 1 Then RS.close
    RS.Open "SELECT BOOKNAME,RATE,BOOKCODE FROM bOOKS WHERE " & stringyear & " and BOOKCODE='" & PopUpValue2 & "'"
    If RS.EOF = False Then
        vs.TextMatrix(vs.RowSel, 0) = RS!Bookcode
        vs.TextMatrix(vs.RowSel, 1) = RS!Bookname
        vs.TextMatrix(vs.RowSel, 2) = RS!rate
        PopUpValue1 = ""
    End If
    End If

End If



End Sub
Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)

If (vs.Col = 1) Then
If KeyCode = 13 Then
   popuplist10 "Select BOOKNAME,BOOKCODE from BOOKS order by BOOKNAME", con
End If
ElseIf (vs.Col = 0) Then
If KeyCode = 112 Then
   popuplist10 "Select BOOKNAME,BOOKCODE from BOOKS order by BOOKNAME", con
End If
End If






If vs.Col = 2 Then
    If KeyCode = 13 Then
       
       If RS.State = 1 Then RS.close
       RS.Open "select BookName,BookCode,Rate from Books where " & stringyear & " and BookCode='" & vs.TextMatrix(vs.RowSel, 0) & "'", con
       If RS.EOF = False Then
          vs.TextMatrix(vs.RowSel, 1) = RS.Fields(0).value
          vs.TextMatrix(vs.RowSel, 2) = RS.Fields(2).value
          sendkeys "{right}"
       End If
       
    End If
End If




If KeyCode = 13 Then
If vs.Col = 1 Then
        sendkeys "{RIGHT}"
        If PopUpValue2 <> "" Then
        If PopUpValue2 <> "" Then
           vs.TextMatrix(vs.RowSel, 0) = PopUpValue2
        End If
        
        If PopUpValue1 <> "" Then
           vs.TextMatrix(vs.RowSel, 1) = PopUpValue1
        End If
        
        End If
        PopUpValue2 = ""
        PopUpValue1 = ""
        
End If
End If



End Sub


Private Sub vs_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    
    If Col = 3 Then
      b = Grid_Validation(KeyAscii)
      If b = False Then KeyAscii = 0
    End If

End Sub

Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then

Dim s1, s2, s3 As String
Dim q As Long
Dim r As Double
Dim D As Double

'====================================================


If vs.Col = 0 Then
    If KeyCode = 13 Then
       
       If RS.State = 1 Then RS.close
       RS.Open "select BookName,BookCode,Rate,discount from Books where " & stringyear & " and BookCode='" & vs.TextMatrix(vs.RowSel, 0) & "'", con
       If RS.EOF = False Then
          vs.TextMatrix(vs.RowSel, 1) = RS.Fields(0).value
          vs.TextMatrix(vs.RowSel, 2) = RS.Fields(2).value
          vs.TextMatrix(vs.RowSel, 5) = RS.Fields("discount").value
          
          r = RS.Fields(2).value
          D = RS.Fields("discount").value
          
          sendkeys "{right}"
          sendkeys "{right}"
          ''SendKeys "{right}"
       Else
          
          If Val(txtTotal1.text) > 0 Then
          If cmdSave.Enabled = True Then
          If vs.TextMatrix(vs.RowSel, 0) = "" And vs.TextMatrix(vs.RowSel, 2) = "" And vs.TextMatrix(vs.RowSel, 2) = "" Then
          If save_1 = False Then
             cmdSave_Click
             save_1 = True
             cmdSave.Enabled = False
          End If
             'cmdRef_Click
          End If
          End If
          End If
          
       End If
       
    End If
End If

'===================================================


If vs.Col = 2 Then
      
       If RS.State = 1 Then RS.close
       RS.Open "select BookName,BookCode,Rate from Books where " & stringyear & " and BookCode='" & vs.TextMatrix(vs.RowSel, 0) & "'", con
       If RS.EOF = False Then
          vs.TextMatrix(vs.RowSel, 1) = RS.Fields(0).value
          vs.TextMatrix(vs.RowSel, 2) = RS.Fields(2).value
          
          
          
          
          
          sendkeys "{right}"
       End If
 
End If

If vs.Col = 3 Then

'Format(Round((q * r) * (D / 100), 2), "0.00")
r = vs.TextMatrix(vs.RowSel, 2)
D = vs.TextMatrix(vs.RowSel, 5)
q = vs.TextMatrix(vs.RowSel, 3)
vs.TextMatrix(vs.RowSel, 4) = (r * q)

vs.TextMatrix(vs.RowSel, 4) = Format(Round((r * q) - ((q * r) * (D / 100)), 2), "0.00")


If vs.TextMatrix(vs.RowSel, 3) <> "" Then
    
 '------------------
 
 s1 = ""
 s2 = ""
 s3 = ""
 
 
 
    
If IssueBook <> "StockTransfar" Then
    
 If OptionIssue.value = True Then
    s1 = "Issue"
    s2 = cbogodown1.text
    s3 = "-"
 Else
    s1 = "Receive"
    s2 = "-"
    s3 = cbogodown1.text
 End If
 
 
 

Else

 
    If IssueBook <> "Issue" Then
      If cboGodown_in.text = cboGodown_Out.text Then
         MsgBox "Destination and source godown can'nt be same ...", vbCritical
         
         Exit Sub
      End If
   End If



 If OptionIssue.value = True Then
    s1 = "Issue"
    s2 = cboGodown_Out.text
    s3 = cboGodown_in.text
 Else
    s1 = "Receive"
    s2 = cboGodown_Out.text
    s3 = cboGodown_in.text
 End If
 
    
    
End If
    
    '=======================================
    
    
    
    sendkeys "{Home}"
    sendkeys "{down}"
    
    
    
    
End If


              
End If

Total

End If

End Sub
Private Sub vs_SelChange()
    vs.TextMatrix(vs.RowSel, 0) = UCase(vs.TextMatrix(vs.RowSel, 0))
    
    
If vs.Col = 1 Then
vs.Editable = flexEDNone
Else
vs.Editable = flexEDKbdMouse
End If

End Sub
