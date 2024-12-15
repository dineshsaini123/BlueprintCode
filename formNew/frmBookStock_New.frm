VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBookStock_New 
   ClientHeight    =   9636
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   13212
   Icon            =   "frmBookStock_New.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9636
   ScaleWidth      =   13212
   Begin VB.Frame panel 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9405
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   12930
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Closing-ORD"
         Height          =   375
         Left            =   9180
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   2160
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdOp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Closing-CH"
         Height          =   375
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   2160
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdStock_Order 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Today Stock Reflect With Order"
         Height          =   735
         Left            =   7920
         Picture         =   "frmBookStock_New.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   810
         Width           =   2400
      End
      Begin VB.CheckBox Check1_selectDate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Select Date For Sold Book"
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   5220
         TabIndex        =   44
         Top             =   1980
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Frame Frame1_dateselection 
         BackColor       =   &H00FFFFFF&
         Height          =   2265
         Left            =   1440
         TabIndex        =   33
         Top             =   0
         Visible         =   0   'False
         Width           =   2760
         Begin VB.CommandButton cmdOK 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&OK"
            Height          =   330
            Left            =   1305
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   1890
            Width           =   1275
         End
         Begin MSComCtl2.DTPicker fromDate1 
            Height          =   330
            Left            =   1305
            TabIndex        =   34
            Top             =   360
            Width           =   1305
            _ExtentX        =   2307
            _ExtentY        =   593
            _Version        =   393216
            Format          =   240386049
            CurrentDate     =   39795
         End
         Begin MSComCtl2.DTPicker toDate1 
            Height          =   330
            Left            =   1305
            TabIndex        =   35
            Top             =   675
            Width           =   1305
            _ExtentX        =   2307
            _ExtentY        =   593
            _Version        =   393216
            Format          =   240386049
            CurrentDate     =   39795
         End
         Begin MSComCtl2.DTPicker fromDate2 
            Height          =   330
            Left            =   1305
            TabIndex        =   36
            Top             =   1260
            Width           =   1305
            _ExtentX        =   2307
            _ExtentY        =   593
            _Version        =   393216
            Format          =   240386049
            CurrentDate     =   39795
         End
         Begin MSComCtl2.DTPicker toDate2 
            Height          =   330
            Left            =   1305
            TabIndex        =   37
            Top             =   1575
            Width           =   1305
            _ExtentX        =   2307
            _ExtentY        =   593
            _Version        =   393216
            Format          =   240386049
            CurrentDate     =   39795
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FF8080&
            Caption         =   " Last Session :"
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   45
            Top             =   90
            Width           =   2985
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "To Date :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Index           =   4
            Left            =   225
            TabIndex        =   42
            Top             =   1575
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "From Date :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Index           =   3
            Left            =   225
            TabIndex        =   41
            Top             =   1305
            Width           =   1095
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FF8080&
            Caption         =   " Current Session :"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   40
            Top             =   1035
            Width           =   2985
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "To Date :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Index           =   2
            Left            =   225
            TabIndex        =   39
            Top             =   675
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "From Date :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Index           =   1
            Left            =   225
            TabIndex        =   38
            Top             =   405
            Width           =   1095
         End
      End
      Begin VB.CheckBox Check1_crm 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Noida Stock (CRM)"
         Height          =   390
         Left            =   5220
         TabIndex        =   32
         Top             =   1710
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CommandButton Command1_excel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Export To Excel"
         Height          =   600
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1590
         Width           =   2400
      End
      Begin VB.CommandButton cmdBookTransfer 
         BackColor       =   &H00FFFFC1&
         Caption         =   "&Closing Transfar"
         Height          =   645
         Left            =   10425
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   1515
         Width           =   1860
      End
      Begin VB.CheckBox Check1_sp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Stock For Noida  Godown"
         Height          =   390
         Left            =   5220
         TabIndex        =   29
         Top             =   1395
         Visible         =   0   'False
         Width           =   2535
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
         Left            =   2925
         TabIndex        =   28
         Top             =   1200
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1515
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
         Left            =   2925
         TabIndex        =   27
         Top             =   1470
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton cmdExit1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Exit"
         Height          =   645
         Left            =   10425
         Picture         =   "frmBookStock_New.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   825
         Width           =   1860
      End
      Begin VB.CommandButton cmdPrint_7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Print"
         Height          =   645
         Left            =   10440
         Picture         =   "frmBookStock_New.frx":1468
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   135
         Width           =   1860
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   465
         Left            =   12915
         TabIndex        =   9
         Top             =   1080
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.CommandButton cmdOpening 
         Caption         =   "Data Update"
         Height          =   465
         Left            =   13095
         TabIndex        =   8
         Top             =   600
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.CommandButton LaVolpeButton1 
         Caption         =   "Book Wise List"
         Height          =   600
         Left            =   12960
         TabIndex        =   7
         Top             =   360
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdView 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Book Wise Balance List"
         Height          =   645
         Left            =   7920
         Picture         =   "frmBookStock_New.frx":204C
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   90
         Width           =   2400
      End
      Begin VB.ComboBox cboGroup 
         Height          =   315
         ItemData        =   "frmBookStock_New.frx":2C30
         Left            =   5355
         List            =   "frmBookStock_New.frx":2C32
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   2445
      End
      Begin VB.TextBox txtBookCode 
         Height          =   315
         Left            =   5355
         TabIndex        =   3
         Text            =   "All"
         Top             =   675
         Width           =   2445
      End
      Begin VB.ComboBox cboBinder_Godown 
         Height          =   288
         ItemData        =   "frmBookStock_New.frx":2C34
         Left            =   1530
         List            =   "frmBookStock_New.frx":2C41
         Style           =   1  'Simple Combo
         TabIndex        =   2
         Top             =   765
         Width           =   2310
      End
      Begin VB.ComboBox cboCategory 
         Height          =   315
         ItemData        =   "frmBookStock_New.frx":2C5A
         Left            =   1530
         List            =   "frmBookStock_New.frx":2C67
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   2310
      End
      Begin MSComCtl2.DTPicker fromdate 
         Height          =   285
         Left            =   13050
         TabIndex        =   11
         Top             =   1035
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2561
         _ExtentY        =   508
         _Version        =   393216
         Format          =   240451585
         CurrentDate     =   39795
      End
      Begin MSComCtl2.DTPicker dateAson 
         Height          =   330
         Left            =   1530
         TabIndex        =   12
         Top             =   1200
         Width           =   1305
         _ExtentX        =   2286
         _ExtentY        =   593
         _Version        =   393216
         Format          =   240451585
         CurrentDate     =   39795
      End
      Begin MSComCtl2.DTPicker todate 
         Height          =   285
         Left            =   13050
         TabIndex        =   13
         Top             =   405
         Visible         =   0   'False
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   508
         _Version        =   393216
         Format          =   240451585
         CurrentDate     =   39795
      End
      Begin MSComCtl2.DTPicker datefrom 
         Height          =   285
         Left            =   13095
         TabIndex        =   14
         Top             =   270
         Visible         =   0   'False
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   487
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   16777215
         CalendarTrailingForeColor=   16777215
         Format          =   240451585
         CurrentDate     =   39795
      End
      Begin Crystal.CrystalReport CR 
         Left            =   0
         Top             =   0
         _ExtentX        =   593
         _ExtentY        =   593
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VSFlex7Ctl.VSFlexGrid vs 
         Height          =   6585
         Left            =   45
         TabIndex        =   25
         Top             =   2700
         Width           =   12795
         _cx             =   22569
         _cy             =   11615
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
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
         ForeColorSel    =   -2147483647
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
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
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
      Begin VB.ComboBox cboserName 
         Height          =   315
         ItemData        =   "frmBookStock_New.frx":2C80
         Left            =   5355
         List            =   "frmBookStock_New.frx":2C82
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   990
         Width           =   2445
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "SerName"
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
         Height          =   285
         Left            =   4365
         TabIndex        =   49
         Top             =   990
         Width           =   1050
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
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
         Height          =   285
         Left            =   4365
         TabIndex        =   24
         Top             =   360
         Width           =   870
      End
      Begin VB.Label bookbalance 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11715
         TabIndex        =   23
         Top             =   1110
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Upload From"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   4
         Left            =   12285
         TabIndex        =   22
         Top             =   270
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "As On"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   11700
         TabIndex        =   21
         Top             =   1080
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "F2 For Search Book"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   5355
         TabIndex        =   20
         Top             =   135
         Width           =   1995
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Book "
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
         Index           =   0
         Left            =   4365
         TabIndex        =   19
         Top             =   675
         Width           =   870
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Height          =   150
         Left            =   45
         TabIndex        =   18
         Top             =   2460
         Width           =   14325
      End
      Begin VB.Label binderlbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   135
         TabIndex        =   17
         Top             =   765
         Width           =   690
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Binder/Godown"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   135
         TabIndex        =   16
         Top             =   360
         Width           =   1905
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "As On "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   0
         Left            =   135
         TabIndex        =   15
         Top             =   1200
         Width           =   690
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   0
      X1              =   0
      X2              =   15120
      Y1              =   -945
      Y2              =   -945
   End
End
Attribute VB_Name = "frmBookStock_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sel_key
Dim Sel_Parant_Key
Dim d_from, d_to As Date
Dim Sort_Key As String
Dim bb_ As Boolean
Dim ddd As Integer
Private Sub cboBinder_Godown_GotFocus()

If PopUpValue1 <> "" Then
    cboBinder_Godown.text = PopUpValue1
    PopUpValue1 = ""
End If

End Sub
Private Sub cboBinder_Godown_KeyDown(KeyCode As Integer, Shift As Integer)

If cboCategory.ListIndex = 0 Then

    If KeyCode = 113 Then
        popuplistModel10 "select Godwn as [Binder Name],Address from Godownmaster where " & stringyear & " and binder_printer='b' order by Godwn", con
    End If
    
ElseIf cboCategory.ListIndex = 1 Then
    
    popuplistModel10 "select Godwn as [Binder Name],Address from Godownmaster where " & stringyear & " and binder_printer='g' order by Godwn", con
    
End If


End Sub
Private Sub cboCategory_Click()
If cboCategory.ListIndex = 2 Then
   cboBinder_Godown.Enabled = False
   binderlbl.Enabled = False
   'cboBinder_Godown.Clear
Else
   cboBinder_Godown.Enabled = True
   binderlbl.Enabled = True

If cboCategory.ListIndex = 0 Then
Else
End If


End If
End Sub

Private Sub cboGroup_Click()

cboserName.Clear
cboserName.AddItem ""
If RS.State = 1 Then RS.Close
RS.Open "select serName from BOOKS where GROUPCODE='" & cboGroup.text & "' group by serName order by serName", con
While RS.EOF = False
  If Not IsNull(RS!sername) Then
  cboserName.AddItem RS!sername
  End If
  RS.MoveNext
Wend


'cboGroup.ListIndex = cboGroup.ListCount - 1
'========================================
End Sub

Private Sub cboserName_Click()
'  If RS.State = 1 Then RS.close
'  RS.Open "select groupcode from BOOKS where sername='" & cboserName.text & "'", con
'  If RS.EOF = False Then
'  cboGroup.text = RS(0)
'  txtBookCode.text = "All"
'  End If
End Sub

Private Sub Check1_crm_Click()
If Check1_crm.value = 1 Then

   cboBinder_Godown.text = "NS"
   cboCategory.text = "Godown"
   'cboGroup.Text = "All"
   cboBinder_Godown.Enabled = False
   cboCategory.Enabled = False
   cmdBookTransfer.Enabled = False

Else
   
   cboBinder_Godown.text = ""
   cboCategory.ListIndex = -1
   cboBinder_Godown.Enabled = True
   cboCategory.Enabled = True
   cmdBookTransfer.Enabled = True
   
End If

End Sub

Private Sub Check1_selectDate_Click()
   
   If Check1_selectDate.value = 1 Then
      Frame1_dateselection.Visible = True
      
      fromDate2.value = from_date
      toDate2.value = to_date
        
      fromDate1.value = Mid(from_date, 1, 9) & "" & Int(Mid(from_date, 10)) - 1
      toDate1.value = Mid(to_date, 1, 9) & "" & Int(Mid(to_date, 10)) - 1
      
      cmdView.Enabled = False
      cmdStock_Order.Enabled = False
      
      
   Else
      Frame1_dateselection.Visible = False
      
      cmdView.Enabled = True
      cmdStock_Order.Enabled = True

   End If
   
End Sub
Private Sub cmdBookTransfer_Click()

Dim contr As New ADODB.Connection
Dim db, db1 As String
db = ""
db1 = ""

If cboGroup = "" Then
   MsgBox "Please Select Group Name...", vbInformation
   cboGroup.SetFocus
   Exit Sub
End If


If MsgBox("Want to transfar..", vbQuestion + vbYesNo) = vbYes Then

    I = Right(session, 2) + 1
    J = I - 1
    db1 = J & "" & I
    db = "RLData_" & db1
    Set contr = New ADODB.Connection
    
    
    If LCase(server_) = "server" Then
       contr.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverName_ & "; DATABASE=" & db & "; uid=" & sql_user & "; PWD=" & sql_pass
       contr.Open
    Else
       contr.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverName_ & "; DATABASE=" & db & "; UID=; PWD=;"
       contr.Open
    End If

    contr.Execute "delete from BookOpening where Godown='" & cboBinder_Godown.text & "' and bookgp='" & cboGroup.text & "'"
    
    For I = 1 To vs.rows - 1
        If vs.TextMatrix(I, 0) <> "" Then
           If cboBinder_Godown.text <> "" Then
            contr.Execute "insert into BookOpening(BOOKCODE,ItemName,Balance,Godown,bookgp) values('" & vs.TextMatrix(I, 0) & "','" & vs.TextMatrix(I, 1) & "','" & vs.TextMatrix(I, 2) & "','" & cboBinder_Godown.text & "','" & cboGroup.text & "')"
           End If
        End If
    Next
    
    MsgBox "Book Closing Transfar..", vbInformation
End If

End Sub

Private Sub cmdExit_Click()
panel.Visible = False
End Sub

Private Sub cmdGodown_Click()

End Sub

Private Sub cmdExit1_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()

CCON.Execute "update  saleDate set fromDate='" & fromDate1.value & "' , toDate='" & toDate1.value & "' where setupid=1"
CCON.Execute "update  saleDate set fromDate='" & fromDate2.value & "' , toDate='" & toDate2.value & "' where setupid=2"






Frame1_dateselection.Visible = False
End Sub

Private Sub cmdOp_Click()

If cboBinder_Godown.text = "NS" Then

con.Execute "delete from BookDiff where gp='" & cboGroup.text & "' and godown='NS' and stockType='CH'"

If MsgBox("want to Closing ? ", vbQuestion + vbYesNo) = vbNo Then
   Exit Sub
End If


For I = 1 To vs.rows - 1

If vs.TextMatrix(I, 1) <> "" Then
      
  If Val(vs.TextMatrix(I, 2)) > 0 Then
      q = vs.TextMatrix(I, 2)
  Else
     q = vs.TextMatrix(I, 2)
  End If
      
   ''con.Execute "insert into BookDiff(BOOKCODE,Godown,stockType,gp,balance) values ('" & vs.TextMatrix(I, 0) & "','" & cboBinder_Godown & "','CH','" & cboGroup.Text & "'," & q & ")"
      
End If

Next




End If


End Sub

Private Sub cmdOpening_Click()

Screen.MousePointer = vbHourglass

txtBookCode.text = "All"

uploadData datefrom

Screen.MousePointer = vbDefault

End Sub

Private Sub cmdOpening1_Click()

End Sub

Private Sub cmdPrint_7_Click()


 con.Execute "delete from tmpbook where login='" & UserName & "'"

 For I = 1 To vs.rows - 1
   If vs.TextMatrix(I, 0) <> "" Then
      con.Execute "insert into tmpbook(bcode,bname,qty,login,head,ason,issueQty,BalanceQty) values('" & vs.TextMatrix(I, 0) & "','" & vs.TextMatrix(I, 1) & "'," & _
      "'" & vs.TextMatrix(I, 2) & "','" & UserName & "','" & cboBinder_Godown & "','" & dateAson.value & "','" & vs.TextMatrix(I, 3) & "','" & vs.TextMatrix(I, 4) & "')"
   End If
 Next


DSNNew

DoEvents
DoEvents

'If MsgBox("Want To View ?", vbQuestion + vbYesNo) = vbYes Then
   CR.Reset
   CR.ReportFileName = rptPath & "/BookList.rpt"
   CR.Connect = "filedsn=chitradsn;uid=" & sql_user & ";pwd=" & sql_pass
   CR.ReplaceSelectionFormula "{bookmaster.login}='" & UserName & "'"
   CR.WindowShowPrintSetupBtn = True
   CR.WindowState = crptMaximized
   CR.Action = 1

'End If


End Sub
Sub fillGrid()

Dim Str As String
Dim rs_data As New ADODB.Recordset


vs.Cols = 3
vs.FormatString = "** BookCode ** |Book Name|*** Balance ***"
vs.rows = 2
vs.ColWidth(1) = 5000

Str = ""

If Str = "" Then
If cboGroup.text <> "All" And txtBookCode = "All" Then
   Str = "groupCODE = '" & cboGroup & "'"
ElseIf cboGroup <> "All" And txtBookCode <> "All" Then
   Str = "groupCODE = '" & cboGroup & "' and bookcode = '" & txtBookCode & "'"
ElseIf cboGroup = "All" And txtBookCode = "All" Then
   Str = ""
End If
End If


If rs_data.State = 1 Then rs_data.Close
If Str = "" Then
    If cboBinder_Godown.text = "All" Then
       rs_data.Open "select BOOKCODE,bookname from  Books where kitcode='n' and " & stringyear & "  group by BOOKCODE,bookname", con, adOpenKeyset
    Else
       rs_data.Open "select BOOKCODE,bookname from  Books where kitcode='n' and " & Str & " and " & stringyear & "  group by BOOKCODE,bookname", con, adOpenKeyset
    End If

Else
    If cboBinder_Godown.text = "All" Then
        rs_data.Open "select BOOKCODE,bookname from  Books where kitcode='n' and " & stringyear & "  group by BOOKCODE,bookname", con, adOpenKeyset
    Else
        rs_data.Open "select BOOKCODE,bookname from  Books where kitcode='n' and " & Str & " and " & stringyear & "  group by BOOKCODE,bookname", con, adOpenKeyset
    End If
End If


sum1 = 0
sum2 = 0


While rs_data.EOF = False


If RS.State = 1 Then RS.Close

str1 = ""
If RS.State = 1 Then RS.Close
If cboBinder_Godown.text = "All" Then
   str1 = "Godown<>'Z'"
Else
   str1 = "Godown='" & cboBinder_Godown.text & "'"
End If

godown_rec = 0
godown_issue = 0



If RS.State = 1 Then RS.Close
RS.Open "select sum(Quantity) from stocksummaryQry where convert(smalldatetime,Dates,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
" BookCode='" & rs_data.Fields("BOOKCODE").value & "' and " & _
" " & str1 & " and issue_Receive='Receive'", con, adOpenKeyset
If Not IsNull(RS(0)) Then godown_rec = RS(0)

If RS.State = 1 Then RS.Close
RS.Open "select sum(Quantity) from stocksummaryQry where convert(smalldatetime,Dates,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
" BookCode='" & rs_data.Fields("BOOKCODE").value & "' and " & _
" " & str1 & " and issue_Receive='Issue'", con, adOpenKeyset
If Not IsNull(RS(0)) Then godown_issue = RS(0)




'-----------------------------------------------------------------------------

If (godown_rec - godown_issue) <> 0 Then
K = K + 1
vs.rows = vs.rows + 1

vs.TextMatrix(K, 0) = rs_data!Bookcode
vs.TextMatrix(K, 1) = rs_data!Bookname
vs.TextMatrix(K, 2) = (godown_rec - godown_issue)

End If

sum1 = sum1 + godown_rec
sum2 = sum2 + IIf(IsNull(godown_issue), 0, godown_issue)
rs_data.MoveNext
Wend

vs.TextMatrix(K + 1, 1) = "                                                                   * Total "
vs.TextMatrix(K + 1, 2) = (sum1 - sum2)

vs.Cell(flexcpBackColor, K + 1, 0, K + 1, 2) = &HB0D8FF

Screen.MousePointer = vbDefault

End Sub
Private Sub cmdStock_Order_Click()


If cboBinder_Godown.text <> "" Then

Screen.MousePointer = vbHourglass

Dim sm1, sm2 As Long

sm1 = 0
sm2 = 0

vs.Cols = 5
Set rs1 = con.Execute("exec PendingOrder_Sale_Specimen_New1 '" & cboBinder_Godown.text & "','" & cboGroup.text & "'")

For k2 = 1 To vs.rows - 1
   
If (vs.TextMatrix(k2, 0) <> "") Then
    
    
        rs1.MoveFirst
        rs1.Find "bookcode='" & vs.TextMatrix(k2, 0) & "'"
        If rs1.EOF = False Then
           
           If rs1!balQty < 0 Then
           vs.TextMatrix(k2, 3) = 0
           Else
           vs.TextMatrix(k2, 3) = rs1!balQty
           End If
           
           
           vs.TextMatrix(k2, 4) = vs.TextMatrix(k2, 2) - Val(vs.TextMatrix(k2, 3))
         Else
           vs.TextMatrix(k2, 3) = 0
           vs.TextMatrix(k2, 4) = vs.TextMatrix(k2, 2) - vs.TextMatrix(k2, 3)
         End If
         
    
 
End If
 
Next


sm1 = 0
sm2 = 0
sm2 = 0

For k2 = 1 To vs.rows - 1
    
If (vs.TextMatrix(k2, 0) <> "") Then
   sm1 = sm1 + Val(vs.TextMatrix(k2, 2))
   sm2 = sm2 + Val(vs.TextMatrix(k2, 3))
   sm3 = sm3 + Val(vs.TextMatrix(k2, 4))
   
End If

Next

vs.TextMatrix(k2 - 1, 2) = sm1
vs.TextMatrix(k2 - 1, 3) = sm2
vs.TextMatrix(k2 - 1, 4) = sm3



vs.TextMatrix(0, 3) = "Pending Ord.Qty"
vs.TextMatrix(0, 4) = "Today Bal.Qty"

vs.ColWidth(2) = 2000
vs.ColWidth(3) = 2000
vs.ColWidth(4) = 2000


vs.Cell(flexcpBackColor, k2 - 1, 0, k2 - 1, 2) = &HB0D8FF
vs.Cell(flexcpBackColor, k2 - 1, 3, k2 - 1, 4) = &HB0D8FF


End If

Screen.MousePointer = vbDefault

Exit Sub



'=========================================================




Screen.MousePointer = vbHourglass
'=========================================================
'Noda=====================================================

ddd = 2
PopUpValue6 = 2

If cboBinder_Godown.text <> "NS" Then
   MsgBox "Only for NS ......", vbCritical
   Exit Sub
End If



Dim rs_kit As New ADODB.Recordset
Dim rs_data As New ADODB.Recordset
Dim K As Integer
K = 0

'-------------------------------
vs.Cols = 3
vs.FormatString = "** BookCode ** |Book Name|*** Balance ***"
vs.rows = 2
vs.ColWidth(1) = 5000


con.Execute "delete from tmpNSStock"


''Sale Details-----------

''con.Execute "insert into tmpNSStock(bookcode,Qty,[status],dates,status_) SELECT Distinct [Bookcode] ,(QUANTITY * -1) as Qty,'Out' as Issue_Rec,invoicedate,'Stock Out(By Invoice)' as status_ FROM invoiceBQry where (NsChallanNo is null and Godown ='NS')"

If Not Mid(session, 6) >= 19 Then
   con.Execute "insert into tmpNSStock(bookcode,Qty,[status],dates,status_) SELECT Distinct [Bookcode] ,(QUANTITY) as Qty,'In' as Issue_Rec,invoicedate,'Stock In(By Cr.NotItem' as status_ FROM CreditbQry where (NsChallanNo is null and Godown ='NS')"
   con.Execute "insert into tmpNSStock(bookcode,QtySp,[status],dates,status_) SELECT Distinct [Bookcode] ,Qty,'In' as Issue_Rec,invoicedate,'Stock In(By Sp.Ret)' as status_ FROM SpecimenReturnRegister where (NsChallanNo is null and Godown ='NS')"
End If


If server_ = "client" Then
   con.Execute "insert into tmpNSStock(bookcode,Qty,QtySp,status,dates,status_)  select BOOKCODE,sum(QUANTITY) as QUANTITY,sum(SpQty) as SpQty,'IN' as Issue_Rec,invoiceDate,'Stock IN By Challan' FROM CHALANB_ret group by BOOKCODE,invoiceDate"
   con.Execute "update tmpNSStock set Qty=0 where Qty is null"
   con.Execute "update tmpNSStock set QtySp=0 where QtySp is null"

Else
   con.Execute "insert into tmpNSStock(bookcode,Qty,QtySp,status,dates,status_)  SELECT [BOOKCODE],sum(QUANTITY),sum(SpQty),'IN' as Issue_Rec,invoiceDate,'Stock IN By Challan' FROM CHALANB_ret  group by BOOKCODE,invoiceDate"
   con.Execute "update tmpNSStock set Qty=0 where Qty is null"
   con.Execute "update tmpNSStock set QtySp=0 where QtySp is null"
End If


con.Execute "insert into tmpNSStock(bookcode,Qty,status,dates,status_)  SELECT [BOOKCODE],(sum(Qty)*-1) as Qty,'Out' as Issue_Rec,dates,'BookStock Out' FROM BookStock where (Godown_out ='NS') group by BOOKCODE,dates"
con.Execute "insert into tmpNSStock(bookcode,Qty,status,dates,status_)  SELECT [BOOKCODE],sum(Qty) as Qty,'IN' as Issue_Rec,dates,'BookStock IN' FROM BookStock where (Godown_in ='NS') group by BOOKCODE,dates"
con.Execute "insert into tmpNSStock(bookcode,Qty,QtySp,status,dates,status_)  SELECT [BOOKCODE],(sum(QUANTITY)*-1) as Qty,(sum(SpQty)*-1) as SpQty,'Out' as Issue_Rec,INVOICEDATE,'Stock Out By Order' FROM OrderBookList  where Godown ='NS' group by BOOKCODE,invoiceDate"
con.Execute "insert into tmpNSStock(bookcode,Qty,status,dates,status_)  SELECT  BOOKCODE,Balance,'Out' as Issue_Rec,'" & fromDate_setup & "','Opening Qty in (-)' FROM BookOpening where (Godown ='NS' and Balance<0)"
con.Execute "insert into tmpNSStock(bookcode,Qty,status,dates,status_)  SELECT  BOOKCODE,Balance,'IN' as Issue_Rec,'" & fromDate_setup & "','Opening Qty in (+)' FROM BookOpening where (Godown ='NS' and Balance>0)"

'---------------------------------------------------------------------

DoEvents
DoEvents
DoEvents

If rs_kit.State = 1 Then rs_kit.Close
rs_kit.Open "select  BOOKNAME,bookcode  from BOOKS", con, adOpenDynamic, adLockOptimistic
DoEvents
DoEvents

If rs_data.State = 1 Then rs_data.Close
rs_data.Open "select  tmpNSStock.BookCode,sum(Qty),sum(QtySP),BOOKS.GROUPCODE from tmpNSStock  tmpNSStock INNER JOIN books ON tmpNSStock.BookCode = BOOKS.BOOKCODE " & _
 " where BOOKS.GROUPCODE='" & cboGroup.text & "' and convert(smalldatetime,Dates,103)<= convert(smalldatetime,'" & dateAson.value & "',103) group by tmpNSStock.BookCode,books.groupcode order by tmpNSStock.BookCode", con, adOpenDynamic, adLockOptimistic
While rs_data.EOF = False

DoEvents
DoEvents

K = K + 1
vs.rows = vs.rows + 1

vs.TextMatrix(K, 0) = rs_data!Bookcode

rs_kit.MoveFirst
rs_kit.Find "bookcode='" & rs_data!Bookcode & "'"
If rs_kit.EOF = False Then
   vs.TextMatrix(K, 1) = rs_kit(0)
End If


vs.TextMatrix(K, 2) = (IIf(IsNull(rs_data(1)), 0, rs_data(1)) + IIf(IsNull(rs_data(2)), 0, rs_data(2)))


'------------------------------------------
'sum1 = sum1 + Binder_Rec
'sum2 = sum2 + Binder_Issue
'------------------------------------------


DoEvents
DoEvents

rs_data.MoveNext
Wend

Screen.MousePointer = vbDefault


'==========================================================
'==========================================================



End Sub

Private Sub cmdView_Click()

Screen.MousePointer = vbHourglass

vs.Clear

bb_ = False
Dim sum1, sum2 As Long
Dim Opening, Binder_Rec, godown_rec, SalesReturn, Spec_Return, TotalRec As Long
Dim Binder_Issue, godown_issue, sales, Spec_Issue, Damage, TotalIssue As Long

Binder_Rec = 0: godown_rec = 0: SalesReturn = 0: Spec_Return = 0: TotalRec = 0
Binder_Issue = 0: godown_issue = 0: sales = 0: Spec_Issue = 0: TotalIssue = 0
Opening = 0
Damage = 0
K = 0

Dim d1 As Double
d1 = 0

Dim Str, str_go, str_go_issue As String
Str = ""
str_go = ""
str_go_issue = ""

ddd = 1
PopUpValue6 = 1

Dim rs_kit As New ADODB.Recordset
Dim rs_data As New ADODB.Recordset
Dim rs_issue As New ADODB.Recordset
Dim rs_cash As New ADODB.Recordset
Dim newoldbk As String
Dim Qty1 As Integer
Qty1 = 0

If cboCategory.text = "" Then
   MsgBox "Select Binder/Godwon ..", vbCritical
   Screen.MousePointer = vbDefault
   Exit Sub
End If

'=========================================================
'Noda=====================================================

If Check1_crm.value = 1 Then

'-------------------------------
vs.Cols = 3
vs.FormatString = "** BookCode ** |Book Name|*** Balance ***"
vs.rows = 2
vs.ColWidth(1) = 5000


con.Execute "delete from tmpNSStock"


''Sale Details-----------



If Not Mid(session, 6) >= 19 Then
    con.Execute "insert into tmpNSStock(bookcode,Qty,[status],dates,status_) SELECT Distinct [Bookcode] ,(QUANTITY * -1) as Qty,'Out' as Issue_Rec,invoicedate,'Stock Out(By Invoice)' as status_ FROM invoiceBQry where (NsChallanNo is null and Godown ='NS')"
    con.Execute "insert into tmpNSStock(bookcode,Qty,[status],dates,status_) SELECT Distinct [Bookcode] ,(QUANTITY) as Qty,'In' as Issue_Rec,invoicedate,'Stock In(By Cr.NotItem' as status_ FROM CreditbQry where (NsChallanNo is null and Godown ='NS')"
    con.Execute "insert into tmpNSStock(bookcode,QtySp,[status],dates,status_) SELECT Distinct [Bookcode] ,Qty,'In' as Issue_Rec,invoicedate,'Stock In(By Sp.Ret)' as status_ FROM SpecimenReturnRegister where (NsChallanNo is null and Godown ='NS')"
    con.Execute "insert into tmpNSStock(bookcode,QtySp,[status],dates,status_) SELECT Distinct [Bookcode] ,Qty*-1 as Qty,'Out' as Issue_Rec,invoicedate,'Stock Out(By Sp.)' as status_ FROM SpecimenRegister where (NsChallanNo is null and Godown ='NS')"
    'con.Execute "insert into tmpNSStock(bookcode,Qty,[status],dates,status_) SELECT Distinct [Bookcode] ,(QUANTITY * -1) as Qty,'Out' as Issue_Rec,invoicedate,'Stock Out(By CashSale)' as status_ FROM cashBQry where (Godown ='NS')"
End If


con.Execute "insert into tmpNSStock(bookcode,Qty,status,dates,status_)  SELECT distinct [BOOKCODE],[Qty]*-1,'Out' as Issue_Rec,dates,'BookStock Out' FROM BookStock where (Godown_out ='NS')"
con.Execute "insert into tmpNSStock(bookcode,Qty,status,dates,status_)  SELECT distinct [BOOKCODE],[Qty],'IN' as Issue_Rec,dates,'BookStock IN' FROM BookStock where (Godown_in ='NS')"


If server_ = "client" Then
   'con.Execute "insert into tmpNSStock(bookcode,Qty,QtySp,status,dates,status_)  SELECT distinct [BOOKCODE],QUANTITY,SpQty,'IN' as Issue_Rec,invoiceDate,'Stock IN By Challan' FROM CHALANB_ret"
   'con.Execute "insert into tmpNSStock(bookcode,Qty,QtySp,status,dates,status_)  SELECT distinct [BOOKCODE],QUANTITY*-1,SpQty*-1,'Out' as Issue_Rec,INVOICEDATE,'Stock Out By Challan' FROM CHALANB"
   con.Execute "insert into tmpNSStock(bookcode,Qty,QtySp,status,dates,status_) SELECT  CHALANB_ret.BOOKCODE,sum(CHALANB_ret.QUANTITY) as QUANTITY,sum(CHALANB_ret.SpQty) as SpQty,'IN' as Issue_Rec,CHALANA_ret.invoiceDate,'Stock IN By Challan' FROM  CHALANA_ret INNER JOIN  CHALANB_ret ON CHALANA_ret.INVOICENO = CHALANB_ret.INVOICENO group by CHALANB_ret.BOOKCODE,CHALANA_ret.invoiceDate"
   con.Execute "insert into tmpNSStock(bookcode,Qty,QtySp,status,dates,status_) SELECT  CHALANB.BOOKCODE,sum(CHALANB.QUANTITY*-1),sum(CHALANB.SpQty*-1),'Out' as Issue_Rec,CHALANA.invoiceDate,'Stock Out By Challan' FROM  CHALANA INNER JOIN  CHALANB ON CHALANA.INVOICENO = CHALANB.INVOICENO group by CHALANB.BOOKCODE,CHALANA.invoiceDate"


Else
   'con.Execute "insert into tmpNSStock(bookcode,Qty,QtySp,status,dates,status_)  SELECT distinct [BOOKCODE],iif(QUANTITY is null,0,QUANTITY) as QUANTITY,iif(SpQty is null,0,SpQty) as SpQty,'IN' as Issue_Rec,invoiceDate,'Stock IN By Challan' FROM CHALANB_ret"
   con.Execute "insert into tmpNSStock(bookcode,Qty,QtySp,status,dates,status_) SELECT  CHALANB_ret.BOOKCODE,iif(sum(CHALANB_ret.QUANTITY) is null,0,sum(CHALANB_ret.QUANTITY)) as QUANTITY,iif(sum(CHALANB_ret.SpQty) is null,0,sum(CHALANB_ret.SpQty)) as SpQty,'IN' as Issue_Rec,CHALANA_ret.invoiceDate,'Stock IN By Challan' FROM  CHALANA_ret INNER JOIN  CHALANB_ret ON CHALANA_ret.INVOICENO = CHALANB_ret.INVOICENO group by CHALANB_ret.BOOKCODE,CHALANA_ret.invoiceDate"
   'con.Execute "insert into tmpNSStock(bookcode,Qty,QtySp,status,dates,status_) SELECT distinct [BOOKCODE],iif(QUANTITY is null,0,QUANTITY*-1) as QUANTITY,iif(SpQty is null,0,SpQty*-1) as SpQty,'Out' as Issue_Rec,INVOICEDATE,'Stock Out By Challan' FROM CHALANB"
   con.Execute "insert into tmpNSStock(bookcode,Qty,QtySp,status,dates,status_) SELECT  CHALANB.BOOKCODE,iif(sum(CHALANB.QUANTITY) is null,0,sum(CHALANB.QUANTITY*-1)) as QUANTITY,iif(sum(CHALANB.SpQty) is null,0,sum(CHALANB.SpQty*-1)) as SpQty,'Out' as Issue_Rec,CHALANA.invoiceDate,'Stock Out By Challan' FROM  CHALANA INNER JOIN  CHALANB ON CHALANA.INVOICENO = CHALANB.INVOICENO group by CHALANB.BOOKCODE,CHALANA.invoiceDate"
   
End If

con.Execute "insert into tmpNSStock(bookcode,Qty,status,dates,status_)  SELECT distinct BOOKCODE,Balance,'Out' as Issue_Rec,'" & fromDate_setup & "','Opening Qty in (-)' FROM BookOpening where (Godown ='NS' and Balance<0)"
con.Execute "insert into tmpNSStock(bookcode,Qty,status,dates,status_)  SELECT distinct BOOKCODE,Balance,'IN' as Issue_Rec,'" & fromDate_setup & "','Opening Qty in (+)' FROM BookOpening where (Godown ='NS' and Balance>0)"

'con.Execute "insert into tmpNSStock(bookcode,Qty,status,dates,status_)  SELECT distinct BOOKCODE,abs(Balance),'In' as Issue_Rec,'" & fromDate_setup & "','Diff Qty in (+)' FROM Bookdiff where (Godown ='NS' and Balance<0 and stockType='CH')"
'con.Execute "insert into tmpNSStock(bookcode,Qty,status,dates,status_)  SELECT distinct BOOKCODE,(Balance*-1) as bal_,'Out' as Issue_Rec,'" & fromDate_setup & "','Diff Qty in (-)' FROM Bookdiff where (Godown ='NS' and Balance>0 and stockType='CH')"


'------------------------------
DoEvents
DoEvents
DoEvents

If rs_kit.State = 1 Then rs_kit.Close
rs_kit.Open "select  BOOKNAME,bookcode  from BOOKS", con, adOpenDynamic, adLockOptimistic
DoEvents
DoEvents

'MsgBox "" & Str

If rs_data.State = 1 Then rs_data.Close

If cboserName.text = "" Then
    rs_data.Open "select  tmpNSStock.BookCode,sum(Qty),sum(QtySP),BOOKS.GROUPCODE from tmpNSStock  tmpNSStock INNER JOIN books ON tmpNSStock.BookCode = BOOKS.BOOKCODE " & _
    " where BOOKS.GROUPCODE='" & cboGroup.text & "' and convert(smalldatetime,Dates,103)<= convert(smalldatetime,'" & dateAson.value & "',103) group by tmpNSStock.BookCode,books.groupcode order by tmpNSStock.BookCode", con, adOpenDynamic, adLockOptimistic
ElseIf cboserName.text <> "" Then
    rs_data.Open "select  tmpNSStock.BookCode,sum(Qty),sum(QtySP),BOOKS.GROUPCODE from tmpNSStock  tmpNSStock INNER JOIN books ON tmpNSStock.BookCode = BOOKS.BOOKCODE " & _
    " where (BOOKS.serName='" & cboserName.text & "' and BOOKS.GROUPCODE='" & cboGroup.text & "' and convert(smalldatetime,Dates,103)<= convert(smalldatetime,'" & dateAson.value & "',103)) group by tmpNSStock.BookCode,books.groupcode order by tmpNSStock.BookCode", con, adOpenDynamic, adLockOptimistic

End If

While rs_data.EOF = False

DoEvents
DoEvents

K = K + 1
vs.rows = vs.rows + 1

vs.TextMatrix(K, 0) = rs_data!Bookcode

rs_kit.MoveFirst
rs_kit.Find "bookcode='" & rs_data!Bookcode & "'"
If rs_kit.EOF = False Then
   vs.TextMatrix(K, 1) = rs_kit(0)
End If


vs.TextMatrix(K, 2) = (IIf(IsNull(rs_data(1)), 0, rs_data(1)) + IIf(IsNull(rs_data(2)), 0, rs_data(2)))



'------------------------------------------
'sum1 = sum1 + Binder_Rec
'sum2 = sum2 + Binder_Issue
'------------------------------------------


DoEvents
DoEvents

rs_data.MoveNext
Wend

Screen.MousePointer = vbDefault

Exit Sub
End If

'==========================================================
'==========================================================

If Option1_new.value = True Then
   newoldbk = "NEW"
Else
   newoldbk = "OLD"
End If


If Str = "" Then
If cboGroup.text <> "All" And txtBookCode = "All" Then
   Str = "groupCODE = '" & cboGroup & "'"
ElseIf cboGroup <> "All" And txtBookCode <> "All" Then
   Str = "groupCODE = '" & cboGroup & "' and bookcode = '" & txtBookCode & "'"
ElseIf cboGroup = "All" And txtBookCode = "All" Then
   Str = ""
End If
End If

If cboserName.text <> "" Then
   Str = Str & " and sername = '" & cboserName & "'"
End If


If cboCategory.ItemData(cboCategory.ListIndex) = 1 Then    'Binder


vs.Cols = 5
vs.FormatString = "** BookCode ** |Book Name|Receive From Binder|Issue To Binder|*** Balance ***"
vs.rows = 2
vs.ColWidth(1) = 5000
'----------------------------------------------------------------------------------

If rs_data.State = 1 Then rs_data.Close
If Str = "" Then
     rs_data.Open "select BOOKCODE,bookname from  Books where kitcode='n' and " & stringyear & "  group by BOOKCODE,bookname", con, adOpenKeyset
Else
     rs_data.Open "select BOOKCODE,bookname from  Books where kitcode='n' and " & Str & " and " & stringyear & "  group by BOOKCODE,bookname", con, adOpenKeyset
End If


sum1 = 0
sum2 = 0

While rs_data.EOF = False


If RS.State = 1 Then RS.Close

Binder_Rec = 0
Binder_Issue = 0

If RS.State = 1 Then RS.Close
RS.Open "select sum(Qty) from BinderReceiveRegister where convert(smalldatetime,invoiceDate,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
" Book_Code='" & rs_data.Fields("BOOKCODE").value & "' and " & _
" subledger='" & cboBinder_Godown.text & "' and " & stringyear & "", con, adOpenKeyset
If Not IsNull(RS(0)) Then Binder_Rec = RS(0)
If Not IsNull(RS(0)) Then
   Binder_Rec = RS(0)
Else
   Binder_Rec = 0
End If


'======================================================================================
If RS.State = 1 Then RS.Close
RS.Open "select sum(CONVERT(INT, Qty)) from BookOrderPrint_Qry where convert(smalldatetime,Ord_Date,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
" BCode='" & rs_data.Fields("BOOKCODE").value & "' and " & _
" binder='" & cboBinder_Godown.text & "' and " & stringyear & " ", con, adOpenKeyset
If Not IsNull(RS(0)) Then Binder_Issue = RS(0)


'-------------------------------------------------------------------------------------

If RS.State = 1 Then RS.Close
RS.Open "select sum(netbook) from BinderIssueRegister where convert(smalldatetime,invoiceDate,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
" Book_Code='" & rs_data.Fields("BOOKCODE").value & "' and " & _
" subledger='" & cboBinder_Godown.text & "' and " & stringyear & "", con, adOpenKeyset
'If Not IsNull(RS(0)) Then Binder_Issue = Binder_Issue + RS(0)
If RS.EOF = False Then Binder_Issue = Binder_Issue + RS(0)


If (Binder_Rec - Binder_Issue) <> 0 Then

K = K + 1
vs.rows = vs.rows + 1

vs.TextMatrix(K, 0) = rs_data!Bookcode
vs.TextMatrix(K, 1) = rs_data!Bookname
vs.TextMatrix(K, 2) = Binder_Rec
vs.TextMatrix(K, 3) = Binder_Issue
vs.TextMatrix(K, 4) = (Binder_Rec - Binder_Issue)


End If

sum1 = sum1 + Binder_Rec
sum2 = sum2 + Binder_Issue

rs_data.MoveNext
Wend

vs.TextMatrix(K + 1, 1) = "                                                                   * Total "
vs.TextMatrix(K + 1, 2) = sum1
vs.TextMatrix(K + 1, 3) = sum2
vs.TextMatrix(K + 1, 4) = (sum1 - sum2)

vs.Cell(flexcpBackColor, K + 1, 0, K + 1, 4) = &HB0D8FF

'*******************************************************************************************************************
'-------------------------------------------------- Godown Balance List---------------------------------------------
'*******************************************************************************************************************
'*******************************************************************************************************************
'-------------------------------------------------- Godown Balance List---------------------------------------------
'*******************************************************************************************************************


ElseIf cboCategory.ItemData(cboCategory.ListIndex) = 2 Then     ' Godown Balance List



vs.Cols = 3
vs.FormatString = "** BookCode ** |Book Name|*** Balance ***"
vs.rows = 2
vs.ColWidth(1) = 5000



If rs_data.State = 1 Then rs_data.Close
If Str = "" Then
    If cboBinder_Godown.text = "All" Then
       rs_data.Open "select BOOKCODE,bookname from  Books where kitcode='n' and " & stringyear & "  group by BOOKCODE,bookname", con, adOpenKeyset
    Else
       rs_data.Open "select BOOKCODE,bookname from  Books where kitcode='n' and " & Str & " and " & stringyear & "  group by BOOKCODE,bookname", con, adOpenKeyset
    End If

Else
    If cboBinder_Godown.text = "All" Then
        rs_data.Open "select BOOKCODE,bookname from  Books where kitcode='n' and " & stringyear & "  group by BOOKCODE,bookname", con, adOpenKeyset
    Else
    
    If Check1_sp.value = 1 Then
        rs_data.Open "select BOOKCODE,bookname from  Books where " & Str & " and " & stringyear & "  group by BOOKCODE,bookname", con, adOpenKeyset
    Else
        rs_data.Open "select BOOKCODE,bookname from  Books where kitcode='n' and " & Str & " and " & stringyear & "  group by BOOKCODE,bookname", con, adOpenKeyset
    End If
    
    End If
End If


sum1 = 0
sum2 = 0



While rs_data.EOF = False


If RS.State = 1 Then RS.Close

str1 = ""
If RS.State = 1 Then RS.Close
If LCase(cboBinder_Godown.text) = "all" Then
   str1 = "Godown<>'Z'"
   str_go = "Godown_In<>'Z'"
   str_go_issue = "Godown_Out<>'Z'"
Else
   str1 = "Godown='" & cboBinder_Godown.text & "'"
   str_go = "Godown_In='" & cboBinder_Godown.text & "'"
   str_go_issue = "Godown_Out='" & cboBinder_Godown.text & "'"
End If

godown_rec = 0
godown_issue = 0





If RS.State = 1 Then RS.Close
RS.Open "select sum(QUANTITY) from Salestbl where convert(smalldatetime,invoicedate,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
" BookCode='" & rs_data.Fields("BOOKCODE").value & "' and catagory='saleret' and " & _
" " & str1 & "", con, adOpenKeyset
If Not IsNull(RS(0)) Then godown_rec = RS(0)

'for Kit
If RS.State = 1 Then RS.Close
RS.Open "SELECT sum(a.[QUANTITY]) FROM Salestbl as a inner join BOOKS_KIT as b on (a.BOOKCODE =b.KITCODE) where convert(smalldatetime,invoicedate,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
" b.BookCode='" & rs_data.Fields("BOOKCODE").value & "' and a.catagory='saleret' and " & _
" " & str1 & "", con, adOpenKeyset
If Not IsNull(RS(0)) Then godown_rec = RS(0)




'If RS.State = 1 Then RS.close
'RS.Open "select sum(QUANTITY) from CashBQry_estimateRet where convert(smalldatetime,invoicedate,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
'" BookCode='" & rs_data.Fields("BOOKCODE").value & "' and " & _
'" " & str1 & " and " & stringyear & "", con, adOpenKeyset
'If Not IsNull(RS(0)) Then godown_rec = godown_rec + RS(0)

'For Kit

'If RS.State = 1 Then RS.close
'RS.Open "SELECT sum(a.[QUANTITY]) FROM cashBQry_estimateRet as a inner join BOOKS_KIT as b on (a.BOOKCODE =b.KITCODE) where convert(smalldatetime,invoicedate,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
'" b.BookCode='" & rs_data.Fields("BOOKCODE").value & "' and " & _
'" " & str1 & "", con, adOpenKeyset
'If Not IsNull(RS(0)) Then godown_rec = RS(0)



If RS.State = 1 Then RS.Close
RS.Open "select sum(Qty) from SaleReturnRegister where convert(smalldatetime,invoiceDate,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
" BookCode='" & rs_data.Fields("BOOKCODE").value & "' and " & _
" " & str1 & " and " & stringyear & "", con, adOpenKeyset
If Not IsNull(RS(0)) Then godown_rec = godown_rec + RS(0)


'For Kit
If RS.State = 1 Then RS.Close
RS.Open "SELECT sum(a.[QTY]) FROM SaleReturnRegister as a inner join BOOKS_KIT as b on (a.BOOKCODE =b.KITCODE) where convert(smalldatetime,invoicedate,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
" b.BookCode='" & rs_data.Fields("BOOKCODE").value & "' and " & _
" " & str1 & "", con, adOpenKeyset
If Not IsNull(RS(0)) Then godown_rec = RS(0)



'''If RS.State = 1 Then RS.close
'''RS.Open "select sum(Qty) from SaleReturnRegister_Free where convert(smalldatetime,invoiceDate,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
'''" BookCode='" & rs_data.Fields("BOOKCODE").value & "' and " & _
'''" " & str1 & " and " & stringyear & "", con, adOpenKeyset
'''If Not IsNull(RS(0)) Then godown_rec = godown_rec + RS(0)
''

'-----------------------------------------------------------------------------

If RS.State = 1 Then RS.Close
RS.Open "select sum(Qty) from SpecimenReturnRegister where convert(smalldatetime,invoiceDate,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
" BookCode='" & rs_data.Fields("BOOKCODE").value & "' and " & _
" " & str1 & " and " & stringyear & "", con, adOpenKeyset
If Not IsNull(RS(0)) Then godown_rec = godown_rec + RS(0)


'For Kit
If RS.State = 1 Then RS.Close
RS.Open "SELECT sum(a.[QTY]) FROM SpecimenReturnRegister as a inner join BOOKS_KIT as b on (a.BOOKCODE =b.KITCODE) where convert(smalldatetime,invoicedate,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
" b.BookCode='" & rs_data.Fields("BOOKCODE").value & "' and " & _
" " & str1 & "", con, adOpenKeyset
If Not IsNull(RS(0)) Then godown_rec = godown_rec + RS(0)



If RS.State = 1 Then RS.Close
RS.Open "select sum(Qty) from BinderReceiveRegister where convert(smalldatetime,invoiceDate,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
" Book_Code='" & rs_data.Fields("BOOKCODE").value & "' and " & _
" " & str1 & " and " & stringyear & "", con, adOpenKeyset
If Not IsNull(RS(0)) Then godown_rec = godown_rec + RS(0)



'''Stock Transfer
'''Receive Qty

If RS.State = 1 Then RS.Close
RS.Open "select sum(Qty) from BookStock where convert(smalldatetime,Dates,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
" BOOKCODE='" & rs_data.Fields("BOOKCODE").value & "' and " & _
" " & str_go & " and " & stringyear & "", con, adOpenKeyset
If Not IsNull(RS(0)) Then godown_rec = godown_rec + RS(0)



If RS.State = 1 Then RS.Close
RS.Open "select sum(Balance) from BookOpening where BOOKCODE='" & rs_data.Fields("BOOKCODE").value & "' and Godown='" & cboBinder_Godown & "'", con, adOpenKeyset
If Not IsNull(RS(0)) Then godown_rec = godown_rec + RS(0)


'=================================================Issue Code Start=====================================


str1 = ""
If RS.State = 1 Then RS.Close
If cboBinder_Godown.text = "All" Then
   str1 = "Godown<>'Z'"
Else
   str1 = "Godown='" & cboBinder_Godown.text & "'"
End If


If RS.State = 1 Then RS.Close
RS.Open "select sum(Qty) from SaleRegister where convert(smalldatetime,invoiceDate,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
" BookCode='" & rs_data.Fields("BOOKCODE").value & "' and " & _
" " & str1 & " and " & stringyear & "", con, adOpenKeyset
If Not IsNull(RS(0)) Then
   godown_issue = RS(0)
Else
   godown_issue = 0
End If



'For Kit

If RS.State = 1 Then RS.Close
RS.Open "SELECT sum(a.[QTY]) FROM SaleRegister as a inner join BOOKS_KIT as b on (a.BOOKCODE =b.KITCODE) where convert(smalldatetime,invoicedate,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
" b.BookCode='" & rs_data.Fields("BOOKCODE").value & "' and " & _
" " & str1 & "", con, adOpenKeyset
If Not IsNull(RS(0)) Then godown_issue = godown_issue + RS(0)




If rs_cash.State = 1 Then rs_cash.Close
rs_cash.Open "select sum(Qty) from CashSaleRegister where convert(smalldatetime,invoicedate,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
" BookCode='" & rs_data.Fields("BOOKCODE").value & "' and " & _
" " & str1 & " and " & stringyear & "", con, adOpenKeyset
If Not IsNull(rs_cash(0)) Then godown_issue = godown_issue + rs_cash(0)


'For Kit

If RS.State = 1 Then RS.Close
RS.Open "SELECT sum(a.[QTY]) FROM CashSaleRegister as a inner join BOOKS_KIT as b on (a.BOOKCODE =b.KITCODE) where convert(smalldatetime,invoicedate,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
" b.BookCode='" & rs_data.Fields("BOOKCODE").value & "' and " & _
" " & str1 & "", con, adOpenKeyset
If Not IsNull(RS(0)) Then godown_issue = godown_issue + RS(0)




If rs_cash.State = 1 Then rs_cash.Close
rs_cash.Open "select sum(QUANTITY) from Salestbl where convert(smalldatetime,invoicedate,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
" BookCode='" & rs_data.Fields("BOOKCODE").value & "' and catagory='sale' and " & _
" " & str1 & "", con, adOpenKeyset
If Not IsNull(rs_cash(0)) Then godown_issue = godown_issue + rs_cash(0)


'For Kit

If RS.State = 1 Then RS.Close
RS.Open "SELECT sum(a.QUANTITY) FROM Salestbl as a inner join BOOKS_KIT as b on (a.BOOKCODE =b.KITCODE) where convert(smalldatetime,invoicedate,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
" b.BookCode='" & rs_data.Fields("BOOKCODE").value & "' and a.catagory='sale' and " & _
" " & str1 & "", con, adOpenKeyset
If Not IsNull(RS(0)) Then godown_issue = godown_issue + RS(0)





'If rs_cash.State = 1 Then rs_cash.close
'rs_cash.Open "select sum(Qty) from CashSaleRegister_basil where convert(smalldatetime,invoicedate,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
'" BookCode='" & rs_data.Fields("BOOKCODE").value & "' and " & _
'" " & str1 & " and " & stringyear & "", con, adOpenKeyset
'If Not IsNull(rs_cash(0)) Then godown_issue = godown_issue + rs_cash(0)



'For Kit

'If RS.State = 1 Then RS.close
'RS.Open "SELECT sum(a.[QTY]) FROM CashSaleRegister_basil as a inner join BOOKS_KIT as b on (a.BOOKCODE =b.KITCODE) where convert(smalldatetime,invoicedate,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
'" b.BookCode='" & rs_data.Fields("BOOKCODE").value & "' and " & _
'" " & str1 & "", con, adOpenKeyset
'If Not IsNull(RS(0)) Then godown_issue = godown_issue + RS(0)




If rs_issue.State = 1 Then rs_issue.Close
rs_issue.Open "select sum(Qty) from SpecimenRegister where convert(smalldatetime,invoicedate,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
" BookCode='" & rs_data.Fields("BOOKCODE").value & "' and " & _
" " & str1 & " and " & stringyear & "", con, adOpenKeyset
If Not IsNull(rs_issue(0)) Then godown_issue = godown_issue + rs_issue(0)



'For Kit

If RS.State = 1 Then RS.Close
RS.Open "SELECT sum(a.[QTY]) FROM SpecimenRegister as a inner join BOOKS_KIT as b on (a.BOOKCODE =b.KITCODE) where convert(smalldatetime,invoicedate,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
" b.BookCode='" & rs_data.Fields("BOOKCODE").value & "' and " & _
" " & str1 & "", con, adOpenKeyset
If Not IsNull(RS(0)) Then godown_issue = godown_issue + RS(0)




If RS.State = 1 Then RS.Close
RS.Open "select netbook from BinderIssueRegister where convert(smalldatetime,invoiceDate,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
" Book_Code='" & rs_data.Fields("BOOKCODE").value & "' and " & _
" " & str1 & " and " & stringyear & "", con, adOpenKeyset
While RS.EOF = False
    godown_issue = godown_issue + RS(0)
RS.MoveNext
Wend



If RS.State = 1 Then RS.Close
RS.Open "select sum(Qty) from BookStock where convert(smalldatetime,Dates,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
" BOOKCODE='" & rs_data.Fields("BOOKCODE").value & "' and " & _
" " & str_go_issue & " and " & stringyear & "", con, adOpenKeyset
If Not IsNull(RS(0)) Then godown_issue = godown_issue + RS(0)

d1 = IIf(IsNull(godown_rec), 0, godown_rec) - IIf(IsNull(godown_issue), 0, godown_issue)

If (d1) <> 0 Then

K = K + 1
vs.rows = vs.rows + 1

vs.TextMatrix(K, 0) = rs_data!Bookcode
vs.TextMatrix(K, 1) = rs_data!Bookname
vs.TextMatrix(K, 2) = (godown_rec - godown_issue)

End If

sum1 = sum1 + godown_rec
sum2 = sum2 + IIf(IsNull(godown_issue), 0, godown_issue)

rs_data.MoveNext
Wend

vs.TextMatrix(K + 1, 1) = "                                                                   * Total "
vs.TextMatrix(K + 1, 2) = (sum1 - sum2)

vs.Cell(flexcpBackColor, K + 1, 0, K + 1, 2) = &HB0D8FF

'===================================================================================================================
'-------------------------------------------------All Over All Books Balance----------------------------------------
'===================================================================================================================

ElseIf cboCategory.ItemData(cboCategory.ListIndex) = 3 Then     ' All Balance

'-------------------------------------------------- End Code -------------------------------------------------------
Else
'==========================Godown Stock=====================================================
End If

Screen.MousePointer = vbDefault

End Sub
Private Sub Command1_Click()
If cboBinder_Godown.text = "NS" Then

con.Execute "delete from BookDiff where gp='" & cboGroup.text & "' and godown='NS' and stockType='ORD'"

If MsgBox("want to Closing ? ", vbQuestion + vbYesNo) = vbNo Then
   Exit Sub
End If



'If RS.State = 1 Then RS.close
'RS.Open "SELECT [BOOKCODE],[Balance],[Godown],[stockType],[gp] from BookDiff where gp='" & cboGroup.Text & "' and godown='NS'", con, adOpenDynamic, adLockOptimistic

For I = 1 To vs.rows - 1

If vs.TextMatrix(I, 1) <> "" Then
      
  If Val(vs.TextMatrix(I, 2)) > 0 Then
      q = vs.TextMatrix(I, 2)
  Else
     q = vs.TextMatrix(I, 2)
  End If
      
   'con.Execute "insert into BookDiff(BOOKCODE,Godown,stockType,gp,balance) values ('" & vs.TextMatrix(I, 0) & "','" & cboBinder_Godown & "','ORD','" & cboGroup.Text & "'," & q & ")"
      
End If

Next




End If


End Sub

Private Sub Command1_excel_Click()

Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim I As Integer
Dim xl As Excel.Application
Dim rs_1 As New ADODB.Recordset
Dim rs_2 As New ADODB.Recordset
Dim str_ As String
Dim soldTillDate As Long
Dim from_date, last_date As Date
Dim db As String
Dim conLastData As ADODB.Connection


On Error GoTo err:


If xl Is Nothing Then
   Set xl = CreateObject("Excel.Application")
End If

xl.Visible = True
Set xlBook = xl.Workbooks.Add
Set xlSheet = xlBook.Worksheets.Add

Dim c, r, r1 As Long
Dim Q1, q2, J As Integer

c = 1
r = 1

r1 = 1


xl.Columns("A:H").ColumnWidth = 8
J = 2



If ddd = 1 Then

        '---------------------------------------------------------------------------
        
        If RS.State = 1 Then RS.Close
        RS.Open "select fromDate,toDate from saleDate where setupid=1", CCON
        If RS.EOF = False Then
              fromDate1.value = RS!fromdate
              toDate1.value = RS!todate
        End If
        
        If RS.State = 1 Then RS.Close
        RS.Open "select fromDate,toDate from saleDate where setupid=2", CCON
        If RS.EOF = False Then
              fromDate2.value = RS!fromdate
              toDate2.value = RS!todate
        End If
        
        
        con.Execute "exec netsale '" & fromDate1.value & "','" & toDate1.value & "','" & fromDate2.value & "','" & toDate2.value & "'"
        
        '---------------------------------------------------------------------------
        
        
        
        xlSheet.Cells(1, 1).value = "STOCK SUMMARY "
           
        
        
        For r = 1 To vs.rows - 1
               
        If vs.TextMatrix(r, 0) <> "" Then
               
               r1 = r1 + 1
               
                xlSheet.Cells(1, 1).value = "Book Code"
               xlSheet.Cells(1, 2).value = "Book Name"
               xlSheet.Cells(1, 3).value = "Stock As On "
               
               xlSheet.Cells(1, 4).value = "Pending Ord.Qty "
               xlSheet.Cells(1, 5).value = "Total Bal.Qty"
               
               xlSheet.Cells(1, 6).value = "Sold Last Year"
               xlSheet.Cells(1, 7).value = "Sold Current Year"
               
               
               xlSheet.Cells(r1, 1).value = vs.TextMatrix(r, 0)
               xlSheet.Cells(r1, 2).value = vs.TextMatrix(r, 1)
               xlSheet.Cells(r1, 3).value = vs.TextMatrix(r, 2)
               xlSheet.Cells(r1, 4).value = vs.TextMatrix(r, 3)
               xlSheet.Cells(r1, 5).value = vs.TextMatrix(r, 4)
               
               
               
               
               
               
               soldTillDate = 0
                      
               
               
               If RS.State = 1 Then RS.Close
               RS.Open "select sum(Qty_sale) as Tqty from NetSale_Fullyear where (convert(smalldatetime,invoiceDate,103)>= convert(smalldatetime,'" & fromDate1.value & "',103) and convert(smalldatetime,invoiceDate,103)<= convert(smalldatetime,'" & toDate1.value & "',103)) and BookCode='" & vs.TextMatrix(r, 0) & "'", con, adOpenKeyset
                   If Not IsNull(RS(0)) Then soldTillDate = RS(0)
                      xlSheet.Cells(r1, 6).value = soldTillDate
                  End If
               '----------------
               soldTillDate = 0
               
               If RS.State = 1 Then RS.Close
               RS.Open "select sum(Qty_sale) as Tqty from NetSale_Fullyear where (convert(smalldatetime,invoiceDate,103)>= convert(smalldatetime,'" & fromDate2.value & "',103) and convert(smalldatetime,invoiceDate,103)<= convert(smalldatetime,'" & toDate2.value & "',103)) and BookCode='" & vs.TextMatrix(r, 0) & "'", con, adOpenKeyset
                    
               If Not IsNull(RS(0)) Then soldTillDate = RS(0)
                xlSheet.Cells(r1, 7).value = soldTillDate
               
                
               
        Next
        

'======================================================================================================
ElseIf ddd = 2 Then
'======================================================================================================
              
xlSheet.Cells(1, 1).value = "Book Code"
xlSheet.Cells(1, 2).value = "Bool Name"
xlSheet.Cells(1, 3).value = "Stock Qty"
               
   

For r = 1 To vs.rows - 1
       
If vs.TextMatrix(r, 0) <> "" Then
       
       r1 = r1 + 1
       xlSheet.Cells(r1, 1).value = vs.TextMatrix(r, 0)
       xlSheet.Cells(r1, 2).value = vs.TextMatrix(r, 1)
       xlSheet.Cells(r1, 3).value = vs.TextMatrix(r, 2)
 End If
       
Next

        
        
        
        
End If







Screen.MousePointer = vbDefault


Exit Sub

Screen.MousePointer = vbDefault

err:

MsgBox err.Description



End Sub

Private Sub Form_Load()
Me.top = 0
Me.Left = 0


Me.Width = 13200
Me.Height = 10200

If fromDate_setup >= "01/04/2017" Then
   Check1_crm.Enabled = True
Else
   Check1_crm.Enabled = True
End If


If RS.State = 1 Then RS.Close
RS.Open "select cname,yarfrom,yarto from setup1 where " & stringyear & "", con
If RS.EOF = False Then

d_from = RS!yarfrom
d_to = RS!yarto

dateAson.value = RS!yarto

If LCase(UserName) = "v" Then
    If (Month(Date) = 4 Or Month(Date) <= 10) Then
       cmdBookTransfer.Visible = True
    Else
       cmdBookTransfer.Visible = False
    End If
End If


End If

Set RS = Nothing
Sort_Key = ""

fromdate.value = Date
todate.value = Date

cboCategory.Clear
If RS.State = 1 Then RS.Close
RS.Open "select * from StockHead order by id", con
While RS.EOF = False
  cboCategory.AddItem RS!head
  cboCategory.ItemData(cboCategory.NewIndex) = RS!id
  RS.MoveNext
Wend

If RS.State = 1 Then RS.Close
RS.Open "select GROUPCODE from BOOKS where " & stringyear & " group by GROUPCODE", con
While RS.EOF = False
  cboGroup.AddItem RS!groupcode
  RS.MoveNext
Wend

'cboGroup.AddItem "All"
'========================================
cboserName.Clear
cboserName.AddItem ""
If RS.State = 1 Then RS.Close
RS.Open "select serName from BOOKS group by serName order by serName", con
While RS.EOF = False
  If Not IsNull(RS!sername) Then
  cboserName.AddItem RS!sername
  End If
  RS.MoveNext
Wend


cboGroup.ListIndex = cboGroup.ListCount - 1
'========================================
BackColorFrom Me

'========================================
''CON.Execute "update books set KITCODE='n'"
''If RS.State = 1 Then RS.close
''RS.Open "SELECT KITCODE FROM BOOKS_KIT group by KITCODE", CON
''While RS.EOF = False
''   CON.Execute "update books set KITCODE='y' where bookcode='" & RS!Kitcode & "'"
''   RS.MoveNext
''Wend

End Sub


Private Sub Godown_Click()

End Sub

Private Sub Form_Resize()
panel.Left = (Me.ScaleWidth - panel.Width) / 2
panel.top = (Me.ScaleHeight - panel.Height) / 2

End Sub

Private Sub LaVolpeButton1_Click()
Screen.MousePointer = vbHourglass

vs.Clear
Dim I As Integer
Dim b As Boolean
Dim rs_data As New ADODB.Recordset
b = False
Dim Opening, Receive, Issue, sales As Long
Dim godown_rec, godown_issue, sum1, sum2, sum3 As String
pening = 0: Receive = 0: Issue = 0: sales = 0
sum1 = 0: sum2 = 0: sum3 = 0




If txtBookCode.text = "All" And cboGroup.text <> "All" Then

K = 0
vs.Cols = 5
vs.FormatString = "** BookCode ** |Book Name|Total Received|Total Sold|*** Balance ***"
vs.rows = 2
vs.ColWidth(1) = 5000



If rs_data.State = 1 Then rs_data.Close
rs_data.Open "select st.BOOKCODE,b.bookname from StockRegister as st " & _
"inner join Books as b on st.BOOKCODE = b.BOOKCODE where " & stringyear & " and b.GROUPCODE='" & cboGroup & "'" & _
" group by st.BOOKCODE,b.bookname", con, adOpenKeyset

sum1 = 0
sum2 = 0

While rs_data.EOF = False



If RS.State = 1 Then RS.Close
RS.Open "select sum(Qty) from StockRegister where " & stringyear & " and Dates <= datevalue('" & dateAson.value & "') and " & _
" Issue_Receive='Receive' and BookCode='" & rs_data.Fields("BOOKCODE").value & "'" & _
" ", con, adOpenKeyset
If Not IsNull(RS(0)) Then
   godown_rec = RS(0)
Else
   godown_rec = 0
End If


If RS.State = 1 Then RS.Close
RS.Open "select sum(Qty) from StockRegister where " & stringyear & " and Dates<= datevalue('" & dateAson.value & "') and " & _
" Issue_Receive='Issue' and BookCode='" & rs_data.Fields("BOOKCODE").value & "' and Category<>'Sales'" & _
" ", con, adOpenKeyset
If Not IsNull(RS(0)) Then
   Issue = RS(0)
Else
   Issue = 0
End If


If RS.State = 1 Then RS.Close
RS.Open "select sum(Qty) from StockRegister where " & stringyear & " and Dates<= datevalue('" & dateAson.value & "') and " & _
" Issue_Receive='Issue' and BookCode='" & rs_data.Fields("BOOKCODE").value & "' and Category='Sales'" & _
" ", con, adOpenKeyset
If Not IsNull(RS(0)) Then
   sales = RS(0)
Else
   sales = 0
End If




K = K + 1
vs.rows = vs.rows + 1


vs.TextMatrix(K, 0) = rs_data!Bookcode
vs.TextMatrix(K, 1) = rs_data!Bookname
vs.TextMatrix(K, 2) = godown_rec
vs.TextMatrix(K, 3) = sales

sum1 = sum1 + godown_rec
sum2 = sum2 + Issue
sum3 = sum3 + sales
Issue = (Issue - sales)
If Issue < 0 Then
   Issue = (-1 * Issue)
End If

vs.TextMatrix(K, 4) = (godown_rec - Issue)


rs_data.MoveNext
Wend

vs.TextMatrix(K + 1, 1) = "                                                                   * Total "
vs.TextMatrix(K + 1, 2) = (sum1 - (sum2 + sum3))

vs.Cell(flexcpBackColor, K + 1, 0, K + 1, 4) = &HB0D8FF
'===================================================================================================
ElseIf txtBookCode.text = "All" And cboGroup.text = "All" Then

K = 0
vs.Cols = 5
vs.FormatString = "** BookCode ** |Book Name|Total Received|Total Sold|*** Balance ***"
vs.rows = 2
vs.ColWidth(1) = 5000



If rs_data.State = 1 Then rs_data.Close
rs_data.Open "select st.BOOKCODE,b.bookname from StockRegister as st " & _
"inner join Books as b on st.BOOKCODE = b.BOOKCODE " & _
" group by st.BOOKCODE,b.bookname", con, adOpenKeyset

sum1 = 0
sum2 = 0

While rs_data.EOF = False



If RS.State = 1 Then RS.Close
RS.Open "select sum(Qty) from StockRegister where " & stringyear & " and Dates <= datevalue('" & dateAson.value & "') and " & _
" Issue_Receive='Receive' and BookCode='" & rs_data.Fields("BOOKCODE").value & "'" & _
" ", con, adOpenKeyset
If Not IsNull(RS(0)) Then
   godown_rec = RS(0)
Else
   godown_rec = 0
End If


If RS.State = 1 Then RS.Close
RS.Open "select sum(Qty) from StockRegister where " & stringyear & " and Dates<= datevalue('" & dateAson.value & "') and " & _
" Issue_Receive='Issue' and BookCode='" & rs_data.Fields("BOOKCODE").value & "' and Category<>'Sales'" & _
" ", con, adOpenKeyset
If Not IsNull(RS(0)) Then
   Issue = RS(0)
Else
   Issue = 0
End If


If RS.State = 1 Then RS.Close
RS.Open "select sum(Qty) from StockRegister where " & stringyear & " and Dates<= datevalue('" & dateAson.value & "') and " & _
" Issue_Receive='Issue' and BookCode='" & rs_data.Fields("BOOKCODE").value & "' and Category='Sales'" & _
" ", con, adOpenKeyset
If Not IsNull(RS(0)) Then
   sales = RS(0)
Else
   sales = 0
End If




K = K + 1
vs.rows = vs.rows + 1


vs.TextMatrix(K, 0) = rs_data!Bookcode
vs.TextMatrix(K, 1) = rs_data!Bookname
vs.TextMatrix(K, 2) = godown_rec
vs.TextMatrix(K, 3) = sales

sum1 = sum1 + godown_rec
sum2 = sum2 + Issue
sum3 = sum3 + sales
Issue = (Issue - sales)
If Issue < 0 Then
   Issue = (-1 * Issue)
End If

vs.TextMatrix(K, 4) = (godown_rec - Issue)


rs_data.MoveNext
Wend

vs.TextMatrix(K + 1, 1) = "                                                                   * Total "
vs.TextMatrix(K + 1, 2) = (sum1 - (sum2 + sum3))

vs.Cell(flexcpBackColor, K + 1, 0, K + 1, 4) = &HB0D8FF

'====================================================================================================
ElseIf txtBookCode.text <> "All" And cboGroup.text <> "All" Then




K = 0
vs.Cols = 6
vs.FormatString = "Date|>Quantity|Issue/Receive|Category|Godown|Binder Name  "
vs.rows = 1
vs.ColWidth(0) = 1200
vs.ColWidth(1) = 1500
vs.ColWidth(2) = 2000
vs.ColWidth(3) = 2000
vs.ColWidth(4) = 2000
vs.ColWidth(5) = 4000

If rs_data.State = 1 Then rs_data.Close
rs_data.Open "select st.Dates,st.Qty,st.Issue_Receive,st.Category,st.Issue_ReceveFrom, " & _
"st.BinderName from StockRegister as st inner join Books as b on st.BOOKCODE = b.BOOKCODE where " & stringyear & " and st.bookcode='" & txtBookCode.text & "' order by st.Dates" & _
" ", con, adOpenKeyset
sum1 = 0

While rs_data.EOF = False

K = K + 1
vs.rows = vs.rows + 1

vs.TextMatrix(K, 0) = rs_data(0)
vs.TextMatrix(K, 1) = rs_data(1)

If rs_data(2) = "Receive" Then
   sum1 = sum1 + rs_data(1)
Else
   sum1 = sum1 - rs_data(1)
End If

vs.TextMatrix(K, 2) = rs_data(2)
vs.TextMatrix(K, 3) = rs_data(3)
vs.TextMatrix(K, 4) = rs_data(4) & ""
vs.TextMatrix(K, 5) = IIf(IsNull(rs_data(5)), "----", rs_data(5))


rs_data.MoveNext
Wend
bookbalance.Caption = " Balance  :  " & sum1

End If


Screen.MousePointer = vbDefault

End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
'Create_Grid
'Load_Station
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 2
    frmIssue_ReceiceMaster.Show
Case 3
    BinderMaster.Show
Case 4
    CustomerMaster.Show
Case 6

    IssueBook = "StockTransfar"
    Unload frmBookIssue
    frmBookIssue.Show 1
Case 5

    IssueBook = "Issue"
    Unload frmBookIssue
    frmBookIssue.Show 1
Case 7
    'FRMInvoice.Show
    frmIssue1.Show
Case 1
    frmLoginPass.Show
Case 10
   End
Case 13
    End
End Select
End Sub


Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.Index
Case 1
   ' FRMInvoiceRpt.Show
Case 2
     frmConnectionPath.Show
Case 3

''---------------------

If RS.State = 1 Then RS.Close
RS.Open "select * from updatefrom", con
If RS.EOF = False Then
datefrom.value = RS(0)
End If

'----------------------


    StockFrame.Visible = True
Case 5

End Select
End Sub

Private Sub txtBookCode_GotFocus()
   txtBookCode.text = ""
   If PopUpValue2 <> "" Then
     txtBookCode.text = PopUpValue1
     PopUpValue2 = ""
     PopUpValue1 = ""
   End If
End Sub

Private Sub txtBookCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then

If cboGroup <> "All" Then
   
   searchType = "books"
   popuplist10 "select BOOKCODE,BOOKNAME from BOOKS where GROUPCODE ='" & cboGroup & "' and " & stringyear & "  order by BOOKCODE", con
Else
    
   searchType = "books"
   popuplist10 "select BOOKCODE,BOOKNAME from BOOKS where " & stringyear & "  order by BOOKCODE", con

End If


End If

End Sub
Private Sub txtBookCode_LostFocus()
 If txtBookCode.text = "" Then
    txtBookCode.text = "All"
 End If
End Sub
Private Sub vs_DblClick()
   PopUpValue1 = vs.TextMatrix(vs.RowSel, 0)
   PopUpValue2 = vs.TextMatrix(vs.RowSel, 1)
   
   Unload frmBookSt
   frmBookSt.Show 1
End Sub
