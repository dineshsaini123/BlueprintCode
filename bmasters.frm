VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form bookmaster 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5955
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   9585
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "bmasters.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   9585
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   435
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   9285
      TabIndex        =   41
      Top             =   5490
      Width           =   9345
      Begin VB.CommandButton CommandmasterReturn 
         Caption         =   "Return"
         Height          =   375
         Left            =   8220
         TabIndex        =   21
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton CommandmasterPrint 
         Caption         =   "&Print"
         Height          =   375
         Left            =   7200
         TabIndex        =   20
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton Commandmastersearch 
         Caption         =   "Sea&rch"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6210
         TabIndex        =   19
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton Commandmasterdelete 
         Caption         =   "De&lete"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5235
         TabIndex        =   18
         Top             =   15
         Width           =   800
      End
      Begin VB.CommandButton Commandmasterabandon 
         Caption         =   "Aba&ndon"
         Height          =   375
         Left            =   4245
         TabIndex        =   17
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton Commandmastersave 
         Caption         =   "Sa&ve"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3225
         TabIndex        =   11
         Top             =   -15
         Width           =   800
      End
      Begin VB.CommandButton Commandmasteredit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   2190
         TabIndex        =   16
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton Commandmasteradd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   1200
         TabIndex        =   15
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton Commandmasterhelp 
         Caption         =   "Help"
         Height          =   375
         Left            =   200
         TabIndex        =   42
         Top             =   0
         Visible         =   0   'False
         Width           =   800
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5385
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   9499
      _Version        =   393216
      Tabs            =   6
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      WordWrap        =   0   'False
      TabCaption(0)   =   "&Items"
      TabPicture(0)   =   "bmasters.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "booksmaster"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Item Gro&ups"
      TabPicture(1)   =   "bmasters.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "booksgroupmaster"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Distric&ts"
      TabPicture(2)   =   "bmasters.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "districtsmaster"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Agent Master"
      TabPicture(3)   =   "bmasters.frx":0060
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Agent"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Report Item Group"
      TabPicture(4)   =   "bmasters.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Bgroup"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Manufacture"
      TabPicture(5)   =   "bmasters.frx":0098
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "manufacture"
      Tab(5).ControlCount=   1
      Begin VB.Frame manufacture 
         Height          =   3915
         Left            =   -74550
         TabIndex        =   80
         Top             =   1080
         Width           =   8385
         Begin VB.ListBox listmanufacture 
            Enabled         =   0   'False
            Height          =   2400
            ItemData        =   "bmasters.frx":00B4
            Left            =   2190
            List            =   "bmasters.frx":00BB
            Sorted          =   -1  'True
            TabIndex        =   82
            Top             =   840
            Width           =   3915
         End
         Begin VB.TextBox txtmname 
            Height          =   285
            Left            =   2190
            MaxLength       =   30
            TabIndex        =   81
            Top             =   510
            Width           =   3915
         End
         Begin VB.Label Label18 
            Caption         =   "Name"
            Height          =   255
            Left            =   510
            TabIndex        =   83
            Top             =   480
            Width           =   825
         End
      End
      Begin VB.Frame booksmaster 
         Height          =   3960
         Left            =   -74610
         TabIndex        =   33
         Top             =   1170
         Width           =   8550
         Begin VB.ComboBox cboready 
            Height          =   315
            ItemData        =   "bmasters.frx":00D0
            Left            =   2760
            List            =   "bmasters.frx":00DA
            TabIndex        =   10
            Top             =   3570
            Width           =   1095
         End
         Begin VB.TextBox txtper 
            Height          =   285
            Left            =   2760
            TabIndex        =   9
            Top             =   3240
            Width           =   3000
         End
         Begin VB.TextBox txtquality 
            Height          =   285
            Left            =   2760
            TabIndex        =   7
            Top             =   2580
            Width           =   3060
         End
         Begin VB.TextBox txtunit2 
            Height          =   285
            Left            =   2760
            TabIndex        =   6
            Top             =   2235
            Width           =   3075
         End
         Begin VB.TextBox txtsize2 
            Height          =   270
            Left            =   2760
            TabIndex        =   5
            Top             =   1920
            Width           =   3120
         End
         Begin VB.TextBox txtunit1 
            Height          =   300
            Left            =   2745
            TabIndex        =   4
            Top             =   1575
            Width           =   3120
         End
         Begin VB.TextBox txtsize1 
            Height          =   300
            Left            =   2730
            TabIndex        =   3
            Top             =   1230
            Width           =   3150
         End
         Begin VB.ComboBox Comboname 
            Height          =   315
            Left            =   7620
            Sorted          =   -1  'True
            TabIndex        =   45
            Top             =   1635
            Visible         =   0   'False
            Width           =   750
         End
         Begin VB.ComboBox Combocode 
            Height          =   315
            Left            =   2730
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   150
            Visible         =   0   'False
            Width           =   3135
         End
         Begin VB.ComboBox Combobgroupname 
            Height          =   315
            Left            =   6450
            Sorted          =   -1  'True
            TabIndex        =   13
            Top             =   120
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.ComboBox Combobgroupcode 
            Height          =   315
            ItemData        =   "bmasters.frx":00E7
            Left            =   2730
            List            =   "bmasters.frx":00F1
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   510
            Width           =   705
         End
         Begin VB.TextBox Textbbookname 
            Height          =   285
            Left            =   2730
            MaxLength       =   40
            TabIndex        =   2
            Top             =   870
            Width           =   3135
         End
         Begin VB.TextBox Textfindbookcode 
            Height          =   285
            Left            =   120
            TabIndex        =   34
            Top             =   150
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox Textbrate 
            Height          =   285
            Left            =   2760
            MaxLength       =   10
            TabIndex        =   8
            Top             =   2910
            Width           =   3015
         End
         Begin VB.TextBox Textbdiscount 
            Height          =   285
            Left            =   7545
            MaxLength       =   10
            TabIndex        =   14
            Top             =   1125
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.TextBox Textbbookcode 
            Height          =   315
            Left            =   3465
            MaxLength       =   10
            TabIndex        =   1
            Top             =   510
            Width           =   2415
         End
         Begin VB.Label Label12 
            Caption         =   "Per"
            Height          =   165
            Left            =   660
            TabIndex        =   70
            Top             =   3315
            Width           =   585
         End
         Begin VB.Label Label9 
            Caption         =   "Quality"
            Height          =   285
            Left            =   645
            TabIndex        =   69
            Top             =   2610
            Width           =   930
         End
         Begin VB.Label Label8 
            Caption         =   "Unit2"
            Height          =   165
            Left            =   660
            TabIndex        =   68
            Top             =   2325
            Width           =   885
         End
         Begin VB.Label Label7 
            Caption         =   "Size2"
            Height          =   210
            Left            =   630
            TabIndex        =   67
            Top             =   1965
            Width           =   870
         End
         Begin VB.Label Label6 
            Caption         =   "Unit1"
            Height          =   180
            Left            =   645
            TabIndex        =   66
            Top             =   1590
            Width           =   900
         End
         Begin VB.Label Label5 
            Caption         =   "Size1"
            Height          =   210
            Left            =   630
            TabIndex        =   65
            Top             =   1245
            Width           =   975
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Rate"
            Height          =   195
            Left            =   675
            TabIndex        =   40
            Top             =   2955
            Width           =   345
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Discount "
            Height          =   195
            Left            =   6780
            TabIndex        =   39
            Top             =   1215
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Ready Made"
            Height          =   195
            Left            =   645
            TabIndex        =   38
            Top             =   3630
            Width           =   1440
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Item Name"
            Height          =   195
            Left            =   600
            TabIndex        =   37
            Top             =   960
            Width           =   765
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Item Code"
            Height          =   195
            Left            =   600
            TabIndex        =   36
            Top             =   570
            Width           =   720
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Group Name"
            Height          =   195
            Left            =   6645
            TabIndex        =   35
            Top             =   705
            Visible         =   0   'False
            Width           =   900
         End
      End
      Begin VB.Frame Agent 
         Height          =   4215
         Left            =   450
         TabIndex        =   46
         Top             =   780
         Width           =   8475
         Begin VB.TextBox aphone 
            Height          =   285
            Left            =   1830
            TabIndex        =   74
            Top             =   2775
            Width           =   2610
         End
         Begin VB.TextBox acity 
            Height          =   315
            Left            =   1800
            TabIndex        =   73
            Top             =   2265
            Width           =   2595
         End
         Begin VB.TextBox aadd2 
            Height          =   315
            Left            =   1785
            TabIndex        =   72
            Top             =   1770
            Width           =   2520
         End
         Begin VB.TextBox aadd1 
            Height          =   285
            Left            =   1815
            TabIndex        =   71
            Top             =   1305
            Width           =   2580
         End
         Begin VB.TextBox TextfindAgentmaster 
            Height          =   285
            Left            =   75
            TabIndex        =   53
            Top             =   240
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.CommandButton Addcd 
            Height          =   540
            Left            =   6810
            Picture         =   "bmasters.frx":00FE
            Style           =   1  'Graphical
            TabIndex        =   52
            Top             =   1365
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.CommandButton Removecd 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   6870
            Picture         =   "bmasters.frx":0440
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   2250
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.ListBox ListDis1 
            Height          =   2400
            ItemData        =   "bmasters.frx":0782
            Left            =   5940
            List            =   "bmasters.frx":0784
            Sorted          =   -1  'True
            TabIndex        =   49
            Top             =   1140
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.ListBox ListDis2 
            Height          =   2400
            ItemData        =   "bmasters.frx":0786
            Left            =   7590
            List            =   "bmasters.frx":0788
            Sorted          =   -1  'True
            TabIndex        =   48
            Top             =   1185
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.ComboBox comboAgentMaster 
            Height          =   315
            Left            =   1770
            Sorted          =   -1  'True
            TabIndex        =   47
            Top             =   780
            Width           =   2652
         End
         Begin VB.Label Label16 
            Caption         =   "Phone"
            Height          =   270
            Left            =   450
            TabIndex        =   78
            Top             =   2820
            Width           =   1200
         End
         Begin VB.Label Label15 
            Caption         =   "City"
            Height          =   225
            Left            =   435
            TabIndex        =   77
            Top             =   2340
            Width           =   1170
         End
         Begin VB.Label Label14 
            Caption         =   "Add2"
            Height          =   240
            Left            =   420
            TabIndex        =   76
            Top             =   1815
            Width           =   1380
         End
         Begin VB.Label Label13 
            Caption         =   "Add1"
            Height          =   240
            Left            =   420
            TabIndex        =   75
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Agent Name "
            Height          =   195
            Left            =   480
            TabIndex        =   50
            Top             =   810
            Width           =   930
         End
      End
      Begin VB.Frame booksgroupmaster 
         Height          =   3495
         Left            =   -73920
         TabIndex        =   28
         Top             =   1770
         Width           =   6885
         Begin VB.TextBox textbgfindcode 
            Height          =   345
            Left            =   3330
            TabIndex        =   43
            Top             =   510
            Visible         =   0   'False
            Width           =   3195
         End
         Begin VB.TextBox Textbggroupname 
            Height          =   285
            Left            =   3345
            MaxLength       =   49
            TabIndex        =   25
            Top             =   1740
            Width           =   3135
         End
         Begin VB.TextBox Textbggroupcode 
            Height          =   285
            Left            =   3330
            MaxLength       =   7
            TabIndex        =   24
            Top             =   1170
            Width           =   3135
         End
         Begin VB.Label Label17 
            Caption         =   "Group Name"
            Height          =   375
            Left            =   600
            TabIndex        =   32
            Top             =   1710
            Width           =   2295
         End
         Begin VB.Label Label22 
            Height          =   585
            Left            =   600
            TabIndex        =   31
            Top             =   1770
            Width           =   2865
         End
         Begin VB.Label Label23 
            Caption         =   "Group Code"
            Height          =   255
            Left            =   600
            TabIndex        =   30
            Top             =   1200
            Width           =   2955
         End
         Begin VB.Label Label24 
            Height          =   585
            Left            =   600
            TabIndex        =   29
            Top             =   2070
            Width           =   2865
         End
      End
      Begin VB.Frame Bgroup 
         Enabled         =   0   'False
         Height          =   4215
         Left            =   -74820
         TabIndex        =   54
         Top             =   810
         Width           =   8895
         Begin VB.CommandButton cmdm1 
            Height          =   540
            Left            =   2670
            Picture         =   "bmasters.frx":078A
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   1410
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton cmdm2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   2670
            Picture         =   "bmasters.frx":0ACC
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   2040
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.ListBox List1 
            Height          =   2400
            ItemData        =   "bmasters.frx":0E0E
            Left            =   180
            List            =   "bmasters.frx":0E10
            Sorted          =   -1  'True
            TabIndex        =   59
            Top             =   1080
            Width           =   2205
         End
         Begin VB.ListBox List2 
            Height          =   2400
            ItemData        =   "bmasters.frx":0E12
            Left            =   3270
            List            =   "bmasters.frx":0E14
            Sorted          =   -1  'True
            TabIndex        =   58
            Top             =   1140
            Width           =   2205
         End
         Begin VB.CommandButton cmdm3 
            Height          =   540
            Left            =   5790
            Picture         =   "bmasters.frx":0E16
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   1380
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton cmdm4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   5790
            Picture         =   "bmasters.frx":1158
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   2010
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.ListBox List3 
            Height          =   2400
            ItemData        =   "bmasters.frx":149A
            Left            =   6450
            List            =   "bmasters.frx":149C
            Sorted          =   -1  'True
            TabIndex        =   55
            Top             =   1140
            Width           =   2205
         End
         Begin VB.Label Label2 
            Caption         =   "Group1"
            Height          =   315
            Left            =   240
            TabIndex        =   64
            Top             =   690
            Width           =   1275
         End
         Begin VB.Label Label3 
            Caption         =   "Group2 "
            Height          =   285
            Left            =   3300
            TabIndex        =   63
            Top             =   690
            Width           =   1245
         End
         Begin VB.Label Label4 
            Caption         =   "Group3"
            Height          =   345
            Left            =   6450
            TabIndex        =   62
            Top             =   660
            Width           =   735
         End
      End
      Begin VB.Frame districtsmaster 
         Height          =   3585
         Left            =   -73650
         TabIndex        =   26
         Top             =   1230
         Width           =   6885
         Begin VB.TextBox TXTCODE 
            Height          =   285
            Left            =   2100
            MaxLength       =   30
            TabIndex        =   79
            Top             =   330
            Width           =   915
         End
         Begin VB.ListBox listdistrictmaster 
            Enabled         =   0   'False
            Height          =   2400
            ItemData        =   "bmasters.frx":149E
            Left            =   2100
            List            =   "bmasters.frx":14A0
            Sorted          =   -1  'True
            TabIndex        =   23
            Top             =   660
            Width           =   3915
         End
         Begin VB.TextBox Textddistrictname 
            Height          =   285
            Left            =   3030
            MaxLength       =   30
            TabIndex        =   22
            Top             =   330
            Width           =   2985
         End
         Begin VB.Label Label30 
            Caption         =   "District Name"
            Height          =   255
            Left            =   420
            TabIndex        =   27
            Top             =   300
            Width           =   1605
         End
      End
   End
End
Attribute VB_Name = "bookmaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim CON As ADODB.Connection
Dim RS As ADODB.Recordset
Dim ctl As Control
Public addmaster As Boolean
Private Sub Comboslgenledgerdiscription_LostFocus()
Set RS = New ADODB.Recordset
    RS.Open "select gledger from gledger where slf= 1 and gledger='" + Trim(Comboslgenledgerdiscription.Text) + "' and  " & stringyear & " ", CON, adOpenKeyset, adLockReadOnly, adCmdText
    If RS.EOF Then
        MsgBox Comboslgenledgerdiscription.Text + " is not configed for subledger "
        Comboslgenledgerdiscription.SetFocus
    End If
End Sub

Private Sub ComboSPECIALCATEGORY_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Textglgeneralledgerdiscription.SetFocus
End If
End Sub
Private Sub Addcd_Click()
If ListDis1.ListCount > 0 Then
    Dim I
    Dim delitem As Integer
    For I = 0 To ListDis1.ListCount - 1
        If ListDis1.Selected(I) Then
                ListDis2.AddItem ListDis1.List(I)
                delitem = I
         End If
    Next
    ListDis1.RemoveItem delitem
End If

End Sub
Sub cmdm1_Click()
If List1.ListCount > 0 Then
    Dim I
    Dim delitem As Integer
    For I = 0 To List1.ListCount - 1
        If List1.Selected(I) Then
                List2.AddItem List1.List(I)
                delitem = I
         End If
    Next I
    List1.RemoveItem delitem
End If
End Sub

Sub cmdm2_Click()
If List2.ListCount > 0 Then
    Dim I
    Dim delitem As Integer
    For I = 0 To List2.ListCount - 1
        If List2.Selected(I) Then
                List1.AddItem List2.List(I)
                delitem = I
         End If
    Next
    List2.RemoveItem delitem
    
End If
End Sub

Sub cmdm3_Click()
If List1.ListCount > 0 Then
    Dim I
    Dim delitem As Integer
    For I = 0 To List1.ListCount - 1
        If List1.Selected(I) Then
                List3.AddItem List1.List(I)
                delitem = I
         End If
    Next
    List1.RemoveItem delitem
End If
End Sub

Sub cmdm4_Click()
If List3.ListCount > 0 Then
    Dim I
    Dim delitem As Integer
    For I = 0 To List3.ListCount - 1
        If List3.Selected(I) Then
                List1.AddItem List3.List(I)
                delitem = I
         End If
    Next
    List3.RemoveItem delitem
End If


End Sub

Private Sub comboAgentMaster_Change()
Dim trs As New ADODB.Recordset
Dim tRS1 As New ADODB.Recordset

tRS1.Open "Select districtname from  districts  where " & stringyear, CON, adOpenStatic, adLockPessimistic
If tRS1.RecordCount > 0 Then
    ListDis1.Clear
    tRS1.MoveFirst
    While Not tRS1.EOF
         If IsNull(tRS1(0)) = False Then
               ListDis1.AddItem tRS1(0)
         End If
         If Not tRS1.EOF Then
                tRS1.MoveNext
         End If
   Wend
   ListDis1.Selected(0) = True
End If

trs.Open "Select  districtname  from districts where  " & stringyear & " and agentname= '" + Trim(Me.comboAgentMaster.Text) + "'", CON, adOpenStatic, adLockPessimistic

If trs.RecordCount > 0 Then
    ListDis2.Clear
    trs.MoveFirst
    While Not trs.EOF
         If IsNull(trs(0)) = False Then
            ListDis2.AddItem trs(0)
         End If
         If Not trs.EOF Then
            trs.MoveNext
         End If
   Wend
Else
  ListDis2.Clear
End If

End Sub

Private Sub comboAgentMaster_Click()
comboAgentMaster_Change
End Sub

Private Sub comboAgentMaster_KeyPress(KeyAscii As Integer)
''''''If KeyAscii = 13 Then
''''''
''''''Dim trs As New ADODB.Recordset
''''''Dim tRS1 As New ADODB.Recordset
''''''comboAgentMaster.Text = UCase(comboAgentMaster.Text)
''''''
''''''tRS1.Open "Select districtname from  districts  where  " & stringyear & " and agentname=''", CON, adOpenStatic, adLockPessimistic
''''''If tRS1.RecordCount > 0 Then
''''''    ListDis1.Clear
''''''    tRS1.MoveFirst
''''''    While Not tRS1.EOF
''''''         If IsNull(tRS1(0)) = False Then
''''''               ListDis1.AddItem tRS1(0)
''''''         End If
''''''         If Not tRS1.EOF Then
''''''                tRS1.MoveNext
''''''         End If
''''''   Wend
''''''  ListDis1.Selected(0) = True
''''''   ListDis1.SetFocus
''''''End If
''''''
''''''trs.Open "Select  districtname  from districts where  " & stringyear & " and agentname= '" + Trim(Me.comboAgentMaster.Text) + "'", CON, adOpenStatic, adLockPessimistic
''''''
''''''If trs.RecordCount > 0 Then
''''''    ListDis2.Clear
''''''    trs.MoveFirst
''''''    While Not trs.EOF
''''''         If IsNull(trs(0)) = False Then
''''''            ListDis2.AddItem trs(0)
''''''         End If
''''''         If Not trs.EOF Then
''''''            trs.MoveNext
''''''         End If
''''''   Wend
''''''Else
''''''  ListDis2.Clear
''''''End If
''''''
''''''End If

End Sub

Private Sub comboAgentMaster_LostFocus()
comboAgentMaster.Text = UCase(comboAgentMaster.Text)
End Sub

Private Sub Combobgroupcode_Change()
    
''    Dim temp As ADODB.Recordset
''    Set temp = New ADODB.Recordset
''    temp.Open "select * from groups where " & stringyear & " ", CON, adOpenKeyset, adLockReadOnly, adCmdText
''    If Not temp.EOF Then
''        temp.Find "groupcode='" + Trim(Me.Combobgroupcode.Text) + "'"
''        If Not temp.EOF Then
''          Me.Combobgroupname.Text = temp(1)
''        End If
''    End If
''    temp.Close
End Sub

Private Sub Combobgroupcode_Click()
''    Dim temp As ADODB.Recordset
''    Set temp = New ADODB.Recordset
''    temp.Open "select * from groups where " & stringyear & " ", CON, adOpenKeyset, adLockReadOnly, adCmdText
''    If Not temp.EOF Then
''        temp.Find "groupcode='" + Trim(Me.Combobgroupcode.Text) + "'"
''        If Not temp.EOF Then
''          Me.Combobgroupname.Text = temp(1)
''        End If
''    End If
''    temp.Close
If addmaster = True Then
    Dim temp As ADODB.Recordset
    Set temp = New ADODB.Recordset
    temp.Open "select MAX(CONVERT(INT,RIGHT(BOOKCODE,LEN(BOOKCODE)-" & Len(Combobgroupcode.Text) & "))) AS MAXID from BOOKS where " & stringyear & "  AND BOOKCODE LIKE '" & UCase(Combobgroupcode.Text) & "%'", CON, adOpenKeyset, adLockReadOnly, adCmdText
    If temp.EOF = False Then
        If temp!maxId > 0 Then
            Textbbookcode.Text = Format(temp!maxId + 1, "000")
        Else
            Textbbookcode.Text = "001"
        End If
    Else
        Textbbookcode.Text = "001"
    End If
    temp.close
End If
End Sub

Private Sub Combobgroupcode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Combobgroupcode = UCase(Combobgroupcode)

   SendKeys "{TAB}"
End If
End Sub

Private Sub Combobgroupcode_LostFocus()
''    If Len(Combobgroupcode.Text) > 7 Then
''       MsgBox "Enter maximum 7 character"
''       Combobgroupcode.SetFocus
''    End If
''
''    rs.Open "select * from groups where " & stringyear & " ", CON, adOpenKeyset, adLockReadOnly, adCmdText
''    If Not rs.EOF Then
''        rs.Find "groupcode='" + Trim(Me.Combobgroupcode.Text) + "'"
''        If Not rs.EOF Then
''            Me.Combobgroupname.Text = rs(1)
''        End If
''    Else
''      MsgBox "Please Enter Valid Group Code...."
''      Combobgroupcode.SetFocus
''    End If
''    rs.Close
End Sub

Private Sub Combobgroupname_Change()
   
    Dim temp As ADODB.Recordset
    Set temp = New ADODB.Recordset
    temp.Open "select * from groups where " & stringyear & " ", CON, adOpenKeyset, adLockReadOnly, adCmdText
    If Not temp.EOF Then
        temp.Find "groupname='" + Trim(Me.Combobgroupname.Text) + "'"
        If Not temp.EOF Then
            Me.Combobgroupcode.Text = temp(0)
        End If
    End If
    temp.close
End Sub

Private Sub Combobgroupname_Click()
    Dim temp As ADODB.Recordset
    Set temp = New ADODB.Recordset
    temp.Open "select * from groups where " & stringyear & " ", CON, adOpenKeyset, adLockReadOnly, adCmdText
    If Not temp.EOF Then
        temp.Find "groupname='" + Trim(Me.Combobgroupname.Text) + "'"
        If Not temp.EOF Then
            Me.Combobgroupcode.Text = temp(0)
        End If
    End If
    temp.close
End Sub

Private Sub Combobgroupname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{TAB}"
   Combobgroupname = UCase(Combobgroupname)
End If
End Sub

Private Sub Combobgroupname_LostFocus()
    RS.Open "select * from groups " & stringyear & " ", CON, adOpenKeyset, adLockReadOnly, adCmdText
    If Not RS.EOF Then
        RS.Find "groupname='" + Trim(Me.Combobgroupname.Text) + "'"
        If Not RS.EOF Then
            Me.Combobgroupcode.Text = RS(0)
        End If
     Else
       MsgBox "Not Valid Group Name"
       Me.Combobgroupname.SetFocus
    End If
    RS.close
End Sub

Private Sub Commandmasterabandon_Click()
    For I = 0 To 2
           Me.SStab1.TabEnabled(I) = True
    Next
    For Each ctl In Me.Controls
        If TypeOf ctl Is textbox Then
                If ctl.Text <> "" And ctl.Enabled = True Then ctl.Text = ""
                ctl.Enabled = False
            End If
            If TypeOf ctl Is ComboBox Then
            If ctl.Style <> 2 Then
             ctl.Text = ""
             End If

            End If
            If TypeOf ctl Is CheckBox Then
                ctl.value = 0
                ctl.Enabled = False
            End If
            If TypeOf ctl Is ListBox Then
                ctl.Enabled = False
            End If
    Next
Commandmasteradd.Enabled = True
Commandmasteredit.Enabled = True
CommandmasterPrint.Enabled = True
If Me.SStab1.Tab <> 2 Then
    Commandmastersearch.Enabled = True
End If
Commandmastersave.Enabled = False
'Commandmasterabandon.Enabled = False
CommandmasterReturn.Enabled = True
Me.Commandmasterdelete.Enabled = False
ListDis1.Clear
ListDis2.Clear

End Sub

Private Sub Commandmasteradd_Click()
    addmaster = True
    
    If SStab1.Tab = 0 Then
    '/**  deactivate other tabs**/
        For I = 0 To 2
            If I <> 0 Then
                SStab1.TabEnabled(I) = False
            End If
        Next
        
        For Each ctl In Me.Controls
              If UCase(ctl.CONTAINER.Name) = UCase("booksmaster") Then
                If TypeOf ctl Is textbox Or TypeOf ctl Is ComboBox Or TypeOf ctl Is CheckBox Or TypeOf ctl Is ListBox Then
                    ctl.Enabled = True
                 
                End If
            End If
        Next
      'Textbbookcode.SetFocus
      cboready.ListIndex = 0
    
    End If
    
    
    If SStab1.Tab = 1 Then
    '/**  deactivate other tabs**/
        For I = 0 To 2
            If I <> 1 Then
                SStab1.TabEnabled(I) = False
            End If
        Next
        For Each ctl In Me.Controls
            If UCase(ctl.CONTAINER.Name) = UCase("booksgroupmaster") Then
                If TypeOf ctl Is textbox Or TypeOf ctl Is ComboBox Or TypeOf ctl Is CheckBox Or TypeOf ctl Is ListBox Then
                    ctl.Enabled = True
                End If
            End If
        Next
        Textbggroupcode.SetFocus
    End If
    If SStab1.Tab = 2 Then
    '/**  deactivate other tabs**/
        For I = 0 To 2
            If I <> 2 Then
                SStab1.TabEnabled(I) = False
            End If
        Next
        For Each ctl In Me.Controls
            If UCase(ctl.CONTAINER.Name) = UCase("districtsmaster") Then
                If TypeOf ctl Is textbox Or TypeOf ctl Is ComboBox Or TypeOf ctl Is CheckBox Or TypeOf ctl Is ListBox Then
                    ctl.Enabled = True
                End If
            End If
        Next
        Textddistrictname.SetFocus
    End If


  If SStab1.Tab = 3 Then
    '/**  deactivate other tabs**/
        For I = 0 To 3
            If I <> 3 Then
                SStab1.TabEnabled(I) = False
            End If
        Next
        For Each ctl In Me.Controls
            If UCase(ctl.CONTAINER.Name) = UCase("Agent") Then
                If TypeOf ctl Is textbox Or TypeOf ctl Is ComboBox Or TypeOf ctl Is CheckBox Or TypeOf ctl Is ListBox Then
                    ctl.Enabled = True
                End If
            End If
        Next
        Me.comboAgentMaster.Clear
        'ListDis1.Clear
        ListDis2.Clear
        Me.Agent.Enabled = True
        Me.comboAgentMaster.Enabled = True
        Me.comboAgentMaster.SetFocus
        Me.Addcd.Enabled = True
        Me.Removecd.Enabled = True
    End If
  
    
   If SStab1.Tab = 4 Then
    '/**  deactivate other tabs**/
        For I = 0 To 4
            If I <> 4 Then
                SStab1.TabEnabled(I) = False
            End If
        Next
        For Each ctl In Me.Controls
            If UCase(ctl.CONTAINER.Name) = UCase("Bgroup") Then
                If TypeOf ctl Is textbox Or TypeOf ctl Is ComboBox Or TypeOf ctl Is CheckBox Or TypeOf ctl Is ListBox Then
                    ctl.Enabled = True
                End If
            End If
        Next
        List1.Clear
        List2.Clear
        List3.Clear
        
        Me.Bgroup.Enabled = True
        List1.SetFocus
        Me.cmdm1.Enabled = True
        Me.cmdm2.Enabled = True
        Me.cmdm3.Enabled = True
        Me.cmdm4.Enabled = True
        
        
    End If
    
    If SStab1.Tab = 5 Then
    '/**  deactivate other tabs**/
        For I = 0 To 5
            If I <> 5 Then
                SStab1.TabEnabled(I) = False
            End If
        Next
        For Each ctl In Me.Controls
            If UCase(ctl.CONTAINER.Name) = UCase("manufacture") Then
                If TypeOf ctl Is textbox Or TypeOf ctl Is ComboBox Or TypeOf ctl Is CheckBox Or TypeOf ctl Is ListBox Then
                    ctl.Enabled = True
                End If
            End If
        Next
        'txtmname.SetFocus
    End If
    
    Commandmasteradd.Enabled = False
    Commandmasteredit.Enabled = False
    CommandmasterPrint.Enabled = False
    Commandmastersave.Enabled = True
    Commandmasterabandon.Enabled = True
    CommandmasterReturn.Enabled = True
    Commandmastersearch.Enabled = True
End Sub
    
Private Sub Commandmasterdelete_Click()
'=====================rohan==============
'
'    If SSTab1.Tab = 0 Then
'        If Me.Textbbookcode.Text <> "" Then
'            rs.Open "select * from books where  " & stringyear & " order by bookcode", CON, adOpenKeyset, adLockOptimistic, adCmdText
'            rs.Find "bookcode='" + Combobgroupcode.Text & Trim(UCase(Me.Textbbookcode.Text)) + "'"
'            If Not rs.EOF Then
'                X = MsgBox("Are You Sure...!", vbYesNo, "Warning.....")
'                If X = 6 Then
'                    rs.Delete
'                    rs.Update
'                    Me.Textbbookcode.Text = ""
'                    Me.Textbbookname.Text = ""
'                    'Me.Textbdiscount.Text = ""
'                    Me.Textbrate.Text = ""
'                    Me.txtper.Text = ""
'                    Me.txtquality.Text = ""
'                    Me.txtsize1.Text = ""
'                    Me.txtsize2.Text = ""
'                    Me.txtunit1.Text = ""
'                    Me.txtunit2.Text = ""
'
'                    'Me.Combobgroupcode.Text = ""
'                    'Me.Combobgroupname.Text = ""
'
'                End If
'            Else
'                MsgBox "Not Found"
'            End If
'            rs.Close
'        End If
'    End If
'
''********** book  group
'
'
'
' If SSTab1.Tab = 1 Then
'        If Me.Textbggroupcode.Text <> "" Then
'             rs.CancelUpdate
'            If rs.State = 1 Then rs.Close
'            rs.Open "select * from groups where  " & stringyear & " order by groupcode", CON, adOpenKeyset, adLockOptimistic, adCmdText
'            rs.Find "groupcode='" + Trim(UCase(Me.Textbggroupcode.Text)) + "'"
'            If Not rs.EOF Then
'               X = MsgBox("Are You Sure...!", vbYesNo, "Warning.....")
'               If X = 6 Then
'               On Error GoTo tt
'                  rs.Delete
'tt:               If Err.Number = -2147217887 Then
'                     MsgBox " This group used in Book or Category.."
'                     Exit Sub
'               End If
'               rs.Update
'               rs.Close
'
'
'
'                    Me.Textbggroupcode = ""
'                    Me.Textbggroupname = ""
'                End If
'            Else
'                MsgBox "Not Found"
'            End If
'
'        End If
'    End If
'
'
''/*********************************
''   district
''/*********************************
'    If SSTab1.Tab = 2 Then
'        If Me.Textddistrictname.Text <> "" Then
'            rs.Open "select * from DISTRICTS where  " & stringyear & "", CON, adOpenKeyset, adLockOptimistic, adCmdText
'            rs.Find "districtname='" + Trim(UCase(Me.Textddistrictname.Text)) + "'"
'            If Not rs.EOF Then
'                X = MsgBox("Are You Sure...!", vbYesNo, "Warning.....")
'                If X = 6 Then
'                    rs.Delete
'                    rs.Update
'                    rs.Close
'                    rs.Open "select * from  DISTRICTS where " & stringyear & " ", CON, adOpenKeyset, adLockOptimistic, adCmdText
'                    Me.Textddistrictname.Text = ""
'                    Me.listdistrictmaster.Clear
'                    Do While Not rs.EOF
'                        Me.listdistrictmaster.AddItem Trim(UCase(rs(0)))
'                        If Not rs.EOF Then
'                            rs.MoveNext
'                        End If
'                    Loop
'                End If
'            Else
'                MsgBox "Not Found"
'            End If
'            rs.Close
'        End If
'    End If
'
'
'
''///////////////////////////////
''          AGENT
''///////////////////////////////
'
' If SSTab1.Tab = 3 Then
'        If Me.comboAgentMaster.Text <> "" Then
'            rs.Open "select * from AgentMaster where  " & stringyear & "  order by agentname", CON, adOpenKeyset, adLockOptimistic, adCmdText
'            rs.Find "agentname='" + Trim(UCase(Me.comboAgentMaster.Text)) + "'"
'            If Not rs.EOF Then
'                X = MsgBox("Are You Sure...!", vbYesNo, "Warning.....")
'                If X = 6 Then
'                    CON.Execute "Delete  from AgentMaster where agentname ='" + Trim(UCase(Me.comboAgentMaster.Text)) + "' and " & stringyear & " "
'                    'CON.Execute "Update Districts Set  agentname = '' where  " & stringyear & " and agentname  ='" + Trim(UCase(Me.comboAgentMaster.Text)) + "'"
'                    Me.comboAgentMaster.Text = ""
'                    Me.aadd1.Text = ""
'                    Me.aadd2.Text = ""
'                    Me.acity.Text = ""
'                    Me.aphone.Text = ""
'                    'ListDis1.Clear
'                    'ListDis2.Clear
'
'                End If
'            Else
'                MsgBox "Not Found"
'            End If
'            rs.Close
'        End If
'    End If

End Sub

Private Sub Commandmasteredit_Click()
    addmaster = False

    If Me.SStab1.Tab = 0 Then
        For I = 0 To 2
            If I <> 0 Then
                SStab1.TabEnabled(I) = False
            End If
        Next
        For Each ctl In Me.Controls
            If UCase(ctl.CONTAINER.Name) = UCase("booksmaster") Then
                If TypeOf ctl Is textbox Or TypeOf ctl Is ComboBox Or TypeOf ctl Is CheckBox Or TypeOf ctl Is ListBox Then
                    ctl.Enabled = True
                End If
            End If
        Next
        Me.Commandmasteradd.Enabled = False
        Me.Commandmasteredit.Enabled = False
        Me.Commandmasterabandon.Enabled = True
        Me.Commandmastersave.Enabled = True
        Me.Commandmasterdelete.Enabled = True
    End If
    If Me.SStab1.Tab = 1 Then
        For I = 0 To 2
            If I <> 1 Then
                SStab1.TabEnabled(I) = False
            End If
        Next
        For Each ctl In Me.Controls
            If UCase(ctl.CONTAINER.Name) = UCase("booksgroupmaster") Then
                If TypeOf ctl Is textbox Or TypeOf ctl Is ComboBox Or TypeOf ctl Is CheckBox Or TypeOf ctl Is ListBox Then
                    ctl.Enabled = True
                End If
            End If
        Next
        Me.Commandmasteradd.Enabled = False
        Me.Commandmasteredit.Enabled = False
        Me.Commandmasterabandon.Enabled = True
        Me.Commandmastersave.Enabled = True
        Me.Commandmasterdelete.Enabled = True
    End If
    If Me.SStab1.Tab = 2 Then
        For I = 0 To 2
            If I <> 2 Then
                SStab1.TabEnabled(I) = False
            End If
        Next
        For Each ctl In Me.Controls
            If UCase(ctl.CONTAINER.Name) = UCase("districtsmaster") Then
                If TypeOf ctl Is textbox Or TypeOf ctl Is ComboBox Or TypeOf ctl Is CheckBox Or TypeOf ctl Is ListBox Then
                    ctl.Enabled = True
                End If
            End If
        Next
        Me.Commandmasteradd.Enabled = False
        Me.Commandmasteredit.Enabled = False
        Me.Commandmasterabandon.Enabled = True
        Me.Commandmastersave.Enabled = True
        Me.Commandmasterdelete.Enabled = True
    End If
    
    If Me.SStab1.Tab = 3 Then
        For I = 0 To 3
            If I <> 3 Then
                SStab1.TabEnabled(I) = False
            End If
        Next
        For Each ctl In Me.Controls
            If UCase(ctl.CONTAINER.Name) = UCase("Agent") Then
                If TypeOf ctl Is textbox Or TypeOf ctl Is ComboBox Or TypeOf ctl Is CheckBox Or TypeOf ctl Is ListBox Then
                    ctl.Enabled = True
                End If
            End If
        Next
        Me.Commandmasteradd.Enabled = False
        Me.Commandmasteredit.Enabled = False
        Me.Commandmasterabandon.Enabled = True
        Me.Commandmastersave.Enabled = True
        Me.Commandmasterdelete.Enabled = True
        TextfindAgentmaster.Text = comboAgentMaster.Text
        Me.Agent.Enabled = True
    End If
  If Me.SStab1.Tab = 4 Then
        For I = 0 To 4
            If I <> 4 Then
                SStab1.TabEnabled(I) = False
            End If
        Next
        For Each ctl In Me.Controls
            If UCase(ctl.CONTAINER.Name) = UCase("Bgroup") Then
                If TypeOf ctl Is textbox Or TypeOf ctl Is ComboBox Or TypeOf ctl Is CheckBox Or TypeOf ctl Is ListBox Then
                    ctl.Enabled = True
                End If
            End If
        Next
        Me.Commandmasteradd.Enabled = False
        Me.Commandmasteredit.Enabled = False
        Me.Commandmasterabandon.Enabled = True
        Me.Commandmastersave.Enabled = True
        Me.Commandmasterdelete.Enabled = True
        Bgroup.Enabled = True
      
    End If
If Me.SStab1.Tab = 5 Then
        For I = 0 To 5
            If I <> 5 Then
                SStab1.TabEnabled(I) = False
            End If
        Next
        For Each ctl In Me.Controls
            If UCase(ctl.CONTAINER.Name) = UCase("manufacture") Then
                If TypeOf ctl Is textbox Or TypeOf ctl Is ComboBox Or TypeOf ctl Is CheckBox Or TypeOf ctl Is ListBox Then
                    ctl.Enabled = True
                End If
            End If
        Next
        Me.Commandmasteradd.Enabled = False
        Me.Commandmasteredit.Enabled = False
        Me.Commandmasterabandon.Enabled = True
        Me.Commandmastersave.Enabled = True
        Me.Commandmasterdelete.Enabled = True
    End If

    
    
End Sub

Private Sub CommandmasterPrint_Click()
If SStab1.Tab = 0 Then
        MainMenu.cr1.Connect = constr
        MainMenu.cr1.ReportFileName = strrptpath & "\rEPORTS\ITEMLIST.RPT"
        MainMenu.cr1.SelectionFormula = "{books.fyear}='" & main.session & "' AND {books.SETUPID}=" & main.setupid
        MainMenu.cr1.WindowShowPrintBtn = True
        MainMenu.cr1.WindowShowPrintSetupBtn = True
        MainMenu.cr1.WindowState = crptMaximized
        MainMenu.cr1.Action = 1
ElseIf SStab1.Tab = 1 Then
        MainMenu.cr1.Connect = constr
        MainMenu.cr1.ReportFileName = strrptpath & "\rEPORTS\ITEMGROUPLIST.RPT"
        MainMenu.cr1.SelectionFormula = "{GROUPS.fyear}='" & main.session & "' AND {GROUPS.SETUPID}=" & main.setupid
        MainMenu.cr1.WindowShowPrintBtn = True
        MainMenu.cr1.WindowShowPrintSetupBtn = True
        MainMenu.cr1.WindowState = crptMaximized
        MainMenu.cr1.Action = 1
ElseIf SStab1.Tab = 2 Then
        MainMenu.cr1.Connect = constr
        MainMenu.cr1.ReportFileName = strrptpath & "\rEPORTS\DISTRICTLIST.RPT"
        MainMenu.cr1.SelectionFormula = "{DISTRICTS.fyear}='" & main.session & "'  AND {DISTRICTS.SETUPID}=" & main.setupid
        MainMenu.cr1.WindowShowPrintBtn = True
        MainMenu.cr1.WindowShowPrintSetupBtn = True
        MainMenu.cr1.WindowState = crptMaximized
        MainMenu.cr1.Action = 1
ElseIf SStab1.Tab = 3 Then
        MainMenu.cr1.Connect = constr
        MainMenu.cr1.ReportFileName = strrptpath & "\rEPORTS\AGENTLIST.RPT"
        MainMenu.cr1.SelectionFormula = "{agentmaster.fyear}='" & main.session & "' AND {AGENTMASTER.SETUPID}=" & main.setupid
        MainMenu.cr1.WindowShowPrintBtn = True
        MainMenu.cr1.WindowShowPrintSetupBtn = True
        MainMenu.cr1.WindowState = crptMaximized
        MainMenu.cr1.Action = 1
End If
End Sub

Private Sub CommandmasterReturn_Click()
''''MainMenu.Toolbar1.Visible = True
Unload Me

End Sub

Private Sub Commandmastersave_Click()
    
    
    If SStab1.Tab = 0 Then
        'If Me.Textbbookcode.Text <> "" And Me.Textbbookname <> "" And Me.Combobgroupcode <> "" And Me.Textbrate <> "" Then
        If Me.Textbbookcode.Text <> "" And Combobgroupcode.Text <> "" And Me.Textbbookname <> "" Then
            RS.Open "select * from books where " & stringyear & " ", CON, adOpenKeyset, adLockOptimistic, adCmdText
            
            If addmaster = True Then
                
                If Not RS.BOF Then
                    RS.MoveFirst
                End If
                
                RS.Find "bookcode='" + Combobgroupcode.Text & Trim(Me.Textbbookcode.Text) + "'"
                If Not RS.EOF Then
                    MsgBox "ALREADY EXIST....."
                    'On Error Resume Next
                    bookmaster.booksmaster.Enabled = True
                    bookmaster.Textbbookcode.Enabled = True
                    Textbbookcode.SetFocus
                Else
                    For J = 0 To UBound(arycname)
                    '       If complist.Selected(I) = True Then
                            RS.AddNew
                            RS(0) = Combobgroupcode.Text & Trim(Me.Textbbookcode.Text)
                            RS(1) = Trim(Me.Textbbookname.Text)
                            RS!Size1 = Val(Trim(Me.txtSize1.Text))
                            RS!Size2 = Val(Trim(Me.txtSize2.Text))
                            RS!unit1 = Trim(Me.txtunit1.Text)
                            RS!unit2 = Trim(Me.txtunit2.Text)
                            RS!quality = Trim(Me.txtQuality.Text)
                            RS!rate = Val(Trim(Me.Textbrate.Text))
                            RS!per = Trim(Me.txtper.Text)
                            'If cboready.ListCount = 2 Then
                            'rs!groupcode = "Yes"
                            'Else
                            RS!groupcode = cboready.Text
                            'End If
                            RS!fyear = main.session: RS!setupid = Val(Left(arycname(J), InStr(arycname(J), " (")))
                            RS!createdby = main.username
                            RS!createdon = Now
                            RS.update
                     Next
                            Me.Textbbookcode.Text = ""
                            Me.Textbbookname.Text = ""
                            Me.txtSize1.Text = ""
                            Me.txtSize2.Text = ""
                            Me.txtunit1.Text = ""
                            Me.txtunit2.Text = ""
                            Me.txtQuality.Text = ""
                            Me.Textbrate.Text = ""
                            Me.txtper.Text = ""
                    '    End If
                    
                End If
            
            
            
            Else
            
                If Not RS.BOF Then
                    RS.MoveFirst
                End If
                RS.Find "bookcode='" + Combobgroupcode.Text & (Me.Textfindbookcode.Text) + "'"
                If RS.EOF Then
                    MsgBox "NOT EXIST...."
                Else
                    
                    
                    
                    
''''                    rs(0) = Trim(Me.Textbbookcode.Text)
''''                    rs(1) = Trim(Me.Textbbookname.Text)
''''                    rs(2) = Trim(Me.Combobgroupcode.Text)
''''                    rs(3) = Val(Trim(Me.Textbrate.Text))
''''                    rs(4) = Val(Trim(Me.Textbdiscount.Text))
                    
                    
                    RS(0) = Trim(Combobgroupcode.Text) & (Me.Textbbookcode.Text)
                    RS(1) = Trim(Me.Textbbookname.Text)
                    RS!Size1 = Trim(Me.txtSize1.Text)
                    RS!Size2 = Trim(Me.txtSize2.Text)
                    RS!unit1 = Trim(Me.txtunit1.Text)
                    RS!unit2 = Trim(Me.txtunit2.Text)
                    RS!quality = Trim(Me.txtQuality.Text)
                    RS!rate = Val(Trim(Me.Textbrate.Text))
                    RS!per = Trim(Me.txtper.Text)
                    RS!groupcode = cboready.Text
                    RS!updatedby = main.username
                    RS!updatedon = Now
                    RS.update
                    
                    Me.Textbbookcode.Text = ""
                    Me.Textbbookname.Text = ""
                    Me.txtSize1.Text = ""
                    Me.txtSize2.Text = ""
                    Me.txtunit1.Text = ""
                    Me.txtunit2.Text = ""
                    Me.txtQuality.Text = ""
                    Me.Textbrate.Text = ""
                    Me.txtper.Text = ""
                    
'''                    Me.Textbbookcode.Text = ""
'''                    Me.Textbbookname.Text = ""
'''                    Me.Combobgroupcode.Text = ""
'''                    Me.Combobgroupname = ""
'''                    Me.Textbrate.Text = ""
'''                    Me.Textbdiscount.Text = ""
                End If
            End If
            RS.close
        End If
    End If

'******************************
'    book group
'/*****************************
    If SStab1.Tab = 1 Then
    
        If Me.Textbggroupcode.Text <> "" And Me.Textbggroupname <> "" Then
        If RS.State = adStateOpen Then RS.close
            RS.Open "select * from groups where " & stringyear & " ", CON, adOpenKeyset, adLockOptimistic, adCmdText
            If addmaster = True Then
                If Not RS.BOF Then
                    RS.MoveFirst
                End If
                RS.Find "groupcode='" + Trim(Me.Textbggroupcode.Text) + "'"
                If Not RS.EOF Then
                    MsgBox "ALREADY EXIST....."
                Else
                    For J = 0 To UBound(arycname)
                    RS.AddNew
                    RS!groupcode = Trim(Me.Textbggroupcode.Text)
                    RS!GroupName = Trim(Me.Textbggroupname.Text)
                    RS!group1 = 0
                    RS!group2 = 0
                    RS!fyear = main.session
                    RS!setupid = Val(Left(arycname(J), InStr(arycname(J), " (")))
                    RS!createdby = main.username
                    RS!createdon = Now
                    RS.update
                    Next
                    Me.Textbggroupcode.Text = ""
                    Me.Textbggroupname.Text = ""
                End If
            Else
                If Not RS.BOF Then
                    RS.MoveFirst
                End If
                RS.Find "groupcode='" + Trim(Me.textbgfindcode.Text) + "'"
                If RS.EOF Then
                    MsgBox "NOT EXIST...."
                Else
                    RS(0) = Trim(Me.Textbggroupcode.Text)
                    RS(1) = Trim(Me.Textbggroupname.Text)
                    RS!updatedby = main.username
                    RS!updatedon = Now
                    RS.update
                    Me.Textbggroupcode.Text = ""
                    Me.Textbggroupname.Text = ""
                End If
            End If
            RS.close
        End If
    End If
'******************************
'/*****************************
    If SStab1.Tab = 2 Then
        If Trim(Me.Textddistrictname.Text) <> "" And Trim(txtCode.Text) <> "" Then
            If addmaster = True Then
                RS.Open "select * from DISTRICTS where  " & stringyear & " ", CON, adOpenKeyset, adLockOptimistic, adCmdText
                If Not RS.BOF Then
                    RS.MoveFirst
                End If
                RS.Find "DISTCODE='" & Trim(txtCode.Text) & "'"
                If Not RS.EOF Then
                    MsgBox "District Code Already Exist....."
                    txtCode.SetFocus
                    Exit Sub
                End If
                RS.MoveFirst
                RS.Find "DISTRICTNAME='" & Trim(Textddistrictname.Text) & "'"
                If Not RS.EOF Then
                    MsgBox "District Name Already Exist....."
                    Textddistrictname.SetFocus
                    Exit Sub
                Else
                    For J = 0 To UBound(arycname)
                    RS.AddNew
                    RS!distcode = Trim(txtCode.Text)
                    RS!DISTRICTNAME = Trim(UCase(Me.Textddistrictname))
                    RS!agentname = " "
                    If Val(Left(arycname(J), 2)) = setupid Then
                    Me.listdistrictmaster.AddItem Trim(UCase(Me.Textddistrictname))
                    End If
                    RS!fyear = main.session
                    RS!setupid = Val(Left(arycname(J), InStr(arycname(J), " (")))
                    RS!createdby = main.username
                    RS!createdon = Now
                    RS.update
                    Next
                    Me.Textddistrictname.Text = ""
                    Me.txtCode.Text = ""
                End If
            Else
                RS.Open "select * from DISTRICTS where  " & stringyear & " and distcode<>'" & Trim(Left(listdistrictmaster.Text, InStr(1, listdistrictmaster.Text, " "))) & "' and DISTRICTNAME<>'" & Mid(listdistrictmaster.Text, InStr(1, listdistrictmaster.Text, " ") + 1) & "'", CON, adOpenKeyset, adLockOptimistic, adCmdText
                If Not RS.BOF Then
                    RS.MoveFirst
                End If
                RS.Find "DISTCODE='" & txtCode.Text & "'"
                If RS.EOF = False Then
                    MsgBox "District Code Already Exist...."
                    txtCode.SetFocus
                    Exit Sub
                End If
                RS.MoveFirst
                RS.Find "DISTRICTNAME='" & Textddistrictname.Text & "'"
                If RS.EOF = False Then
                    MsgBox "District Name Already Exist...."
                Else
                If RS.State = 1 Then RS.close
                RS.Open "select * from DISTRICTS where  " & stringyear & "", CON, adOpenKeyset, adLockOptimistic, adCmdText
                If Not RS.BOF Then
                    RS.MoveFirst
                End If
                RS.Find "DISTCODE='" & Trim(Left(listdistrictmaster.Text, InStr(1, listdistrictmaster.Text, " "))) & "'"
                If RS.EOF Then
                    MsgBox "NOT EXIST...."
                Else
                    RS!distcode = Trim(txtCode.Text)
                    RS!DISTRICTNAME = Trim(UCase(Me.Textddistrictname))
                    RS!updatedby = main.username
                    RS!updatedon = Now
                    RS.update
                    Me.Textddistrictname.Text = ""
                    Me.txtCode.Text = ""
                    RS.close
                    Me.listdistrictmaster.Clear
                    RS.Open "select * from DISTRICTS where  " & stringyear & " ", CON, adOpenKeyset, adLockOptimistic, adCmdText
                    RS.MoveFirst
                    Do While Not RS.EOF
                        Me.listdistrictmaster.AddItem RS!distcode & " " & RS!DISTRICTNAME
                        'Me.listdistrictmaster.AddItem rs!DISTRICTNAME
                        If Not RS.EOF Then
                            RS.MoveNext
                        End If
                    Loop
                    End If
                End If
            End If
            RS.close
        End If
    End If

'*************************************
'                  Agent master
'*************************************


If SStab1.Tab = 3 Then
  
Me.Agent.Enabled = True
  Dim I As Integer
  Dim Drs1 As New ADODB.Recordset
  
  If Me.comboAgentMaster.Text <> "" Then
     If RS.State = 1 Then RS.close
     RS.Open "select * from AgentMaster where " & stringyear & " ", CON, adOpenKeyset, adLockPessimistic, adCmdText
     If addmaster = True Then
        If Not RS.BOF Then RS.MoveFirst
        RS.Find "Agentname='" + Trim(UCase(Me.comboAgentMaster.Text)) + "'"
        If Not RS.EOF Then
            MsgBox "ALREADY EXIST....."
        Else
            For J = 0 To UBound(arycname)
            RS.AddNew
            RS!agentname = comboAgentMaster.Text
            RS!add1 = Me.aadd1.Text
            RS!add2 = Me.aadd2.Text
            RS!city = Me.acity.Text
            RS!phone = Me.aphone.Text
            RS!fyear = main.session: RS!setupid = Val(Left(arycname(J), InStr(arycname(J), " (")))
            RS!createdby = main.username
            RS!createdon = Now
            RS.update
            Next
        End If
        
        
        'If rs.State = 1 Then rs.Close
        'rs.Open "Districts", CON, adOpenKeyset, adLockPessimistic, adcmdtext
        'If Not rs.BOF Then rs.MoveFirst
        'rs.Find "Agentname='" + Trim(UCase(Me.comboAgentMaster.Text)) + "'"
        'If Not rs.EOF Then
         '   MsgBox "ALREADY EXIST....."
        'Else
            'For I = 0 To ListDis2.ListCount - 1
                 'If Drs1.State = 1 Then Drs1.Close
                 'Drs1.Open "Districts", CON, adOpenDynamic, adLockOptimistic, adcmdtext
                 'Drs1.Find "Districtname='" + Trim(UCase(Me.ListDis2.List(I))) + "'"
                 'If Not Drs1.EOF Then
                    'Drs1!Agentname = Trim(UCase(Me.comboAgentMaster.Text))
                    'Drs1.Update
                  'End If
              'Next I
              'For I = 0 To ListDis1.ListCount - 1
                      'CON.Execute "Update districts Set Agentname = '' where   " & stringyear & " and districtname= '" & ListDis1.List(I) & "'"
              'Next I
              'Me.comboAgentMaster.Text = ""
              'ListDis2.Clear
              'ListDis1.Clear
         'End If
    Else
        If Not RS.BOF Then RS.MoveFirst
        RS.Find "Agentname='" + Trim(UCase(Me.TextfindAgentmaster.Text)) + "'"
        If Not RS.EOF Then
             'CON.Execute "Delete from agentmaster where Agentname='" + Trim(UCase(Me.TextfindAgentmaster.Text)) + "' and " & stringyear & " "
             'rs.AddNew
             RS!agentname = comboAgentMaster.Text
             RS!add1 = Me.aadd1.Text
                RS!add2 = Me.aadd2.Text
                RS!city = Me.acity.Text
                RS!phone = Me.aphone.Text
                RS!updatedby = main.username
                RS!updatedon = Now
                   
             RS.update
        End If
        
        
        
        
        'For I = 0 To ListDis2.ListCount - 1
         '   If rs.State = 1 Then rs.Close
          '  rs.Open "Districts", CON, adOpenDynamic, adLockOptimistic, adcmdtext
           ' rs.Find "Districtname='" + Trim(UCase(Me.ListDis2.List(I))) + "'"
            'If Not rs.EOF Then
             '  rs!Agentname = Trim(UCase(Me.comboAgentMaster.Text))
               
              ' rs.Update
            'End If
        'Next I
        'For I = 0 To ListDis1.ListCount - 1
         '    CON.Execute "Update districts Set Agentname = '' where   " & stringyear & " and districtname = '" & ListDis1.List(I) & "'"
        'Next I
        'If rs.State = 1 Then rs.Close
        'rs.Open "Select distinct Agentname from Districts where  " & stringyear & " ", CON, adOpenKeyset, adLockOptimistic, adCmdText
        'rs.MoveFirst
        'comboAgentMaster.Clear
        'Do While Not rs.EOF
         '    If IsNull(rs(0)) = False Then Me.comboAgentMaster.AddItem rs(0)
             'rs.MoveNext
        'Loop
    
    
    End If

End If

'ListDis1.Clear
'ListDis2.Clear
Me.Agent.Enabled = False
Me.Addcd.Enabled = False
Me.Removecd.Enabled = False
    Commandmasterabandon_Click
    If RS.State = 1 Then RS.close
End If

If SStab1.Tab = 4 Then
   CON.Execute "Update groups Set group1 = false where " & stringyear & " "
   CON.Execute "Update groups Set group2 = false where " & stringyear & " "
   For I = 0 To List2.ListCount - 1
       CON.Execute "Update groups Set group1 = true where  groupcode = '" & List2.List(I) & "' and " & stringyear & " "
   Next I
   
   If List2.ListCount <= 0 Then
         CON.Execute "Update groups Set group1 = false where " & stringyear & " "
   End If
   
   For I = 0 To List3.ListCount - 1
       CON.Execute "Update groups Set group2 = true where  groupcode = '" & List3.List(I) & "' and " & stringyear & " "
       
   
  Next I
   
   If List3.ListCount <= 0 Then
         CON.Execute "Update groups Set group2 = false where " & stringyear & " "
   End If
   
   
   List1.Clear
   List2.Clear
   List3.Clear
   Me.Bgroup.Enabled = False
End If
'******************************
'/*****************************
    If SStab1.Tab = 5 Then
        If txtmname.Text <> "" Then
            If addmaster = True Then
                RS.Open "select * from manufacture where  " & stringyear & " ", CON, adOpenKeyset, adLockOptimistic, adCmdText
                If Not RS.BOF Then
                    RS.MoveFirst
                End If
                RS.Find "mname='" & Trim(txtmname.Text) & "'"
                If Not RS.EOF Then
                    MsgBox "Manufacture Name Already Exist....."
                    txtmname.SetFocus
                    Exit Sub
                Else
                    For J = 0 To UBound(arycname)
                    RS.AddNew
                    RS!mname = Trim(txtmname.Text)
                    If Val(Left(arycname(J), 2)) = setupid Then
                    Me.listmanufacture.AddItem Trim(UCase(Me.txtmname))
                    End If
                    RS!fyear = main.session
                    RS!setupid = Val(Left(arycname(J), InStr(arycname(J), " (")))
                    RS!createdby = main.username
                    RS!createdon = Now
                    RS.update
                    Next
                    Me.txtmname.Text = ""
                End If
            Else
                RS.Open "select * from manufacture where  " & stringyear & " and mname<>'" & Trim(listmanufacture.Text) & "'", CON, adOpenKeyset, adLockOptimistic, adCmdText
                If Not RS.BOF Then
                    RS.MoveFirst
                End If
                RS.Find "mname='" & txtmname.Text & "'"
                If RS.EOF = False Then
                    MsgBox "Manufacture Name Already Exist...."
                    txtmname.SetFocus
                    Exit Sub
                Else
                If RS.State = 1 Then RS.close
                RS.Open "select * from manufacture where  " & stringyear & "", CON, adOpenKeyset, adLockOptimistic, adCmdText
                If Not RS.BOF Then
                    RS.MoveFirst
                End If
                RS.Find "mname='" & Trim(listmanufacture.Text) & "'"
                If RS.EOF Then
                    MsgBox "NOT EXIST...."
                Else
                    RS!mname = Trim(txtmname.Text)
                    RS!updatedby = main.username
                    RS!updatedon = Now
                    RS.update
                    Me.txtmname.Text = ""
                    RS.close
                    Me.listmanufacture.Clear
                    RS.Open "select * from manufacture where  " & stringyear & " ", CON, adOpenKeyset, adLockOptimistic, adCmdText
                    RS.MoveFirst
                    Do While Not RS.EOF
                        Me.listmanufacture.AddItem RS!mname
                        If Not RS.EOF Then
                            RS.MoveNext
                        End If
                    Loop
                    End If
                End If
            End If
            RS.close
        End If
    End If


End Sub

Private Sub Commandmastersearch_Click()
      
    Me.Enabled = False
   ' searchscreen.Grid1.row = 0
    'searchscreen.Grid1.col = 0
    Call searchscreen.tempr(Me.SStab1.Tab, Me.Name)
End Sub



Private Sub Form_Activate()
'bookmaster.Commandmasteradd.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub Form_Load()
' /****      FRAMEINI      ****/
    Me.Top = 20
    Me.Left = 200
    Dim TMPA As Control
    For Each TMPA In Me.Controls
        If TypeOf TMPA Is frame Then
            TMPA.Top = 800
            TMPA.Left = 800
            TMPA.Width = 7515
            TMPA.Height = 4005
            
            If SStab1.Tab = 4 Then
            
                 TMPA.Top = 800
                 TMPA.Left = 210
                 TMPA.Width = 9000
                 TMPA.Height = 4005
            
            End If
        End If
        If TypeOf TMPA Is textbox Then
            TMPA.Enabled = False
        End If
        If TypeOf TMPA Is CheckBox Then
            TMPA.Enabled = False
        End If
        If TypeOf TMPA Is ComboBox Then
            TMPA.Enabled = False
        End If
    Next
    Set RS = New ADODB.Recordset
    RS.Open "select * from  DISTRICTS  where  agentname='' and " & stringyear & " ", CON, adOpenKeyset, adLockReadOnly
    If Not RS.EOF Then
        Do While Not RS.EOF
            ListDis1.AddItem RS!DISTRICTNAME
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.close
   
    RS.Open "select * from  DISTRICTS where " & stringyear & " ORDER BY DISTRICTNAME", CON, adOpenKeyset, adLockReadOnly
    If Not RS.EOF Then
        Do While Not RS.EOF
        'listdistrictmaster.AddItem RS!distcode & " " & RS!DISTRICTNAME
        listdistrictmaster.AddItem RS!DISTRICTNAME
        If Not RS.EOF Then
            RS.MoveNext
        End If
        Loop
    End If
    RS.close
    
'''    RS.Open "select * from  manufacture where " & stringyear & " ORDER BY mname", CON, adOpenKeyset, adLockReadOnly
'''    listmanufacture.Clear
'''    If Not RS.EOF Then
'''        Do While Not RS.EOF
'''         listmanufacture.AddItem RS!mname
'''            If Not RS.EOF Then
'''                RS.MoveNext
'''            End If
'''        Loop
'''    End If
'''    RS.close
    
    
    '*******Agent master combo fill
    RS.Open "select Agentname from Agentmaster where  " & stringyear & " order by agentname", CON, adOpenKeyset, adLockReadOnly, adCmdText
    If Not RS.EOF Then
        Do While Not RS.EOF
          If IsNull(RS(0)) = False Then
            Me.comboAgentMaster.AddItem RS(0)
          End If
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.close

    
    
    '/ ***** Combobgroupcode
    RS.Open "select * from GROUPS where " & stringyear & " ", CON, adOpenKeyset, adLockReadOnly, adCmdText
    Combobgroupcode.Clear
    If Not RS.EOF Then
        Do While Not RS.EOF
            Me.Combobgroupcode.AddItem RS(0)
            Me.Combobgroupname.AddItem RS(1)
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If

    RS.close
       
    
   Commandmastersearch.Enabled = True
   SetButton Commandmasteradd, Commandmasteredit, Commandmastersave, Commandmasterdelete
End Sub

Private Sub List1_Click()
'cmdm1_Click
End Sub

Private Sub List1_DblClick()
'cmdm3_Click
End Sub

Private Sub List2_DblClick()
cmdm2_Click
End Sub

Private Sub List3_DblClick()
cmdm4_Click
End Sub


Private Sub ListDis1_DblClick()
Addcd_Click
End Sub

Private Sub ListDis2_DblClick()
Removecd_Click
End Sub

'Private Sub Textglyearopeningbalance_KeyPress(KeyAscii As Integer)
 '   If KeyAscii >= 48 And KeyAscii <= 57 Then
  '
   ' Else
   '     If KeyAscii <> 46 Then
   '         KeyAscii = 0
   '     End If
  '  End If
'End Sub



Private Sub listdistrictmaster_Click()
Me.txtCode.Text = Trim(VBA.Left(Me.listdistrictmaster.Text, InStr(1, Me.listdistrictmaster.Text, " ")))
Me.Textddistrictname.Text = Mid(Me.listdistrictmaster.Text, InStr(1, Me.listdistrictmaster.Text, " ") + 1)
'Me.Textddistrictname.Text = Me.listdistrictmaster.Text

End Sub

Private Sub okcd_Click()

End Sub

Private Sub listmanufacture_Click()
txtmname.Text = listmanufacture.Text
End Sub

Private Sub Removecd_Click()
If ListDis2.ListCount > 0 Then
    Dim I
    Dim delitem As Integer
    For I = 0 To ListDis2.ListCount - 1
        If ListDis2.Selected(I) Then
                ListDis1.AddItem ListDis2.List(I)
                delitem = I
         End If
    Next
    ListDis2.RemoveItem delitem
End If


End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SStab1.Tab = 2 Then
        Me.Commandmastersearch.Enabled = False
    Else
        Me.Commandmastersearch.Enabled = True
    End If
End Sub

Private Sub SSTab1_DblClick()
    If SStab1.Tab = 2 Then
        Me.Commandmastersearch.Enabled = False
    Else
        Me.Commandmastersearch.Enabled = True
    End If
End Sub


Private Sub SSTab1_GotFocus()
If SStab1.Tab = 4 Then
Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset
rs1.Open "Select groupcode from groups where group1= 0 and group2=0 and " & stringyear & " ", CON, adOpenStatic, adLockOptimistic
If rs1.RecordCount > 0 Then
       rs1.MoveFirst
       While Not rs1.EOF
           List1.AddItem rs1!groupcode
           rs1.MoveNext
       Wend
End If
If rs1.State = 1 Then rs1.close

rs1.Open "Select groupcode from groups where group1= 1 and group2=0 and " & stringyear & " ", CON, adOpenStatic, adLockOptimistic
If rs1.RecordCount > 0 Then
       rs1.MoveFirst
       While Not rs1.EOF
           List2.AddItem rs1!groupcode
           rs1.MoveNext
       Wend
End If

If rs1.State = 1 Then rs1.close
rs1.Open "Select groupcode from groups where group1= 0 and group2=1 and " & stringyear & " ", CON, adOpenStatic, adLockOptimistic
If rs1.RecordCount > 0 Then
       rs1.MoveFirst
       While Not rs1.EOF
           List3.AddItem rs1!groupcode
           rs1.MoveNext
       Wend
End If



End If





End Sub

Private Sub Textbbookcode_GotFocus()
    'Combocode.Visible = True
    'Combocode.ZOrder
    'Combocode.SetFocus
    
End Sub

Private Sub Textbbookcode_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then SendKeys "{TAB}"

End Sub

Private Sub Textbbookcode_LostFocus()
' If Textbbookcode.Text = "" Then
'    MsgBox "Please Enter Item code"
'        Me.Textbbookcode.SetFocus
'
'Exit Sub
' End If

End Sub

Private Sub Textbbookname_KeyPress(KeyAscii As Integer)

If Textbbookname.Text = "" And KeyAscii = 13 Then
MsgBox "Please Enter Item Name"
Me.Textbbookname.SetFocus
Exit Sub
 End If


'If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Textbbookname_LostFocus()
'''If Textbbookname.Text = "" Then
'''MsgBox "Please Enter Item Name"
'''Me.Textbbookname.SetFocus
'''Exit Sub
''' End If
End Sub

Private Sub Textbdiscount_KeyPress(KeyAscii As Integer)
    
        If KeyAscii = 13 Then bookmaster.Commandmastersave.SetFocus
    
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        
    Else
        If KeyAscii <> 46 Then
            If KeyAscii <> 8 Then
                KeyAscii = 0
            End If
        End If
    End If
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Textbggroupcode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Textbggroupcode_LostFocus()
Textbggroupcode.Text = UCase(Textbggroupcode.Text)
End Sub

Private Sub Textbggroupname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then bookmaster.Commandmastersave.SetFocus
End Sub

Private Sub Textbrate_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        
    Else
        If KeyAscii <> 46 Then
            If KeyAscii <> 8 Then
                KeyAscii = 0
            End If
        End If
    End If
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Textddistrictname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Textddistrictname_LostFocus()
  Textddistrictname.Text = UCase(Textddistrictname.Text)
End Sub

