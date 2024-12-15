VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Voucherform 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   10056
   ClientLeft      =   276
   ClientTop       =   1716
   ClientWidth     =   16092
   ClipControls    =   0   'False
   Icon            =   "Voucher.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   10056
   ScaleWidth      =   16092
   Begin VB.Frame panel 
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
      Height          =   9825
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   13920
      Begin VB.TextBox txtchecked 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   2388
         MaxLength       =   100
         TabIndex        =   38
         Top             =   7776
         Width           =   540
      End
      Begin VB.TextBox txtpaytype 
         Height          =   285
         Left            =   12312
         TabIndex        =   36
         Top             =   324
         Visible         =   0   'False
         Width           =   240
      End
      Begin MSMask.MaskEdBox crno1 
         Height          =   370
         Left            =   1125
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1440
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1778
         _ExtentY        =   656
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin VB.TextBox crno 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   28
         Top             =   1440
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.ComboBox Gencombo 
         Height          =   1872
         Left            =   540
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   24
         Top             =   1872
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.ComboBox SubCombo 
         CausesValidation=   0   'False
         Height          =   1872
         ItemData        =   "Voucher.frx":000C
         Left            =   3180
         List            =   "Voucher.frx":000E
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   23
         Top             =   1845
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   915
         Left            =   150
         ScaleHeight     =   912
         ScaleWidth      =   8880
         TabIndex        =   8
         Top             =   8565
         Width           =   8880
         Begin VB.CommandButton Commandmasterhelp 
            Caption         =   "Help"
            Height          =   375
            Left            =   -1290
            TabIndex        =   17
            Top             =   0
            Width           =   800
         End
         Begin VB.CommandButton Commandmasteradd 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Add"
            Height          =   735
            Left            =   90
            Picture         =   "Voucher.frx":0010
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   120
            Width           =   1065
         End
         Begin VB.CommandButton Commandedit 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Edit"
            Height          =   735
            Left            =   1170
            Picture         =   "Voucher.frx":0BF4
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   120
            Width           =   1104
         End
         Begin VB.CommandButton Commandsave 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Sa&ve"
            Enabled         =   0   'False
            Height          =   735
            Left            =   2300
            Picture         =   "Voucher.frx":1036
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   108
            Width           =   1065
         End
         Begin VB.CommandButton Commandabandon 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Aba&ndon"
            Height          =   735
            Left            =   3390
            Picture         =   "Voucher.frx":1C1A
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   120
            Width           =   1065
         End
         Begin VB.CommandButton Commanddelete 
            BackColor       =   &H00FFFFFF&
            Caption         =   "De&lete"
            Height          =   735
            Left            =   4480
            Picture         =   "Voucher.frx":21A4
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   120
            Width           =   1065
         End
         Begin VB.CommandButton Commandsearch 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Search"
            Height          =   735
            Left            =   5570
            Picture         =   "Voucher.frx":2D88
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   120
            Width           =   1065
         End
         Begin VB.CommandButton CommandPrint 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Print"
            Height          =   735
            Left            =   6660
            Picture         =   "Voucher.frx":396C
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   120
            Width           =   1065
         End
         Begin VB.CommandButton CommandReturn 
            BackColor       =   &H00FFFFFF&
            Caption         =   "E&xit"
            Height          =   735
            Left            =   7755
            Picture         =   "Voucher.frx":4550
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   120
            Width           =   1065
         End
      End
      Begin VB.ComboBox vtype 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         ItemData        =   "Voucher.frx":5134
         Left            =   1560
         List            =   "Voucher.frx":5141
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   360
         Width           =   960
      End
      Begin VB.CommandButton Command1 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   825
         TabIndex        =   5
         Top             =   7605
         Width           =   675
      End
      Begin VB.CommandButton Command2 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   105
         TabIndex        =   4
         Top             =   7605
         Width           =   675
      End
      Begin VB.TextBox totaldebit 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6765
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   7635
         Width           =   1416
      End
      Begin VB.TextBox totalcredit 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8235
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   7635
         Width           =   1416
      End
      Begin VB.ComboBox FindCombo 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   7.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   4284
         Left            =   2010
         Style           =   1  'Simple Combo
         TabIndex        =   1
         Top             =   1620
         Visible         =   0   'False
         Width           =   2835
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid1 
         Height          =   6432
         Left            =   48
         TabIndex        =   6
         Top             =   1032
         Width           =   13548
         _ExtentX        =   23897
         _ExtentY        =   11345
         _Version        =   393216
         Cols            =   5
         BackColorFixed  =   7917545
         GridColorFixed  =   8438015
         MergeCells      =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSMask.MaskEdBox vno 
         Height          =   348
         Left            =   7020
         TabIndex        =   20
         Top             =   300
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   614
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox vdate 
         Height          =   384
         Left            =   4368
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   300
         Width           =   1248
         _ExtentX        =   2201
         _ExtentY        =   677
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medit 
         Height          =   285
         Left            =   570
         TabIndex        =   19
         Top             =   3705
         Width           =   3105
         _ExtentX        =   5461
         _ExtentY        =   508
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox debit 
         Height          =   285
         Left            =   2010
         TabIndex        =   21
         Top             =   5055
         Width           =   1185
         _ExtentX        =   2074
         _ExtentY        =   508
         _Version        =   393216
         Appearance      =   0
         AutoTab         =   -1  'True
         Format          =   "###0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox credit 
         Height          =   255
         Left            =   660
         TabIndex        =   22
         Top             =   4560
         Width           =   1155
         _ExtentX        =   2032
         _ExtentY        =   445
         _Version        =   393216
         Appearance      =   0
         AutoTab         =   -1  'True
         Format          =   "###0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox DESCRIPTION1 
         Height          =   255
         Left            =   3870
         TabIndex        =   25
         Top             =   3735
         Width           =   1155
         _ExtentX        =   2032
         _ExtentY        =   466
         _Version        =   393216
         Appearance      =   0
         AutoTab         =   -1  'True
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin VB.TextBox tmpsubledger 
         Height          =   285
         Left            =   3510
         TabIndex        =   27
         Top             =   3465
         Width           =   3045
      End
      Begin VB.TextBox DESCRIPTION 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   930
         MaxLength       =   50
         TabIndex        =   26
         Top             =   5055
         Width           =   1725
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Checked :"
         Height          =   252
         Left            =   1620
         TabIndex        =   39
         Top             =   7812
         Width           =   684
      End
      Begin MSForms.Label Label5 
         Height          =   372
         Left            =   108
         TabIndex        =   37
         Top             =   8136
         Visible         =   0   'False
         Width           =   336
         BackColor       =   16777215
         Caption         =   "R"
         Size            =   "593;656"
         MousePointer    =   14
         FontEffects     =   1073741829
         FontHeight      =   240
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CheckBox CheckBox1_createDesc 
         Height          =   372
         Left            =   8352
         TabIndex        =   35
         Top             =   276
         Visible         =   0   'False
         Width           =   3840
         BackColor       =   16777215
         ForeColor       =   4210752
         DisplayStyle    =   4
         Size            =   "6773;661"
         Value           =   "0"
         Caption         =   "Use Bill No and Date for Created Narration Automatically"
         FontName        =   "Arial Narrow"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H0078CFE9&
         BorderWidth     =   4
         Height          =   1050
         Left            =   120
         Top             =   8520
         Width           =   8970
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher Type"
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
         Left            =   210
         TabIndex        =   33
         Top             =   390
         Width           =   1275
      End
      Begin VB.Label Total 
         Caption         =   "Total :- "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5985
         TabIndex        =   32
         Top             =   7635
         Width           =   825
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5895
         TabIndex        =   31
         Top             =   360
         Width           =   1365
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher Date"
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
         Left            =   3105
         TabIndex        =   30
         Top             =   330
         Width           =   1275
      End
      Begin VB.Label Label4 
         Caption         =   "Press F2 for Search A Voucher"
         Height          =   255
         Left            =   3405
         TabIndex        =   29
         Top             =   7665
         Width           =   2475
      End
   End
End
Attribute VB_Name = "Voucherform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim CON As ADODB.Connection
Dim rs1 As ADODB.Recordset
Dim RS As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rs2_ As ADODB.Recordset
Dim cmd As New ADODB.Command
Dim lastrow, lastcol As Integer
Dim maxrow As Integer
Dim clicksave As Boolean
Dim tmplistindex As Long
Dim Edit As Boolean
Dim addmode As Boolean
Dim keyEdit As Boolean
Dim rs_S As New ADODB.Recordset
Dim searchmode As Boolean
Dim varVtype As String
Dim varVdate As Variant
Dim VarVno As Integer
Dim Line1 As Integer
Dim dr_New, cr_New As Double
Dim flagsearch As Boolean
Sub refreshTotal()

  totalcredit.text = ""
  totaldebit.text = ""
  For I = 1 To grid1.rows - 1
     totaldebit.text = Val(totaldebit.text) + Val(grid1.TextMatrix(I, 2))
     totalcredit.text = Val(totalcredit.text) + Val(grid1.TextMatrix(I, 3))
  Next
  totaldebit.text = Format(totaldebit.text, "0.00")
  totalcredit.text = Format(totalcredit.text, "0.00")



End Sub

Private Sub Command1_Click()

'
'
'addmode = False
'vtype.Enabled = False
'vdate.Enabled = False
'vno.Enabled = False
'Commandsearch.Enabled = True
'grid1.Enabled = False
'searchmode = True
'If Not rs2.EOF Then
'     rs2.MoveNext
'     If rs2.EOF And rs2.RecordCount > 0 Then
'           rs2.MoveLast
'           Exit Sub
'     End If
'     totaldebit.Text = ""
'     totalcredit.Text = ""
'     vtype = rs2!VoucherType
'     vdate = rs2!VoucherDate
'     vno = rs2!VoucherNumber
'     Me.vtype_LostFocus
'     Me.vdate_LostFocus
'     Me.vno_LostFocus
'
'
'
'End If
'


End Sub
Private Sub Command2_Click()

'If Not rs2.BOF Then
'    addmode = False
'    searchmode = True
'
'    rs2.MovePrevious
'    If rs2.BOF And rs2.RecordCount > 0 Then
'           rs2.MoveFirst
'           Exit Sub
'    End If
'     totaldebit.Text = ""
'     totalcredit.Text = ""
'
'     vtype = rs2!VoucherType
'     vdate = rs2!VoucherDate
'     vno = rs2!VoucherNumber
'     Me.vtype_LostFocus
'     Me.vdate_LostFocus
'     Me.vno_LostFocus
'
'End If


End Sub
Sub Commandabandon_Click()


'If MsgBox("Are you Sure ? want to Abandon ...", vbQuestion + vbYesNo) = vbYes Then
'
'    If addmode = True Then
'       If totaldebit.Text <> "" Then
'         createLog UserName, Trim(vno.Text), "voucher" & vtype, "Abandon Without save:" & totaldebit.Text, Date
'       End If
'    End If
'
'End If



v_vtype = ""
v_vdate = ""
v_vnumber = ""


txtchecked.text = ""
'txtVNumber.text = ""

'If SAVED Then
    For I = 1 To grid1.rows - 1
        
        grid1.CellBackColor = vbWhite
    
        grid1.Row = I
          If I Mod 2 = 0 Then
               grid1.text = "."
               grid1.TextMatrix(I, 7) = ""
          Else
          For J = 0 To 6
            grid1.Col = J
            grid1.text = ""
            grid1.text = ""
            txtpaytype.text = ""
            
          
        Next J
          grid1.Col = 0
       End If
       
 
       
    Next I
    
    dr_New = 0
    cr_New = 0
   
    grid1.Refresh
    totalcredit = ""
    totaldebit = ""
    Gencombo.Visible = False
    SubCombo.Visible = False
    credit.Visible = False
    debit.Visible = False
    crno.Visible = False
    grid1.Row = 1
    grid1.Col = 0
    vno.Enabled = True
    addmode = False
    searchmode = True
    
vtype.Enabled = True
vtype.SetFocus


grid1.Enabled = False
Commandmasteradd.Enabled = True
Commandedit.Enabled = True
Commanddelete.Enabled = True
Commandsearch.Enabled = True
CommandPrint.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
mnuMenu_ = "menujournalvoucher"
SetButton Commandmasteradd, Commandedit, Commandsave, Commanddelete
Commandsave.Enabled = False
Edit = False
addmode = False



End Sub

Private Sub Commanddelete_Click()

On Error GoTo Del


If MsgBox("Are You Sure, Delete it ?", vbYesNo) = vbNo Then
        Exit Sub
Else
    
    tmpVoucher ("del")
    BackupVoucher ("del")
    
    voucherMainBK "Delete"
    
    
    
    
    If (AuditTrail = "y") Then
    
    'If (txtchecked.text = "y") Then
        'con.Execute "update VOUCHERS_bk set EntryNumber='1' where (voucherType='" + Trim(vtype) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & vdate & "',103)) And Vouchernumber = " + Trim(vno)

        actionType_ = "Delete"
        vtype1_ = "v"
        vdate_ = vdate
        vno_ = vno
        vtypeNew = vtype
        
        frmAuditTrailLog_Rem.Show 1
        
    'End If
    
    End If
    
    createLog UserName, Trim(vno.text), "voucher" & vtype, " Delete : " & totaldebit.text, Date
    
    
    
    
    con.Execute "delete  from vouchers where " & stringyear & " and  vouchertype='" + Trim(vtype.text) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & vdate & "',103) And vouchernumber = " + Trim(vno.text)
    con.Execute "delete  from VOUCHERS_Main where " & stringyear & " and  vouchertype='" + Trim(vtype.text) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & vdate & "',103) And vouchernumber = " + Trim(vno.text)
    
    
    
    
    Commandabandon_Click
               
End If


Exit Sub
Del:
MsgBox "" & err.DESCRIPTION

End Sub
Sub updateV_Main(ActionType As String, VoucherType As String, vdate As Date, vno As String, particullar_ As String, amt As Double)
   
 Dim Checked_YesNo  As Integer
   
    If (txtchecked.text = "y") Then
         Checked_YesNo = 1
    Else
         Checked_YesNo = 0
    End If
   
Dim ss_

ss_ = Right(session, 2)

If Val(ss_) >= 24 Then


Dim rs22_ As ADODB.Recordset
Set rs22_ = New ADODB.Recordset
   
'If (ActionType = "add") Then
'Else



rs22_.Open "select VoucherID from  VOUCHERS_Main where " & stringyear & " and  vouchertype='" + Trim(vtype.text) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & vdate & "',103) And vouchernumber = " + vno, con

If rs22_.EOF = False Then
        

        con.Execute "update VOUCHERS_Main set " & _
                "VoucherType='" & VoucherType & "'," & _
                "VoucherDate='" & Format(vdate, "MM/dd/yyyy") & "'," & _
                "VoucherNumber='" & vno & "'," & _
                "Amount=" & amt & "," & _
                "Particular='" & particullar_ & "'," & _
                "ModifyBy='" & UserName & "'" & _
                " where VoucherID=" & rs22_!VoucherID & ""
                
        If (AuditTrail = "y") Then
                
                con.Execute "update VOUCHERS_Main set " & _
                "Checked_YesNo='" & Checked_YesNo & "'" & _
                " where VoucherID=" & rs22_!VoucherID & ""
                
        End If
        

Else

   con.Execute "insert into VOUCHERS_Main(VoucherType,VoucherDate,VoucherNumber,setupid,fyear,userName,CreatedBy,Particular,Amount)" & _
   " values('" & VoucherType & "','" & Format(vdate, "MM/dd/yyyy") & "','" & vno & "'," & setupid & ",'" & session & "','" & UserName & "','" & UserName & "','" & particullar_ & "','" & amt & "')"


        
End If
        
End If
   
'End If
   
End Sub
Sub voucherMainBK(ActionType As String)
   
Dim ss_

ss_ = Right(session, 2)

If Val(ss_) >= 24 Then

Dim rs22_ As ADODB.Recordset
Set rs22_ = New ADODB.Recordset

rs22_.Open "select VoucherID from  VOUCHERS_Main where " & stringyear & " and  vouchertype='" + Trim(vtype.text) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & vdate & "',103) And vouchernumber = " + vno, con
If rs22_.EOF = False Then
        
con.Execute "insert into VOUCHERS_MainBk(VoucherID,VoucherType,VoucherDate,VoucherNumber,setupid,fyear,userName,Particular,Checked_YesNo,CheckedBy,Amount,ActionType) " & _
"SELECT VoucherID,VoucherType,VoucherDate,VoucherNumber,setupid,fyear,userName,Particular,Checked_YesNo,CheckedBy,Amount,'" & ActionType & "' from VOUCHERS_Main where VoucherID=" & rs22_!VoucherID & ""
    
        
End If
   
End If
   
End Sub
Private Sub Commandedit_Click()
    
   
    
    
    DoEvents
    
    If MsgBox("Are You Sure.. , Want to Edit ?", vbYesNo) = vbNo Then
      Exit Sub
    Else
      grid1.Enabled = True
      BackupVoucher ("edit")
      createLog UserName, Trim(vno.text), "voucher" & vtype, " Edit : " & totaldebit.text, Date
    End If
   
    
    searchmode = False
    Gencombo.Clear
    
    Set RS = New ADODB.Recordset
    RS.Open "Select * from gledger where " & stringyear & " order by gledger", con, adOpenDynamic, adLockReadOnly, adCmdText
    If Not RS.EOF Then
        Do While Not RS.EOF
            Gencombo.AddItem RS(1)
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.close
    Commandedit.Enabled = False
    Edit = True
    
    
    mnuMenu_ = "menujournalvoucher"
    SetButton Commandmasteradd, Commandedit, Commandsave, Commanddelete

    editView
    
    
    varVtype = Trim(vtype)
    varVdate = Format(vdate.text, "dd/mm/yyyy")
    VarVno = Val(vno)
    grid1.Enabled = True
    vtype.Enabled = True
    vdate.Enabled = True
    vtype.SetFocus
    vno.Enabled = True
    Commandsave.Enabled = True
    addmode = False
    Commandmasteradd.Enabled = False
    Commandedit.Enabled = False
    Commanddelete.Enabled = True
    Commandsearch.Enabled = False
    'CommandPrint.Enabled = False
    Command1.Enabled = False
    Command2.Enabled = False
    varVtype = Trim(vtype)
    varVdate = Format(vdate.text, "dd/mm/yyyy")
    VarVno = Val(vno)
    searchmode = False
    refreshTotal
    
    
End Sub
Private Sub Commandmasteradd_Click()
  
  On Error Resume Next
  
  searchmode = False
  Commandabandon_Click
  Dim tRS1 As New ADODB.Recordset
  
  
  If tRS1.State = 1 Then tRS1.close
  tRS1.Open "SELECT top 100 VoucherNumber FROM vouchers where " & stringyear & " order by vsno", con, adOpenStatic, adLockReadOnly
  If tRS1.RecordCount <= 0 Then
      vtype.text = "J"
      vdate = Format(Date, "dd/MM/yyyy")
  End If
    Command1.Enabled = False
    Command2.Enabled = False
    addmode = True
    Edit = False
    
    Set rs1 = New ADODB.Recordset
    Set RS = New ADODB.Recordset
    rs1.Open "Select max(VoucherNumber) from vouchers where " & stringyear & " and vouchertype='" + Trim(vtype.text) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & vdate & "',103)", con, adOpenDynamic, adLockReadOnly, adCmdText
    If Not rs1.BOF Then
        If Not rs1(0) Then
            c = rs1(0)
        Else
            c = 0
        End If
    Else
        c = 0
    End If
    rs1.close
    vno.text = c + 1
    maxrow = 0
    
    Gencombo.Width = Gencombo.Width + 180
    SubCombo.Width = SubCombo.Width + 180
    debit.Width = credit.Width
    debit.Height = credit.Height
    crno.Width = credit.Width
    crno.Height = credit.Height
    medit.Width = Gencombo.Width
    medit.Height = Gencombo.Height
    
    
    If rs1.State = 1 Then
        rs1.close
    End If
    
    
    
    grid1.Row = 0
    grid1.Col = 0
    
    grid1.SelectionMode = flexSelectionFree
    grid1.ColWidth(0, 1) = 11111
    X = grid1.MergeCol(1)
    
    DESCRIPTION.Width = grid1.Width - 300
    DESCRIPTION.Left = 200
    DESCRIPTION.Height = crno.Height
    
    vtype.Enabled = True
    vdate.Enabled = True
    vno.Enabled = True
    grid1.Enabled = True
    vtype.SetFocus
    
    Commandsave.Enabled = True
    Commandmasteradd.Enabled = False
    Commandedit.Enabled = False
    Commanddelete.Enabled = False
    Commandsearch.Enabled = False
    'CommandPrint.Enabled = False
    addmode = True
    vno.Enabled = False
    
    

End Sub
Sub addfunction()
  searchmode = False
  ''Commandabandon_Click
  Dim tRS1 As New ADODB.Recordset
  
  
'  If tRS1.State = 1 Then tRS1.close
'  tRS1.Open "SELECT top 100 VoucherNumber FROM vouchers where " & stringyear & " order by vsno", con, adOpenStatic, adLockReadOnly
'  If tRS1.RecordCount <= 0 Then
      vtype.text = "J"
      vdate = Format(Date, "dd/MM/yyyy")
 ' End If
  
  
    Command1.Enabled = False
    Command2.Enabled = False
    addmode = True
    Edit = False
    
    Dim vdate1
    
    
    
    Set rs1 = New ADODB.Recordset
    Set RS = New ADODB.Recordset
    rs1.Open "Select max(VoucherNumber) from vouchers where " & stringyear & " and vouchertype='J' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & vdate & "',103)", con, adOpenDynamic, adLockReadOnly, adCmdText
    If Not rs1.BOF Then
        If Not rs1(0) Then
            c = rs1(0)
        Else
            c = 0
        End If
    Else
        c = 0
    End If
    rs1.close
    vno.text = Str(c + 1)
    maxrow = 0
    
    Gencombo.Width = Gencombo.Width + 180
    SubCombo.Width = SubCombo.Width + 180
    debit.Width = credit.Width
    debit.Height = credit.Height
    crno.Width = credit.Width
    crno.Height = credit.Height
    medit.Width = Gencombo.Width
    medit.Height = Gencombo.Height
    
    
    If rs1.State = 1 Then
        rs1.close
    End If
    
    
    
    grid1.Row = 0
    grid1.Col = 0
    
    grid1.SelectionMode = flexSelectionFree
    grid1.ColWidth(0, 1) = 11111
    X = grid1.MergeCol(1)
    
    DESCRIPTION.Width = grid1.Width - 300
    DESCRIPTION.Left = 200
    DESCRIPTION.Height = crno.Height
    
    vtype.Enabled = True
    vdate.Enabled = True
    vno.Enabled = True
    vtype.Enabled = True
    grid1.Enabled = True
    ''s = vtype.Enabled
    ''vtype.SetFocus
    
    Commandsave.Enabled = True
    Commandmasteradd.Enabled = False
    Commandedit.Enabled = False
    Commanddelete.Enabled = False
    Commandsearch.Enabled = False
    'CommandPrint.Enabled = False
    addmode = True
    vno.Enabled = False
    
End Sub
Private Sub CommandPrint_Click()

Dim sss1 As String


DSNNew

sss1 = vno.text & "-" & vdate.text & "-" & vtype.text
con.Execute "update vouchers set  printId='" & sss1 & "' where  vouchertype='" + Trim(vtype.text) + "' and convert(smalldatetime,voucherdate,103)= convert (smalldatetime,'" & vdate & "',103) And vouchernumber = " + Trim(vno.text)

If MsgBox("Want to View  ? ", vbQuestion + vbYesNo) = vbYes Then

    MainMenu.cr1.Reset
    MainMenu.cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    MainMenu.cr1.ReportFileName = rptPath & "\" & main.directory & "\Voucher.rpt"
    MainMenu.cr1.ReplaceSelectionFormula "{vouchers.printid}='" & sss1 & "'"
    MainMenu.cr1.WindowShowPrintBtn = True
    MainMenu.cr1.WindowShowPrintSetupBtn = True
    MainMenu.cr1.WindowShowSearchBtn = True
    MainMenu.cr1.WindowState = crptMaximized
    MainMenu.cr1.WindowShowRefreshBtn = True
    MainMenu.cr1.WindowShowExportBtn = True
    MainMenu.cr1.Action = 1

End If



End Sub

Private Sub CommandReturn_Click()

''MainMenu.Toolbar1.Visible = True
    
    
If MsgBox("Are you Sure ? want to exit ...", vbQuestion + vbYesNo) = vbYes Then
    
    
If addmode = True Then
   createLog UserName, Trim(vno.text), "voucher" & vtype, "Exit Without save:" & totaldebit.text, Date
End If
    
Unload Me
    
End If




    
End Sub
Sub updateVoucher()

Dim id As String
Dim paytype As String
Dim party_ As String

Dim drcr As String
Dim CBND As String
Dim billdate
Dim uname As String
Dim gledger As String
Dim amount

Dim description_ As String

For I = 1 To grid1.rows Step 2

  grid1.Row = I
  grid1.Col = 0
  
  If grid1.text <> "" Then
  
  
    gledger = grid1.text
   
    grid1.Col = 1
    If grid1.text = "" Then
       party_ = ""
    Else
       party_ = grid1.text
       
    End If
    
    
    
    
    grid1.Col = 2
    If grid1.text <> "" Then
        amount = Val(grid1.text)
        drcr = "D"
    Else
        grid1.Col = 3
        amount = Val(grid1.text)
        drcr = "C"
    End If

    paytype = IIf(txtpaytype.text = "", "n", txtpaytype.text)
    
    
    
    grid1.Col = 4
    CBND = grid1.text
            
    grid1.Col = 5
    If IsDate(grid1.text) Then
       ''billDate = grid1.Text
       billdate = Format(grid1.text, "MM/dd/yyyy")
    Else
       billdate = ""
    End If
            
    grid1.Col = 6
    uname = grid1.text
     
    grid1.Col = 0
    grid1.Row = grid1.Row + 1
    If grid1.text <> "" Then
       description_ = grid1.text
    End If
    
    id = grid1.TextMatrix(I + 1, 7)
    Line1 = grid1.Row
    
    If id <> "" Then
    
        con.Execute "update VOUCHERS set " & _
        "VoucherType='" & vtype.text & "'," & _
        "VoucherDate='" & Format(Me.vdate.text, "MM/dd/yyyy") & "'," & _
        "VoucherNumber='" & vno.text & "'," & _
        "Genledger='" & gledger & "'," & _
        "SubLedger='" & party_ & "'," & _
        "Amount=" & amount & "," & _
        "DebitorCredit='" & drcr & "'," & _
        "CBND='" & CBND & "'," & _
        "description='" & description_ & "'," & _
        "UserName='" & uname & "'," & _
        "paytype='" & paytype & "'" & _
        " where vsno=" + id + ""
        
        If billdate <> "" Then
           con.Execute "update VOUCHERS set billDate='" & billdate & "' where vsno=" + id + ""
        Else
           con.Execute "update VOUCHERS set billDate=null where vsno=" + id + ""
        End If
    
    Else
    
    
        con.Execute "insert into VOUCHERS(VoucherType,VoucherDate,VoucherNumber,Genledger," & _
        "SubLedger,Amount,DebitorCredit,CBND,description,setupid,fyear,userName,paytype) " & _
        "values('" & vtype.text & "','" & Format(Me.vdate.text, "MM/dd/yyyy") & "','" & vno.text & "'," & _
        "'" & gledger & "','" & party_ & "'," & amount & ",'" & drcr & "','" & CBND & "'," & _
        "'" & description_ & "','" & setupid & "','" & session & "','" & UserName & "','" & paytype & "')"
        
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        
        If rs1.State = 1 Then rs1.close
        rs1.Open "select vsno from VOUCHERS where  vouchertype='" + Trim(vtype.text) + "' and convert(smalldatetime,voucherdate,103)= convert (smalldatetime,'" & vdate & "',103) And vouchernumber = " + Trim(vno.text), con
        If rs1.EOF = False Then
            If billdate <> "" Then
               con.Execute "update VOUCHERS set billDate='" & billdate & "' where vsno=" & rs1(0) & ""
            Else
               con.Execute "update VOUCHERS set billDate=null where vsno=" & rs1(0) & ""
            End If
        End If
        
        
        
        
        
    End If
    
    
    


  End If

Next





Dim dr_ As Double
Dim cr_ As Double

b10 = 0


dr_ = ReturnAmt_Voucher("D")
cr_ = ReturnAmt_Voucher("C")




If (Val(dr_) <> Val(cr_)) Then
   createLog UserName, Trim(vno.text), "voucher" & vtype, "NotUpdated:" & totaldebit.text, Date
   MsgBox "Voucher correctly not saved.... "
Else
   createLog UserName, Trim(vno.text), "voucher" & vtype, "VEdit:" & totaldebit.text, Date
End If




End Sub
Sub BackupVoucher(type_ As String)

On Error GoTo bk_

sss1 = vno.text & "-" & vdate.text & "-" & vtype.text
con.Execute "insert into VOUCHERS_bk" & _
"(VoucherType,VoucherDate,VoucherNumber,GenLedger,SubLedger,Amount,DebitorCredit,CBND," & _
"EntryNumber, [DESCRIPTION], vsno, CashCheck, setupid" & _
",fyear,userName,printId,billDate,PayType) " & _
"SELECT VoucherType,VoucherDate,VoucherNumber,GenLedger,SubLedger,Amount,DebitorCredit,CBND,EntryNumber," & _
"[DESCRIPTION],vsno,CashCheck,setupid,fyear,'" & UserName & "','" & sss1 & "',billDate,'" & type_ & "' from VOUCHERS " & _
"where  vouchertype='" + Trim(vtype.text) + "' and convert(smalldatetime,voucherdate,103)= convert (smalldatetime,'" & vdate & "',103) And vouchernumber = " + Trim(vno.text)



Exit Sub
bk_:
MsgBox "" & err.DESCRIPTION
Commandsave.Enabled = True


End Sub
Sub tmpVoucher(type_ As String)

On Error GoTo bk_

sss1 = vno.text & "-" & vdate.text & "-" & vtype.text

con.Execute "insert into tmpVOUCHERS_del" & _
"(VoucherType,VoucherDate,VoucherNumber,GenLedger,SubLedger,Amount,DebitorCredit,CBND," & _
"EntryNumber, [DESCRIPTION], vsno, CashCheck, setupid" & _
",fyear,userName,printId,billDate,PayType) " & _
"SELECT VoucherType,VoucherDate,VoucherNumber,GenLedger,SubLedger,Amount,DebitorCredit,CBND,EntryNumber," & _
"[DESCRIPTION],vsno,CashCheck,setupid,fyear,'" & UserName & "','" & sss1 & "',billDate,'" & type_ & "' from VOUCHERS " & _
"where  vouchertype='" + Trim(vtype.text) + "' and convert(smalldatetime,voucherdate,103)= convert (smalldatetime,'" & vdate & "',103) And vouchernumber = " + Trim(vno.text)


Exit Sub
bk_:
MsgBox "" & err.DESCRIPTION
Commandsave.Enabled = True


End Sub
Sub tmpVoucher_row_wise(rowid As String)

On Error GoTo bk_

sss1 = vno.text & "-" & vdate.text & "-" & vtype.text

con.Execute "insert into tmpVOUCHERS_del" & _
"(VoucherType,VoucherDate,VoucherNumber,GenLedger,SubLedger,Amount,DebitorCredit,CBND," & _
"EntryNumber, [DESCRIPTION], vsno, CashCheck, setupid" & _
",fyear,userName,printId,billDate,PayType) " & _
"SELECT VoucherType,VoucherDate,VoucherNumber,GenLedger,SubLedger,Amount,DebitorCredit,CBND,EntryNumber," & _
"[DESCRIPTION],vsno,CashCheck,setupid,fyear,'" & UserName & "','" & sss1 & "',billDate,'" & type_ & "' from VOUCHERS " & _
"where vsno=" & rowid & ""


Exit Sub
bk_:
MsgBox "" & err.DESCRIPTION
Commandsave.Enabled = True


End Sub
Private Sub Commandsave_Click()

On Error GoTo save_


    Dim rs5 As New ADODB.Recordset
    Dim SAVED As Boolean
    SAVED = False
    
    If Val(totalcredit.text) = 0 Or Val(totaldebit.text) = 0 Then
       MsgBox "Empty Voucher Not Saved... "
       Exit Sub
    End If
    
    
    
    If (Val(totalcredit.text) <> Val(totaldebit.text)) Then
       MsgBox "Please Check That Total Credit and Total Debit is Differ ", vbCritical
       Exit Sub
    End If
    
    Dim particullar_ As String
    Dim amt As Double
    amt = 0
    particullar_ = ""
     
    
    If addmode = True Then
    
    If MsgBox("Do you want to save it now ?", vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    
    End If
    
    Commandsave.Enabled = False
    Commandsearch.Enabled = True
    Commandabandon.Enabled = False
    
If Val(totalcredit.text) = Val(totaldebit.text) And totalcredit.text <> "" And totaldebit.text <> "" Then 'And Me.Commandedit.Enabled = True Then
    If RS.State = 1 Then
        RS.close
    End If
    
    If Edit = False Then
        RS.Open "select * from vouchers where " & stringyear & " and VoucherNumber<=0", con, adOpenDynamic, adLockOptimistic
    Else
        
        ''con.Execute "Delete  from vouchers where " & stringyear & " and VoucherType='" + Trim(varVtype) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & varVdate & "',103) and vouchernumber=" + Trim(VarVno) + ""
        ''con.Execute "Delete  from vouchers where " & stringyear & " and VoucherType='" + Trim(varVtype) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & varVdate & "',103) and vouchernumber=" + Trim(VarVno) + ""
        
        updateVoucher
        
       
        
        If grid1.TextMatrix(1, 1) <> "" Then
             particullar_ = grid1.TextMatrix(1, 1)
          Else
             particullar_ = grid1.TextMatrix(1, 0)
          End If
            
        amt = Val(totaldebit.text)
        updateV_Main "edit", vtype, vdate, vno, particullar_, amt
        
        voucherMainBK "Edit"
        
        
        
                
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        
        MsgBox " RECORD Updated ... ", vbInformation
        

        
        
        ''RS.Open "select * from vouchers where " & stringyear & " and VoucherNumber<=0", con, adOpenDynamic, adLockPessimistic
    
    End If
  
  If addmode = True Then
  
     rs5.Open "Select max(VoucherNumber) from vouchers where " & stringyear & " and vouchertype='" + Trim(vtype.text) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & vdate & "',103)", con, adOpenDynamic, adLockReadOnly, adCmdText
     If Not rs5.BOF Then
        If Not rs5(0) Then
            c = rs5(0)
        Else
            c = 0
       End If
    Else
        c = 0
    End If
    If rs5.State = 1 Then rs5.close
    vno.text = Str(c + 1)
    
 
  
    grid1.Row = I
    grid1.Col = 0
    For I = 1 To grid1.rows Step 2
        grid1.Row = I
        grid1.Col = 0
        If grid1.text <> "" Then
            RS.AddNew
            RS(0) = vtype.text
            RS(1) = vdate.text
            RS(2) = vno.text
            RS(3) = grid1.text
            RS!paytype = IIf(txtpaytype.text = "", "n", txtpaytype.text)
            
            grid1.Col = 1
            If grid1.text = "" Then
               RS(4) = ""
            Else
               RS(4) = grid1.text
            End If
            grid1.Col = 2
            If grid1.text <> "" Then
                RS(5) = Val(grid1.text)
                RS(6) = "D"
            Else
                grid1.Col = 3
                RS(5) = Val(grid1.text)
                RS(6) = "C"
            End If
            grid1.Col = 4
            RS(7) = grid1.text
            
            grid1.Col = 5
            If IsDate(grid1.text) Then
               RS!billdate = grid1.text
            End If
            
            grid1.Col = 6
            RS!UserName = grid1.text

            
            grid1.Col = 0
            grid1.Row = grid1.Row + 1
            If grid1.text <> "" Then
                RS(9) = grid1.text
            End If
            
            Line1 = grid1.Row
            
            SAVED = True
            RS!setupid = setupid
            RS!fyear = session
            
            RS.update
        
            
        
        End If
        
    Next
    
    
    
  
     If grid1.TextMatrix(1, 1) <> "" Then
        particullar_ = grid1.TextMatrix(1, 1)
     Else
        particullar_ = grid1.TextMatrix(1, 0)
     End If
     
         
     amt = Val(totaldebit.text)
     updateV_Main "add", vtype, vdate, vno, particullar_, amt
      
     voucherMainBK "Insert"
    
    
    Dim dr_ As Double
    Dim cr_ As Double
    
    
    
    dr_ = ReturnAmt_Voucher("D")
    cr_ = ReturnAmt_Voucher("C")
    
    If (Val(dr_) <> Val(cr_)) Then
       createLog UserName, Trim(vno.text), "voucher" & vtype, "Notsaved:" & totaldebit.text, Date
       MsgBox "Voucher correctly not saved.... "
       Exit Sub
    Else
        createLog UserName, Trim(vno.text), "voucher" & vtype, "saved:" & totaldebit.text, Date
        BackupVoucher ("Save")
    End If
    
    
    
    MsgBox " RECORD SAVED ... ", vbInformation
    grid1.Enabled = True
   
    
    End If
    
    
    
    CommandPrint.Visible = True
    
    If RS.State = 1 Then RS.close
  
    
    DoEvents
    DoEvents
    
    Commandsave.Enabled = False
    Commandmasteradd.Enabled = True
    Commandmasteradd.SetFocus
    Commandedit.Enabled = True
    Commanddelete.Enabled = True
    Commandsearch.Enabled = True
    CommandPrint.Enabled = True
    Command1.Enabled = True
    Command2.Enabled = True
    Commandabandon.Enabled = True
    
    vtype.Enabled = False
    vdate.Enabled = False
    vno.Enabled = False
    Commandsearch.Enabled = True
    grid1.Enabled = False
    SAVED = False
    
    mnuMenu_ = "menujournalvoucher"
    SetButton Commandmasteradd, Commandedit, Commandsave, Commanddelete
    
    Commandsave.Enabled = False
    
    Else
    
    'If Commandedit.Enabled = True Then
    MsgBox "Please Check That Total Credit and Total Debit is Differ "
    'End If
    End If
    addmode = False
    
    
    
    
    If (AuditTrail = "y") Then
    
    If rs1.State = 1 Then rs1.close
    rs1.Open "select * from  AuditTrail_Log where vouchertype='" + Trim(vtype) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & vdate & "',103) And Vouchernumber = " + Trim(vno), con
    If rs1.EOF = False Then
 
    'If (txtchecked.text = "y") Then

        actionType_ = "Edit"
        vtype1_ = "v"
        vdate_ = vdate
        vno_ = vno.text
        vtypeNew = vtype

        frmAuditTrailLog_Rem.Show 1
    
    Else
        
        If rs1.State = 1 Then rs1.close
        rs1.Open "select * from  VOUCHERS_main  where vouchertype='" + Trim(vtype) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & vdate & "',103) And Vouchernumber = " + Trim(vno), con
        If rs1.EOF = False Then
            con.Execute "insert into AuditTrail_Log(VoucherID,VoucherType,ActionType,VoucherDate,VoucherNumber,[Description],Amount,ReasionForEdit,UserName) " & _
            " values ('" & rs1!VoucherID & "','" & rs1!VoucherType & "','" & "Insert" & "','" & Format(rs1!voucherDATE, "MM/dd/yyyy") & "','" & rs1!VOUCHERNUMBER & "','" & rs1!Particular & "','" & rs1!amount & "','" & "Add Voucher" & "','" & UserName & "')"
        End If

    End If

    End If
  
    
Exit Sub

save_:

createLog UserName, Trim(vno.text), "voucher" & vtype, "Error : " & totaldebit.text, Date

For k1 = 0 To 6
    grid1.CellBackColor = vbGreen
    DoEvents
Next


MsgBox "Error" & err.DESCRIPTION
Commandsave.Enabled = False
Commandmasteradd.Enabled = True
Commandmasteradd.SetFocus
Commandedit.Enabled = True
Commanddelete.Enabled = True
Commandsearch.Enabled = True
CommandPrint.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Commandabandon.Enabled = True



End Sub
Function ReturnAmt_Voucher(dr_cr As String) As Double
     
     Dim rr As New ADODB.Recordset
     Dim amt As Double
     
     Set rr = New ADODB.Recordset
     rr.Open "SELECT sum(Amount) from VOUCHERS where DebitorCredit='" & dr_cr & "' and  vouchertype='" + Trim(vtype.text) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & vdate & "',103) And vouchernumber = " + Trim(vno.text), con
     If Not IsNull(rr(0)) Then
        amt = rr(0)
     End If
     
     ReturnAmt_Voucher = amt
     
End Function
Private Sub Commandsearch_Click()
    searchmode = True
    Edit = False
    vno.Enabled = True
    addmode = False
    Set rs1 = New ADODB.Recordset
    Set RS = New ADODB.Recordset
    vno.Enabled = True
    vtype.Enabled = True
    vdate.Enabled = True
    vtype.SetFocus
 
End Sub

Private Sub credit_Change()
    If Val(credit.text) <> 0 Then
       grid1.text = Val(credit.text)
       
       grid1.text = Format(grid1.text, "0.00")
    Else
        If grid1.Col <> 4 And grid1.Col <> 0 Then
           'Grid1.Text = ""
        End If
    End If
End Sub
Private Sub credit_GotFocus()
    
     cl = grid1.Col
     grid1.Col = 0
     Dim Cashbankbook As String
     Cashbankbook = 0
     Set rs2_ = New ADODB.Recordset
     If rs2_.State = 1 Then rs2_.close
     rs2_.Open "SELECT cashbankbook FROM GLEDGER where gledger='" & grid1.text & "'", con
     If rs2_.EOF = False Then
        If rs2_(0) = True Then
          Cashbankbook = 1
        End If
     End If
   
   
   If grid1.text <> "" Then
     
     ''If grid1.Text = "CASH-IN-HAND" And vtype.Text = "P" Then
     If Cashbankbook = "1" And vtype.text = "R" Then
        grid1.Col = 3
        grid1.text = ""
        grid1.Col = 2
        debit.Visible = False
        Grid1_Click
    
    Else
       
       grid1.Col = cl
  
       sendkeys "{END}"
       sendkeys "+{HOME}"
       
    End If
    
  End If
    
    ''=================================================
    
    
    
    If maxrow < grid1.Row Then
       maxrow = grid1.Row
    End If
    If searchmode = False Then
        'HIT
        'credit.SetFocus
        sendkeys "{END}"
        
        
        
        sendkeys "+{HOME}"
  End If
  
  If Edit = True Then
    If Val(credit) > 0 Then
       cr_New = credit
    End If
  End If
        
End Sub

Private Sub credit_KeyDown(KeyCode As Integer, Shift As Integer)
   If (KeyCode >= 48 And KeyCode <= 57) Then
      grid1.TextMatrix(grid1.RowSel, 6) = UserName
   End If

End Sub

Private Sub credit_KeyPress(KeyAscii As Integer)


If KeyAscii = 13 Then
'    If Val(credit) > 0 Or Val(debit) > 0 Then
    Dim changed As Boolean
    Dim HIT As Boolean
    HIT = False
    changed = False
    RRR = grid1.Row
    CCC = grid1.Col
    grid1.Col = 3
    totalcredit.text = ""
    For I = 1 To maxrow Step 1
        grid1.Row = I
        totalcredit.text = Val(totalcredit.text) + Val(grid1.text)
    Next
    totalcredit.text = Format(totalcredit.text, "0.00")
    grid1.Row = RRR
    grid1.Col = CCC
    grid1.Col = 3
    If credit.text = "" Then
            grid1.text = ""
            grid1.Col = 2
            Grid1_Click
            Exit Sub
     End If
    If grid1.Col = lastcol And grid1.Row = lastrow Then
        grid1.Col = 0
        If Trim(grid1.Col) <> "" Then
            HIT = True
        End If
        grid1.Col = 3
        grid1.Col = grid1.Col + 1
        refreshTotal
        changed = True
    End If
    RRR = grid1.Row
    CCC = grid1.Col
    grid1.Col = 2
    If grid1.Row <> lastrow Or grid1.Col <> lastcol Then
        grid1.text = ""
    End If
   
    
    grid1.Row = RRR
    grid1.Col = CCC
    
    
    
    If changed And HIT Then
        Grid1_Click
    End If
    Else
    ''credit.Visible = False
   ' 'Grid1.col = 2
'    SendKeys "{left}{left}"
    ''Grid1_Click
'    End If
End If
End Sub
Private Sub credit_LostFocus()
    ''credit = ""
    ''credit.Visible = False
    RRR = grid1.Row
CCC = grid1.Col
grid1.Row = lastrow

grid1.Col = lastcol
If Trim(grid1.text) <> "" Then
    grid1.Row = RRR
    grid1.Col = CCC
    credit.Visible = False
Else
'  Grid1.col = Grid1.col + 1
'  debit.Visible = False
'  Grid1_Click
'Grid1.Text = credit.Text

     If Edit = True Then
        If Val(credit) <> Val(cr_New) Then
           grid1.TextMatrix(grid1.RowSel, 6) = UserName
        End If
     End If

End If
credit = ""
    
End Sub
Private Sub crno_Change()
'crno = UCase(crno)



If grid1.ColSel = lastcol And grid1.Row = lastrow Then
   grid1.text = crno.text
End If



End Sub

Private Sub crno_GotFocus()
crno.text = grid1.text
If searchmode = False Then
    HIT
End If

End Sub
Function checkDuplicateBill(billno As String, sledger As String) As String
  Dim rs5 As ADODB.Recordset
  Set rs5 = New ADODB.Recordset
  
  rs5.Open "select voucherNumber,VoucherDate from VOUCHERS where (genledger='SUNDRY CREDITORS' and subledger='" & sledger & "' and cbnd='" & billno & "')", con
  If rs5.EOF = False Then
     checkDuplicateBill = "This Bill No Already Exist in Voucher No : " & rs5!VOUCHERNUMBER & " and " & rs5!voucherDATE
  Else
     checkDuplicateBill = "non"
  End If

End Function

Private Sub Crno_KeyPress(KeyAscii As Integer)
On Error GoTo err_

If KeyAscii = 13 Then
    Dim changed As Boolean
    Dim HIT As Boolean
    HIT = False
    changed = False
    
 
   
    
    If grid1.Col = lastcol And grid1.Row = lastrow Then
        crno = UCase(crno)
        grid1.Col = 0
        If Trim(grid1.text) <> "" Then
            HIT = True
        End If
        
       
        
        If (grid1.TextMatrix(grid1.RowSel, 6) = "" Or grid1.TextMatrix(grid1.RowSel, 6) = ".") Then
        grid1.TextMatrix(grid1.RowSel, 6) = UserName
        End If
        
        If grid1.text <> "" Then
           grid1.Col = 5
        End If
        
        'grid1.Row = grid1.Row + 1
        
        changed = True
    End If
    
    If Edit = False Then
     If (grid1.TextMatrix(grid1.RowSel, 4) <> "" And grid1.TextMatrix(grid1.RowSel, 1) <> "") Then
        stt1 = checkDuplicateBill(grid1.TextMatrix(grid1.RowSel, 4), grid1.TextMatrix(grid1.RowSel, 1))
        If stt1 <> "non" Then
           MsgBox " " & stt1
           grid1.Col = 4
           Exit Sub
        End If
     End If
     
    End If
    
    If changed And HIT Then
        Grid1_Click
    End If
End If


Exit Sub
err_:
MsgBox "" & err.DESCRIPTION


End Sub

Private Sub Crno_LostFocus()
    crno.Visible = False
End Sub

Private Sub crno1_Change()
If IsDate(crno1.text) Then
grid1.text = crno1.text
End If
End Sub
Private Sub crno1_GotFocus()
If IsDate(grid1.text) Then
crno1.text = grid1.text
End If

If searchmode = False Then
    HIT
End If
End Sub

Private Sub crno1_KeyPress(KeyAscii As Integer)
On Error GoTo err_

If KeyAscii = 13 Then
    Dim changed As Boolean
    Dim HIT As Boolean
    HIT = False
    changed = False
    
    
    If grid1.Col = lastcol And grid1.Row = lastrow Then
        crno1 = UCase(crno1)
        grid1.Col = 0
        If Trim(grid1.text) <> "" Then
            HIT = True
        End If
        
        
        If (grid1.TextMatrix(grid1.RowSel, 6) = "" Or grid1.TextMatrix(grid1.RowSel, 6) = ".") Then
        grid1.TextMatrix(grid1.RowSel, 6) = UserName
        End If
        
        grid1.Row = grid1.Row + 1
        
        changed = True
    End If
    If changed And HIT Then
        Grid1_Click
    End If
End If


Exit Sub
err_:
MsgBox "" & err.DESCRIPTION

End Sub
Private Sub crno1_LostFocus()

    'sss = grid1.TextMatrix(grid1.RowSel - 1, 5)

    If Not IsDate(crno1.text) Then
       grid1.TextMatrix(grid1.RowSel - 1, 5) = "__/__/____"
    End If

End Sub

Private Sub debit_Change()
If Val(debit.text) <> 0 Then
    grid1.text = Val(debit.text)
    grid1.text = Format(grid1.text, "0.00")
Else
'If Grid1.col <> 4 Then  Grid1.Text = ""
 If grid1.Col <> 4 And grid1.Col <> 0 Then
   ' Grid1.Text = ""
 End If
End If

End Sub
Private Sub debit_GotFocus()
   
     cl = grid1.Col
     grid1.Col = 0
     Dim Cashbankbook As String
     Cashbankbook = 0
     Set rs2_ = New ADODB.Recordset
     If rs2_.State = 1 Then rs2_.close
     rs2_.Open "SELECT cashbankbook FROM GLEDGER where gledger='" & grid1.text & "'", con
     If rs2_.EOF = False Then
         If rs2_(0) = True Then
          Cashbankbook = 1
        End If
     End If
   
   
   If grid1.text <> "" Then
     
     ''If grid1.Text = "CASH-IN-HAND" And vtype.Text = "P" Then
     If Cashbankbook = "1" And vtype.text = "P" Then
        grid1.Col = 2
        grid1.text = ""
        grid1.Col = 3
        debit.Visible = False
        Grid1_Click
        
     
        
    
    Else
       
       grid1.Col = cl
             
       sendkeys "{END}"
       sendkeys "+{HOME}"
       
    End If
   

    If maxrow < grid1.Row Then
        maxrow = grid1.Row
    End If
    
    If searchmode = False Then
       
        sendkeys "{END}"
        sendkeys "+{HOME}"
    End If
    Else
        debit.Visible = False
    End If
  
  
  If Edit = True Then
   
  If Val(debit) > 0 Then
     dr_New = debit
  End If
  
  End If
   
 
End Sub
Private Sub debit_KeyDown(KeyCode As Integer, Shift As Integer)
   If (KeyCode >= 48 And KeyCode <= 57) Then
      grid1.TextMatrix(grid1.RowSel, 6) = UserName
   End If
End Sub

Private Sub debit_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim changed As Boolean
    Dim HIT As Boolean
    changed = False
    HIT = False
    RRR = grid1.Row
    CCC = grid1.Col
    grid1.Col = 2
    totaldebit.text = ""
    For I = 1 To maxrow
        grid1.Row = I
        totaldebit.text = Val(totaldebit.text) + Val(grid1.text)
    Next
    totaldebit.text = Format(totaldebit.text, "0.00")
    grid1.Row = RRR
    grid1.Col = CCC
    If grid1.Col = lastcol And grid1.Row = lastrow Then
        grid1.Col = 0
        If Trim(grid1.text) <> "" Then
            HIT = True
        End If
        
        grid1.Col = 2
        If debit.text = "" Then grid1.text = ""

        If grid1.text <> "" Then
            grid1.Col = 4
        Else
            grid1.Col = 0
            If grid1.text = "CASH-IN-HAND" And vtype.text = "R" Then
                 grid1.Col = 4
                  grid1.Col = 3
                  If grid1.text = "" Then
                      debit.Visible = True
                      credit.Visible = False
                      grid1.Col = 2
                      debit.SetFocus
                     Exit Sub
                  End If
                 
            Else
                 grid1.Col = 2
                 If debit.text = "" Then grid1.text = ""
                 grid1.Col = 3
            End If
              
            
            
        End If
        changed = True
    End If
    RRR = grid1.Row
    CCC = grid1.Col
    grid1.Col = 2
    If grid1.Row <> lastrow Or grid1.Col <> lastcol Then
        grid1.text = ""
    End If
    grid1.Row = RRR
    grid1.Col = CCC
    If changed And HIT Then
        Grid1_Click
    End If
End If
End Sub

Private Sub debit_LostFocus()
RRR = grid1.Row
CCC = grid1.Col
grid1.Row = lastrow
grid1.Col = lastcol
If Trim(grid1.text) <> "" Then
    grid1.Row = RRR
    grid1.Col = CCC
   
    debit.Visible = False
Else
     refreshTotal
     
     If Edit = True Then
     
        If Val(debit) <> Val(dr_New) Then
           grid1.TextMatrix(grid1.RowSel, 6) = UserName
        End If
        
     End If
     
     
End If
debit = ""
End Sub
Private Sub DESCRIPTION_GotFocus()
 
Dim NARR As String
Dim billno, bdate, pname As String
NARR = ""
billno = ""
bdate = ""
pname = ""

If CheckBox1_createDesc.value = True Then


   billno = grid1.TextMatrix(grid1.RowSel - 1, 4)
   bdate = grid1.TextMatrix(grid1.RowSel - 1, 5)
   pname = grid1.TextMatrix(grid1.RowSel - 1, 1)
   
   NARR = "B/N " & billno & " DT " & bdate & " " & pname
   If Len(billno) > 0 And Len(bdate) > 0 Then
      DESCRIPTION.text = NARR
   Else
   
   If Trim(grid1.text) <> Trim(".") Then
    DESCRIPTION.text = grid1.text
   End If
   
   End If
   
Else
 If Trim(grid1.text) <> Trim(".") Then
    DESCRIPTION.text = grid1.text
 End If
 
End If
 
 
 HIT
End Sub

Private Sub Description_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DESCRIPTION = UCase(DESCRIPTION)
        grid1.text = DESCRIPTION.text
    Dim changed As Boolean
    Dim HIT As Boolean
    HIT = False
    changed = False
    If grid1.Col = lastcol And grid1.Row = lastrow Then
        grid1.Col = 0
        HIT = True
        grid1.Row = grid1.Row + 1
        changed = True
    End If
    'DESCRIPTION.Text = ""
    If changed And HIT Then
        Grid1_Click
    End If
    End If
End Sub
Private Sub DESCRIPTION_LostFocus()
    DESCRIPTION = UCase(DESCRIPTION)
    DESCRIPTION.Visible = False
End Sub

Private Sub FindCombo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If FindCombo.text <> "" Then
vtype.Enabled = False
   vdate.Enabled = False
   vno.Enabled = False
   vtype = Trim(Mid(FindCombo.text, 1, 1))
   vdate = Trim(Mid(FindCombo.text, 2, 12))
   vno = Trim(Mid(FindCombo.text, 14, 5))
   
   vno_LostFocus
   FindCombo.Visible = False
 Else
   FindCombo.Visible = False
 End If
 End If
 
If KeyAscii = 27 Then FindCombo.Visible = False
   
End Sub

Private Sub Form_Activate()

'Command1.SetFocus

mnuMenu_ = "menujournalvoucher"
SetButton Commandmasteradd, Commandedit, Commandsave, Commanddelete
Commandsave.Enabled = False
Commanddelete.Enabled = False
'Commandedit.Enabled = False
Commandsearch.Enabled = True


End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)



If KeyCode = 113 Then

If addmode = False Then
  
  FindCombo.Visible = True
  FindCombo.SetFocus

  Dim rs7 As New ADODB.Recordset
  Set rs7 = con.Execute("exec voucherQry '" & session & "','" & main.setupid & "'")
  
  If rs7.EOF = False Then
   FindCombo.Clear
   rs7.MoveFirst
   While Not rs7.EOF
       FindCombo.AddItem rs7(0) & Space(1) & rs7(1) & Space(1) & rs7(2)
       rs7.MoveNext
   Wend
  End If
  
End If

End If

If KeyCode = 27 Then

If MsgBox("Are you Sure ? want to exit ...", vbQuestion + vbYesNo) = vbYes Then
    
    If addmode = True Then
       If totaldebit.text <> "" Then
         createLog UserName, Trim(vno.text), "voucher" & vtype, "Exit Without save:" & totaldebit.text, Date
       
      
 
      
       End If
    End If
    
    Unload Me
    
End If

End If


End Sub

Private Sub Form_Resize()

panel.Left = (Me.ScaleWidth - panel.Width) / 2
panel.top = (Me.ScaleHeight - panel.Height) / 2

End Sub

Private Sub Gencombo_Change()
 If grid1.Col = lastcol And grid1.Row = lastrow Then
        grid1.text = Gencombo.text
        grid1.Col = 0
 End If
End Sub

Private Sub genCombo_Click()
    grid1.text = Gencombo.text
    If vtype.text = "R" And grid1.text = "CASH-IN-HAND" Then
        grid1.Col = 2
        db = grid1.text
        grid1.Col = 3
        CR = grid1.text
        If db = "" Then
            grid1.Col = 2
            grid1.text = CR
            grid1.Col = 3
            grid1.text = ""
       End If
    End If
   If vtype.text = "P" And grid1.text = "CASH-IN-HAND" Then
        grid1.Col = 2
        db = grid1.text
        grid1.Col = 3
        CR = grid1.text
        If CR = "" Then
            grid1.Col = 3
            grid1.text = db
            grid1.Col = 2
            grid1.text = ""
       End If
    End If
    grid1.Col = 1
    grid1.text = ""
    grid1.Col = 0
End Sub
Private Sub Form_Load()

Me.top = 0
Me.Left = 0

Me.Width = 14350
Me.Height = 10300

Me.Caption = "Voucher Entry"

dr_New = 0
cr_New = 0

varVtype = ""
varVdate = ""
VarVno = 0

searchmode = False
addmode = False
vtype.Enabled = False
vdate.Enabled = False
vno.Enabled = False

Commandsearch.Enabled = True
'grid1.Enabled = False
Edit = False
    'Grid1.Top = 450
    'Grid1.Left = 60
    Me.top = 0
    Me.Left = 0
    grid1.rows = 300
    grid1.Cols = 0
    grid1.Cols = 8
    
    For I = 0 To grid1.rows - 1
       grid1.RowHeight(I) = 250
    Next
    
    grid1.Row = 0
    grid1.Col = 0
    grid1.MergeCells = flexMergeRestrictRows
    For I = 1 To grid1.rows - 1
        grid1.RowHeight(I) = 350
        If I Mod 2 = 0 Then
            grid1.Row = I
            For J = 0 To 5
                grid1.Col = J
                grid1.text = "."
                grid1.MergeRow(I) = True
            Next
        End If
    Next
    
   
   'Grid1.MergeCells = flexMergeRestrictRows
    grid1.Row = 0
    grid1.Col = 0
    grid1.ColAlignment(0) = 0
    grid1.ColAlignment(1) = 0
    grid1.ColWidth(0) = 3000
    grid1.ColWidth(1) = 3600
    grid1.ColWidth(2) = 1350
    grid1.ColWidth(3) = 1350
    grid1.ColWidth(4) = 1100
    grid1.ColWidth(5) = 1100
    
    grid1.text = "Gen. Ledger"
    grid1.Col = 1
    grid1.text = "Sub. Ledger"
    grid1.Col = 2
    grid1.text = "Amount. (Dr.)"
    'totaldebit.Left = Grid1.CellLeft + 50
    grid1.Col = 3
    grid1.text = "Amount. (Cr.)"
    'totalcredit.Left = Grid1.CellLeft + 60
    grid1.Col = 4
    grid1.text = "C/B No."
    
    
    grid1.TextMatrix(0, 5) = "Date"
    grid1.TextMatrix(0, 6) = "User"
    grid1.TextMatrix(0, 7) = "Id"
    grid1.ColWidth(7) = 0
    
    crno.Height = grid1.CellHeight
    
    
    If (UserName = "y" Or UserName = "v" Or LCase(UserName) = "admin") Then
       grid1.ColWidth(6) = 550
    Else
        grid1.ColWidth(6) = 0
    End If

    
    
    
'    Set RS = New ADODB.Recordset
'    RS.PageSize = 2
'    RS.Open "Select top 100 * from vouchers where " & stringyear & "  order by vouchertype,voucherdate,vouchernumber", con, adOpenStatic
'    If Not RS.BOF Then
'        ''vtype = RS!vouchertype
'        ''vdate = RS!voucherDATE
'        ''vno = RS!VOUCHERNUMBER
'        ''Me.vtype_LostFocus
'        ''Me.vdate_LostFocus
'        ''Me.vno_LostFocus
'        addfunction
'    End If
'
    
 
SubCombo.Clear
Gencombo.Clear

Set cmd = New ADODB.Command
cmd.ActiveConnection = con
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "fatch_gledger"

Set rs_S = cmd.Execute
    If Not rs_S.EOF Then
        Do While Not rs_S.EOF
            Gencombo.AddItem rs_S!gledger
            If Not rs_S.EOF Then
                rs_S.MoveNext
            End If
        Loop
    End If
    rs_S.close



Set cmd.ActiveConnection = Nothing
'=================================================
    
mnuMenu_ = "menujournalvoucher"
SetButton Commandmasteradd, Commandedit, Commandsave, Commanddelete
Commandsave.Enabled = False


If (LCase(UserName) = "admin" Or LCase(UserName) = "v" Or LCase(UserName) = "y") Then
    Label5.Visible = True
Else
    Label5.Visible = False
End If




'Set rs_S = New ADODB.Recordset
'rs_S.Open "Select SUBLEDGER,gledger  from  SLEDGER where " & stringyear & "" & _
'" ORDER BY gledger,SUBLEDGER", con, adOpenKeyset, adLockOptimistic


BackColorFrom Me

addfunction

End Sub
Private Sub Gencombo_GotFocus()
    grid1.Col = 0
'    sss = grid1.TextMatrix(grid1.RowSel, 5)
'
'    If IsDate(crno1.text) Then
'       crno1.text = grid1.text
'    Else
'       crno1.text = "__/__/____"
'    End If
    
    Gencombo.text = grid1.text
    If Gencombo.Enabled = True Then
    If searchmode = False Then
    HIT
    End If
    End If
End Sub
Private Sub Gencombo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 115 Then

   If MsgBox("Are You Sure, Delete it ?", vbYesNo) = vbYes Then
   
      If grid1.Row >= 1 Then
      
           
           If (grid1.TextMatrix(grid1.RowSel + 1, 7) <> "") Then
               rid = grid1.TextMatrix(grid1.RowSel + 1, 7)
               tmpVoucher_row_wise (rid)
               con.Execute "delete  FROM VOUCHERS where vsno=" & grid1.TextMatrix(grid1.RowSel + 1, 7) & ""
               'amt1 = IIf(grid1.TextMatrix(grid1.RowSel, 2) = "", grid1.TextMatrix(grid1.RowSel, 3), grid1.TextMatrix(grid1.RowSel, 2))
               'con.Execute "update VOUCHERS_bk set EntryNumber='1' where (amount='" & amt1 & "' and voucherType='" + Trim(vtype) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & vdate & "',103)) And Vouchernumber = " + Trim(vno)
               actionType_ = "Delete"
               
               
           End If
      
      
           grid1.RemoveItem grid1.Row
           grid1.RemoveItem grid1.Row
           'Gencombo.Text = ""
           Gencombo.Visible = False
           Me.DESCRIPTION.Visible = False
           refreshTotal
           grid1.SetFocus
       End If
   End If
End If

 

End Sub
Private Sub Gencombo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        Dim HIT As Boolean
        Dim changed As Boolean
        changed = False
        HIT = False
        If RS.State = 1 Then RS.close
        RS.Open "GLEDGER", con, adOpenStatic, adLockReadOnly, adCmdTable
        If Trim(Gencombo.text) <> "" Then
        'SendKeys "{DOWN}"
        If Not RS.BOF Then
                RS.MoveFirst
                RS.Find "GLEDGER='" + Trim(Gencombo.text) + "'"
                If RS.EOF Then
                    RS.close
                    Exit Sub
                End If
            End If
        Else
           Gencombo.Visible = False
        End If
        RS.close
        
        If grid1.Col = lastcol And grid1.Row = lastrow Then
            If Trim(grid1.text) <> "" Then
                HIT = True
            End If
            grid1.Col = grid1.Col + 1
            changed = True
        End If
        Gencombo.text = ""
        If changed And HIT Then
            Grid1_Click
       Else
         If changed Then
             Commandsave.Enabled = True
             Commandsave.SetFocus
         End If
        End If

    End If
End Sub

Private Sub Gencombo_LostFocus()
   ' rs.Open "GLEDGER", CON, adOpenDynamic, adLockReadOnly, adCmdTable
   ' If Trim(Gencombo.Text) <> "" Then
   '     If Not rs.BOF Then
   '         rs.MoveFirst
   '         rs.Find "GLEDGER='" + Trim(Gencombo.Text) + "'"
   '         If rs.EOF Then
   '             Gencombo.Visible = True
   '             Gencombo.SetFocus
   '         Else
            
  '              Gencombo.Visible = False
  '              Me.Commandsave.SetFocus
  '          End If
   '     End If
   ' Else
    '    Gencombo.Visible = False
        
    'End If
    'rs.Close
End Sub
Private Sub Grid1_Click()
If grid1.Enabled = True Then
Dim check As Boolean
Dim Row, Col As Integer
Row = grid1.Row
Col = grid1.Col
check = True

 
If grid1.Row Mod 2 <> 0 Then
    DESCRIPTION.Visible = False
    Dim gr As MSHFlexGrid
    Set gr = grid1
    If grid1.Row >= 3 Then
        check = False
        gr.Row = grid1.Row - 2
        gr.Col = 0
        If gr.text <> "" Then
            gr.Col = 2
            If gr.text <> "" Then
                check = True
            Else
                gr.Col = 3
                If gr.text <> "" Then
                    check = True
                End If
            End If
        End If
    End If
  
    grid1.Row = Row
    grid1.Col = Col
    credit.Visible = False
    debit.Visible = False
    SubCombo.Visible = False
    crno.Visible = False
    crno1.Visible = False
    Gencombo.Visible = False
    
    If check Then
       If grid1.Col = 0 Then
          credit.Visible = False
          debit.Visible = False
          SubCombo.Visible = False
          crno.Visible = False
          crno1.Visible = False
          
          Gencombo.text = grid1.text
          Gencombo.Visible = True
          Gencombo.Left = grid1.CellLeft + 30
          Gencombo.top = grid1.top + grid1.CellTop - 50
          Gencombo.Width = grid1.ColWidth(grid1.Col)
          
          Gencombo.ZOrder
          Gencombo.SetFocus
       End If
        
        
        If grid1.Col = 1 Then
            tmpsubledger = ""
            SubCombo.text = ""
            If grid1.text <> "" Then
               tmpsubledger = grid1.text
            End If
            
            Gencombo.Visible = False
            credit.Visible = False
            debit.Visible = False
            crno.Visible = False
            crno1.Visible = False
            
            SubCombo.Visible = True
            SubCombo.text = grid1.text
            SubCombo.Left = grid1.CellLeft + 30
            SubCombo.top = grid1.top + grid1.CellTop - 50
            SubCombo.ZOrder
            SubCombo.SetFocus
            SubCombo.Width = grid1.ColWidth(grid1.Col)
            'SubCombo.Height = Grid1.RowHeight(Grid1.Row)
        End If
        If grid1.Col = 2 Then
            Gencombo.Visible = False
            SubCombo.Visible = False
            credit.Visible = False
            crno.Visible = False
            crno1.Visible = False
            
            grid1.Col = 3
            If Trim(grid1.text) = "" Then
                grid1.Col = 2
                If grid1.text <> "" Then
                    debit.text = grid1.text
                End If
                debit.Visible = True
                debit.Left = grid1.CellLeft + 30
                debit.top = grid1.top + grid1.CellTop - 20
                debit.Width = grid1.ColWidth(grid1.Col)
                debit.Height = grid1.CellHeight
                debit.ZOrder
                debit.SetFocus
                
           Else
                       
              grid1.Col = 3
          
            End If
          End If
        If grid1.Col = 3 Then
            Gencombo.Visible = False
            SubCombo.Visible = False
            debit.Visible = False
            crno.Visible = False
            crno1.Visible = False
            
            grid1.Col = 2
            If Trim(grid1.text) = "" Then
                grid1.Col = 3
                If grid1.text <> "" Then
                credit.text = grid1.text
                End If
                credit.Visible = True
                credit.Left = grid1.CellLeft + 30
                credit.top = grid1.top + grid1.CellTop - 25
                credit.Width = grid1.ColWidth(grid1.Col)
                
                credit.Height = grid1.CellHeight
                
                credit.ZOrder
                credit.SetFocus
                
            End If
        End If
        If grid1.Col = 4 Then
            Gencombo.Visible = False
            SubCombo.Visible = False
            debit.Visible = False
            credit.Visible = False
            crno1.Visible = False
            
            crno.text = grid1.text
            crno.Visible = True
            crno.Left = grid1.CellLeft + 30
            crno.top = grid1.top + grid1.CellTop - 25
            crno.Width = grid1.ColWidth(grid1.Col)
            crno.Height = grid1.CellHeight
            crno.ZOrder
            crno.SetFocus
        End If
        
'        If Grid1.Col = 5 Then
'            Gencombo.Visible = False
'            SubCombo.Visible = False
'            debit.Visible = False
'            credit.Visible = False
'            crno1.Visible = False
'
'            crno.Text = Grid1.Text
'            crno.Visible = True
'            crno.Left = Grid1.CellLeft + 30
'            crno.Top = Grid1.Top + Grid1.CellTop - 25
'            crno.Width = Grid1.ColWidth(Grid1.Col)
'            crno.Height = Grid1.CellHeight
'            crno.ZOrder
'            crno.SetFocus
'        End If
        
        If grid1.Col = 5 Then
            Gencombo.Visible = False
            SubCombo.Visible = False
            debit.Visible = False
            credit.Visible = False
            crno.Visible = False
            If IsDate(grid1.text) Then
               crno1.text = grid1.text
            Else
               crno1.text = "__/__/____"
            End If
            crno1.Visible = True
            crno1.Left = grid1.CellLeft + 30
            crno1.top = grid1.top + grid1.CellTop - 25
            crno1.Width = grid1.ColWidth(grid1.Col)
            crno1.Height = grid1.CellHeight
            crno1.ZOrder
            crno1.SetFocus
        End If
        
    lastrow = grid1.Row
    lastcol = grid1.Col
    Else
    MsgBox "Previous entry didn't complete"
    End If
Else
    If grid1.Row >= 2 Then
        RRR = grid1.Row
        CCC = grid1.Col
        grid1.Row = grid1.Row - 1
        grid1.Col = 0
        If Trim(grid1.text) <> "" Then
            grid1.Row = RRR
            grid1.Col = CCC
            If grid1.Col = 0 Then
                DESCRIPTION.Left = grid1.CellLeft + 30
                DESCRIPTION.top = grid1.top + grid1.CellTop - 25
                DESCRIPTION.Width = grid1.CellWidth  '' Grid1.ColWidth(Grid1.col)
                crno.Visible = False
                Gencombo.Visible = False
                SubCombo.Visible = False
                credit.Visible = False
                debit.Visible = False
                DESCRIPTION.Visible = True
                DESCRIPTION.ZOrder
                'DESCRIPTION.Height = Grid1.CellHeight
                DESCRIPTION.SetFocus
            End If
        End If
        lastrow = grid1.Row
        lastcol = grid1.Col
    End If
End If
End If

End Sub


Private Sub Label5_Click()
frmVoucherHistory.Show
End Sub

Private Sub SubCombo_Change()
    
    
    If grid1.Col = lastcol And grid1.Row = lastrow Then
        grid1.text = SubCombo.text
    End If
    
    
    
    
End Sub
Private Sub SubCombo_Click()
    grid1.text = SubCombo.text
End Sub
Private Sub SubCombo_GotFocus()

   tmplistindex = 0
   Dim X(5000) As String
   Gencombo.Visible = False
   SubCombo.Clear
   tmpsubledger.text = grid1.text
   
   
      
    grid1.Col = 0
   If grid1.text <> "" Then
       
       
       If rs_S.State = 1 Then rs_S.close
       rs_S.Open "Select SUBLEDGER,gledger  from  SLEDGER where " & stringyear & " and gledger='" + Trim(grid1.text) + "'  ORDER BY gledger,SUBLEDGER", CCON, adOpenStatic, adLockReadOnly
       If Not rs_S.EOF Then
        Dim X1  As Integer
          X1 = 0
          CV = 0
          Do While Not rs_S.EOF
             
           If rs_S!gledger = Trim(grid1.text) Then
             SubCombo.AddItem rs_S(0)
           End If
             
             
             
             If rs_S(0) = tmpsubledger.text Then
                 tmplistindex = CV
                 CV = X1
             End If
             X1 = X1 + 1
             rs_S.MoveNext
         Loop

    End If
    
    
    grid1.Col = CCC
    grid1.Col = 1
    If SubCombo.ListCount = 0 Then
        grid1.Col = 2
        Grid1_Click
    End If
    SubCombo.text = tmpsubledger.text
    If CV >= 1 Then
       SubCombo.ListIndex = CV
    End If
    
     Set cmd.ActiveConnection = Nothing
    
End If

End Sub


Private Sub SubCombo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim changed As Boolean
    Dim valid As Boolean
    Dim HIT As Boolean
    changed = False
    HIT = False
    valid = False
    RRR = grid1.Row
    CCC = grid1.Col
    grid1.Row = lastrow
    grid1.Col = 0
    If RS.State = 1 Then
        RS.close
    End If
    
    If Trim(SubCombo.text) <> "" And Trim(grid1.text) <> "" Then
        sendkeys "{down}"
        RS.Open "select * from SLEDGER where " & stringyear & " and subledger='" + Trim(SubCombo.text) + "' and gledger='" + Trim(grid1.text) + "'", con, adOpenForwardOnly, adLockReadOnly, adCmdText
        If RS.BOF Or RS.EOF Then
            SubCombo.Visible = True
            SubCombo.SetFocus
        Else
            valid = True
        End If
        RS.close
    Else
        If Trim(SubCombo.text) = "" Then
            SubCombo.Visible = True
            SubCombo.SetFocus
       End If
    End If
    grid1.Row = RRR
    grid1.Col = CCC
    If valid Then
        If grid1.Col = lastcol And grid1.Row = lastrow Then
            grid1.Col = 0
            If Trim(grid1.Col) <> "" Then
                HIT = True
            End If
            grid1.Col = 3
            If grid1.text <> "" Then
            
            Else
                grid1.Col = 2
            End If
            changed = True
        End If
            SubCombo.text = ""
        If changed And HIT Then
            Grid1_Click
        End If
    End If
End If
End Sub
Private Sub SubCombo_LostFocus()
Dim RRR, CCC As Integer
'    If rs.State = 1 Then
'        rs.Close
'    End If
'    RRR = Grid1.row
'    CCC = Grid1.col
'    Grid1.row = lastrow
'    Grid1.col = 0
'    If Trim(SubCombo.Text) <> "" And Trim(Grid1.Text) <> "" Then
'        rs.Open "select * from SLEDGER where " & stringyear & " and subledger='" + Trim(SubCombo.Text) + "' and gledger='" + Trim(Grid1.Text) + "'", CON, adOpenDynamic, adLockReadOnly, adCmdText
'        If rs.BOF Or rs.EOF Then
'            SubCombo.Visible = True
'            SubCombo.SetFocus
'        Else
         SubCombo.Visible = False
'        End If
'        rs.Close
'    Else
'        If Trim(SubCombo.Text) <> "" Then
'            SubCombo.Visible = True
'            SubCombo.SetFocus
'        End If
'    End If
'    Grid1.row = RRR
'    Grid1.col = CCC
  SubCombo.text = ""
End Sub
Sub checkData()
    Set RS = New ADODB.Recordset
    RS.Open "Select * from setup1 where " & stringyear & " and yarfrom>datevalue('" + vdate.text + "') and yarto<=datevalue('" + vdate.text + "')", con, adOpenDynamic, adLockReadOnly, adCmdText
    If RS.BOF = True Then
       vdate.SetFocus
    End If
End Sub
Private Sub vdate_GotFocus()
HIT
End Sub

Private Sub vdate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        Set RS = New ADODB.Recordset
        RS.Open "setup1", con
        
        If CDate(vdate.text) >= RS!yarfrom And CDate(vdate.text) <= RS!yarto Then
        Else
           MsgBox "Enter Valid Date ....", vbInformation
           vdate.SetFocus
           Exit Sub
        End If
        RS.close
 
    
    
        If addmode = True Or Edit = True Then
           grid1.Row = 1
           grid1.Col = 0
           grid1.SetFocus
           'Print 1111
           Grid1_Click
        Else
           'SendKeys "{tab}"
           vno.SetFocus
        End If
    End If
End Sub
    Sub vdate_LostFocus()
         
    
    If Not IsDate(vdate.text) Then
       vdate.SetFocus
       Exit Sub
    End If
    
    
    If addmode = True Then
        If RS.State = 1 Then RS.close
        Set RS = New ADODB.Recordset
        RS.Open "Select max(VoucherNumber) from vouchers where " & stringyear & " and vouchertype='" + Trim(vtype.text) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & vdate & "',103)", con, adOpenDynamic, adLockReadOnly, adCmdText
        If Not RS.BOF Then
            If Not RS(0) Then
                c = RS(0)
            Else
                c = 0
            End If
        Else
            c = 0
        End If
        RS.close
        vno.text = c + 1
    Else
       If varVtype = "" Then Exit Sub
       If varVdate = "" Then Exit Sub
       If searchmode = True Then Exit Sub
       If Trim(vtype.text) <> Trim(varVtype) Or CDate(Format(varVdate, "dd/mm/yyyy")) <> Format(vdate, "dd/mm/yyyy") Then
            If RS.State = 1 Then RS.close
            Set RS = New ADODB.Recordset
            RS.Open "Select max(VoucherNumber) from vouchers where " & stringyear & " and vouchertype='" + Trim(vtype.text) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & vdate & "',103)", con, adOpenDynamic, adLockReadOnly, adCmdText
            If Not RS.BOF Then
                If Not RS(0) Then
                    c = RS(0)
                Else
                    c = 0
                End If
            Else
                c = 0
            End If
            RS.close
            vno.text = c + 1
       End If
   End If
   
   
   If Val(vno) <> 0 Then
   
    If rs1.State = 1 Then rs1.close
    rs1.Open "Select top 1 * from vouchers where " & stringyear & " and VoucherType='" + Trim(vtype.text) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & vdate & "',103)  and vouchernumber=" & Trim(vno.text) & " order by vsno", con, adOpenDynamic, adLockOptimistic, adCmdText  'and VoucherNumber=" + Val(Trim(vno.Text)) + "", con, adOpenDynamic, adLockOptimistic, adCmdText
    If rs1.EOF = False Then
       v_vtype = rs1!VoucherType
       v_vdate = rs1!voucherDATE
       v_vnumber = rs1!VOUCHERNUMBER
     End If
   
   End If

   
End Sub
Private Sub vno_GotFocus()
      vno.SetFocus
      
      sendkeys "{END}"
      sendkeys "+{HOME}"
End Sub

Private Sub vno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If searchmode = True Then
    vtype.SetFocus
    Exit Sub
  End If
  
  
  'SendKeys ("{TAB}")
  
End If
End Sub
Sub vno_LostFocus()

If grid1.Enabled = True Then
grid1.SetFocus
sendkeys "{UP}"
End If

If Val(vno) <> 0 Then
   
   If RS.State = 1 Then RS.close
   RS.Open "Select top 1000 * from vouchers where " & stringyear & " and VoucherType='" + Trim(vtype.text) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & vdate & "',103)  and vouchernumber=" & Trim(vno.text) & " order by vsno", con, adOpenDynamic, adLockOptimistic, adCmdText  'and VoucherNumber=" + Val(Trim(vno.Text)) + "", con, adOpenDynamic, adLockOptimistic, adCmdText
   If RS.EOF = False Then
      
      'grid1.Enabled = False
      Commanddelete.Enabled = False
   
      txtpaytype.text = RS!paytype & ""
      vtype.text = RS!VoucherType & ""
      
      v_vtype = RS!VoucherType
      v_vdate = RS!voucherDATE
      v_vnumber = RS!VOUCHERNUMBER
      
   End If
     
     For I = 1 To grid1.rows - 1
       grid1.Row = I
       If I Mod 2 = 0 Then
          grid1.text = "."
       Else
          For J = 0 To 6
             grid1.Col = J
            grid1.text = ""
          Next J
          grid1.Col = 0
         End If
         
   
         
   Next I

   

   totaldebit.text = 0
   totalcredit.text = 0

   Dim rows As Integer
   rows = 1
   
   If Not RS.BOF Then
      
      I = 1
    
      Do While Not RS.EOF
      
      
      
      
        grid1.Row = I
        grid1.Col = 0
        grid1.text = RS(3)
        grid1.Col = 1
         If IsNull(RS(4)) Then
              grid1.text = ""
         Else
              grid1.text = RS(4)
         End If
         grid1.Col = 2
         If RS(6) = "D" Then
            grid1.text = Format(RS(5), "0.00")
            totaldebit.text = Val(totaldebit.text) + RS(5)
            totaldebit.text = Format(totaldebit.text, "0.00")
         Else
            grid1.Col = 3
            grid1.text = Format(RS(5), "0.00")
            totalcredit.text = Val(totalcredit.text) + RS(5)
            totalcredit.text = Format(totalcredit.text, "0.00")
         End If
         grid1.Col = 4
         grid1.text = IIf(IsNull(RS(7)), ".", RS(7))
         
         
         grid1.Col = 5
         
         If Not IsNull(RS!billdate) Then
           If IsDate(RS!billdate) Then
              grid1.text = RS!billdate
           End If
         End If
            
          
         grid1.Col = 6
         grid1.text = IIf(IsNull(RS!UserName), ".", RS!UserName)
         
         grid1.Row = grid1.Row + 1
         grid1.Col = 0
         grid1.text = RS(9) & ""
         grid1.TextMatrix(grid1.Row, 7) = RS!vsno
         
      
         
         I = I + 2
         If Not RS.EOF Then
            RS.MoveNext
         End If
         
         
     Loop
     
     
     
     '---------------------------------------------------------------------------
     
     Dim ss_

     ss_ = Right(session, 2)

     If Val(ss_) >= 24 Then

       Dim rs22_ As ADODB.Recordset
       Set rs22_ = New ADODB.Recordset

       rs22_.Open "select Checked_YesNo,VoucherID from  VOUCHERS_Main where " & stringyear & " and  vouchertype='" + Trim(vtype.text) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & vdate & "',103) And vouchernumber = " + vno, con
       If rs22_.EOF = False Then
          If rs22_(0) = False Then
             txtchecked.text = "n"
             ''txtVNumber.text = rs22_(1)
          Else
             txtchecked.text = "y"
          End If
       End If
   
     End If
     
     '---------------------------------------------------------------------------
     
     
     Gencombo.Visible = False
     SubCombo.Visible = False
     credit.Visible = False
     debit.Visible = False
     crno.Visible = False
     grid1.Row = 1
     grid1.Col = 0
  
  End If
  RS.close
  grid1.Row = 1
  grid1.Col = 0
  
  ''If searchmode = False Then Grid1_Click
     'Print 1
  ''Else
  ''   vno.SetFocus
  
  End If

End Sub
Sub editView()


If Val(vno) <> 0 Then
   
   If RS.State = 1 Then RS.close
   RS.Open "Select top 1000 * from vouchers where " & stringyear & " and VoucherType='" + Trim(vtype.text) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & vdate & "',103)  and vouchernumber=" & Trim(vno.text) & " order by vsno", con, adOpenDynamic, adLockOptimistic, adCmdText  'and VoucherNumber=" + Val(Trim(vno.Text)) + "", con, adOpenDynamic, adLockOptimistic, adCmdText
   If RS.EOF = False Then
      txtpaytype.text = RS!paytype & ""
      vtype.text = RS!VoucherType & ""
   End If
     
     For I = 1 To grid1.rows - 1
       grid1.CellBackColor = vbWhite
       grid1.Row = I
       If I Mod 2 = 0 Then
          grid1.text = "."
       Else
          For J = 0 To 6
             grid1.Col = J
            grid1.text = ""
          Next J
          grid1.Col = 0
         End If
         
   Next I

   

   totaldebit.text = 0
   totalcredit.text = 0

   Dim rows As Integer
   rows = 1
   
   If Not RS.BOF Then
      
      I = 1
    
      Do While Not RS.EOF
      
      
      
      
        grid1.Row = I
        grid1.Col = 0
        grid1.text = RS(3)
        grid1.Col = 1
         If IsNull(RS(4)) Then
              grid1.text = ""
         Else
              grid1.text = RS(4)
         End If
         grid1.Col = 2
         If RS(6) = "D" Then
            grid1.text = Format(RS(5), "0.00")
            totaldebit.text = Val(totaldebit.text) + RS(5)
            totaldebit.text = Format(totaldebit.text, "0.00")
         Else
            grid1.Col = 3
            grid1.text = Format(RS(5), "0.00")
            totalcredit.text = Val(totalcredit.text) + RS(5)
            totalcredit.text = Format(totalcredit.text, "0.00")
         End If
         grid1.Col = 4
         grid1.text = IIf(IsNull(RS(7)), ".", RS(7))
         
         
         grid1.Col = 5
         
         If Not IsNull(RS!billdate) Then
           If IsDate(RS!billdate) Then
              grid1.text = RS!billdate
           End If
         End If
            
          
         grid1.Col = 6
         grid1.text = IIf(IsNull(RS!UserName), ".", RS!UserName)
         
         grid1.Row = grid1.Row + 1
         grid1.Col = 0
         grid1.text = RS(9) & ""
         grid1.TextMatrix(grid1.Row, 7) = RS!vsno
         
      
         
         I = I + 2
         If Not RS.EOF Then
            RS.MoveNext
         End If
         
         
     Loop
     
     
     
     
     
     
     Gencombo.Visible = False
     SubCombo.Visible = False
     credit.Visible = False
     debit.Visible = False
     crno.Visible = False
     grid1.Row = 1
     grid1.Col = 0
  
  End If
  
  RS.close
  grid1.Row = 1
  grid1.Col = 0
 
  
End If


   If Val(vno) <> 0 Then
   
    If rs1.State = 1 Then rs1.close
    rs1.Open "Select top 1 * from vouchers where " & stringyear & " and VoucherType='" + Trim(vtype.text) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & vdate & "',103)  and vouchernumber=" & Trim(vno.text) & " order by vsno", con, adOpenDynamic, adLockOptimistic, adCmdText  'and VoucherNumber=" + Val(Trim(vno.Text)) + "", con, adOpenDynamic, adLockOptimistic, adCmdText
    If rs1.EOF = False Then
       v_vtype = rs1!VoucherType
       v_vdate = rs1!voucherDATE
       v_vnumber = rs1!VOUCHERNUMBER
     End If
   
   End If


End Sub
Private Sub vtype_Change()
        If RS.State = 1 Then RS.close
        Set RS = New ADODB.Recordset
        RS.Open "Select max(VoucherNumber) from vouchers where " & stringyear & " and vouchertype='" + Trim(vtype.text) + "' and voucherdate= CDate ('" + Format(vdate.text, "dd/mm/yy") + "')", con, adOpenDynamic, adLockReadOnly, adCmdText
        If Not RS.BOF Then
           If Not RS(0) Then
                c = RS(0)
           Else
                c = 0
           End If
        Else
           c = 0
        End If
        RS.close
        vno.text = Str(c + 1)
End Sub

Private Sub vtype_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.vdate.SetFocus
    End If
End Sub
Public Function getval()
    Grid1_Click
End Function

Sub vtype_LostFocus()

End Sub

