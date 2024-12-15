VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form CRNOTE 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7110
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   9510
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "CRNOTE.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   Begin MSMask.MaskEdBox textbox 
      Height          =   315
      Left            =   5700
      TabIndex        =   5
      Top             =   150
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.CommandButton Commandall 
      Caption         =   "All Books"
      Height          =   375
      Left            =   1050
      TabIndex        =   54
      Top             =   5190
      Width           =   945
   End
   Begin VB.CommandButton Commandother 
      Caption         =   "&Other"
      Height          =   375
      Left            =   210
      TabIndex        =   26
      Top             =   5190
      Width           =   795
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2985
      Left            =   60
      TabIndex        =   47
      Top             =   1200
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   5265
      _Version        =   393216
      FillStyle       =   1
      Appearance      =   0
   End
   Begin VB.ComboBox Bookname 
      Height          =   960
      Left            =   3420
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   42
      Top             =   2550
      Width           =   2295
   End
   Begin VB.ComboBox Bookcode 
      Height          =   765
      Left            =   480
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   41
      Top             =   2580
      Width           =   2355
   End
   Begin VB.PictureBox Picture5 
      Height          =   435
      Left            =   -30
      ScaleHeight     =   375
      ScaleWidth      =   9495
      TabIndex        =   16
      Top             =   5610
      Width           =   9555
      Begin VB.CommandButton Commandadd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   1230
         TabIndex        =   0
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton CommandReturn 
         Caption         =   "&Return"
         Height          =   375
         Left            =   8190
         TabIndex        =   23
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton CommandPrint 
         Caption         =   "&Print"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7230
         TabIndex        =   22
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton Commandsearch 
         Caption         =   "&Search"
         Height          =   375
         Left            =   6200
         TabIndex        =   21
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton Commanddelete 
         Caption         =   "De&lete"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5200
         TabIndex        =   20
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton Commandabandon 
         Caption         =   "Aba&ndon"
         Height          =   375
         Left            =   4260
         TabIndex        =   19
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton Commandsave 
         Caption         =   "Sa&ve"
         Height          =   375
         Left            =   3270
         TabIndex        =   28
         Top             =   30
         Width           =   800
      End
      Begin VB.CommandButton Commandedit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   2250
         TabIndex        =   18
         Top             =   30
         Width           =   800
      End
      Begin VB.CommandButton Commandhelp 
         Caption         =   "Help"
         Height          =   375
         Left            =   180
         TabIndex        =   17
         Top             =   0
         Visible         =   0   'False
         Width           =   800
      End
   End
   Begin MSMask.MaskEdBox through 
      Height          =   285
      Left            =   2970
      TabIndex        =   9
      Top             =   1020
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   50
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox I_DTOB 
      Height          =   285
      Left            =   4020
      TabIndex        =   4
      Top             =   420
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox weight 
      Height          =   285
      Left            =   8280
      TabIndex        =   15
      Top             =   1590
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox freight 
      Height          =   285
      Left            =   6810
      TabIndex        =   14
      Top             =   1590
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   15
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox bdated 
      Height          =   285
      Left            =   5550
      TabIndex        =   13
      Top             =   1590
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox biltno 
      Height          =   285
      Left            =   3660
      TabIndex        =   12
      Top             =   1590
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox bundles 
      Height          =   285
      Left            =   1080
      TabIndex        =   8
      Top             =   1020
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   15
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox I_OB 
      Height          =   285
      Left            =   2250
      TabIndex        =   3
      Top             =   420
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   15
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox i_dt 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   285
      Left            =   1230
      TabIndex        =   2
      Top             =   420
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox station 
      Height          =   285
      Left            =   60
      TabIndex        =   11
      Top             =   1590
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   50
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox tempmeb 
      Height          =   285
      Left            =   480
      TabIndex        =   43
      Top             =   2040
      Visible         =   0   'False
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox rate 
      Height          =   285
      Left            =   960
      TabIndex        =   44
      Top             =   4080
      Visible         =   0   'False
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox amount 
      Height          =   285
      Left            =   60
      TabIndex        =   45
      Top             =   4470
      Visible         =   0   'False
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox I_NO 
      Height          =   285
      Left            =   90
      TabIndex        =   1
      Top             =   420
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox through1 
      Height          =   285
      Left            =   6210
      TabIndex        =   10
      Top             =   1020
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   50
      PromptChar      =   "_"
   End
   Begin VB.ComboBox customercode 
      Height          =   1155
      Left            =   5700
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   6
      Top             =   150
      Visible         =   0   'False
      Width           =   3795
   End
   Begin VB.ComboBox Genledger 
      Height          =   315
      Left            =   5040
      Sorted          =   -1  'True
      TabIndex        =   46
      Top             =   -120
      Visible         =   0   'False
      Width           =   645
   End
   Begin MSMask.MaskEdBox marka 
      Height          =   285
      Left            =   60
      TabIndex        =   7
      Top             =   1005
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      PromptChar      =   "_"
   End
   Begin VB.Label labelbybanklbl 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "By Bank : "
      Height          =   255
      Left            =   2340
      TabIndex        =   56
      Top             =   5250
      Width           =   1200
   End
   Begin VB.Label labelbybank 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   3480
      TabIndex        =   55
      Top             =   5250
      Width           =   1200
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dated : "
      Height          =   285
      Left            =   5550
      TabIndex        =   27
      Top             =   1305
      Width           =   1290
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Marka : "
      Height          =   285
      Left            =   60
      TabIndex        =   33
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label18 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Total Quantity : "
      Height          =   255
      Left            =   3300
      TabIndex        =   53
      Top             =   4560
      Width           =   1470
   End
   Begin VB.Label tqu 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   3150
      TabIndex        =   52
      Top             =   4830
      Width           =   840
   End
   Begin VB.Label mgd 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   7980
      TabIndex        =   50
      Top             =   4860
      Width           =   1200
   End
   Begin VB.Label mna 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   7980
      TabIndex        =   49
      Top             =   5220
      Width           =   1200
   End
   Begin VB.Label mga 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   6780
      TabIndex        =   48
      Top             =   4860
      Width           =   1200
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dated : "
      Height          =   285
      Left            =   1230
      TabIndex        =   40
      Top             =   135
      Width           =   1020
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Invoice No. : "
      Height          =   285
      Left            =   60
      TabIndex        =   39
      Top             =   135
      Width           =   1155
   End
   Begin VB.Label label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cust Code : "
      Height          =   585
      Left            =   5070
      TabIndex        =   38
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Net Amount : "
      Height          =   255
      Left            =   6780
      TabIndex        =   37
      Top             =   5220
      Width           =   1200
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Gross Amount : "
      Height          =   255
      Left            =   6660
      TabIndex        =   36
      Top             =   4410
      Width           =   1260
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Order By : "
      Height          =   285
      Left            =   2250
      TabIndex        =   35
      Top             =   135
      Width           =   1785
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dated : "
      Height          =   285
      Left            =   4020
      TabIndex        =   34
      Top             =   135
      Width           =   1050
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bundle(s) : "
      Height          =   285
      Left            =   1050
      TabIndex        =   32
      Top             =   720
      Width           =   1905
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Through : "
      Height          =   285
      Left            =   2970
      TabIndex        =   31
      Top             =   720
      Width           =   6495
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Railway/Station : "
      Height          =   285
      Left            =   60
      TabIndex        =   30
      Top             =   1290
      Width           =   3630
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bilty No. : "
      Height          =   285
      Left            =   3660
      TabIndex        =   29
      Top             =   1305
      Width           =   1905
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Freight : "
      Height          =   285
      Left            =   6810
      TabIndex        =   25
      Top             =   1290
      Width           =   1455
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Weight : "
      Height          =   285
      Left            =   8250
      TabIndex        =   24
      Top             =   1290
      Width           =   1230
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Total Discount : "
      Height          =   255
      Left            =   7980
      TabIndex        =   51
      Top             =   4350
      Width           =   1290
   End
   Begin VB.Menu dd 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu Delete 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "CRNOTE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim kk As ADODB.Recordset
Dim con As ADODB.Connection
Dim I As Integer
Dim lastrow, lastcol As Integer
Dim VALIDRATE As Boolean
Dim maxrow As Integer
Public totalamount, totaldiscount As Double
Public otheramount, otherdiscount As Double
Dim autoscroll As Boolean
Public edit As Boolean
Dim addmode As Boolean
Dim addoredit As Boolean
Sub CREDITCalc()
'OTHERSALES.calc
     mga.Caption = Format(Round(totalamount, 2), "0.00")
     mgd.Caption = Format(Round(totaldiscount, 2), "0.00")
     mna.Caption = Format(Round((totalamount - totaldiscount + otheramount - otherdiscount), 2), "0.00")
End Sub
Sub CREDITAbandon()
 Me.Commandadd.Enabled = True
        Me.Commandedit.Enabled = True
        Me.Commandsearch.Enabled = True
        Me.Commandsave.Enabled = False
        Me.Commanddelete.Enabled = True
        Me.Commandabandon.Enabled = True
        Me.CommandPrint.Enabled = True
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
If kk.State = 1 Then
   kk.Close
End If
If edit = False Then
  '  kk.Open "select max(invoiceno) from CREDITA", con, adOpenStatic, adLockReadOnly, adCmdText
   ' If kk(0) > 0 Then
   '     kk.MoveLast
   '     Me.I_NO.Text = Trim(Str(kk(0) + 1))
   ' Else
    '    Me.I_NO.Text = "1"
    'End If
    'kk.Close
    End If
        Dim ctl As Control
        For Each ctl In Me.Controls
            If TypeOf ctl Is MaskEdBox Or TypeOf ctl Is textbox Or TypeOf ctl Is ComboBox Then
                If UCase(Trim(ctl.Name)) <> UCase(Trim("I_NO")) And UCase(Trim(ctl.Name)) <> UCase(Trim("Genledger")) And UCase(Trim(ctl.Name)) <> UCase(Trim("I_DTOB")) And UCase(Trim(ctl.Name)) <> UCase(Trim("I_DT")) And UCase(Trim(ctl.Name)) <> UCase(Trim("bdated")) Then
                    ctl.Text = ""
                End If
                ctl.Enabled = False
            End If
        Next
        For I = 1 To maxrow
           grid1.row = I
            For j = 1 To 8
                grid1.col = j
               grid1.Text = ""
           Next
        Next
       I_DTOB = "__/__/____"
        bdated = "__/__/____"
        tqu.Caption = ""
        mga.Caption = ""
        mgd.Caption = ""
        mna.Caption = ""
        labelbybank.Caption = ""
        maxrow = 0
        addoredit = False
        Unload OTHERSALES
End Sub
Public Function templost() As Boolean
    Dim check As Boolean
    Dim row, col As Integer
    Dim RRR, CCC As Integer
    Dim r, q, d As Double
    Dim mprevcol As Integer
    Dim mq As Currency, mr As Currency, mrot As Currency
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    If lastrow <= 0 Then
        templost = True
        Exit Function
    End If
    RRR = grid1.row
    CCC = grid1.col
    grid1.row = lastrow
    grid1.col = lastcol
    mprevcol = grid1.col
    Select Case grid1.col
            Case 1
                grid1.Text = tempmeb.Text
                '/*************************
                If rs.State = 1 Then
                    rs.Close
                End If
                rs.Open "books", con, adOpenKeyset, adLockReadOnly, adCmdTable
                row = grid1.row
                col = grid1.col
                If Trim(grid1.Text) <> "" Then
                    If Not rs.BOF Then
                        rs.MoveFirst
                        rs.Find "bookcode='" + Trim(grid1.Text) + "'"
                        If rs.EOF Then
                            tempmeb.Visible = True
                            tempmeb.SetFocus
                            rs.Close
                            templost = False
                            Exit Function
                        Else
                            grid1.Text = rs(0)
                            grid1.col = 2
                            grid1.Text = rs(1)
                         '   If Not edit Then
                                grid1.col = 3
                                If Trim(grid1.Text) = "" Then
                                    grid1.Text = 0
                                End If
                                q = Val(grid1.Text)
                                grid1.col = 5
                                If Trim(grid1.Text) = "" Then
                                grid1.Text = Format(rs(3), "0.00")            'rs(3)
                                r = rs(3)
                                End If
                                '/******************
                                Set kk = con.Execute("select DISCATEGORY from sledger where subledger='" + Trim(customercode.Text) + "'")
                                grid1.col = 6
                                If Trim(kk(0)) <> "" Then
                                    tempstr = Trim(kk(0))
                                    kk.Close
                                    Set kk = con.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(tempstr) + "' and groupcode='" + Trim(rs(2)) + "'")
                                    grid1.col = 4
                                    If kk.BOF Then
                                        GoTo abc
                                    End If
                                    grid1.Text = Format(kk(0), "0.00")
                                    grid1.col = 6
                                    grid1.Text = Format(kk(0), "0.00")
                                    d = kk(0)
                                Else
abc:
                                    grid1.col = 4
                                    grid1.Text = Format(rs(4), "0.00")
                                    grid1.col = 6
                                    grid1.Text = Format(rs(4), "0.00")
                                    d = rs(4)
                                End If
                                grid1.col = 7
                                grid1.Text = Format(Round(q * r, 2), "0.00")
                                grid1.col = 8
                                grid1.Text = Format(Round((q * r) * (d / 100), 2), "0.00")
                          '  End If
                            grid1.col = col
                            rs.Close
                        End If
                    End If
                End If
            Case 3, 5, 6
                If grid1.col <> 3 Then
                    grid1.Text = Format(Trim(tempmeb.Text), "0.00")
                Else
                    grid1.Text = Format(Trim(tempmeb.Text), "0")
                End If
                If Trim(grid1.Text) = "" Then
                    grid1.Text = 0
                End If
                row = grid1.row
                col = grid1.col
                grid1.col = 3
                q = Val(Trim(grid1.Text))
                grid1.col = 5
                r = Val(Trim(grid1.Text))
                grid1.col = 6
                d = Val(Trim(grid1.Text))
                grid1.col = 7
                grid1.Text = Format(Round(q * r, 2), "0.00")
                grid1.col = 8
                grid1.Text = Format(Round((q * r) * (d / 100), 2), "0.00")
                grid1.col = col
            Case 4
                grid1.Text = tempmeb.Text
                If Trim(grid1.Text) = "" Then
                    grid1.Text = 0
                End If
        End Select
        row = grid1.row
        col = grid1.col
        totalamount = 0
        totaldiscount = 0
        For I = 1 To maxrow
            grid1.row = I
            grid1.col = 7
            totalamount = totalamount + Val(Trim(grid1.Text))
            grid1.col = 8
            totaldiscount = totaldiscount + Val(Trim(grid1.Text))
        Next
        CREDITCalc
        Me.tqu.Caption = ""
        For I = 1 To maxrow
            grid1.col = 3
            grid1.row = I
            Me.tqu.Caption = Val(Trim(Me.tqu.Caption)) + Val(Trim(grid1.Text))
        Next
        grid1.row = RRR
        grid1.col = CCC
        templost = True
End Function

Private Sub bdated_LostFocus()
If Trim(bdated.Text) <> Trim("__/__/____") Then
   If Not checkdate(Trim(bdated.Text), bdated) Then
         bdated.SetFocus
    End If
End If
End Sub

Private Sub biltno_LostFocus()
biltno = UCase(biltno)
End Sub

Private Sub Bookcode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
End If

End Sub

Private Sub Bookname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        Dim mprevcol As Integer
        Dim mq As Currency, mr As Currency, mrot As Currency
        mprevcol = grid1.col
        Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
        Select Case grid1.col
            Case 2
                Dim row, col As Integer
                row = grid1.row
                col = grid1.col
                If Trim(Bookname.Text) = "" Then
                    grid1.col = 1
                    If Trim(grid1.Text) = "" Then
                        grid1.Text = Bookname.Text
                           Bookname.SetFocus
                        
                           'Commandother.SetFocus
                        Exit Sub
                    End If
                End If
                grid1.row = row
                grid1.col = col
                grid1.Text = Bookname.Text
                '/*************************
                If rs.State = 1 Then
                    rs.Close
                End If
                rs.Open "books", con, adOpenKeyset, adLockReadOnly, adCmdTable
                row = grid1.row
                col = grid1.col
                If Trim(grid1.Text) <> "" Then
                    If Not rs.BOF Then
                        rs.MoveFirst
                        rs.Find "bookname='" + Trim(grid1.Text) + "'"
                        If rs.EOF Then
                            Bookname.Visible = True
                            Bookname.SetFocus
                            rs.Close
                            Exit Sub
                        Else
                            
                            grid1.col = 1
                            grid1.Text = rs(0)
                            grid1.col = 2
                            grid1.Text = rs(1)
                        '    If Not edit Then
                                 grid1.col = 3
                                If Trim(grid1.Text) = "" Then
                                        grid1.Text = 0
                                End If
                                q = Val(grid1.Text)
                                grid1.col = 5
                                grid1.Text = Format(rs(3), "0.00")
                                r = rs(3)
                                '/******************
                                Set kk = con.Execute("select DISCATEGORY from sledger where subledger='" + Trim(customercode.Text) + "'")
                                grid1.col = 6
                                If Trim(kk(0)) <> "" Then
                                    tempstr = Trim(kk(0))
                                    kk.Close
                                    Set kk = con.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(tempstr) + "' and groupcode='" + Trim(rs(2)) + "'")
                                    grid1.col = 4
                                    If kk.BOF Then
                                        GoTo abc
                                    End If
                                    grid1.Text = Format(kk(0), "0.00")
                                    grid1.col = 6
                                    grid1.Text = Format(kk(0), "0.00")
                                    d = kk(0)
                                Else
abc:
                                    grid1.col = 4
                                    grid1.Text = Format(rs(4), "0.00")
                                    grid1.col = 6
                                    grid1.Text = Format(rs(4), "0.00")
                                    d = rs(4)
                                End If
                                grid1.col = 7
                                grid1.Text = Round(q * r, 2)
                                grid1.col = 8
                                grid1.Text = Round((q * r) * (d / 100), 2)
                         '   End If
                            grid1.col = col
                            rs.Close
                        End If
                    End If
                End If
        End Select
        row = grid1.row
        col = grid1.col
        totalamount = 0
        totaldiscount = 0
        For I = 1 To maxrow
            grid1.row = I
            grid1.col = 7
            totalamount = totalamount + Val(Trim(grid1.Text))
            grid1.col = 8
            totaldiscount = totaldiscount + Val(Trim(grid1.Text))
        Next
        CREDITCalc
        grid1.row = row
        grid1.col = col
        Select Case grid1.col
            Case 1
                grid1.col = 3
                grid1.SetFocus
                Grid1_Click
            Case 2
                grid1.col = 3
                grid1.SetFocus
                Grid1_Click
            Case 3, 4, 5
                grid1.col = grid1.col + 1
                grid1.SetFocus
                Grid1_Click
            Case 6
                grid1.col = 1
                grid1.row = grid1.row + 1
                grid1.SetFocus
                Grid1_Click
        End Select
    End If
End Sub
Private Sub Bookname_LostFocus()
    Bookname.Visible = False
End Sub
Private Sub bundles_LostFocus()
bundles = UCase(bundles)
End Sub
Private Sub Commandabandon_Click()
CREDITAbandon
End Sub
Private Sub Commandadd_Click()
CREDITAbandon
    Dim rs As ADODB.Recordset
    addoredit = True
    addmode = True
    Set rs = New ADODB.Recordset
    Dim TEMPNUM As Integer
    If edit = False Then
    If con.Execute("Select max(invoiceno) from CREDITA")(0) >= Val(Trim(Me.I_NO.Text)) Then
        Me.I_NO.Text = con.Execute("Select max(invoiceno) from CREDITA")(0) + 1
    End If
    End If
    Dim ctl As Control
    For Each ctl In Me.Controls
        If Not TypeOf ctl Is CommandButton Then
                ctl.Enabled = True
        End If
        'If UCase(Trim(ctl.Name)) = UCase(Trim(Me.I_NO.Name)) Then
        '   ctl.Enabled = False
        'End If
    Next
    Picture5.Enabled = True
    Commandother.Enabled = True
    Commandadd.Enabled = False
    Commanddelete.Enabled = False
    Commandedit.Enabled = False
    CommandPrint.Enabled = False
    Commandall.Enabled = True
    Commandsave.Enabled = True
    Commandsearch.Enabled = False
    grid1.Enabled = True
    Me.customercode.Enabled = True
End Sub
Private Sub Commandall_Click()
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim myvalue As String

If Trim(Me.customercode.Text) = "" Then
    MsgBox "Please Fill the customer detail "
    Exit Sub
End If

myvalue = InputBox("Please enter the quantity ", "Enter the quantity: ", "1")
    
If Len(myvalue) > 0 And Val(myvalue) > 0 Then
    
    
    
    grid1.Rows = 1
    grid1.Rows = 2
    grid1.col = 1
    grid1.row = 1
    If rs.State = 1 Then
        rs.Close
    End If
    rs.Open "select * from books order by BOOKCODE", con, adOpenKeyset, adLockReadOnly, adCmdText
    row = grid1.row
    col = grid1.col
    If Not rs.BOF Then
        rs.MoveFirst
        Do While Not rs.EOF
            grid1.col = 1
            grid1.Text = rs(0)
            grid1.col = 2
            grid1.Text = rs(1)
            grid1.col = 3
            If Trim(grid1.Text) = "" Then
                grid1.Text = Val(myvalue)
            End If
            q = Val(grid1.Text)
            grid1.col = 5
            grid1.Text = Format(rs(3), "0.00")            'rs(3)
            r = rs(3)
            '/******************
            Set kk = con.Execute("select DISCATEGORY from sledger where subledger='" + Trim(customercode.Text) + "'")
            grid1.col = 6
            If Trim(kk(0)) <> "" Then
                tempstr = Trim(kk(0))
                kk.Close
                Set kk = con.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(tempstr) + "' and groupcode='" + Trim(rs(2)) + "'")
                grid1.col = 4
                If kk.BOF Then
                    GoTo abc
                End If
                grid1.Text = Format(kk(0), "0.00")
                grid1.col = 6
                grid1.Text = Format(kk(0), "0.00")
                d = kk(0)
            Else
abc:
                grid1.col = 4
                grid1.Text = Format(rs(4), "0.00")
                grid1.col = 6
                grid1.Text = Format(rs(4), "0.00")
                d = rs(4)
            End If
            grid1.col = 7
            grid1.Text = Format(Round(q * r, 2), "0.00")
            grid1.col = 8
            grid1.Text = Format(Round((q * r) * (d / 100), 2), "0.00")
            If Not rs.EOF Then
                grid1.Rows = grid1.Rows + 1
                grid1.row = grid1.row + 1
                rs.MoveNext
            End If
        Loop
        '/**fghfghgh
        '    Grid1.col = col
    End If
    rs.Close
   ' row = Grid1.row
   ' col = Grid1.col
    totalamount = 0
    totaldiscount = 0
    Me.tqu.Caption = ""
    For I = 1 To grid1.Rows - 1
            grid1.row = I
            grid1.col = 7
            totalamount = totalamount + Val(Trim(grid1.Text))
            grid1.col = 8
            totaldiscount = totaldiscount + Val(Trim(grid1.Text))
            grid1.col = 3
            Me.tqu.Caption = Val(Trim(Me.tqu.Caption)) + Val(Trim(grid1.Text))
     Next
     maxrow = grid1.Rows - 1
Else
'Grid1_Click
Exit Sub
End If

CREDITCalc

End Sub

Private Sub Commanddelete_Click()
If MsgBox("Are You Sure, Delete it ?", vbYesNo) = vbNo Then
        Exit Sub
Else
                con.Execute ("delete * from CREDITA where INVOICENO = " + Trim(I_NO.Text))
                con.Execute ("delete * from CREDITB where INVOICENO = " + Trim(I_NO.Text))
                con.Execute ("delete * from CREDITC where INVOICENO = " + Trim(I_NO.Text))
                CREDITAbandon
End If
End Sub

Private Sub Commandedit_Click()
    Commandadd.Enabled = False
    Me.Commandedit.Enabled = False
    Picture5.Enabled = True
    Commandother.Enabled = True
    Commandadd.Enabled = False
    Commandedit.Enabled = False
    Commandall.Enabled = True
    Commandsave.Enabled = True
    Commanddelete.Enabled = False
    Commandsearch.Enabled = False
    CommandPrint.Enabled = False
    grid1.Enabled = True
    Me.customercode.Enabled = True
    edit = True
    I_NO_LostFocus
    i_dt.Enabled = True
    i_dt.SetFocus
    ' CREDITCtmp creation start
    con.Execute ("delete * from CREDITCtmp")
    con.Execute ("insert into CREDITCtmp  select * from CREDITC where INVOICENO = " + Trim(I_NO.Text))
    ' CREDITTMP creation end
    addoredit = False
    HIT
End Sub
Private Sub Commandother_Click()
    Me.Enabled = False
    OTHERSALES.Top = 0
    OTHERSALES.Left = 0
    'Unload OTHERSALES
    'Load OTHERSALES
       
    OTHERSALES.Show
End Sub
Private Sub CommandPrint_Click()
 Me.Commandadd.Enabled = True
        Me.Commandedit.Enabled = True
        Me.Commandsearch.Enabled = True
        Me.Commandsave.Enabled = False
        Me.Commanddelete.Enabled = True
        Me.Commandabandon.Enabled = True
        Me.CommandPrint.Enabled = True
    Dim called1, called2 As Boolean
    Dim maxline As Integer
    Dim netamount As Double
    Dim totalquantity As Long
    Dim T1, T2, T3, t4, t5, t6, t7, t8 As Integer
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    T1 = 10
    T2 = 25
    T3 = 40
    t4 = 55
    t5 = 70
    t6 = 85
    t7 = 100
    t8 = 115
    netamount = 0
    totalquantity = 0
    maxline = 50
    called1 = False
    called2 = False
    Dim line As Integer
    Dim rs1 As ADODB.Recordset
    Dim kkk As ADODB.Recordset
    Set kkk = New ADODB.Recordset
    Set rs1 = New ADODB.Recordset
    
    Open "" + VB.App.Path + "\print.TXT" For Output As #1
    line = 0
header:
    If kkk.State = 1 Then
        kkk.Close
    End If
    kkk.Open "select * from setup1", con, adOpenKeyset, adLockReadOnly, adCmdText
    If Not kkk.BOF Then
        Print #1, Chr(27) + Chr(15) + Chr(14)
        Print #1, Tab(T1); Chr(27) + Chr(15) + Chr(14); dspace(Trim(kkk!CNAME)) '
        Print #1, Tab(T2 - 7); Chr(27) + Chr(15); dspace(Trim(kkk!add1))
        Print #1, Tab(T3); Trim(kkk!phone1)
        line = line + 4
    End If
    If rs1.State = 1 Then
        rs1.Close
    End If
    rs1.Open "CREDITA", con, adOpenKeyset, adLockReadOnly, adCmdTable
    rs1.Find "invoiceno='" + Trim(Me.I_NO.Text) + "'"
    Print #1, repli("-", 150)
    line = line + 1
    If Not rs1.EOF Then
        Print #1, " To,"; Tab(T1 - 3); rs1!subledger; Tab(t5); "Invoice No. : "; Trim(rs1!INVOICENO); Tab(t8 + 5); "Dt. "; Tab(t8 + 12); rs1!INVOICEDATE 'Chr(27) + Chr(15);
            If kkk.State = 1 Then
                kkk.Close
            End If
            kkk.Open "select * from sledger where subledger='" + Trim(rs1!subledger) + "'", con, adOpenKeyset, adLockReadOnly, adCmdText
            If Not kkk.EOF Then
                Print #1, Tab(3); kkk!DESCFORINVOICE
                Print #1, Tab(3); kkk!ADDRESS1; Tab(t5); "Order by : "; Trim(rs1!ORDERBY); Tab(t8 + 5); "Dt. "; Tab(t8 + 12); rs1!ORDERDATE
                Print #1, Tab(3); kkk!DISTCODE; Tab(t5); "Bilty No.: "; Trim(rs1!biltyno); Tab(t8 + 5); "Dt. "; Tab(t8 + 12); rs1!BILTYDATE
                kkk.Close
                Print #1,
                Print #1, "Through  :"; Tab(12); Trim(rs1!through) + ", " + Trim(rs1!through1)
                Print #1, "Station  :"; Tab(12); Trim(rs1!station); Tab(t5); "Pvt. Mark   : "; Trim(rs1!marka)
                Print #1, "Freight  :"; Tab(12); Trim(rs1!freight); Tab(t5); "Weight      : "; Trim(rs1!weight); Tab(t7 + 7); "Bundle(s)   : "; Trim(rs1!bundles)
                Print #1, repli("-", 150)
                Print #1, "S.No."; Tab(11); "Book Description"; Tab(t5 - 3); "Quantity"; Tab(t6 + 4); "Rate"; Tab(t7 + 4); "Amount"; Tab(t8 + 9); "Net Amount"
                Print #1, repli("-", 150)
                line = line + 11
            End If
            If called1 Then
                GoTo printagain1
            End If
            If called2 Then
                GoTo printagain2
            End If
            If kk.State = 1 Then
                kk.Close
            End If
            kk.Open "select * from CREDITB where invoiceno=" + Trim(rs1!INVOICENO) + " order by discount,printorder", con, adOpenKeyset, adLockReadOnly, adCmdText
            If Not kk.BOF Then
                kk.MoveFirst
                Dim cdiscount As Double
                Dim sno As Integer
                Dim tdata As ADODB.Recordset
                Set tdata = New ADODB.Recordset
                sno = 1
                Do While Not kk.EOF
                    cdiscount = kk!discount
                    Do While kk!discount = cdiscount
                        tdata.Open "Select bookname from books where bookcode='" + Trim(kk!Bookcode) + "'", con, adOpenKeyset, adLockReadOnly, adCmdText
                        Print #1, rsets(Trim(Str(sno)), 4); Tab(11); Trim(tdata!Bookname); Tab(t5 - 4); rsets(Trim(Str(kk!quantity)), 5); Tab(t6 + 1); rsets(Trim(Format(Str(kk!rate), "0.00")), 8); Tab(t7 - 1); rsets(Trim(Format(Str(kk!amount), "0.00")), 12)
                        totalquantity = totalquantity + kk!quantity
                        line = line + 1
                        If line > maxline Then
                            called1 = True
                            line = 0
                            Print #1, Chr(12)
                            GoTo header
printagain1:
                            called1 = False
                        End If
                        tdata.Close
                        If Not kk.EOF Then
                            sno = sno + 1
                            kk.MoveNext
                        End If
                        If kk.EOF Then
                            Exit Do
                        End If
                    Loop
                        If line > maxline - 4 Then
                            called2 = True
                            line = 0
                            Print #1, Chr(12)
                            GoTo header
printagain2:
                            called2 = False
                        End If
                        
                        Print #1, Tab(t7); repli("-", 22)
                        tdata.Open "select sum(amount) from CREDITB where invoiceno=" + Trim(rs1!INVOICENO) + " and discount=" + Trim(Str(cdiscount)) + " group by discount", con, adOpenKeyset, adLockReadOnly, adCmdText
                        If Not tdata.BOF Then
                            Print #1, Tab(t7 - 1); rsets(Trim(Format(Str(Round(tdata(0), 2)), "0.00")), 12)
                            Print #1, Tab(t5); "Less Discount @ " + Trim(Format(Str(Round(cdiscount, 2)), "0.00")) + " %"; Tab(t7 - 1); rsets(Trim(Format(Str(Round(tdata(0) * cdiscount / 100, 2)), "0.00")), 12); Tab(t8 + 5); rsets(Trim(Format(Str(Round(tdata(0) - (tdata(0) * cdiscount / 100), 2)), "0.00")), 12)
                            netamount = netamount + Round(tdata(0) - (tdata(0) * cdiscount / 100), 2)
                        End If
                        tdata.Close
                        Print #1, Tab(t7); repli("-", 22)
                Loop
            End If
           End If
           Print #1, Tab(t5 - 4); rsets(Trim(Str(totalquantity)), 5); Tab(t8 + 5); rsets(Trim(Format(Str(Round(netamount, 2)), "0.00")), 12)
           Print #1, Tab(t6); repli("-", 22)
           If kk.State = 1 Then
                kk.Close
           End If
           kk.Open "Select * from CREDITC where invoiceno=" + Trim(Me.I_NO.Text), con, adOpenKeyset, adLockReadOnly, adCmdText
           If Not kk.BOF Then
            Do While Not kk.EOF
                If kk!amount > 0 Then
                    If Trim(kk!DebitorCredit) = Trim("Credit") Then
                        netamount = netamount + kk!amount
                    Else
                        netamount = netamount - kk!amount
                    End If
                    If kk!rate > 0 Then
                        Print #1, Tab(t5); Trim(kk!Text) + "    " + Trim(Format(Str(Round(kk!rate, 2)), "0.00")); Tab(t8 + 5); rsets(Trim(Format(Str(Round(kk!amount, 2)), "0.00")), 12)
                    Else
                        Print #1, Tab(t5); Trim(kk!Text); Tab(t8 + 5); rsets(Trim(Format(Str(Round(kk!amount, 2)), "0.00")), 12)
                    End If
                End If
                If Not kk.EOF Then
                    kk.MoveNext
                End If
            Loop
            Print #1, Tab(t6); repli("-", 22)
            Print #1, Tab(6); Chr(71) + "NET AMOUNT: "; Tab(t8 + 5); Chr(72) + rsets(Trim(Format(Str(Round(netamount, 2)), "0.00")), 12)
           End If
           kk.Close
           kk.Open "Select * from CREDITA where invoiceno=" + Trim(Me.I_NO.Text), con, adOpenKeyset, adLockReadOnly, adCmdText
           If Not kk.BOF Then
                If kk!txt1a <> 0 Then
                    Print #1, Tab(t5); kk!txt1; Tab(t8 + 5); rsets(Trim(Format(Str(Abs(Round(kk!txt1a, 2))), "0.00")), 12)
                    netamount = netamount + Round(kk!txt1a, 2)
                End If
                If kk!txt2a <> 0 Then
                    Print #1, Tab(t5); kk!txt2; Tab(t8 + 5); rsets(Trim(Format(Str(Abs(Round(kk!txt2a, 2))), "0.00")), 12)
                    netamount = netamount + Round(kk!txt2a, 2)
                End If
                If kk!baa <> 0 Then
                    Print #1, Tab(t5); "BY BANK "; Tab(t8 + 5); rsets(Trim(Format(Str(Abs(Round(kk!baa, 2))), "0.00")), 12)
                    netamount = netamount - Round(kk!baa, 2)
                End If
           End If
           Print #1, Tab(t6); repli("-", 22)
           Print #1, Tab(t5); Chr(71) + "BALANCE : "; Tab(t8 + 5); rsets(Trim(Format(Str(Round(netamount, 2)), "0.00")), 12) + Chr(72)
        
       ' PRINT THE FOOTER IN INVOICE START
            Print #1, Tab(0); repli("-", 120)
            Dim tempdata As ADODB.Recordset
            Set tempdata = New ADODB.Recordset
            Dim LEFTM As Integer
            LEFTM = 5
            tempdata.Open "setup1", con, adOpenKeyset, adLockReadOnly, adCmdTable
            Print #1, Tab(1); "E.& O.E"
            Print #1, Tab(LEFTM); tempdata!COURT; Tab(LEFTM + (paperwidth - ((Len(tempdata!COURT) + Len(tempdata!CNAME)) * 0.75))); "FOR " + Trim(tempdata!CNAME)
            Print #1, Tab(LEFTM); ""

       'PRINT THE FOOTER IN INVOICE END
       
       
        
        
        
        
        
        Close #1
        Me.Enabled = False
        viewinvoice.Left = 0
        viewinvoice.Top = 10
        viewinvoice.Show
End Sub
Private Sub Commandreturn_Click()
    Unload Me
    addoredit = False
     MainMenu.Toolbar1.Visible = True
End Sub
Private Sub Commandsave_Click()
    Dim SAVED As Boolean
    Dim LAMOUNT As Double
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    If MsgBox("Do you want to save it now ?", vbYesNo) = vbNo Then
        Exit Sub
    End If
    SAVED = False
        If Trim(I_NO.Text) <> "" And Trim(i_dt.Text) <> "" And Trim(customercode.Text) <> "" Then
            If edit Then
                con.Execute ("delete * from CREDITA where INVOICENO = " + Trim(I_NO.Text))
                con.Execute ("delete * from CREDITB where INVOICENO = " + Trim(I_NO.Text))
                con.Execute ("delete * from CREDITC where INVOICENO = " + Trim(I_NO.Text))
            End If
            If rs.State = 1 Then
                rs.Close
            End If
            LAMOUNT = 0
            rs.Open "CREDITA", con, adOpenKeyset, adLockOptimistic, adCmdTable
            If Not edit Then
again:
           '     If con.Execute("Select max(invoiceno) from CREDITA")(0) >= Val(Trim(Me.I_NO.Text)) Then
            '        Me.I_NO.Text = Str(Val(Trim(Me.I_NO.Text)) + 1)
            '        GoTo again
            '    End If
            End If
            rs.AddNew
            rs!INVOICENO = Val(Me.I_NO.Text)
            rs!INVOICEDATE = Me.i_dt.Text
            rs!Genledger = Trim(Me.Genledger.Text)
            rs!subledger = Trim(Me.customercode.Text)
            rs!ORDERBY = Trim(Me.I_OB.Text)
            If Trim(Me.I_DTOB) = Trim("__/__/____") Then
                rs!ORDERDATE = Date
            Else
                rs!ORDERDATE = Trim(Me.I_DTOB.Text)
            End If
            rs!marka = Trim(Me.marka.Text)
            rs!bundles = Trim(Me.bundles)
            rs!through = Trim(Me.through.Text)
            rs!through1 = Trim(Me.through1.Text)
            If Trim(Me.through1.Text) = "" Then
                rs!through1 = " "
            End If
            rs!station = Trim(Me.station.Text)
            rs!biltyno = Trim(Me.biltno.Text)
            If Trim(Me.bdated) = Trim("__/__/____") Then
                rs!BILTYDATE = Date
            Else
                rs!BILTYDATE = Me.bdated
            End If
            rs!freight = Me.freight
            rs!weight = Me.weight
            rs!netamount = Round(Val(Trim(Me.mna.Caption)), 2)
            rs!gamount = (Me.totalamount - Me.totaldiscount)
            rs!txt1 = Trim(OTHERSALES.T1TEXT.Text)
            rs!txt1a = Val(Trim(OTHERSALES.T1.Text))
            rs!txt2 = Trim(OTHERSALES.T2TEXT.Text)
            rs!txt2a = Val(Trim(OTHERSALES.T2.Text))
            rs!baa = Val(Trim(OTHERSALES.T3TEXT.Text))
'err1:
          '  If Not edit Then
          '      If con.Execute("Select max(invoiceno) from CREDITA")(0) >= Val(Trim(Me.I_NO.Text)) Then
          '          Me.I_NO.Text = Str(Val(Trim(Me.I_NO.Text)) + 1)
          '      End If
          '  End If
          '  rs!INVOICENO = Val(Me.I_NO.Text)
          '  On Error GoTo err1
            rs.Update
            On Error GoTo 0
            rs.Close
            rs.Open "CREDITB", con, adOpenKeyset, adLockOptimistic, adCmdTable
            Dim I As Integer
            RRRR = grid1.row
            CCCC = grid1.col
            For I = 1 To maxrow
                grid1.row = I
                grid1.col = 1
                If Trim(grid1.Text) <> "" Then
                    grid1.col = 3
                    If Val(Trim(grid1.Text)) > 0 Then
                        rs.AddNew
                        grid1.col = 1
                        rs!INVOICENO = Val(Me.I_NO.Text)
                        rs!INVOICEDATE = Me.i_dt.Text
                        rs!Genledger = Trim(Me.Genledger.Text)
                        rs!subledger = Trim(Me.customercode.Text)
                        rs!Bookcode = Trim(grid1.Text)
                        grid1.col = 3
                        rs!quantity = Trim(grid1.Text)
                        grid1.col = 5
                        rs!rate = Trim(grid1.Text)
                        grid1.col = 7
                        rs!amount = Trim(grid1.Text)
                        LAMOUNT = Val(Trim(grid1.Text))
                        grid1.col = 4
                        rs!PRINTORDER = Trim(grid1.Text)
                        grid1.col = 6
                        rs!discount = Trim(grid1.Text)
                        grid1.col = 8
                        rs!netamount = LAMOUNT - Trim(grid1.Text)
                        LAMOUNT = 0
                        rs.Update
                    End If
                End If
            Next
            rs.Close
            rs.Open "CREDITC", con, adOpenKeyset, adLockOptimistic, adCmdTable
            '/******
                'Dim I, x As Integer
                Dim temprs As ADODB.Recordset
                Set temprs = New ADODB.Recordset
                For I = 1 To OTHERSALES.mrow
                    OTHERSALES.grid1.row = I
                    OTHERSALES.grid1.col = 0
                    If Trim(OTHERSALES.grid1.Text) <> "" Then
                        rs.AddNew
                        rs!INVOICENO = Val(Me.I_NO.Text)
                        rs!INVOICEDATE = Me.i_dt.Text
                        rs!gamount = (Me.totalamount - Me.totaldiscount)
                        rs!Text = Trim(OTHERSALES.grid1.Text)
                        If temprs.State = 1 Then
                            temprs.Close
                        End If
                        If edit Then
                        temprs.Open "select * from CREDITCtmp", con, adOpenKeyset, adLockReadOnly, adCmdText
                        If OTHERSALES.grid1.Text <> "" Then
                                temprs.Find "TEXT='" + Trim(OTHERSALES.grid1.Text) + "'"
                                rs!Genledger = temprs!Genledger
                                rs!subledger = temprs!subledger
                                rs!DebitorCredit = temprs!DebitorCredit
                                rs!RYN = temprs!RYN
                        End If
                        temprs.Close
                        
                        
                        
                        Else
                        
                        temprs.Open "select * from CREDITEND", con, adOpenKeyset, adLockReadOnly, adCmdText
                        If OTHERSALES.grid1.Text <> "" Then
                                temprs.Find "TEXT='" + Trim(OTHERSALES.grid1.Text) + "'"
                                rs!Genledger = temprs!Genledger
                                rs!subledger = temprs!subledger
                                rs!DebitorCredit = temprs!DebitorCredit
                                rs!RYN = temprs!RYN
                        End If
                        temprs.Close
                        End If
                        
                        OTHERSALES.grid1.col = 1
                        rs!rate = Val(Trim(OTHERSALES.grid1.Text))
                        If Val(Trim(OTHERSALES.grid1.Text)) > 0 Then
                            rs!amount = (Me.totalamount - Me.totaldiscount) * (Val(Trim(OTHERSALES.grid1.Text)) / 100)
                        Else
                        OTHERSALES.grid1.col = 2
                            rs!amount = Val(Trim(OTHERSALES.grid1.Text))
                        End If
                    rs.Update
                    End If
                Next
                rs.Close
                If addmode = True Then
                    rs.Open "tempinv", con, adOpenKeyset, adLockOptimistic, adCmdTable
                    If rs.BOF Then
                        rs.AddNew
                    End If
                    rs!In = Val(Me.I_NO.Text)
                    rs.Update
                    rs.Close
                End If
            SAVED = True
        End If
        If SAVED Then
            MsgBox "Record Saved"
        Me.customercode.Enabled = False
        Me.grid1.Enabled = False
        Me.Commandall.Enabled = False
        Me.Commandother.Enabled = False
        Me.Commandadd.Enabled = True
        Me.Commandedit.Enabled = True
        Me.Commandsearch.Enabled = True
        Me.Commandsave.Enabled = False
        Me.Commanddelete.Enabled = True
        Me.Commandabandon.Enabled = True
        Me.CommandPrint.Enabled = True
        End If
        addmode = False
    'If SAVED Then
    '    If MsgBox("Do  You Want To Print it Now", vbYesNo, "Do You Want To Print.....") = 6 Then
    '        CommandPrint_Click
     '   End If
      '  If edit Then
       '     Commandabandon_Click
        '    Exit Sub
      '  End If
      '  I_NO.Text = Val(I_NO.Text) + 1
      '  Unload OTHERSALES
      '  Me.totalamount = 0
      '  Me.totaldiscount = 0
      '  Me.mga.Caption = "0"
      '  Me.mgd.Caption = "0"
      '  Me.mna.Caption = "0"
      '  Me.tqu.Caption = "0"
      '  Dim ctl As Control
      '  For Each ctl In Me.Controls
      '      If TypeOf ctl Is MaskEdBox Or TypeOf ctl Is textbox Or TypeOf ctl Is ComboBox Then
       '         If UCase(Trim(ctl.Name)) <> UCase(Trim("I_NO")) And UCase(Trim(ctl.Name)) <> UCase(Trim("Genledger")) Then
       '             If Trim(UCase(ctl.Name)) <> Trim(UCase("I_DTOB")) And Trim(UCase(ctl.Name)) <> Trim(UCase("I_DT")) And Trim(UCase(ctl.Name)) <> Trim(UCase("bdated")) Then
' ctl.Text = ""
 '                   End If
  '              End If
   '         End If
    '    Next
     '   For I = 1 To maxrow
      '      grid1.row = I
      '      For j = 1 To 8
      '          grid1.col = j
       '         grid1.Text = ""
        '    Next
        'Next
        'maxrow = 0
        'Me.customercode.Enabled = True
    'End If
   End Sub

Private Sub Commandsearch_Click()
    Me.Enabled = False
    searchscreen.grid1.row = 0
    searchscreen.grid1.col = 0
    Call searchscreen.tempr(11, "invoice")
End Sub

Private Sub customercode_KeyPress(KeyAscii As Integer)
   ' If KeyAscii = 13 Then
  '  SendKeys "{DOWN}"
   ' SendKeys "{TAB}"
   ' marka.SetFocus
   ' End If
End Sub
Private Sub customercode_LostFocus()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "select * from sledger where gledger='SUNDRY DEBTORS'", con, adOpenKeyset, adLockReadOnly, adCmdText
    rs.Find "subledger='" + Trim(customercode.Text) + "'"
    If rs.EOF Then
        customercode.SetFocus
        HIT
        rs.Close
        Exit Sub
    End If
    rs.Close
    'Me.customercode.Text = Me.customercode.Index
    Me.textbox.Text = Me.customercode.Text
    Me.customercode.Visible = False
End Sub

Private Sub Delete_Click()
If grid1.row >= 1 Then
    grid1.SetFocus
    grid1.RemoveItem (grid1.row)
    If grid1.row > 1 Then
        grid1.row = grid1.row - 1
    End If
    Grid1_Click
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If UCase(Trim(VB.screen.ActiveControl.Name)) = UCase(Trim("CUSTOMERCODE")) Then
            SendKeys "{DOWN}"
            SendKeys "{TAB}"
        Else
            If UCase(Trim(VB.screen.ActiveControl.Name)) <> UCase(Trim("grid1")) And UCase(Trim(VB.screen.ActiveControl.Name)) <> UCase(Trim("tempmeb")) And UCase(Trim(VB.screen.ActiveControl.Name)) <> UCase(Trim("bookname")) And UCase(Trim(VB.screen.ActiveControl.Name)) <> UCase(Trim("weight")) Then
                 
                SendKeys ("{TAB}")
            End If
        End If
    End If
End Sub
Private Sub Form_Load()
    addmode = False
    edit = False
    autoscroll = True
    VALIDRATE = True
    maxrow = 0
    totalamount = 0
    totaldiscount = 0
    otheramount = 0
    otherdiscount = 0
    Set kk = New ADODB.Recordset
    grid1.Left = 90
    grid1.Top = 1900
    Set con = New ADODB.Connection
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Me.Top = 0
    Me.Left = 0
    grid1.Rows = 2
    grid1.Cols = 1
    grid1.Rows = 10
    grid1.Cols = 9
    grid1.row = 0
    grid1.col = 1
    grid1.Text = "Book Code "
    grid1.col = grid1.col + 1
    grid1.Text = "Book Name"
    grid1.col = grid1.col + 1
    grid1.Text = "Quantity"
    grid1.col = grid1.col + 1
    grid1.Text = "Print. Ord."
    grid1.col = grid1.col + 1
    grid1.Text = "Rate"
    grid1.col = grid1.col + 1
    grid1.Text = "Disc %"
    grid1.col = grid1.col + 1
    grid1.Text = "Amount"
    grid1.col = grid1.col + 1
    grid1.Text = "Disc. Amount"
    grid1.RowHeight(0) = grid1.CellHeight + 50
    grid1.ColWidth(0) = 200
    grid1.ColWidth(1) = 1200
    grid1.ColWidth(2) = 1850
    grid1.ColWidth(3) = 750
    grid1.ColWidth(4) = 750
    grid1.ColWidth(5) = 1200
    grid1.ColWidth(6) = 800
    grid1.ColWidth(7) = 1200
    grid1.ColWidth(8) = 1200
    Me.CommandPrint.Enabled = True
    con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + VB.App.Path + "\" + Trim(main.directory) + "\data.mdb"
    con.Open
    rs.Open "select * from books", con, adOpenKeyset, adLockReadOnly, adCmdText
    If Not rs.BOF Then
        Do While Not rs.EOF
            Me.Bookcode.AddItem rs(0)
            Me.Bookname.AddItem rs(1)
            If Not rs.EOF Then
                rs.MoveNext
            End If
        Loop
    End If
    rs.Close
    rs.Open "select * from gledger", con, adOpenKeyset, adLockReadOnly, adCmdText
    If Not rs.BOF Then
        Do While Not rs.EOF
            Me.Genledger.AddItem rs(1)
            If Not rs.EOF Then
                rs.MoveNext
            End If
        Loop
    End If
    rs.Close
    Genledger.Text = "SUNDRY DEBTORS"
    rs.Open "select * from sledger where gledger='" + Trim(Genledger.Text) + "'", con, adOpenKeyset, adLockReadOnly, adCmdText
    If Not rs.BOF Then
        Do While Not rs.EOF
            Me.customercode.AddItem rs(1)
            If Not rs.EOF Then
                rs.MoveNext
            End If
        Loop
    End If
    rs.Close
    Bookcode.Left = grid1.Left
    Bookcode.Visible = False
    Bookname.Visible = False
    grid1.Rows = 100
    For I = 1 To 99
        grid1.RowHeight(I) = 300
    Next
    Bookcode.Width = 1230
    Bookname.Width = 2830
    amount.Width = rate.Width
  
       rs.Open "tempinv", con, adOpenKeyset, adLockOptimistic, adCmdTable
          If Not rs.BOF Then
              Me.I_NO.Text = rs!In
              INVOICE.Enabled = True
              INVOICE.edit = False
              INVOICE.I_NO_LostFocus
              INVOICE.I_NO.Enabled = False
    lastrow = 0
    lastcol = 1
    Dim ctl As Control
    For Each ctl In INVOICE.Controls
        If Not TypeOf ctl Is CommandButton Then
                ctl.Enabled = False
        End If
        If UCase(Trim(ctl.Name)) = UCase(Trim(INVOICE.Commandother.Name)) Or UCase(Trim(ctl.Name)) = UCase(Trim(INVOICE.Commandall.Name)) Then
           ctl.Enabled = False
        End If
    Next
    INVOICE.Picture5.Enabled = True
    addoredit = False
              SendKeys "{TAB}"
          Else
              kk.Open "SELECT MAX(INVOICENO) FROM CREDITA", con, adOpenKeyset, adLockReadOnly, adCmdText
    If kk(0) <> "" Then
        Me.I_NO.Text = Trim(Str(kk(0) + 1))
    Else
        Me.I_NO.Text = "1"
    End If
    kk.Close
          End If
        rs.Close
    
Commanddelete.Enabled = True
Commandedit.Enabled = True
Commandsave.Enabled = False
    lastrow = 0
    lastcol = 1
  '  Dim ctl As Control
    For Each ctl In Me.Controls
        If Not TypeOf ctl Is CommandButton Then
                ctl.Enabled = False
        End If
        If UCase(Trim(ctl.Name)) = UCase(Trim(Me.Commandother.Name)) Or UCase(Trim(ctl.Name)) = UCase(Trim(Me.Commandall.Name)) Then
           ctl.Enabled = False
        End If
    Next
    Picture5.Enabled = True

End Sub
Private Sub freight_LostFocus()
freight = UCase(freight)
End Sub

Private Sub Grid1_Click()
If Trim(Me.customercode.Text) <> "" Then
Dim PREVROW As Integer
Dim prevcol As Integer
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
prevcol = grid1.col
PREVROW = grid1.row
If grid1.row > 1 Then
    grid1.row = grid1.row - 1
    grid1.col = 1
    If Trim(grid1.Text) <> "" Then
        grid1.row = PREVROW
        grid1.col = prevcol
        If Trim(Me.customercode.Text) <> "" Then
            If Me.customercode.Enabled = True Then
                Me.customercode.Enabled = False
            End If
            grid1.col = 1
            If prevcol > 1 And Trim(grid1.Text) = "" Then
                grid1.col = 2
                SendKeys Chr(13)
            Else
                grid1.col = prevcol
                SendKeys Chr(13)
            End If
        Else
            MsgBox "Please fill the customer detail first"
        End If
    End If
Else
    If Trim(Me.customercode.Text) <> "" Then
        If Me.customercode.Enabled = True Then
            Me.customercode.Enabled = False
        End If
        grid1.col = 1
        If prevcol > 1 And Trim(grid1.Text) = "" Then
            grid1.col = 2
            grid1.SetFocus
            SendKeys Chr(13)
        Else
        'IF GRID1.COL
            grid1.col = prevcol
            grid1.SetFocus
            SendKeys Chr(13)
        End If
        'SendKeys Chr(13)
    End If
End If
End If
End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
If Trim(Me.customercode.Text) <> "" Then
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
    If (KeyAscii = 13) Or (KeyAscii >= 32 And KeyAscii <= 126) Then
        If mwritemode = addmode Or mwritemode = EditMode Then
            Dim mprevcol As Integer
            
            Select Case grid1.col
            Case 1, 3, 4, 5, 6
                Bookname.Visible = False
                tempmeb.Visible = True: tempmeb.Enabled = True
                tempmeb.ZOrder
                If grid1.col <> 1 Then
                    If grid1.col <> 3 Then
                        tempmeb.Text = Format(grid1.Text, "0.00")
                        
                    Else
                        tempmeb.Text = Format(grid1.Text, "0")
                    End If
                   
                Else
                    tempmeb.Text = grid1.Text
                End If
                tempmeb.Width = grid1.ColWidth(grid1.col)
                tempmeb.Left = grid1.CellLeft + 80
                tempmeb.Top = grid1.Top + grid1.CellTop '- 50
            Case 2
                tempmeb.Visible = False
                Bookname.Visible = True: Bookname.Enabled = True
                Bookname.ZOrder
                Bookname.Text = grid1.Text
                Bookname.Top = grid1.Top + grid1.CellTop
                Bookname.Left = grid1.CellLeft + 80
                Bookname.Width = grid1.ColWidth(grid1.col)
            Case 6
                
            Case Else
                Bookname.Visible = False
                tempmeb.Visible = False
            End Select
            Select Case grid1.col
                Case 1, 3, 4, 5, 6
                    tempmeb.Mask = ""
                    tempmeb.MaxLength = 20
                Case 2
                    With Bookname
                        .Visible = True
                        .ZOrder
                    End With
                End Select
            Select Case grid1.col
            Case 2
                Bookname.SetFocus
                If KeyAscii <> 13 Then
                    SendKeys Chr(KeyAscii)
                End If
            Case 1, 3, 4, 5, 6
                mprevcol = grid1.col
                tempmeb.SetFocus
            Case Else
                If KeyAscii = 13 Then
                    SendKeys "{RIGHT}"
                End If
            End Select
        End If
    If maxrow < grid1.row Then
        maxrow = grid1.row
    End If
End If
    lastrow = grid1.row
    lastcol = grid1.col
End If
End Sub

Private Sub Grid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
   PopupMenu dd, , grid1.Left + X, grid1.Top + Y
End If
End Sub

Private Sub Grid1_Scroll()
If tempmeb.Visible <> True And Bookname.Visible <> True Then
        Bookname.Visible = False
        tempmeb.Visible = False
        grid1.SetFocus
End If
autoscroll = True
End Sub

Private Sub i_dt_GotFocus()
'    If Me.Enabled = True Then
'        datedlg.Top = i_dt.Top - 10
'        datedlg.Left = i_dt.Left - 200
'        Me.Enabled = False
'        datedlg.Calendar1.Value = i_dt
'        X = datedlg.GETDATE(Me.Name, i_dt.Name)
'        datedlg.Show
'    End If
End Sub
Private Sub i_dt_LostFocus()
If Trim(i_dt.Text) <> Trim("__/__/____") Then
    If Not checkdate(Trim(i_dt.Text), i_dt) Then
        i_dt.SetFocus
    End If
Else
    i_dt.SetFocus
    HIT
    Exit Sub
End If
End Sub


Private Sub I_DTOB_LostFocus()
If Trim(I_DTOB.Text) <> "__/__/____" Then
    If Not checkdate(Trim(I_DTOB.Text), I_DTOB) Then
        I_DTOB.SetFocus
    End If
End If
End Sub

Private Sub I_NO_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Then
    Else
        If KeyAscii <> 46 Then
            If KeyAscii <> 8 And KeyAscii <> 45 Then
                KeyAscii = 0
            End If
        End If
    End If
End Sub

Sub I_NO_LostFocus()
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
    If Trim(I_NO.Text) = "" Then
        MsgBox "Invoice  cannot be null"
        I_NO.SetFocus
    Else
    If rs.State = 1 Then
           rs.Close
    End If
        ''rs.Open "CREDITA", con, adOpenKeyset, adLockReadOnly, adCmdTable
        rs.Open "Select * from  CREDITA where INVOICENO = " + Trim(I_NO.Text) + "", con, adOpenStatic, adLockReadOnly
        ''rs.Find "INVOICENO='" + Trim(I_NO.Text) + "'"
        
        If rs.EOF Then
            If addoredit = False Then
                MsgBox "Invoice not found"
                Exit Sub
            End If
            Exit Sub
        End If
        If addoredit Then
            MsgBox "Invoice already exist..."
            I_NO.SetFocus
            HIT
            Exit Sub
        End If
        Dim ctl As Control
        For Each ctl In Me.Controls
            If Not TypeOf ctl Is CommandButton Then
                ctl.Enabled = True
            End If
        Next
        Me.Commandother.Enabled = True
        I_NO.Text = rs!INVOICENO
        Me.i_dt.Text = rs!INVOICEDATE
        Me.Genledger.Text = Trim(rs!Genledger)
        Me.customercode.Text = Trim(rs!subledger)
        Me.textbox.Text = Trim(rs!subledger)
        Me.I_OB.Text = Trim(rs!ORDERBY)
        Me.I_DTOB.Text = rs!ORDERDATE
        Me.marka.Text = Trim(rs!marka)
        Me.bundles = Trim(rs!bundles)
        Me.through.Text = rs!through
        Me.through1.Text = rs!through1
        Me.station.Text = rs!station
        Me.biltno.Text = Trim(rs!biltyno)
        Me.bdated = rs!BILTYDATE
        Me.freight = Trim(rs!freight)
        Me.weight = Trim(rs!weight)
        Me.labelbybank = Trim(rs!baa)
        mna.Caption = rs!netamount
        rs.Close
        'OTHERSALES.Show
        'OTHERSALES.Hide
'*/**/*/*/*/*//*/*
If rs.State = 1 Then
    rs.Close
End If
'Commandedit.Enabled = True
       ' 'rs.Open "CREDITB", con, adOpenKeyset, adLockReadOnly, adCmdTable
       rs.Open "Select * from CREDITB where INVOICENO =" + Trim(I_NO.Text) + "", con, adOpenStatic, adLockReadOnly
       '' rs.Find "INVOICENO='" + Trim(I_NO.Text) + "'"
    
        If Not rs.EOF Then
            grid1.row = 1
            grid1.col = 1
            Do While Not rs.EOF
               If Trim(rs!INVOICENO) = Trim(I_NO.Text) Then
               grid1.col = 1
                grid1.Text = Trim(rs!Bookcode)
                If kk.State = 1 Then
                    kk.Close
                End If
                kk.Open "select * from books where bookcode='" + Trim(rs!Bookcode) + "'", con, adOpenStatic, adLockReadOnly, adCmdText
                grid1.col = 2
                grid1.Text = Trim(kk!Bookname)
                grid1.col = 3
                grid1.Text = Trim(rs!quantity)
                grid1.col = 5
                grid1.Text = Format(Round(rs!rate, 2), "0.00")
                grid1.col = 7
                grid1.Text = Format(Round(rs!amount, 2), "0.00")
                grid1.col = 4
                grid1.Text = rs!PRINTORDER
                grid1.col = 6
                grid1.Text = Format(Round(rs!discount, 2), "0.00")
                grid1.col = 8
                grid1.Text = Format(Round(rs!amount * (rs!discount / 100), 2), "0.00")
                End If
                If Not rs.EOF Then
                    rs.MoveNext
                End If
                grid1.row = grid1.row + 1
                grid1.Rows = grid1.Rows + 1
            Loop
            maxrow = grid1.row
        '    Me.i_dt.SetFocus
        End If
                row = grid1.row
        col = grid1.col
        totalamount = 0
        totaldiscount = 0
        For I = 1 To maxrow
            grid1.row = I
            grid1.col = 7
            totalamount = totalamount + Val(Trim(grid1.Text))
            grid1.col = 8
            totaldiscount = totaldiscount + Val(Trim(grid1.Text))
        Next
     mga.Caption = Format(Round(totalamount, 2), "0.00")
     mgd.Caption = Format(Round(totaldiscount, 2), "0.00")
     Me.tqu.Caption = ""
        For I = 1 To maxrow
            grid1.col = 3
            grid1.row = I
            Me.tqu.Caption = Val(Trim(Me.tqu.Caption)) + Val(Trim(grid1.Text))
        Next
        grid1.row = RRR
        grid1.col = CCC
       ' templost = True
    End If
End Sub

Private Sub I_OB_LostFocus()
I_OB = UCase(I_OB)
End Sub

Private Sub marka_LostFocus()
marka = UCase(marka)
End Sub

Private Sub station_LostFocus()
station = UCase(station)
End Sub

Private Sub tempmeb_Change()
If grid1.col = 1 Or grid1.col = 2 Then
    grid1.Text = tempmeb.Text
Else
    If grid1.col = 3 Then
        grid1.Text = Format(tempmeb.Text, "0")
    Else
        grid1.Text = Format(tempmeb.Text, "0.00")
    End If
End If
End Sub
Private Sub tempmeb_GotFocus()
    HIT
End Sub
Private Sub tempmeb_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        Dim rs As ADODB.Recordset
           Set rs = New ADODB.Recordset
            Select Case grid1.col
                Case 1
                    If rs.State = 1 Then
                        rs.Close
                    End If
                    rs.Open "books", con, adOpenKeyset, adLockReadOnly, adCmdTable
                    If Not rs.BOF Then
                        rs.MoveFirst
                        rs.Find "bookcode='" + Trim(grid1.Text) + "'"
                        If rs.EOF And Trim(grid1.Text) <> "" Then
                            rs.Close
                            Exit Sub
                        Else
                            rs.Close
                        If Trim(grid1.Text) <> "" Then
                                grid1.col = 3
                            Else
                                grid1.col = 2
                            End If
                        End If
                    Else
                        If Trim(grid1.Text) <> "" Then
                            grid1.col = 3
                        Else
                            grid1.col = 2
                        End If
                    End If
                    grid1.SetFocus
                    Grid1_Click
                Case 3
                    If Val(tempmeb.Text) > 0 Then
                        grid1.col = grid1.col + 2
                        grid1.SetFocus
                        Grid1_Click
                    End If
                Case 4
                    grid1.col = grid1.col + 2
              '       SendKeys "{LEFT}"
                    grid1.SetFocus
                    Grid1_Click
                Case 5
                    If Val(tempmeb.Text) > 0 Then
                        grid1.col = grid1.col - 1
                        grid1.SetFocus
                        Grid1_Click
                    End If
                Case 6
                    grid1.col = 1
                    grid1.row = grid1.row + 1
                    grid1.Rows = grid1.Rows + 1
                    grid1.SetFocus
                    Grid1_Click
            End Select
        Else
        If grid1.col = 3 Or grid1.col = 4 Or grid1.col = 5 Or grid1.col = 6 Then
            If KeyAscii >= 48 And KeyAscii <= 57 Then
            Else
                If KeyAscii <> 46 Then
                    If KeyAscii <> 8 Then
                        KeyAscii = 0
                    End If
                End If
            End If
        End If
    End If
End Sub
Private Sub tempmeb_LostFocus()
    If templost Then
        tempmeb.Visible = False
    End If
End Sub
Private Sub textbox_GotFocus()
    Me.customercode.Enabled = True
    Me.customercode.Visible = True
  '  Me.customercode.Height = 1100
    Me.customercode.ZOrder
    Me.customercode.SetFocus
End Sub

Private Sub through_LostFocus()
through = UCase(through)
End Sub

Private Sub through1_LostFocus()
through1 = UCase(through1)
End Sub

Private Sub weight_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(Me.customercode.Text) <> "" Then
            grid1.col = 1
            grid1.row = 1
            Grid1_Click
        Else
            Me.textbox.SetFocus
            'Me.customercode.SetFocus
        End If
    End If
End Sub

Private Sub weight_LostFocus()
weight = UCase(weight)

End Sub
