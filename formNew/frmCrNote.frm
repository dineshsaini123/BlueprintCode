VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCrNote 
   Caption         =   "Credit Note"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10125
   Icon            =   "frmCrNote.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7185
   ScaleWidth      =   10125
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pic1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   135
      ScaleHeight     =   810
      ScaleWidth      =   9195
      TabIndex        =   28
      Top             =   6150
      Width           =   9195
      Begin VB.CommandButton cmdNHPrint 
         BackColor       =   &H00FFFFFF&
         Caption         =   "N&HPrint"
         Height          =   675
         Left            =   6675
         Picture         =   "frmCrNote.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   840
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.CommandButton Commandadd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add"
         Height          =   675
         Left            =   45
         Picture         =   "frmCrNote.frx":0BF0
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   60
         Width           =   1065
      End
      Begin VB.CommandButton CommandReturn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Return"
         Height          =   675
         Left            =   8040
         Picture         =   "frmCrNote.frx":17D4
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   60
         Width           =   1065
      End
      Begin VB.CommandButton CommandPrint 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Print"
         Height          =   675
         Left            =   6840
         Picture         =   "frmCrNote.frx":23B8
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   60
         Width           =   1185
      End
      Begin VB.CommandButton Commandsearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Search"
         Height          =   675
         Left            =   5580
         Picture         =   "frmCrNote.frx":2F9C
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   60
         Width           =   1185
      End
      Begin VB.CommandButton Commanddelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "De&lete"
         Height          =   675
         Left            =   4485
         Picture         =   "frmCrNote.frx":3B80
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   60
         Width           =   1065
      End
      Begin VB.CommandButton Commandedit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Edit"
         Height          =   675
         Left            =   1140
         Picture         =   "frmCrNote.frx":4764
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   60
         Width           =   1065
      End
      Begin VB.CommandButton Commandabandon 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Aba&ndon"
         Height          =   675
         Left            =   3390
         Picture         =   "frmCrNote.frx":4BA6
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   60
         Width           =   1065
      End
      Begin VB.CommandButton Commandsave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sa&ve"
         Enabled         =   0   'False
         Height          =   675
         Left            =   2235
         Picture         =   "frmCrNote.frx":5130
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   60
         Width           =   1110
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   5265
      Left            =   60
      TabIndex        =   10
      Top             =   180
      Width           =   9495
      Begin VB.TextBox DrCrText 
         DataField       =   "cnd"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   6300
         MaxLength       =   1
         TabIndex        =   20
         Text            =   "D"
         Top             =   3030
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox AText 
         DataField       =   "cnd"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   5370
         TabIndex        =   19
         Top             =   3030
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox CdateT 
         DataField       =   "Cnd"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   5010
         TabIndex        =   18
         Top             =   270
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.TextBox Text2 
         DataField       =   "n"
         Height          =   555
         Left            =   1770
         MaxLength       =   39
         TabIndex        =   4
         Top             =   1800
         Width           =   7065
      End
      Begin VB.ComboBox Scombo 
         Height          =   315
         Left            =   1770
         TabIndex        =   3
         Top             =   1350
         Width           =   4995
      End
      Begin VB.ComboBox GCombo 
         Height          =   315
         Left            =   1770
         TabIndex        =   2
         Top             =   840
         Width           =   4965
      End
      Begin VB.ComboBox gencombo1 
         Height          =   1545
         Left            =   360
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   17
         Top             =   2670
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.ComboBox Subcombo1 
         Height          =   1545
         Left            =   2790
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   16
         Top             =   2670
         Visible         =   0   'False
         Width           =   2385
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         DataField       =   "na"
         Height          =   315
         Left            =   6450
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   4920
         Width           =   1485
      End
      Begin VB.TextBox SText 
         DataField       =   "psld"
         Height          =   285
         Left            =   7350
         TabIndex        =   14
         Top             =   1410
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.TextBox GText 
         DataField       =   "pgld"
         Height          =   285
         Left            =   7350
         TabIndex        =   13
         Top             =   870
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.TextBox TCNN 
         DataField       =   "Cnn"
         Height          =   315
         Left            =   1800
         TabIndex        =   0
         Top             =   300
         Width           =   795
      End
      Begin VB.ComboBox cmbAgentName 
         Height          =   315
         Left            =   1800
         TabIndex        =   6
         Top             =   4920
         Width           =   2745
      End
      Begin VB.TextBox agn 
         DataField       =   "agentname"
         Enabled         =   0   'False
         Height          =   285
         Left            =   7380
         TabIndex        =   12
         Top             =   300
         Width           =   1395
      End
      Begin VB.ComboBox cmbgroup 
         Height          =   1545
         Left            =   6180
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   11
         Top             =   2580
         Visible         =   0   'False
         Width           =   2415
      End
      Begin MSMask.MaskEdBox Cbdate 
         Height          =   315
         Left            =   3690
         TabIndex        =   1
         Top             =   270
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid DGrid 
         Height          =   2445
         Left            =   120
         TabIndex        =   5
         Top             =   2400
         Width           =   9285
         _ExtentX        =   16378
         _ExtentY        =   4313
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         AllowUserResizing=   2
         RowSizingMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Credit Note No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   27
         Top             =   330
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3090
         TabIndex        =   26
         Top             =   300
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Narration"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   25
         Top             =   1890
         Width           =   795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Sub Ledger"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   24
         Top             =   1380
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Genral Ledger"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         TabIndex        =   23
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Net Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   5250
         TabIndex        =   22
         Top             =   4980
         Width           =   1005
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Agent Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   480
         TabIndex        =   21
         Top             =   4980
         Width           =   1050
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   120
      TabIndex        =   7
      Top             =   5520
      Width           =   2115
      Begin VB.CommandButton Command1 
         Caption         =   ">>"
         Height          =   345
         Left            =   1080
         TabIndex        =   9
         Top             =   60
         Width           =   765
      End
      Begin VB.CommandButton Command2 
         Caption         =   "<<"
         Height          =   345
         Left            =   150
         TabIndex        =   8
         Top             =   60
         Width           =   855
      End
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0078CFE9&
      BorderWidth     =   4
      Height          =   960
      Left            =   0
      Top             =   6060
      Width           =   9390
   End
End
Attribute VB_Name = "frmCrNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RS As ADODB.Recordset

'Dim con As New ADODB.Connection
Dim mvBookMark As Variant
Dim cmdAdd As Boolean
Dim cmdedit As Boolean
Dim LRC As Integer
Dim LCC As Integer
Dim Glastrow As Integer
Dim Datachange As Boolean
Dim Printheader As Boolean
Sub CNbandon()
        Me.Commandadd.Enabled = True
        Me.Commandedit.Enabled = True
        Me.Commandsearch.Enabled = True
        Me.Commandsave.Enabled = False
        Me.Commanddelete.Enabled = True
        Me.Commandabandon.Enabled = True
        Me.CommandPrint.Enabled = True
 End Sub
Sub Setgrid()
  DGrid.Row = 0
  DGrid.Col = 0
  DGrid.Text = "Gen Ledger"
  DGrid.Col = 1
  DGrid.Text = "Sub Ledger"
  DGrid.Col = 2
  'DGrid.Text = Format(DGrid.Text, "0.00")
  DGrid.Text = "Amount"
  DGrid.Col = 3
   DGrid.Text = "Amount"
  DGrid.Col = 4
  DGrid.Text = "Group"
  DGrid.ColWidth(0) = 3000
  DGrid.ColWidth(1) = 3000
  DGrid.Col = 0

    

End Sub

Sub Gridrefresh()
DoEvents
      Dim grs1 As New ADODB.Recordset
      If grs1.State = 1 Then grs1.close
         If TCNN.Text = "" Then
            grs1.Open "SELECT gld as GenLedger,Sld as SubLedger, a as Amount ,Dc as DebitCredit,groupcode FROM  CNF1B WHERE  CNN = 0", con, adOpenStatic
            Set DGrid.DataSource = grs1
            DGrid.Refresh
            Setgrid

         ElseIf cmdedit = True Then
             DoEvents
                grs1.Open "SELECT gld as GenLedger,Sld as SubLedger, a as Amount ,Dc as DebitCredit,groupcode FROM  CNF1B WHERE  CNN = 0", con, adOpenStatic
                DGrid.Rows = 99
                DGrid.TopRow = 1
                Setgrid
                
         ElseIf cmdAdd = True Then
         DoEvents
                grs1.Open "SELECT gld as GenLedger,Sld as SubLedger, a as Amount ,Dc as DebitCredit,groupcode FROM  CNF1B WHERE  CNN = 0", con, adOpenStatic
                DGrid.Rows = 99
                DGrid.TopRow = 1
                Setgrid
         Else
         DoEvents
                grs1.Open "SELECT gld as GenLedger,Sld as SubLedger, a as Amount ,Dc as DebitCredit,groupcode FROM  CNF1B WHERE  CNN = " + Trim(TCNN.Text) + "", con, adOpenStatic
                Set DGrid.DataSource = grs1
                DGrid.Refresh
                Setgrid
         End If
         
     If TCNN.Text <> "" Then
     DoEvents
        For I = 1 To grs1.RecordCount
          DGrid.Row = I
          DGrid.Col = 2
          DGrid.Text = Format(DGrid.Text, "0.00")
          DGrid.Refresh
        Next I
    End If
    DoEvents
    DGrid.Refresh
    DGrid.Col = 0
End Sub


Sub GridEdit()
      Dim grs1 As New ADODB.Recordset
      If grs1.State = 1 Then grs1.close
      If TCNN.Text = "" Then
         grs1.Open "SELECT gld as GenLedger,Sld as SubLedger, a as Amount ,Dc as DebitCredit FROM  TempCNF1B", con, adOpenStatic
      End If
End Sub

Private Sub agn_Change()
cmbAgentName.Text = agn.Text
End Sub

Private Sub AText_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 115 Then
   If MsgBox("Are You Sure, Delete it ?", vbYesNo) = vbYes Then
      If DGrid.Row >= 1 Then
           DGrid.RemoveItem DGrid.Row
           gencombo1.Text = ""
           gencombo1.Visible = False
           AText.Text = ""
           AText.Visible = False
           DGrid.SetFocus
       End If
   End If
End If
End Sub
Private Sub AText_KeyPress(KeyAscii As Integer)

If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 46 Or KeyAscii = 8 Or KeyAscii = 13 Then

Else
   KeyAscii = 0
End If
If KeyAscii = 13 Then
    If Val(AText.Text) <= 0 Then
        AText.SetFocus
        Exit Sub
    End If
     DGrid.Text = Format(AText.Text, "0.00")
     If AText = "" Then DGrid.Text = 0
     DGrid.Col = DGrid.Col + 1
     DGrid_Click
     Glastrow = DGrid.Row
End If


End Sub

Private Sub AText_LostFocus()
If Val(AText.Text) <= 0 Then
        AText.Visible = True
        AText.SetFocus
        Exit Sub
    End If
End Sub

Private Sub Cbdate_Change()

CdateT.Text = Cbdate.Text
End Sub

Private Sub Cbdate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{tab}"
End If


End Sub

Private Sub Cbdate_LostFocus()

If Trim(Cbdate.Text) = "__/__/____" Then
   Cbdate.SetFocus
   Exit Sub
End If

    Dim tRS1 As New ADODB.Recordset
    Dim trs2 As New ADODB.Recordset
    
    
    
    
    If trs2.State = 1 Then trs2.close
    trs2.Open "Select max(cnn) as mid from cnf1a where  cnd <= convert(smalldatetime,'" & Cbdate.Text & "',103)-1", con, adOpenDynamic, adLockOptimistic
    If trs2.RecordCount > 0 Then
        If IsNull(trs2!Mid) <> True Then
            If Val(TCNN.Text) >= trs2!Mid Then
               If tRS1.State = 1 Then tRS1.close
               tRS1.Open "Select  min(cnn)as m2 from cnf1a where cnd >= cdate('" & Cbdate.Text & "')+1", con, adOpenDynamic, adLockOptimistic
               If tRS1.RecordCount > 0 Then
                  If IsNull(tRS1!m2) <> True Then
                     If Val(TCNN.Text) <= tRS1!m2 Then
                     Else
                         MsgBox "Please select valid date.."
                         TCNN.SetFocus
                     End If
                  End If
               End If
            
            Else
            If i_dt.Enabled = True Then
               MsgBox "Please select valid date.."
               TCNN.SetFocus
            End If
            End If
        End If
     End If

If Trim(Cbdate.Text) <> "__/__/____" Then
    If Not checkdate(Trim(Cbdate.Text), Cbdate) Then
        Cbdate.SetFocus
    End If
End If
End Sub

Private Sub CdateT_Change()
 'On Error GoTo er1
 
 If CdateT = "" Then
  
    'Cbdate.Text = "__/__/____"
 Else
   Cbdate.Text = CdateT.Text
 End If
er1:  If err.Number = 380 Then
       Exit Sub
      End If
End Sub

Private Sub cmbAgentName_Change()
agn.Text = cmbAgentName.Text

End Sub

Private Sub cmbAgentName_LostFocus()

If cmbAgentName.Text = "" Then
   MsgBox "Enter a Agent Name.. "
   'cmbAgentName.SetFocus
   Exit Sub
Else
  Dim rs1 As New ADODB.Recordset
  rs1.Open "select rep  from SalesRepQry where rep='" & cmbAgentName.Text & "'", CON_blue
  If rs1.EOF = True Then
     MsgBox "Enter valid Agent Name.. "
     cmbAgentName.SetFocus
  End If
End If
End Sub

Private Sub cmbgroup_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
      Dim rs2 As New ADODB.Recordset
      Dim rs3 As New ADODB.Recordset
      If cmbgroup.Text <> "" Then
          Dim rs4 As New ADODB.Recordset
          rs4.Open "Select* from groups where groupcode='" + Trim(cmbgroup.Text) + "'", con, adOpenForwardOnly, adLockReadOnly, adCmdText
          If rs4.RecordCount <= 0 Then
             MsgBox "No valid group"
             cmbgroup.Visible = True
             cmbgroup.SetFocus
             Exit Sub
          End If
     End If
        DGrid.Text = cmbgroup.Text
        cmbgroup.Visible = False
        Glastrow = DGrid.Row
        DGrid.Row = DGrid.Row + 1
        LRC = LRC + 1
        DGrid.Col = 0
        DGrid_Click
       
        
   End If
End Sub

Private Sub cmdPrint_Click()

     printch = "cnf1a"
     ino = TCNN
     printch1 = "CNN"


Printheader = True
GenrateReport
 
End Sub

Private Sub Command1_Click()
On Error GoTo er
If Not RS.BOF Or Not RS.EOF Then
    RS.MoveNext
End If
If RS.EOF And RS.RecordCount > 0 Then
    Beep
    RS.MoveLast
  End If
  
er:   If err.Number = 3021 Then
         Exit Sub
      End If
  End Sub

Private Sub Command2_Click()
  Dim gs As ADODB.Recordset
  Set gs = New ADODB.Recordset
  If Not RS.BOF Then
     RS.MovePrevious
    
 
End If
 DoEvents
  If RS.BOF And RS.RecordCount > 0 Then
    Beep
    RS.MoveFirst
    

  End If
End Sub

Private Sub Command3_Click()
   RS.CancelBatch
  If mvBookMark > 0 Then
    RS.Bookmark = mvBookMark
  Else
    RS.MoveFirst
  End If
  Gridrefresh
  pic1.Visible = True
  Pic2.Visible = False
End Sub
Public Sub Commandabandon_Click()
 On Error Resume Next
  
''  If RS.RecordCount > 0 Then
''      'RS.CancelUpdate
''      'RS.MoveFirst
''  End If
  
  cmdAdd = False
  cmdedit = False
  Gridrefresh
  pic1.Visible = True
  Frame2.Enabled = True
  gencombo1.Visible = False
  Subcombo1.Visible = False
  AText.Visible = False
  DrCrText.Visible = False
  Frame1.Enabled = False
  CNbandon
' SetButton Commandadd, Commandedit, Commandsave, Commanddelete
End Sub

Private Sub Commandadd_Click()
 Dim rs6 As New ADODB.Recordset
 Me.Commandadd.Enabled = False
 Me.Commandedit.Enabled = False
 Me.Commandsearch.Enabled = False
 Me.Commandsave.Enabled = True
 Me.Commanddelete.Enabled = False
 Me.Commandabandon.Enabled = True
 Me.CommandPrint.Enabled = False
 Dim rs1  As New ADODB.Recordset
cmdAdd = True
cmdedit = False
LRC = 1
LCC = 0
Frame1.Enabled = True
TCNN.Enabled = True
TCNN.SetFocus

'With RS
'    .AddNew
'End With

CdateT.Text = ""
DGrid.Cols = 5
If cmdedit = False Then
     Dim trs As New ADODB.Recordset
     trs.Open "Select max(cnn)as mcnn from cnf1a", con, adOpenStatic, adCmdText
     If trs.RecordCount <= 1 And IsNull(trs!Mcnn) Then
       TCNN.Text = 1
     Else
       TCNN.Text = trs!Mcnn + 1
    End If
End If
DGrid.Rows = 100
DGrid.Cols = 0
DGrid.Cols = 5
Dim I
For I = 0 To 99
   DGrid.RowHeight(I) = 270
Next



Setgrid
DoEvents
GCombo.Text = "SUNDRY DEBTORS"
Dim rs2 As New ADODB.Recordset
rs2.Open "Select * from sledger where GLEDGER='" + Trim(GCombo.Text) + "'", con, adOpenForwardOnly, adLockReadOnly, adCmdText
If Not rs2.EOF Then
       Scombo.Clear
       Do While Not rs2.EOF
          Scombo.AddItem rs2(1)
          If Not rs2.EOF Then
            rs2.MoveNext
          End If
      Loop
Else
   Scombo.Clear
   Scombo.Text = ""
   Text2.SetFocus
   Exit Sub
End If
DGrid.Row = 1
DGrid.Col = 0
DGrid.TopRow = 1
Frame2.Enabled = False

End Sub

Private Sub Commanddelete_Click()
    On Error Resume Next

    '=================================
    

    Dim rs_h As New ADODB.Recordset
    Dim rs1 As ADODB.Recordset
    Set rs1 = New ADODB.Recordset
    If rs1.State = 1 Then rs1.close
    rs1.Open "select * from cnf1a where cnn=" & Trim(TCNN) & "", con
    If rs1.EOF = False Then
        If rs_h.State = 1 Then rs_h.close
        rs_h.Open "select * from cnf1a where cnn=" & Trim(TCNN) & "", con
        'If rs_h.Fields("Print_yes").Value = "y" Then
        If rs1!bAuthorized = True Then
           MsgBox "You can'nt change, Bill Already Locked !!", vbExclamation, "Alert"
           Exit Sub
        End If
        'End If
    End If

    
    '==================================
    



If RS.RecordCount > 0 Then
  If MsgBox("Are you sure.......", vbYesNo) = vbYes Then
  If TCNN.Text <> "" Then
       con.Execute "Delete * FROM  CNF1B WHERE  CNN=" + Trim(TCNN.Text) + ""
  End If
  With RS
    .delete
    
     If Not RS.BOF And Not RS.EOF Then RS.MoveFirst
     

     
     
  End With
    
  Gridrefresh
  Exit Sub


End If

End If


End Sub

 Sub Commandedit_Click()
 If RS.RecordCount > 0 Then
        Me.Commandadd.Enabled = False
        Me.Commandedit.Enabled = False
        Me.Commandsearch.Enabled = False
        Me.Commandsave.Enabled = True
        Me.Commanddelete.Enabled = False
        Me.Commandabandon.Enabled = True
        Me.CommandPrint.Enabled = False
        cmdAdd = False
        cmdedit = True
        TCNN.Enabled = False
        Frame1.Enabled = True
        Cbdate.SetFocus

        Gridrefresh
        
        If cmdedit = True Then
          Dim rs3 As New ADODB.Recordset
        rs3.Open "SELECT * FROM  CNF1B WHERE  CNN=" + Trim(TCNN.Text) + "", con, adOpenStatic, adLockOptimistic, adCmdText
          If rs3.RecordCount >= 0 Then
             DGrid.Rows = 99
             DGrid.TopRow = 1
             DGrid.Row = rs3.RecordCount + 1
          End If
    End If
    Frame2.Enabled = False
    VB.Screen.ActiveForm.ActiveControl.SelLength = 2
 End If
End Sub

Private Sub CommandPrint_Click()
 
     printch = "cnf1a"
     ino = TCNN
     printch1 = "CNN"
 
 
Printheader = False
GenrateReport
 
 
 
End Sub

Private Sub Commandreturn_Click()
'rs.Close
Unload Me
'MainMenu.Toolbar1.Visible = True
End Sub
Private Sub Commandsave_Click()
   
    Dim Grs As New ADODB.Recordset
    
    '==============================
    
    Dim rs_h As New ADODB.Recordset
    Dim rs1 As ADODB.Recordset
    Set rs1 = New ADODB.Recordset
    If rs1.State = 1 Then rs1.close
    rs1.Open "select * from cnf1a where cnn=" & Trim(TCNN) & "", con
    If rs1.EOF = False Then
        If rs_h.State = 1 Then rs_h.close
        rs_h.Open "select * from cnf1a where cnn=" & Trim(TCNN) & "", con
        'If rs_h.Fields("Print_yes").Value = "y" Then
        If rs1!bAuthorized = True Then
           MsgBox "You can'nt change, Bill Already Locked !!", vbExclamation, "Alert"
           Exit Sub
        End If
        'End If
    End If

   
   '==================================
   
   
   'Dim RS As ADODB.Recordset
   'Set RS = New ADODB.Recordset
   
   
   
   If TCNN.Text = "" Then
       Commandabandon_Click
       Exit Sub
   End If
    
   con.Execute "Delete  FROM  CNF1B WHERE  CNN=" + Trim(TCNN.Text) + ""
   TCNN.Enabled = True
   Grs.Open "cnf1b", con, adOpenDynamic, adLockPessimistic
   Dim sum As Double
   sum = 0
   Dim I, J As Integer
     I = 1
     DGrid.Row = 1
     DGrid.Col = 0
     If DGrid.Text = "" Then
         RS.CancelUpdate
         Exit Sub
     End If
  
     While DGrid.Text <> ""
        DGrid.Row = I
        Grs.AddNew
        Grs!cnn = Val(TCNN.Text)
        Grs!Cnd = Cbdate.Text
        Grs!groupcode = DGrid.TextMatrix(I, 4)
        For J = 0 To 3
          DGrid.Col = J
                If J = 0 Then
                   If DGrid.Text = "" Then
                      MsgBox "Please fill Gen Ledger"
                      DGrid_Click
                      Exit Sub
                       
                   Else
                     Grs!gld = DGrid.Text
                   End If
               End If
                If J = 1 Then
                    If DGrid.Text = "" Then
                       
                         Grs!sld = Null
                                               
                    Else
                         Grs!sld = DGrid.Text
                     End If
                   End If
                If J = 2 Then
                     If DGrid.Text = "" Then
                       
                        Grs!a = 0
                        
                     Else
                        Grs!a = Format(Val(DGrid.Text), "0.00")
                    End If
                 End If
                 If J = 3 Then
                 
                   If DGrid.Text = "" Then
                         MsgBox "Please fill Correct Entry  "
                         DGrid_Click
                         Exit Sub
                   Else
                         Grs!dc = DGrid.Text
                   End If
                   If DGrid.Text = "D" Then
                           DGrid.Col = 2
                          sum = sum + Format(Val(DGrid.Text), "0.00")
                    Else
                           DGrid.Col = 2
                           sum = sum - Format(Val(DGrid.Text), "0.00")
      
                   End If
                 End If
          Next J
          Grs.update
         
         I = I + 1
         DGrid.Row = I
         DGrid.Col = 0
    Wend
   
   If cmdAdd = True Then
   RS.AddNew
   RS!cnn = TCNN.Text
   RS!Cnd = Cbdate.Text
   RS!Pgld = GText.Text
   RS!agentname = cmbAgentName.Text
   If SText.Text = "" Then
      RS!psld = Null
   Else
      RS!psld = SText.Text
   End If

   If Text2.Text <> "" Then
   RS!n = Text2.Text
   End If
   If sum >= 0 Then
        RS!dc = "C"
   Else
       RS!dc = "D"
   End If
   
   RS!na = Abs(Format(sum, "0.00"))
   RS.update
   End If
   
    'con.Execute "Delete * FROM  CNF1B WHERE  CNN=" + Trim(TCNN.Text) + ""
    'con.Execute "Insert into cnf1b select * from tempcnf1b"
       Frame1.Enabled = False
       cmdAdd = False
       cmadd = False
       cmdedit = False
        Me.Commandadd.Enabled = True
        Me.Commandedit.Enabled = True
        Me.Commandsearch.Enabled = True
        Me.Commandsave.Enabled = False
        Me.Commanddelete.Enabled = True
        Me.Commandabandon.Enabled = True
        Me.CommandPrint.Enabled = True
        Gridrefresh
        gencombo1.Visible = False
        Subcombo1.Visible = False
        AText.Visible = False
        DrCrText.Visible = False
        
        Frame2.Enabled = True
End Sub

Private Sub Commandsearch_Click()


searchType = "inv"
popuplist10 "select CNN,CND,PSLD,NA from CNF1A order by CNN", con


End Sub

Private Sub DGrid_AfterUpdate()
Dim rs2 As New ADODB.Recordset
rs2.Open "Select sum(a) as tot from  tempcnf1b ", con, adOpenStatic, adCmdText
If rs2.RecordCount > 0 Then
If IsNull(rs2!Tot) = True Then
    Text3.Text = 0
 Else
   Text3.Text = rs2!Tot
End If
End If
End Sub
Private Sub Commandsearch_LostFocus()

If PopUpValue1 <> "" Then
     TCNN.Text = PopUpValue1
     TCNN_LostFocus

     PopUpValue1 = ""
End If
End Sub

Private Sub DGrid_Click()
If DGrid.Row > 0 Then
       Select Case DGrid.Col
    
           Case 0
                gencombo1.Text = DGrid.Text
                gencombo1.Visible = True
                Subcombo1.Visible = False
                AText.Visible = False
                DrCrText.Visible = False
                gencombo1.Move DGrid.CellLeft + DGrid.Left, DGrid.CellTop + DGrid.Top - 15, DGrid.CellWidth
                gencombo1.SetFocus

            Case 1
                Subcombo1.Text = DGrid.Text
                Subcombo1.Visible = True

                gencombo1.Visible = False
                AText.Visible = False
                DrCrText.Visible = False
                If Subcombo1.ListCount > 0 Then
                        'Subcombo1.Text = DGrid.Text
                        'Subcombo1.Visible = True
                    Subcombo1.Move DGrid.CellLeft + DGrid.Left, DGrid.CellTop + DGrid.Top - 15, DGrid.CellWidth
                    Subcombo1.SetFocus
                End If
    
            Case 2
                AText = DGrid.Text
                AText.Visible = True
                gencombo1.Visible = False
                Subcombo1.Visible = False
                DrCrText.Visible = False
                AText.Move DGrid.CellLeft + DGrid.Left, DGrid.CellTop + DGrid.Top - 15, DGrid.CellWidth
                AText.SetFocus
 
            Case 3
                DrCrText.Text = "D"
                DrCrText.Text = DGrid.Text
                If DGrid.Text = "" Then DrCrText.Text = "D"
                DrCrText.Visible = True
                gencombo1.Visible = False
                Subcombo1.Visible = False
                AText.Visible = False
                DrCrText.Move DGrid.CellLeft + DGrid.Left, DGrid.CellTop + DGrid.Top - 15, DGrid.CellWidth
                DrCrText.SetFocus
            
            
             Case 4
                cmbgroup.Text = DGrid.Text
                Subcombo1.Visible = False
                gencombo1.Visible = False
                cmbgroup.Visible = True
              
                AText.Visible = False
                DrCrText.Visible = False
                If cmbgroup.ListCount > 0 Then
                   cmbgroup.Move DGrid.CellLeft + DGrid.Left, DGrid.CellTop + DGrid.Top - 15, DGrid.CellWidth
                   cmbgroup.SetFocus
                End If
    
            
            End Select
   
  End If
    
    
    
    
End Sub

Private Sub DrCrText_Change()
DrCrText.Text = UCase(DrCrText.Text)
End Sub

Private Sub DrCrText_KeyPress(KeyAscii As Integer)




If KeyAscii = 13 Then
     If DrCrText.Text = "" Then
        MsgBox "please Enter  D or C."
        DrCrText.SetFocus
        Exit Sub
   End If
   
     Glastrow = DGrid.Row
     DGrid.Text = DrCrText.Text
     DGrid.Row = DGrid.Row + 1
     
     LRC = LRC + 1
     DGrid.Col = 0
     DGrid_Click
End If
End Sub

Private Sub Form_Activate()
Commandadd.Enabled = True
If Commandadd.Visible = True Then
 Commandadd.SetFocus
End If
End Sub

Private Sub Form_Load()

   Dim rs2 As New ADODB.Recordset
   Me.Top = 400
   Me.Left = 50
   Set RS = New ADODB.Recordset
   cmdAdd = False
  
   cmdedit = False
  'ue
   Datachange = False
   GCombo.Clear
   gencombo1.Clear
   'MainMenu.Toolbar1.Visible = False
   
   
    If RS.State = 1 Then RS.close
    RS.Open "select Rep as Representative from SalesRepQry where (email is not null and len(email)>1) order by Rep", CON_blue
    cmbAgentName.Clear
    If Not RS.EOF Then
       Do While Not RS.EOF
          If IsNull(RS(0)) = False Then
            Me.cmbAgentName.AddItem RS(0)
          End If
          If Not RS.EOF Then RS.MoveNext
        Loop
    End If
    
 
 
 
   
   
 
 
   If RS.State = 1 Then RS.close
   If Val(inviceNo) > 0 Then
     RS.Open "Select *  from CNF1A where cnn=" & inviceNo & " order by cnn  ", con, adOpenDynamic, adLockOptimistic, adCmdText
     Else
     RS.Open "Select *  from CNF1A order by cnn  ", con, adOpenDynamic, adLockOptimistic, adCmdText
   End If
   inviceNo = ""
   
   If RS.RecordCount > 0 Then
     RS.MoveLast
     Set TCNN.DataSource = RS
     Set CdateT.DataSource = RS
     Set Text2.DataSource = RS
     Set Text3.DataSource = RS
     Set GText.DataSource = RS
     Set SText.DataSource = RS
     Set agn.DataSource = RS
     'Set cmbAgentName.DataSource = RS
     
  End If
   
     If rs2.State = 1 Then rs2.close
   
   rs2.Open "Select * from gledger where  slf = 1  order by gledger", con, adOpenStatic, adLockOptimistic
   
   If Not rs2.EOF Then
        Do While Not rs2.EOF
           GCombo.AddItem rs2(1)
            If Not rs2.EOF Then
                rs2.MoveNext
            End If
        Loop
    End If
  
   
   
     If rs2.State = 1 Then rs2.close
   rs2.Open "Select * from groups order by groupcode", con, adOpenStatic, adLockReadOnly, adCmdText
   If Not rs2.EOF Then
        Do While Not rs2.EOF
            cmbgroup.AddItem rs2!groupcode
            If Not rs2.EOF Then
               rs2.MoveNext
            End If
        Loop
   End If
   
    If rs2.State = 1 Then rs2.close
   
   rs2.Open "Select * from gledger  order by gledger", con, adOpenStatic, adLockOptimistic
   
   If Not rs2.EOF Then
        Do While Not rs2.EOF
            gencombo1.AddItem rs2(1)
            
            If Not rs2.EOF Then
                rs2.MoveNext
            End If
        Loop
    End If
    
    
    
    
     cmdedit = False
    If RS.RecordCount <= 0 Then
        Commandedit.Enabled = False
    End If
    
    CNbandon
    
    SetButton Commandadd, Commandedit, Commandsave, Commanddelete
    
 
End Sub

Private Sub GCombo_Change()
GText.Text = GCombo.Text

End Sub

Private Sub GCombo_Click()
GText.Text = GCombo.Text
End Sub

Private Sub GCombo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Dim GEN As String
     Dim SC1 As String
     
     GEN = GCombo.Text
     SC1 = Scombo.Text
     
     Dim rs2 As New ADODB.Recordset
            rs2.Open "Select * from sledger where GLEDGER='" + Trim(GCombo.Text) + "'", con, adOpenStatic, adLockReadOnly, adCmdText
            Scombo.Clear
            If Not rs2.EOF Then
                 Do While Not rs2.EOF
                        Scombo.AddItem rs2(1)
                        If Not rs2.EOF Then
                            rs2.MoveNext
                        End If
                    Loop
      
            Else
                    Scombo.Clear
                    Scombo.Text = ""
                    Text2.SetFocus
                    Exit Sub
            End If
     
        If KeyAscii = 13 Then
                    SendKeys "{tab}"
                    Datachange = False
         End If
  Else
  
     Datachange = True
  
  
  End If
 
  If Datachange = False Then
       GCombo.Text = GEN
       Scombo.Text = SC1
       Datachange = False
  End If
 

End Sub

Private Sub GCombo_LostFocus()
  If GCombo.Text = "" Then
            GCombo.SetFocus
  Else
       If GCombo.Text <> "" Then
            Dim rs4 As New ADODB.Recordset
            rs4.Open "Select* from gledger  where  slf = 1  and GLEDGER = '" + Trim(GCombo.Text) + "'", con, adOpenStatic
            If rs4.RecordCount <= 0 Then
                 MsgBox "No valid Gen.Ledger"
                 GCombo.SetFocus
            End If
        End If
   End If




End Sub

Private Sub gencombo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 115 Then
   If MsgBox("Are You Sure, Delete it ?", vbYesNo) = vbYes Then
      If DGrid.Row >= 1 Then
           DGrid.RemoveItem DGrid.Row
           gencombo1.Text = ""
           gencombo1.Visible = False
           DGrid.SetFocus
       End If
   End If
End If
End Sub

Private Sub gencombo1_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
      Dim rs2 As New ADODB.Recordset
      Dim rs3 As New ADODB.Recordset
    
      If gencombo1.Text = "" Then
        gencombo1.Visible = False
        Commandsave.SetFocus
        Exit Sub
      End If
     
      If gencombo1.Text <> "" Then
            Dim rs4 As New ADODB.Recordset
            rs4.Open "Select* from gledger where GLEDGER='" + Trim(gencombo1.Text) + "'", con, adOpenForwardOnly, adLockReadOnly, adCmdText
            If rs4.RecordCount <= 0 Then
                 MsgBox "No valid Gen.Ledger"
                 gencombo1.Visible = True
                 gencombo1.SetFocus
                 Exit Sub
            End If
       End If
    
       DGrid.Text = gencombo1.Text
       Subcombo1.Clear
       rs2.Open "Select * from sledger where GLEDGER='" + Trim(gencombo1.Text) + "'", con, adOpenForwardOnly, adLockReadOnly, adCmdText
       DGrid.Col = 0
       If rs2.RecordCount > 0 Then
            rs2.MoveFirst
            Do While Not rs2.EOF
                    Subcombo1.AddItem rs2(1)
                    If Not rs2.EOF Then
                        rs2.MoveNext
                    End If
            Loop
            DGrid.Col = DGrid.Col + 1
            DGrid_Click
         Else
             DGrid.Col = DGrid.Col + 1
             DGrid.Text = ""
            DGrid.Col = DGrid.Col + 1
            Subcombo1.Visible = False
            DGrid_Click
           
         End If
           
End If



End Sub

Private Sub GText_Change()
GCombo.Text = GText.Text

End Sub

Private Sub Scombo_Change()
   SText.Text = Scombo.Text
End Sub

Private Sub Scombo_Click()
    SText.Text = Scombo.Text
End Sub

Private Sub Scombo_KeyPress(KeyAscii As Integer)
If VB.Screen.ActiveForm.ActiveControl.SelLength > 0 Then
   SendKeys "{tab}"
   Exit Sub
End If
If KeyAscii = 13 Then
     'If cmdAdd = True Then
        SendKeys "{Down}"
    ' End If
   SendKeys "{tab}"
End If
End Sub

Private Sub Scombo_LostFocus()
If Scombo.Text <> "" And GCombo.Text <> "" Then
  Dim rs4 As New ADODB.Recordset
  Dim rs5 As New ADODB.Recordset
  rs4.Open "Select* from sledger where GLEDGER='" + Trim(GCombo.Text) + "' and SubLedger='" + Trim(Scombo.Text) + "'", con, adOpenStatic
  If rs4.RecordCount <= 0 Then
     MsgBox "No valid Sub Ledger"
     Scombo.SetFocus
  End If
  If rs4!DISTCODE <> "" Then
       rs5.Open "Select * from Districts where Districtname = '" & rs4!DISTCODE & "'", con, adOpenStatic, adLockReadOnly
       If rs5.RecordCount > 0 Then
          Me.cmbAgentName = IIf(IsNull(rs5!agentname), "", rs5!agentname)
       End If
  End If
End If
If Scombo.ListCount > 0 And Scombo.Text = "" Then
     Scombo.SetFocus
End If

End Sub

Private Sub SText_Change()
 Scombo.Text = SText.Text
End Sub

Private Sub Subcombo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then


      If Subcombo1.Text <> "" And gencombo1.Text <> "" Then
                Dim rs4 As New ADODB.Recordset
                rs4.Open "Select* from sledger where GLEDGER='" + Trim(gencombo1.Text) + "' and SubLedger='" + Trim(Subcombo1.Text) + "'", con, adOpenStatic
                  If rs4.RecordCount <= 0 Then
                            MsgBox "No valid Sub Ledger"
                            Subcombo1.SetFocus
                            Exit Sub
                  End If

       End If
       If Subcombo1.ListCount > 0 And Subcombo1.Text = "" Then
               Subcombo1.SetFocus
               Exit Sub
       End If
       DGrid.Col = 0
       If DGrid.Text = "" Then
                 DGrid.Col = 1
                 If DGrid.Text = "" Then
                      Commandsave.SetFocus
                      Exit Sub
                 End If
       End If
       DGrid.Col = 1
       DGrid.Text = Subcombo1.Text
       DGrid.Col = DGrid.Col + 1
       DGrid_Click
 End If
 
 
 
End Sub

Private Sub TCNN_Change()
   
If cmdedit = True Then
    cmdedit = False
   Gridrefresh
End If


    
End Sub

Private Sub TCNN_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Dim rs1 As New ADODB.Recordset
     Dim rs6 As New ADODB.Recordset
     rs6.Open "Select invoiceno from CREDITA   where invoiceno = " + TCNN.Text + "", con, adOpenStatic, adLockOptimistic, adCmdText
     If rs6.RecordCount > 0 Then
            MsgBox "Credit Note(Item) already exist..."
            TCNN.SetFocus
            Exit Sub
     End If
     
     
     
     rs1.Open "Select * from  CNF1A  where cnn = " + TCNN.Text + "", con, adOpenStatic, adLockOptimistic, adCmdText
     If Not rs1.EOF Then
     If cmdAdd Then
            MsgBox "Credit Note already exist..."
            TCNN.SetFocus
            Exit Sub
     End If
  
  End If
  
  SendKeys "{tab}"
End If
End Sub

Private Sub TCNN_LostFocus()
 
 Dim rs6 As New ADODB.Recordset
     If TCNN.Text = "" Then
       TCNN.SetFocus
       Exit Sub
     End If
     


End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   DGrid.Row = 1
   DGrid.Col = 0
   DGrid.SetFocus
   DGrid.Row = 1
   DGrid.Col = 0
   DGrid_Click
End If


KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub

Private Sub Text3_Change()
Text3 = Format(Val(Text3.Text), "0.00")
End Sub




Sub GenrateReport()

   Dim rs7 As ADODB.Recordset
   Dim rs1 As ADODB.Recordset
   Dim kk As ADODB.Recordset
   Dim trs As ADODB.Recordset
   Dim paperWidth As Integer
   Dim kkk As ADODB.Recordset
   Dim Tot As Double
   Dim MaxLine, Pno, Line As Integer
   Dim called1 As Boolean
   Dim Glist1 As String
   Dim ID1 As String
   Dim Gc As String
   Dim Gc1 As String
   Dim FooterYes As Boolean
   Dim NetTotal As Double
   Dim GTotal As Double
   Dim J As Integer
   NetTotal = 0
   I = 0
   
   GTotal = 0
   FooterYes = False
   Set kkk = New ADODB.Recordset
   Set rs1 = New ADODB.Recordset
   Set rs7 = New ADODB.Recordset
   Set kk = New ADODB.Recordset
   Set trs = New ADODB.Recordset
   Tot = 0
   Line = 0
   Pno = 1
   MaxLine = 72
   called1 = False
   called2 = False
   main.reportname = "Dis. Sales"
   main.reportdata
   main.repors.Find "reportname='" + Trim(main.reportname) + "'"
   MaxLine = main.repors!totalline
   If main.repors!comp = True Then
      paperWidth = Int(main.repors!totalcolumn * 1.75)
   Else
      paperWidth = main.repors!totalcolumn
   End If
   Open "" + VB.App.Path + "\vipin.txt" For Output As #1
   MaxLine = 72
   called1 = False
   Pno = 1
   paperWidth = 96
header:
   For I = 1 To main.repors!TopMargin
       Print #1, ""
       Line = Line + 1
   Next
   If FooterYes = True Then
       While Line < 72
           Print #1, ""
           Line = Line + 1
       Wend
       Line = 0
       FooterYes = False
   End If
   If kkk.State = 1 Then kkk.close
   CNSetup
   kkk.Open "select * from setup1", con, adOpenStatic, adLockReadOnly, adCmdText
   If Printheader = True Then
     
   If Not kkk.BOF Then
     Print #1, ""
     Print #1, ""
     Print #1, ""
     Print #1, ""
     Print #1, Chr(27) + Chr(71); Chr(27) + Chr(77) + Chr(14)
     Print #1, Tab(((paperWidth - (Len(Trim(kkk!cname)) * 2)) / 2)); Chr(27) + Chr(77) + Chr(14); Trim(kkk!cname)
     Print #1, Tab((paperWidth - (Len(Trim(kkk!add1)) * 2)) / 2); Chr(27) + Chr(77); dspace(Trim(kkk!add1))
     Print #1, Tab((paperWidth - (Len(Trim(kkk!phone1)) * 2)) / 2); Trim(kkk!phone1) & "," & Trim(kkk!phone2)
     Line = Line + 8
   End If
Else
     Print #1, ""
     Print #1, ""
     Print #1, ""
     Print #1, ""
     Print #1, ""
     Print #1, ""
     Print #1, ""
     Print #1, Chr(27) + Chr(77)
     Line = Line + 8
End If
   Print #1, Chr(27) + Chr(71); Chr(27) + Chr(72); Tab((paperWidth - (Len(Trim("CREDIT NOTE")))) / 2 - 3); Chr(14); "CREDIT NOTE"; Chr(20);
   Line = Line + 1
   If Printheader = True Then
      Print #1, Tab(63); kkk!uptt
      Print #1, Tab(63); kkk!cst
      Line = Line + 2
   End If
   If Printheader = False Then
      Print #1, Tab(63); "CIN-U22110UP2004PTC028474"
      Print #1, ""
      Line = Line + 2
   End If
   Print #1, repli("-", paperWidth)
   Print #1, ""
   Line = Line + 2
   If called1 = True Then
        called1 = False
        GoTo printagain1
    End If
If rs7.State = 1 Then rs7.close
rs7.Open "Select * from cnf1a where cnn =" & TCNN.Text & "  and cnd = cdate('" + Trim(Cbdate.Text) + "')", con
If rs7.RecordCount > 0 Then
   Print #1, Chr(27) + Chr(71); "To,   S.L. Code : "; Tab(20); Mid$(rs7!psld, 1, 5); Tab(50); "Credit Note No. : "; Chr(27) + Chr(72); Trim(rs7!cnn); Tab(83); Chr(27) + Chr(71); "Date : "; Chr(27) + Chr(72); rs7!Cnd
   Line = Line + 1
   If kkk.State = 1 Then kkk.close
   kkk.Open "select * from sledger where subledger='" + Trim(rs7!psld) + "'", con, adOpenStatic, adLockReadOnly, adCmdText
   If Not kkk.EOF Then
      Print #1, Tab(5); "M/s " & kkk!DESCFORINVOICE
      Print #1, Tab(5); IIf(IsNull(kkk!address1), " ", kkk!address1); Tab(50); " Agent Name      : " + Creditnotefile.cmbAgentName
      Print #1, Tab(5); IIf(IsNull(kkk!address2), " ", kkk!address2)
      Print #1, Tab(5); IIf(IsNull(kkk!address3), " ", kkk!address3)
      Print #1, ""
      kkk.close
   End If
   Print #1, ""
   Print #1, Tab(5); "Narration         : "; Tab(30); rs7!n
   Print #1, ""
   Print #1, Tab(0); repli("-", paperWidth)
   Print #1, Tab(5); "GenLedger"; Tab(30); ""; Tab(85); "Amount"
   Print #1, repli("-", paperWidth)
   Line = Line + 11
   If trs.State = 1 Then trs.close
   trs.Open "Select * from cnf1b where cnn =" & TCNN.Text & "  and Cnd = cdate('" + Trim(Cbdate.Text) + "')", con
   If trs.RecordCount > 0 Then
      While Not trs.EOF
            Print #1, Tab(5); trs!gld; Tab(35); IIf(IsNull(trs!sld), "", trs!sld); Tab(80); rsets(Trim(Format(Str(trs!a), "0.00")), 12)
            Line = Line + 1
            If Line > MaxLine - 5 Then
                FooterYes = True
                Pno = Pno + 1
                called1 = True
                GoTo header
printagain1:
                called1 = False
            End If
            trs.MoveNext
        Wend
    End If
    While Line <= 58
         Print #1, ""
         Line = Line + 1
    Wend
    
    Print #1, ""
    Print #1, Tab(1); "Net Amount Cr. In Your A/C : "; Tab(80); rsets(Trim(Format(Str(rs7!na), "0.00")), 12)
    Print #1, ""
    Print #1, Tab(1); toword(rs7!na)
    Print #1, repli("-", paperWidth)
    Dim tempdata As ADODB.Recordset
    Set tempdata = New ADODB.Recordset
    CNSetup
    tempdata.Open "setup1", con, adOpenStatic, adLockReadOnly, adCmdTable
    Print #1, Tab(1); "E.& O.E"
    Print #1, Tab(1); tempdata!COURT; Tab(65); "FOR " + Trim(tempdata!cname)
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    
    Close #1
    PrintOption.Show
End If


End Sub


