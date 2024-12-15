VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Debitnotefile 
   Caption         =   "Debit Note"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   9015
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   90
      TabIndex        =   30
      Top             =   5190
      Width           =   3075
      Begin VB.CommandButton Command1 
         Caption         =   ">>"
         Height          =   345
         Left            =   1350
         TabIndex        =   32
         Top             =   60
         Width           =   765
      End
      Begin VB.CommandButton Command2 
         Caption         =   "<<"
         Height          =   345
         Left            =   360
         TabIndex        =   31
         Top             =   60
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   5265
      Left            =   -30
      TabIndex        =   13
      Top             =   -120
      Width           =   9195
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
         TabIndex        =   28
         Top             =   3030
         Visible         =   0   'False
         Width           =   915
      End
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
         TabIndex        =   27
         Text            =   "C"
         Top             =   3030
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox CdateT 
         DataField       =   "Dnd"
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
         Left            =   5940
         TabIndex        =   24
         Top             =   300
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox Text2 
         DataField       =   "n"
         Height          =   315
         Left            =   1770
         MaxLength       =   39
         TabIndex        =   5
         Top             =   1890
         Width           =   6945
      End
      Begin VB.ComboBox Scombo 
         Height          =   315
         Left            =   1770
         TabIndex        =   4
         Top             =   1350
         Width           =   4965
      End
      Begin VB.ComboBox GCombo 
         Height          =   315
         Left            =   1770
         TabIndex        =   3
         Top             =   840
         Width           =   4965
      End
      Begin VB.ComboBox gencombo1 
         Height          =   1545
         Left            =   360
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   6
         Top             =   2880
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.ComboBox Subcombo1 
         Height          =   1545
         Left            =   2790
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   17
         Top             =   2880
         Visible         =   0   'False
         Width           =   2385
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         DataField       =   "na"
         Height          =   315
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   4770
         Width           =   1065
      End
      Begin VB.TextBox SText 
         DataField       =   "psld"
         Height          =   285
         Left            =   7380
         TabIndex        =   15
         Top             =   1440
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.TextBox GText 
         DataField       =   "pgld"
         Height          =   285
         Left            =   7350
         TabIndex        =   14
         Top             =   1020
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.TextBox TCNN 
         DataField       =   "Dnn"
         Height          =   315
         Left            =   1770
         TabIndex        =   1
         Top             =   330
         Width           =   795
      End
      Begin MSMask.MaskEdBox Cbdate 
         Height          =   315
         Left            =   3630
         TabIndex        =   2
         Top             =   270
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid DGrid 
         Height          =   2325
         Left            =   390
         TabIndex        =   29
         Top             =   2430
         Width           =   8595
         _ExtentX        =   15161
         _ExtentY        =   4101
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
         Caption         =   "Debit Note No."
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
         Left            =   180
         TabIndex        =   23
         Top             =   330
         Width           =   1290
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
         TabIndex        =   22
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
         Left            =   600
         TabIndex        =   21
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
         Left            =   480
         TabIndex        =   20
         Top             =   1290
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
         Left            =   270
         TabIndex        =   19
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
         Left            =   5040
         TabIndex        =   18
         Top             =   4800
         Width           =   1005
      End
   End
   Begin VB.PictureBox pic1 
      Height          =   465
      Left            =   90
      ScaleHeight     =   405
      ScaleWidth      =   8475
      TabIndex        =   0
      Top             =   5730
      Width           =   8535
      Begin VB.CommandButton Commandsave 
         Caption         =   "Sa&ve"
         Enabled         =   0   'False
         Height          =   345
         Left            =   1800
         TabIndex        =   26
         Top             =   30
         Width           =   800
      End
      Begin VB.CommandButton Commandabandon 
         Caption         =   "Aba&ndon"
         Height          =   345
         Left            =   2640
         TabIndex        =   25
         Top             =   30
         Width           =   800
      End
      Begin VB.CommandButton Commandedit 
         Caption         =   "&Edit"
         Height          =   345
         Left            =   930
         TabIndex        =   12
         Top             =   30
         Width           =   800
      End
      Begin VB.CommandButton Commanddelete 
         Caption         =   "De&lete"
         Height          =   345
         Left            =   3510
         TabIndex        =   11
         Top             =   30
         Width           =   800
      End
      Begin VB.CommandButton Commandsearch 
         Caption         =   "&Search"
         Height          =   345
         Left            =   4380
         TabIndex        =   10
         Top             =   30
         Width           =   800
      End
      Begin VB.CommandButton CommandPrint 
         Caption         =   "&Print"
         Enabled         =   0   'False
         Height          =   345
         Left            =   5250
         TabIndex        =   9
         Top             =   30
         Width           =   800
      End
      Begin VB.CommandButton CommandReturn 
         Caption         =   "&Return"
         Height          =   345
         Left            =   6120
         TabIndex        =   8
         Top             =   30
         Width           =   800
      End
      Begin VB.CommandButton Commandadd 
         Caption         =   "&Add"
         Height          =   345
         Left            =   90
         TabIndex        =   7
         Top             =   30
         Width           =   800
      End
   End
End
Attribute VB_Name = "Debitnotefile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs As ADODB.Recordset

'Dim con As New ADODB.Connection
Dim mvBookMark As Variant
Dim cmdAdd As Boolean
Dim cmdEdit As Boolean
Dim LRC As Integer
Dim LCC As Integer
Dim Glastrow As Integer
Dim Datachange As Boolean
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
  DGrid.row = 0
  DGrid.col = 0
  DGrid.Text = "Gen Ledger"
  DGrid.col = 1
  DGrid.Text = "Sub Ledger"
  DGrid.col = 2
  DGrid.Text = Format(DGrid.Text, "0.00")
  DGrid.Text = "Amount"
  DGrid.col = 3
  DGrid.Text = "Debit/Credit"
  DGrid.ColWidth(0) = 3000
  DGrid.ColWidth(1) = 3000
  
    

End Sub
Sub Controlclear()
     CdateT.Text = ""
     GCombo.Text = ""
     Scombo.Text = ""
     Text2.Text = ""
     Text3.Text = ""
End Sub





Sub Gridrefresh()
      Dim grs1 As New ADODB.Recordset
      If grs1.State = 1 Then grs1.Close
         If TCNN.Text = "" Then
            grs1.Open "SELECT gld as GenLedger,Sld as SubLedger, a as Amount ,Dc as DebitCredit FROM  DNFB WHERE  DNN = 0", TCON1, adOpenStatic
            Set DGrid.DataSource = grs1
            DGrid.Refresh
            Setgrid
   
         ElseIf cmdEdit = True Then
                grs1.Open "SELECT gld as GenLedger,Sld as SubLedger, a as Amount ,Dc as DebitCredit FROM  DNFB", TCON1, adOpenStatic
                DGrid.Rows = 99
                Setgrid
         ElseIf cmdAdd = True Then
                grs1.Open "SELECT gld as GenLedger,Sld as SubLedger, a as Amount ,Dc as DebitCredit FROM  DNFB", TCON1, adOpenStatic
                DGrid.Refresh
                DGrid.Rows = 99
                Setgrid
         Else
                grs1.Open "SELECT gld as GenLedger,Sld as SubLedger, a as Amount ,Dc as DebitCredit FROM  DNFB WHERE  DNN = " + Trim(TCNN.Text) + "", TCON1, adOpenStatic
                Set DGrid.DataSource = grs1
                DGrid.Refresh
                Setgrid

         End If
    
       If TCNN.Text <> "" Then
         For I = 1 To grs1.RecordCount
           DGrid.row = I
           DGrid.col = 2
           DGrid.Text = Format(DGrid.Text, "0.00")
           DGrid.Refresh
         Next I
       End If
             
       DGrid.col = 2
End Sub


Sub GridEdit()
      Dim grs1 As New ADODB.Recordset
      If grs1.State = 1 Then grs1.Close
      If TCNN.Text = "" Then
         grs1.Open "SELECT gld as GenLedger,Sld as SubLedger, a as Amount ,Dc as DebitCredit FROM  TempCNF1B", TCON1, adOpenStatic
      End If
     
    
           
End Sub


Private Sub AText_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     DGrid.Text = Format(AText.Text, "0.00")
     
      
     If AText = "" Then DGrid.Text = 0
         
      DGrid.col = DGrid.col + 1
      DGrid_Click
      Glastrow = DGrid.row
       
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
If Trim(Cbdate.Text) <> "__/__/____" Then
    If Not checkdate(Trim(Cbdate.Text), Cbdate) Then
        Cbdate.SetFocus
    End If
End If
End Sub

Private Sub CdateT_Change()
 On Error GoTo er1
 Cbdate.Text = CdateT.Text
er1:  If Err.Number = 380 Then
       Exit Sub
      End If
End Sub

Private Sub Command1_Click()
On Error GoTo er

If Not rs.BOF Or Not rs.EOF Then rs.MoveNext
If rs.EOF And rs.RecordCount > 0 Then
    Beep
     
    rs.MoveLast
  End If
  
er:   If Err.Number = 3021 Then
         
         Exit Sub
       
    End If
  End Sub

Private Sub Command2_Click()
If Not rs.BOF Then rs.MovePrevious
  If rs.BOF And rs.RecordCount > 0 Then
    Beep
    rs.MoveFirst
  End If
End Sub

Private Sub Command3_Click()
   rs.CancelBatch
  If mvBookMark > 0 Then
    rs.Bookmark = mvBookMark
  Else
    rs.MoveFirst
  End If
  Gridrefresh
  pic1.Visible = True
  Pic2.Visible = False
End Sub

Public Sub Commandabandon_Click()
  If rs.RecordCount > 0 Then
     rs.CancelUpdate
     rs.MoveFirst
 End If
  
cmdAdd = False
cmdEdit = False
 
  Gridrefresh
  pic1.Visible = True
  Frame2.Enabled = True
  gencombo1.Visible = False
  Subcombo1.Visible = False
  AText.Visible = False
  DrCrText.Visible = False
  
  
  CNbandon
 
End Sub

Private Sub Commandadd_Click()
 Me.Commandadd.Enabled = False
 Me.Commandedit.Enabled = False
 Me.Commandsearch.Enabled = False
 Me.Commandsave.Enabled = True
 Me.Commanddelete.Enabled = False
 Me.Commandabandon.Enabled = True
 Me.CommandPrint.Enabled = False
 Dim rs1  As New ADODB.Recordset
cmdAdd = True
cmdEdit = False
LRC = 1
LCC = 0
frame1.Enabled = True
TCNN.Enabled = True
TCNN.SetFocus
CdateT.Text = ""




With rs
    .AddNew
End With

If cmdEdit = False Then
     Dim Trs As New ADODB.Recordset
     Trs.Open "Select max(DNN)as mcnn from DNFA", TCON1, adOpenStatic, adCmdText
     If Trs.RecordCount <= 1 And IsNull(Trs!Mcnn) Then
       TCNN.Text = 1
     Else
      TCNN.Text = Trs!Mcnn + 1
       
    End If
    
End If
'TCon1.Execute "Delete from tempcnf1b"
DGrid.Rows = 100
DGrid.Cols = 0
DGrid.Cols = 4
Dim I
For I = 0 To 99
        DGrid.RowHeight(I) = 270
Next

Setgrid
DoEvents
GCombo.Text = "SUNDRY DEBTORS"
Dim rs2 As New ADODB.Recordset
            rs2.Open "Select * from sledger where GLEDGER='" + Trim(GCombo.Text) + "'", TCON1, adOpenStatic, adLockReadOnly, adCmdText
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

DGrid.row = 1
DGrid.col = 0
DGrid.TopRow = 1
Frame2.Enabled = False

End Sub

Private Sub Commanddelete_Click()
'On Error GoTo er1

If rs.RecordCount > 0 Then
  If MsgBox("Are you sure.......", vbYesNo) = vbYes Then
  If TCNN.Text <> "" Then
       TCON1.Execute "Delete * FROM  DNFB WHERE  DNN=" + TCNN.Text + ""
  End If
  With rs
    .Delete
     If Not rs.BOF And Not rs.EOF Then rs.MoveFirst
    
     

     
     
  End With
    
  Gridrefresh
  Exit Sub


End If

End If
'er1:  If Err.Number = 3021 Then
          'Controlclear
 '         Gridrefresh
            
  '         Exit Sub
  '    End If
End Sub

 Sub Commandedit_Click()
        If rs.RecordCount <= 0 Then
             cmdAdd = False
             cmadd = False
             cmdEdit = False
   
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

          
          Exit Sub
        End If
 
        Me.Commandadd.Enabled = False
        Me.Commandedit.Enabled = False
        Me.Commandsearch.Enabled = False
        Me.Commandsave.Enabled = True
        Me.Commanddelete.Enabled = False
        Me.Commandabandon.Enabled = True
        Me.CommandPrint.Enabled = False
        cmdAdd = False
        cmdEdit = True
        TCNN.Enabled = False
        frame1.Enabled = True
        Cbdate.SetFocus

        Gridrefresh
        
        If cmdEdit = True Then
          Dim rs3 As New ADODB.Recordset
        rs3.Open "SELECT * FROM  DNFB WHERE  DNN=" + Trim(TCNN.Text) + "", TCON1, adOpenStatic, adLockOptimistic, adCmdText
          If rs3.RecordCount >= 0 Then
             DGrid.Rows = 99
             DGrid.TopRow = 1
             DGrid.row = rs3.RecordCount + 1
          End If
    End If
    Frame2.Enabled = False
    VB.screen.ActiveForm.ActiveControl.SelLength = 2
 
End Sub

Private Sub Commandreturn_Click()
'rs.Close
Unload Me
MainMenu.Toolbar1.Visible = True
End Sub

Private Sub Commandsave_Click()
   Dim Grs As New ADODB.Recordset
   TCON1.Execute "Delete * FROM  DNFB WHERE  DNN=" + Trim(TCNN.Text) + ""
   TCNN.Enabled = True
   Grs.Open "DNFB", TCON1, adOpenDynamic, adLockOptimistic, adCmdTable
   Dim sum As Double
   sum = 0
   Dim I, j As Integer
     I = 1
     DGrid.row = 1
     DGrid.col = 0
     While DGrid.Text <> ""
        DGrid.row = I
        Grs.AddNew
        Grs!dnn = Val(TCNN.Text)
        Grs!dnd = Cbdate.Text
        For j = 0 To 3
          DGrid.col = j
                If j = 0 Then
                   If DGrid.Text = "" Then
                      MsgBox "Please fill Gen Ledger"
                      DGrid_Click
                      Exit Sub
                       
                   Else
                     Grs!gld = DGrid.Text
                   End If
               End If
                If j = 1 Then
                    If DGrid.Text = "" Then
                       
                         Grs!Sld = Null
                                               
                    Else
                         Grs!Sld = DGrid.Text
                     End If
                   End If
                If j = 2 Then
                     If DGrid.Text = "" Then
                       
                        Grs!a = 0
                        
                     Else
                        Grs!a = DGrid.Text
                    End If
                 End If
                 If j = 3 Then
                 
                   If DGrid.Text = "" Then
                         MsgBox "Please fill Correct Entry  "
                         DGrid_Click
                         Exit Sub
                   Else
                         Grs!DC = DGrid.Text
                   End If
                   If DGrid.Text = "D" Then
                      DGrid.col = 2
                          sum = sum + Format(Val(DGrid.Text), "0.00")
                    Else
                      DGrid.col = 2
                           sum = sum - Format(Val(DGrid.Text), "0.00")
      
                   End If
                 End If
          Next j
          Grs.Update
         
         I = I + 1
         DGrid.row = I
         DGrid.col = 0
    Wend
   rs!dnn = TCNN.Text
   rs!dnd = Cbdate.Text
   rs!pgld = GText.Text
   
   If SText.Text = "" Then
      rs!psld = Null 'SText.Text
   Else
      rs!psld = SText.Text
   End If
   rs!n = Text2.Text
   rs!na = Abs(Format(sum, "0.00"))
   rs!DC = "C"
   rs.Update
    'TCon1.Execute "Delete * FROM  DNFB WHERE  CNN=" + Trim(TCNN.Text) + ""
    'TCon1.Execute "Insert into DNFB select * from tempcnf1b"
       frame1.Enabled = False
       cmdAdd = False
       cmadd = False
       cmdEdit = False
   
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

    Me.Enabled = False
    searchscreen.grid1.row = 0
    searchscreen.grid1.col = 0
    Call searchscreen.tempr(16, Me.Name)
End Sub

Private Sub DGrid_AfterUpdate()
Dim rs2 As New ADODB.Recordset
rs2.Open "Select sum(a) as tot from  tempcnf1b ", TCON1, adOpenStatic, adCmdText
If rs2.RecordCount > 0 Then
If IsNull(rs2!tot) = True Then
    Text3.Text = 0
 Else
   Text3.Text = rs2!tot
End If
End If
End Sub

Private Sub DGrid_Click()
If DGrid.row > 0 Then
       Select Case DGrid.col
    
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
                DrCrText.Text = DGrid.Text
                 DrCrText.Text = "C"
                DrCrText.Visible = True
                gencombo1.Visible = False
                Subcombo1.Visible = False
                AText.Visible = False
                DrCrText.Move DGrid.CellLeft + DGrid.Left, DGrid.CellTop + DGrid.Top - 15, DGrid.CellWidth
                DrCrText.SetFocus
            
            End Select
   
  End If
    
    
    
    
End Sub

Private Sub DGrid_RowColChange()
'Dim lastvalue
'lastvalue = DGrid.row
'DGrid.row = LRC - 2

'DGrid.col = 0
'If DGrid.Text = "" Then
'                gencombo1.Text = DGrid.Text
'                gencombo1.Visible = True
'                Subcombo1.Visible = False
'                AText.Visible = False
'                DrCrText.Visible = False
 '               gencombo1.Move DGrid.CellLeft + DGrid.Left, DGrid.CellTop + DGrid.Top - 15, DGrid.CellWidth
               ' gencombo1.SetFocus
 'Else
 
  '         DGrid.row = lastvalue
  '    End If





End Sub

Private Sub DrCrText_Change()
DrCrText.Text = UCase(DrCrText.Text)
End Sub

Private Sub DrCrText_KeyPress(KeyAscii As Integer)




If KeyAscii = 13 Then
     Glastrow = DGrid.row
     DGrid.Text = DrCrText.Text
     DGrid.row = DGrid.row + 1
     
     LRC = LRC + 1
     DGrid.col = 0
     DGrid_Click
End If
End Sub

Private Sub Form_Activate()
Commandadd.Enabled = True
 Commandadd.SetFocus
End Sub

Private Sub Form_Load()
   Dim rs2 As New ADODB.Recordset
   Set rs = New ADODB.Recordset
 
   cmdAdd = False
   cmdEdit = False
   Datachange = False
   
   GCombo.Clear
   gencombo1.Clear
   MainMenu.Toolbar1.Visible = False
   TCON1.CursorLocation = adUseClient
   'set TCon1 = New ADODB.Connection
   'TCON1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + VB.App.Path + "\" + Trim(main.directory) + "\data.mdb"
   
   'con.Open
   If rs.State = 1 Then rs.Close
   rs.Open " Select *  from DNFA  order by dnn", TCON1, adOpenDynamic, adLockOptimistic, adCmdText
   Set TCNN.DataSource = rs
   Set CdateT.DataSource = rs
   Set Text2.DataSource = rs
   Set Text3.DataSource = rs
   Set GText.DataSource = rs
   Set SText.DataSource = rs
      
   
   
   
   rs2.Open "Select * from gledger where  slf = true order by gledger", TCON1, adOpenStatic, adLockOptimistic
   
   If Not rs2.EOF Then
        Do While Not rs2.EOF
            gencombo1.AddItem rs2(1)

            If Not rs2.EOF Then
                rs2.MoveNext
            End If
        Loop
    End If
    If rs2.State = 1 Then rs2.Close
   rs2.Open "Select * from gledger order by gledger", TCON1, adOpenStatic, adLockOptimistic
   
   If Not rs2.EOF Then
        Do While Not rs2.EOF
            gencombo1.AddItem rs2(1)
               If Not rs2.EOF Then
                rs2.MoveNext
            End If
        Loop
    End If
    
    If rs.RecordCount <= 0 Then
        Commandedit.Enabled = True
    End If
   
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
            rs2.Open "Select * from sledger where GLEDGER='" + Trim(GCombo.Text) + "'", TCON1, adOpenStatic, adLockReadOnly, adCmdText
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
             rs4.Open "Select* from gledger where  slf = true and  GLEDGER='" + Trim(GCombo.Text) + "'", TCON1, adOpenStatic
            If rs4.RecordCount <= 0 Then
                 MsgBox "No valid Gen.Ledger"
                 GCombo.SetFocus
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
            rs4.Open "Select* from gledger where GLEDGER='" + Trim(gencombo1.Text) + "'", TCON1, adOpenStatic
           If rs4.RecordCount <= 0 Then
                 MsgBox "No valid Gen.Ledger"
                 gencombo1.Visible = True
                 gencombo1.SetFocus
                 Exit Sub
            End If
       End If
    
       DGrid.Text = gencombo1.Text
       Subcombo1.Clear
       rs2.Open "Select * from sledger where GLEDGER='" + Trim(gencombo1.Text) + "'", TCON1, adOpenStatic, adLockReadOnly, adCmdText
       DGrid.col = 0
       If rs2.RecordCount > 0 Then
            rs2.MoveFirst
            Do While Not rs2.EOF
                    Subcombo1.AddItem rs2(1)
                    If Not rs2.EOF Then
                        rs2.MoveNext
                    End If
            Loop
            DGrid.col = DGrid.col + 1
            DGrid_Click
         Else
            DGrid.col = DGrid.col + 1
            DGrid.Text = ""
            DGrid.col = DGrid.col + 1
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
If KeyAscii = 13 Then
  SendKeys "{tab}"
End If
End Sub

Private Sub Scombo_LostFocus()
If Scombo.Text <> "" And GCombo.Text <> "" Then
  Dim rs4 As New ADODB.Recordset
  rs4.Open "Select* from sledger where GLEDGER='" + Trim(GCombo.Text) + "' and SubLedger='" + Trim(Scombo.Text) + "'", TCON1, adOpenStatic
  If rs4.RecordCount <= 0 Then
     MsgBox "No valid Sub Ledger"
     Scombo.SetFocus
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
                rs4.Open "Select* from sledger where GLEDGER='" + Trim(gencombo1.Text) + "' and SubLedger='" + Trim(Subcombo1.Text) + "'", TCON1, adOpenStatic
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
       DGrid.col = 0
       If DGrid.Text = "" Then
                 DGrid.col = 1
                 If DGrid.Text = "" Then
                      Commandsave.SetFocus
                      Exit Sub
                 End If
       End If
       DGrid.col = 1
       DGrid.Text = Subcombo1.Text
       DGrid.col = DGrid.col + 1
       DGrid_Click
 End If
 
 
 
End Sub

Private Sub TCNN_Change()
 Gridrefresh
 
    
End Sub

Private Sub TCNN_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{tab}"
End If
End Sub

Private Sub TCNN_LostFocus()
Dim rs1 As New ADODB.Recordset
     If TCNN.Text = "" Then
       TCNN.SetFocus
       Exit Sub
     End If
     rs1.Open "Select * from  DNFA  where DNN = " + TCNN.Text + "", TCON1, adOpenStatic, adLockOptimistic, adCmdText
     If Not rs1.EOF Then
     If cmdAdd Then
            MsgBox "Invoice already exist..."
            TCNN.SetFocus
            Exit Sub
     End If

End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  DGrid.col = 0
  DGrid.SetFocus
  DGrid.row = 1
  DGrid_Click
  
End If

KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Text3_Change()
Text3 = Format(Val(Text3.Text), "0.00")
End Sub
