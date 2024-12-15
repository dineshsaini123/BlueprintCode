VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmGLedger 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   8610
   ClientLeft      =   270
   ClientTop       =   1815
   ClientWidth     =   10215
   ClipControls    =   0   'False
   Icon            =   "frmGLedger.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8610
   ScaleWidth      =   10215
   Begin VB.Frame panel 
      BackColor       =   &H00E0E0E0&
      Height          =   8235
      Left            =   90
      TabIndex        =   8
      Top             =   90
      Width           =   10080
      Begin VB.CheckBox DebitFromRepSale 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00B8E4F1&
         Caption         =   "Debit From RepSale"
         Height          =   255
         Left            =   3780
         TabIndex        =   23
         Top             =   2280
         Width           =   2145
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Height          =   750
         Left            =   225
         ScaleHeight     =   750
         ScaleWidth      =   7770
         TabIndex        =   13
         Top             =   7215
         Width           =   7770
         Begin VB.CommandButton Commandmasterhelp 
            Caption         =   "Help"
            Height          =   345
            Left            =   -45
            TabIndex        =   21
            Top             =   855
            Visible         =   0   'False
            Width           =   800
         End
         Begin VB.CommandButton Commandmasteradd 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Add"
            Height          =   660
            Left            =   0
            Picture         =   "frmGLedger.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   30
            Width           =   1065
         End
         Begin VB.CommandButton Commandmasteredit 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Edit"
            Enabled         =   0   'False
            Height          =   660
            Left            =   1110
            Picture         =   "frmGLedger.frx":0BF0
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   30
            Width           =   1065
         End
         Begin VB.CommandButton Commandmastersave 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Sa&ve"
            Enabled         =   0   'False
            Height          =   660
            Left            =   2220
            Picture         =   "frmGLedger.frx":1032
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   30
            Width           =   1065
         End
         Begin VB.CommandButton Commandmasterabandon 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Aba&ndon"
            Enabled         =   0   'False
            Height          =   660
            Left            =   3330
            Picture         =   "frmGLedger.frx":1C16
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   30
            Width           =   1065
         End
         Begin VB.CommandButton Commandmasterdelete 
            BackColor       =   &H00FFFFFF&
            Caption         =   "De&lete"
            Enabled         =   0   'False
            Height          =   660
            Left            =   4440
            Picture         =   "frmGLedger.frx":21A0
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   30
            Width           =   1065
         End
         Begin VB.CommandButton Commandmastersearch 
            Caption         =   "&Search"
            Enabled         =   0   'False
            Height          =   525
            Left            =   5835
            TabIndex        =   16
            Top             =   1170
            Width           =   975
         End
         Begin VB.CommandButton CommandmasterPrint 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Print"
            Enabled         =   0   'False
            Height          =   660
            Left            =   5550
            Picture         =   "frmGLedger.frx":2D84
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   30
            Width           =   1065
         End
         Begin VB.CommandButton CommandmasterReturn 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Return"
            Height          =   660
            Left            =   6660
            Picture         =   "frmGLedger.frx":3968
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   30
            Width           =   1065
         End
      End
      Begin VB.CheckBox GMASTERSL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00B8E4F1&
         Caption         =   "Contains Sub Ledgers  "
         Height          =   255
         Left            =   90
         TabIndex        =   4
         Top             =   1560
         Width           =   2190
      End
      Begin VB.CheckBox GMASTERPL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00B8E4F1&
         Caption         =   "To be Included in P&&L"
         Height          =   255
         Left            =   90
         TabIndex        =   2
         Top             =   870
         Width           =   2175
      End
      Begin VB.CheckBox GMASTERBS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00B8E4F1&
         Caption         =   "To be included in B\S"
         Height          =   255
         Left            =   90
         TabIndex        =   3
         Top             =   1230
         Width           =   2175
      End
      Begin VB.ComboBox ComboSPECIALCATEGORY 
         Height          =   315
         Left            =   2055
         TabIndex        =   0
         Top             =   135
         Width           =   1590
      End
      Begin VB.TextBox Textglgeneralledgerdiscription 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2070
         MaxLength       =   39
         TabIndex        =   1
         Top             =   510
         Width           =   3525
      End
      Begin VB.TextBox Textglyearopeningbalance 
         Height          =   285
         Left            =   2070
         MaxLength       =   15
         TabIndex        =   5
         Top             =   1905
         Width           =   1815
      End
      Begin VB.TextBox Textfindgl 
         Height          =   345
         Left            =   3675
         TabIndex        =   9
         Top             =   135
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox Cashbankbook 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00B8E4F1&
         Caption         =   "Cash / Bank A/C "
         Height          =   255
         Left            =   75
         TabIndex        =   6
         Top             =   2265
         Width           =   2145
      End
      Begin VSFlex7Ctl.VSFlexGrid VS 
         Height          =   4380
         Left            =   105
         TabIndex        =   22
         Top             =   2640
         Width           =   9945
         _cx             =   17542
         _cy             =   7726
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
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
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
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
      Begin VB.Shape Shape3 
         BorderColor     =   &H0078CFE9&
         BorderWidth     =   4
         Height          =   825
         Left            =   180
         Top             =   7170
         Width           =   7845
      End
      Begin VB.Label Label9 
         BackColor       =   &H00B8E4F1&
         Caption         =   "Year Opening Balance"
         Height          =   270
         Left            =   90
         TabIndex        =   12
         Top             =   1890
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Specify Category"
         Height          =   225
         Left            =   135
         TabIndex        =   11
         Top             =   180
         Width           =   2940
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "General Ledger Description"
         Height          =   255
         Left            =   90
         TabIndex        =   10
         Top             =   555
         Width           =   2955
      End
   End
End
Attribute VB_Name = "frmGLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim CON As ADODB.Connection
Dim RS As ADODB.Recordset
Dim ctl As Control
Public addmaster As Boolean
Dim editing  As Boolean
Dim INVEVar As Integer
Private Sub cashCombocnepcontragenledgerdesc_Click()

If cashCombocnepcontragenledgerdesc.Text <> "" Then
    Set RS = New ADODB.Recordset
    RS.Open "select * from sledger where gledger='" + Trim(cashCombocnepcontragenledgerdesc.Text) + "' and " & stringyear & " ", con, adOpenKeyset, adLockReadOnly, adCmdText
    If Not RS.BOF Then
        RS.MoveFirst
        Do While Not RS.EOF
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.Close
End If

End Sub
Private Sub cashCombocnepcontragenledgerdesc_LostFocus()
If cashCombocnepcontragenledgerdesc.Text = "" Then
End If
If cashCombocnepcontragenledgerdesc.Text <> "" Then
        cashCombocnepcontragenledgerdesc.Text = UCase(cashCombocnepcontragenledgerdesc.Text)

        Set RS = New ADODB.Recordset
        RS.Open "select gledger from gledger where gledger='" + Trim(cashCombocnepcontragenledgerdesc.Text) + "'  and " & stringyear & " ", con, adOpenKeyset, adLockReadOnly, adCmdText
        If RS.EOF Then
            MsgBox cashCombocnepcontragenledgerdesc.Text + " Ledger not found"
            cashCombocnepcontragenledgerdesc.SetFocus
            'Exit Sub
        End If
        RS.Close
        RS.Open "select * from sledger where gledger='" + Trim(cashCombocnepcontragenledgerdesc.Text) + "' and " & stringyear, con, adOpenKeyset, adLockReadOnly, adCmdText
        If Not RS.BOF Then
            RS.MoveFirst
            Do While Not RS.EOF
                If Not RS.EOF Then
                   
                    RS.MoveNext
                End If
            Loop
            cashCombocnepcontrasubledgerdesc.SetFocus
         Else
            
        End If
        RS.Close
    End If
End Sub
Private Sub cashCombocnepcontrasubledgerdesc_LostFocus()
Dim RS As New ADODB.Recordset
      
If cashCombocnepcontragenledgerdesc <> "" And cashCombocnepcontrasubledgerdesc.ListCount > 0 And cashCombocnepcontrasubledgerdesc.Text = "" Then cashCombocnepcontrasubledgerdesc.SetFocus
  If cashCombocnepsubledgerdesc.ListCount > 0 And cashCombocnepcontrasubledgerdesc.Text <> "" Then
      RS.Open "Select* from sledger where GLEDGER='" + Trim(cashCombocnepcontragenledgerdesc.Text) + "' and SubLedger='" + Trim(cashCombocnepcontrasubledgerdesc.Text) + "' and " & stringyear, con, adOpenStatic
      If RS.RecordCount <= 0 Then
           MsgBox "No valid Sub Ledger"
           cashCombocnepcontrasubledgerdesc.SetFocus
      End If
End If
End Sub

Private Sub cashCombocnepdrorcr_LostFocus()
If cashCombocnepdrorcr.Text <> "Debit" And cashCombocnepdrorcr.Text <> "Credit" Then
  MsgBox "Please Enter Debit/Credit.."
 cashCombocnepdrorcr.SetFocus
End If
End Sub

Private Sub cashCombocnepgenledgerdesc_LostFocus()
If cashCombocnepgenledgerdesc.Text = "" Then
    
End If
  If cashCombocnepgenledgerdesc.Text <> "" Then
  cashCombocnepgenledgerdesc.Text = UCase(cashCombocnepgenledgerdesc.Text)
        Set RS = New ADODB.Recordset
        RS.Open "select gledger from gledger where gledger='" + Trim(cashCombocnepgenledgerdesc.Text) + "' and " & stringyear & " ", con, adOpenKeyset, adLockReadOnly, adCmdText
        If RS.EOF Then
            MsgBox cashCombocnepcontragenledgerdesc.Text + " Ledger not found"
            cashCombocnepcontragenledgerdesc.SetFocus
            'Exit Sub
        End If
        RS.Close
        RS.Open "select * from sledger where gledger='" + Trim(cashCombocnepgenledgerdesc.Text) + "' and " & stringyear, con, adOpenKeyset, adLockReadOnly, adCmdText
        If Not RS.BOF Then
            RS.MoveFirst
            Do While Not RS.EOF
                If Not RS.EOF Then
                    RS.MoveNext
                End If
            Loop
            cashCombocnepsubledgerdesc.SetFocus
        Else
        End If
        RS.Close
    End If




End Sub
Private Sub cashCombocnepsubledgerdesc_LostFocus()


If cashCombocnepgenledgerdesc <> "" And cashCombocnepsubledgerdesc.ListCount > 0 And Combocnepsubledgerdesc.Text = "" Then cashCombocnepsubledgerdesc.SetFocus
If cashCombocnepsubledgerdesc.ListCount > 0 And cashCombocnepsubledgerdesc.Text <> "" Then
    RS.Open "Select* from sledger where GLEDGER='" + Trim(cashCombocnepgenledgerdesc.Text) + "' and SubLedger='" + Trim(cashCombocnepsubledgerdesc.Text) + "' and " & stringyear, con, adOpenStatic
    If RS.RecordCount <= 0 Then
       MsgBox "No valid Sub Ledger"
       cashCombocnepsubledgerdesc.SetFocus
  End If
  

    
End If





End Sub

Private Sub cashTextcnep20chartext_LostFocus()
cashTextcnep20chartext.Text = UCase(cashTextcnep20chartext.Text)

End Sub

Private Sub cashTextInvePrintOrder_LostFocus()
If IsNumeric(cashTextInvePrintOrder.Text) = False Then
    MsgBox "Please Enter Any No..."
    cashTextInvePrintOrder.SetFocus
End If

End Sub

Private Sub CBODISTCODE_Change()
CBODISTCODE_Click
End Sub

Private Sub CBODISTCODE_Click()
On Error Resume Next
If addmaster = True Then
    Dim temp As ADODB.Recordset
    Set temp = New ADODB.Recordset
    temp.Open "select MAX(CONVERT(INT,SUBSTRING(SUBLEDGER,CHARINDEX('-',SUBLEDGER,1)+1,3))) AS MAXID from SLEDGER where " & stringyear & "  AND SUBLEDGER LIKE '" & UCase(CBODISTCODE.Text) & "%'", con, adOpenKeyset, adLockReadOnly, adCmdText
    If temp.EOF = False Then
        If temp!maxId > 0 Then
            TXTCUSTCODE.Caption = Format(temp!maxId + 1, "000")
        Else
            TXTCUSTCODE = "001"
        End If
    Else
        TXTCUSTCODE = "001"
    End If
    temp.Close
End If

If RS.State = adStateOpen Then RS.Close
RS.Open "SELECT DISTRICTNAME FROM DISTRICTS WHERE DISTCODE='" & CBODISTCODE.Text & "' AND " & stringyear
If RS.EOF = False Then
Combosldistrictcode.Text = RS!DISTRICTNAME
End If
End Sub

Private Sub CneTextInvePrintOrder_LostFocus()
If IsNumeric(CneTextInvePrintOrder.Text) = False Then
    MsgBox "Please Enter Any No..."
    CneTextInvePrintOrder.SetFocus
End If
End Sub

Private Sub Combobgroupcode_Change()
    Dim temp As ADODB.Recordset
    Set temp = New ADODB.Recordset
    temp.Open "select * from groups where " & stringyear & " ", con, adOpenKeyset, adLockReadOnly, adCmdText
    If Not temp.EOF Then
        If Not temp.EOF Then
        End If
    End If
    temp.Close
End Sub
Private Sub Combocnepcontragenledgerdesc_Click()
    If Combocnepcontragenledgerdesc.Text <> "" Then
        Set RS = New ADODB.Recordset
        RS.Open "select * from sledger where gledger='" + Trim(Combocnepcontragenledgerdesc.Text) + "' and " & stringyear, con, adOpenKeyset, adLockReadOnly, adCmdText
        If Not RS.BOF Then
            RS.MoveFirst
            Do While Not RS.EOF
                If Not RS.EOF Then
                    RS.MoveNext
                End If
            Loop
        End If
        RS.Close
    End If
End Sub

Private Sub Combocnepcontragenledgerdesc_LostFocus()
If Combocnepcontragenledgerdesc.Text = "" Then
    'Combocnepcontragenledgerdesc.SetFocus
End If
If Combocnepcontragenledgerdesc.Text <> "" Then
        Set RS = New ADODB.Recordset
        RS.Open "select gledger from gledger where gledger='" + Trim(Combocnepcontragenledgerdesc.Text) + "' and " & stringyear & " ", con, adOpenKeyset, adLockReadOnly, adCmdText
        If RS.EOF Then
            MsgBox Combocnepcontragenledgerdesc.Text + " Ledger not found"
            Combocnepcontragenledgerdesc.SetFocus
            'Exit Sub
        End If
        RS.Close
        RS.Open "select * from sledger where gledger='" + Trim(Combocnepcontragenledgerdesc.Text) + "' and " & stringyear, con, adOpenKeyset, adLockReadOnly, adCmdText
        'Me.Combocnepcontrasubledgerdesc.Clear
        
        If Not RS.BOF Then
            'Me.Combocnepcontrasubledgerdesc.Enabled = True
            RS.MoveFirst
            Do While Not RS.EOF
                'Me.Combocnepcontrasubledgerdesc.AddItem RS(1)
                If Not RS.EOF Then
                    RS.MoveNext
                End If
            Loop
            'Me.Combocnepcontrasubledgerdesc.SetFocus
        Else
        
          'Me.Combocnepcontrasubledgerdesc.Enabled = False
            
            
        End If
        RS.Close
    End If
End Sub
Private Sub Combocnepcontrasubledgerdesc_LostFocus()

Dim RS As New ADODB.Recordset
      
If Combocnepcontragenledgerdesc <> "" And Combocnepcontrasubledgerdesc.ListCount > 0 And Combocnepcontrasubledgerdesc.Text = "" Then Combocnepcontrasubledgerdesc.SetFocus
If Combocnepsubledgerdesc.ListCount > 0 And Combocnepcontrasubledgerdesc.Text <> "" Then
    
    RS.Open "Select* from sledger where GLEDGER='" + Trim(Combocnepcontragenledgerdesc.Text) + "' and SubLedger='" + Trim(Combocnepcontrasubledgerdesc.Text) + "' and " & stringyear, con, adOpenStatic
    If RS.RecordCount <= 0 Then
       MsgBox "No valid Sub Ledger"
       Combocnepcontrasubledgerdesc.SetFocus
  End If
  

    
End If

End Sub

Private Sub Combocnepdrorcr_LostFocus()
If Combocnepdrorcr.Text <> "Debit" And Combocnepdrorcr.Text <> "Credit" Then
         MsgBox "Please Enter Debit/Credit.."
         'Combocnepdrorcr.SetFocus
         
    End If
    
End Sub

Private Sub Combocnepgenledgerdesc_LostFocus()

   
   

 If Combocnepgenledgerdesc.Text <> "" Then
        Set RS = New ADODB.Recordset
        RS.Open "select gledger from gledger where  gledger='" + Trim(Combocnepgenledgerdesc.Text) + "' and " & stringyear & " ", con, adOpenKeyset, adLockReadOnly, adCmdText
        If RS.EOF Then
            MsgBox Combocnepgenledgerdesc.Text + " Ledger not found"
            Combocnepgenledgerdesc.SetFocus
            'Exit Sub
        End If
        RS.Close
        RS.Open "select * from sledger where gledger='" + Trim(Combocnepgenledgerdesc.Text) + "' and " & stringyear, con, adOpenKeyset, adLockReadOnly, adCmdText
        If Not RS.BOF Then
            RS.MoveFirst
            Do While Not RS.EOF
                If Not RS.EOF Then
                    RS.MoveNext
                End If
                
            Loop
            Combocnepsubledgerdesc.SetFocus
        Else
            'Me.Combocnepsubledgerdesc.Enabled = False
        End If
        RS.Close
    End If
 
 
 
 
End Sub
Private Sub Combocnepsubledgerdesc_LostFocus()
    
   
Dim RS As New ADODB.Recordset
      
If Combocnepgenledgerdesc <> "" And Combocnepsubledgerdesc.ListCount > 0 And Combocnepsubledgerdesc.Text = "" Then Combocnepsubledgerdesc.SetFocus
If Combocnepsubledgerdesc.ListCount > 0 And Combocnepsubledgerdesc.Text <> "" Then
    RS.Open "Select* from sledger where GLEDGER='" + Trim(Combocnepgenledgerdesc.Text) + "' and SubLedger='" + Trim(Combocnepsubledgerdesc.Text) + "' and " & stringyear, con, adOpenStatic
    If RS.RecordCount <= 0 Then
       MsgBox "No valid Sub Ledger"
       Combocnepsubledgerdesc.SetFocus
  End If
  

    
End If

  
End Sub

Private Sub Comboinvepcontragenledgerdesc_LostFocus()
If Comboinvepcontragenledgerdesc.Text = "" Then
    
End If


    If Comboinvepcontragenledgerdesc.Text <> "" Then
        'Me.Comboinvepcontrasubledgerdesc.Enabled = True
        Set RS = New ADODB.Recordset
        RS.Open "select gledger from gledger where gledger='" + Trim(Comboinvepcontragenledgerdesc.Text) + "' and " & stringyear & " ", con, adOpenKeyset, adLockReadOnly, adCmdText
        If RS.EOF Then
            MsgBox Comboinvepcontragenledgerdesc.Text + " Ledger not found"
            Comboinvepcontragenledgerdesc.SetFocus
        End If
        RS.Close
        RS.Open "select * from sledger where gledger='" + Trim(Comboinvepcontragenledgerdesc.Text) + "' and " & stringyear, con, adOpenKeyset, adLockReadOnly, adCmdText
        'Me.Comboinvepcontrasubledgerdesc.Clear
        If Not RS.BOF Then
            RS.MoveFirst
            'Me.Comboinvepcontrasubledgerdesc.Enabled = True
            Do While Not RS.EOF
                'Me.Comboinvepcontrasubledgerdesc.AddItem RS(1)
                If Not RS.EOF Then
                    RS.MoveNext
                End If
            Loop
            Comboinvepcontrasubledgerdesc.SetFocus
        Else
            'Me.Comboinvepcontrasubledgerdesc.Enabled = False
            
        End If
        RS.Close
    End If
End Sub

Private Sub Comboinvepcontrasubledgerdesc_LostFocus()
Dim RS As New ADODB.Recordset
      
If Comboinvepcontragenledgerdesc <> "" And Comboinvepcontrasubledgerdesc.ListCount > 0 And Comboinvepcontrasubledgerdesc.Text = "" Then Comboinvepcontrasubledgerdesc.SetFocus

If Comboinvepcontrasubledgerdesc.ListCount > 0 And Comboinvepcontrasubledgerdesc.Text <> "" Then
    RS.Open "Select* from sledger where GLEDGER='" + Trim(Comboinvepcontragenledgerdesc.Text) + "' and SubLedger='" + Trim(Comboinvepcontrasubledgerdesc.Text) + "'  and " & stringyear, con, adOpenStatic
    If RS.RecordCount <= 0 Then
       MsgBox "No valid Sub Ledger"
       Comboinvepcontrasubledgerdesc.SetFocus
  End If
  

    
End If
End Sub

Private Sub Comboinvepdrorcr_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    ' Commandmastersave.SetFocus
End If
End Sub

Private Sub Comboinvepdrorcr_LostFocus()
If Comboinvepdrorcr.Text <> "Debit" And Comboinvepdrorcr.Text <> "Credit" Then
  MsgBox "Please Enter Debit/Credit.."
  Comboinvepdrorcr.SetFocus
End If
End Sub

Private Sub Comboinvepgenledgerdesc_Click()
'Me.Comboinvepsubledgerdesc.Text = ""
' Me.Comboinvepsubledgerdesc.Enabled = True
If Comboinvepgenledgerdesc.Text <> "" Then
        Set RS = New ADODB.Recordset
        RS.Open "select * from sledger where gledger='" + Trim(Comboinvepgenledgerdesc.Text) + "'  and " & stringyear, con, adOpenKeyset, adLockReadOnly, adCmdText
        'Me.Comboinvepsubledgerdesc.Clear
        If Not RS.BOF Then
            'Me.Comboinvepsubledgerdesc.Enabled = True
            
            RS.MoveFirst
            Do While Not RS.EOF
                'Me.Comboinvepsubledgerdesc.AddItem RS(1)
                If Not RS.EOF Then
                    RS.MoveNext
                End If
            Loop
            Comboinvepsubledgerdesc.SetFocus
        Else
            'Me.Comboinvepsubledgerdesc.Enabled = False
        
            
        End If
        RS.Close
    End If
End Sub
Private Sub Comboinvepgenledgerdesc_LostFocus()
If Comboinvepgenledgerdesc.Text = "" Then
     'MsgBox "Enter Gen. Ledger..."
    ' Comboinvepgenledgerdesc.SetFocus
End If
     

    If Comboinvepgenledgerdesc.Text <> "" Then
        Set RS = New ADODB.Recordset
        RS.Open "select gledger from gledger where  gledger='" + Trim(Comboinvepgenledgerdesc.Text) + "'  and " & stringyear & " ", con, adOpenKeyset, adLockReadOnly, adCmdText
        If RS.EOF Then
            MsgBox Comboinvepgenledgerdesc.Text + " Ledger not found"
            Comboinvepgenledgerdesc.SetFocus
            'Exit Sub
        End If
        RS.Close
        RS.Open "select * from sledger where gledger='" + Trim(Comboinvepgenledgerdesc.Text) + "' and " & stringyear, con, adOpenKeyset, adLockReadOnly, adCmdText
        'Me.Comboinvepsubledgerdesc.Clear
        If Not RS.BOF Then
            'Me.Comboinvepsubledgerdesc.Enabled = True
            RS.MoveFirst
            Do While Not RS.EOF
                'Me.Comboinvepsubledgerdesc.AddItem RS(1)
                If Not RS.EOF Then
                    RS.MoveNext
                End If
                
            Loop
        Else
            'Me.Comboinvepsubledgerdesc.Enabled = False
        End If
        RS.Close
    End If
End Sub

Private Sub Comboinvepsubledgerdesc_LostFocus()
Dim RS As New ADODB.Recordset
If Comboinvepgenledgerdesc <> "" And Comboinvepsubledgerdesc.ListCount > 0 And Comboinvepsubledgerdesc.Text = "" Then
  'Comboinvepsubledgerdesc.Enabled = True
  Comboinvepsubledgerdesc.SetFocus
End If
If Comboinvepsubledgerdesc.ListCount > 0 And Comboinvepsubledgerdesc.Text <> "" Then
    RS.Open "Select* from sledger where GLEDGER='" + Trim(Comboinvepgenledgerdesc.Text) + "' and SubLedger='" + Trim(Comboinvepsubledgerdesc.Text) + "' and " & stringyear, con, adOpenStatic
    If RS.RecordCount <= 0 Then
       MsgBox "No valid Sub Ledger"
       Comboinvepsubledgerdesc.SetFocus
  End If
End If

End Sub

Private Sub Combosldistrictcode_Click()
'On Error Resume Next
If addmaster = True And Combosldistrictcode.Text <> "" Then
    Dim temp As ADODB.Recordset
    Set temp = New ADODB.Recordset
    'temp.Open "select MAX(CONVERT(INT,SUBSTRING(SUBLEDGER,CHARINDEX('-',SUBLEDGER,1)+1,3))) AS MAXID from SLEDGER where " & stringyear & "  AND SUBLEDGER LIKE '" & UCase(Combosldistrictcode.Text) & "%'", CON, adOpenKeyset, adLockReadOnly, adCmdText
    temp.Open "Select distcode from DISTRICTS where " & stringyear & " and  DISTRICTNAME='" & Combosldistrictcode.Text & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
    If temp.EOF = False Then
         txtdistcode.Caption = temp!distcode
    End If
    temp.Close
    End If


End Sub


Private Sub Comboslgenledgerdiscription_LostFocus()
    If Len(Comboslgenledgerdiscription.Text) >= 40 Then
           MsgBox "Enter only 40 Character"
           Comboslgenledgerdiscription.SetFocus
           Exit Sub
    End If
    
    If Comboslgenledgerdiscription.Text <> "" Then
        Set RS = New ADODB.Recordset
        RS.Open "select gledger from gledger where slf= 1 and gledger='" + Trim(Comboslgenledgerdiscription.Text) + "' and " & stringyear, con, adOpenKeyset, adLockReadOnly, adCmdText
        If RS.EOF Then
            MsgBox Comboslgenledgerdiscription.Text + " Ledger not found"
            Comboslgenledgerdiscription.SetFocus
        Else
            Comboslgenledgerdiscription.Text = RS!gledger
        End If
        RS.Close
    End If
    
End Sub

Private Sub ComboSPECIALCATEGORY_Click()
fillGrid
End Sub

Private Sub ComboSPECIALCATEGORY_LostFocus()
'If ComboSPECIALCATEGORY.Text = "" Then MsgBox "Enter Category"
If Len(ComboSPECIALCATEGORY.Text) > 10 Then MsgBox "Enter only 10 Character"
End Sub

Private Sub Commandmasterabandon_Click()
    addmaster = False
    editing = False

SetButton Commandmasteradd, Commandmasteredit, Commandmastersave, Commandmasterdelete
Commandmasteredit.Enabled = False
Commandmastersave.Enabled = False
Commandmasterdelete.Enabled = False

Commandmasteradd.Enabled = True


End Sub
Private Sub Commandmasteradd_Click()
     addmaster = True
     For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
        If ctl.Enabled = True Then ctl.Text = ""
            'ctl.Enabled = False
        End If
        If TypeOf ctl Is ComboBox Then
        If ctl.Style <> 2 Then
        ctl.Text = ""
        End If
        End If
        If TypeOf ctl Is CheckBox Then
            ctl.value = 0
            'ctl.Enabled = False
        End If
        If TypeOf ctl Is ListBox Then
            'ctl.Enabled = False
        End If
    Next
    
    
    ComboSPECIALCATEGORY.Enabled = True
    
    
    ComboSPECIALCATEGORY.Clear
    ComboSPECIALCATEGORY.AddItem "Assets"
    ComboSPECIALCATEGORY.AddItem "Liability"
    ComboSPECIALCATEGORY.AddItem "Income"
    ComboSPECIALCATEGORY.AddItem "Expences"
    Textglgeneralledgerdiscription.Enabled = True

    
    
    'If SSTab1.Tab = 0 Then
    
    '/*  deactivate other tabs */
   
'        For i = 0 To 5
'            If i <> SSTab1.Tab Then
'                SSTab1.TabEnabled(i) = False
'            End If
'        Next
'        For Each ctl In Me.Controls
'            If UCase(ctl.Container.Name) = UCase("gledger") Then
'                If TypeOf ctl Is textbox Or TypeOf ctl Is ComboBox Or TypeOf ctl Is CheckBox Then
'                    ctl.Enabled = True
'                End If
'            End If
'        Next
     ' frmGLedgergledger.Enabled = True
      'frmGLedger.ComboSPECIALCATEGORY.Enabled = True
       'Me.ComboSPECIALCATEGORY.SetFocus
     
 '  End If
    
  '  If SSTab1.Tab = 1 Then
    '/**  deactivate other tabs**/
'        For i = 0 To 5
'            If i <> SSTab1.Tab Then
'                SSTab1.TabEnabled(i) = False
'            End If
'        Next
'        For Each ctl In Me.Controls
'            If ctl.Container.Name = "sledger" Then
'                If TypeOf ctl Is textbox Or TypeOf ctl Is ComboBox Or TypeOf ctl Is CheckBox Then
'                    ctl.Enabled = True
'                End If
'            End If
'        Next
       'frmGLedger.sledger.Enabled = True
 '       Combosldistrictcode_Click
       
 '   End If
  '  If SSTab1.Tab = 2 Then
    '/**  deactivate other tabs**/
'        For i = 0 To 5
'            If i <> SSTab1.Tab Then
'                SSTab1.TabEnabled(i) = False
'            End If
'        Next
        For Each ctl In Me.Controls
            If UCase(ctl.Container.Name) = UCase("invnoteend") Then
                If TypeOf ctl Is TextBox Or TypeOf ctl Is ComboBox Or TypeOf ctl Is CheckBox Then
                    ctl.Enabled = True
                End If
            End If
        Next
       'frmGLedger.invnoteend.Enabled = True
        'Comboinvepcontragenledgerdesc.SetFocus
      ' frmGLedger.TextInvePrintOrder.SetFocus
        
    '    Textinveprate.Text = "0"
  '  End If
    
    
   ' If SSTab1.Tab = 3 Then
    '/**  deactivate other tabs**/
        
       'SSTab1.Tab = 3
        
'        For i = 0 To 5
'            If i <> SSTab1.Tab Then
'                SSTab1.TabEnabled(i) = False
'            End If
'        Next
        For Each ctl In Me.Controls
            If UCase(ctl.Container.Name) = UCase("crenoteend") Then
                If TypeOf ctl Is TextBox Or TypeOf ctl Is ComboBox Or TypeOf ctl Is CheckBox Then
                    ctl.Enabled = True
                End If
            'crenoteend.Enabled = True
            End If
        Next
    
       'Combocnepcontragenledgerdesc.SetFocus
      'frmGLedger.CneTextInvePrintOrder.SetFocus
    '   Textcneprate.Text = "0"
   ' End If
    
    
    'If SSTab1.Tab = 4 Then
    '/**  deactivate other tabs**/
        
        
'        For i = 0 To 5
'            If i <> SSTab1.Tab Then
'                SSTab1.TabEnabled(i) = False
'            End If
'        Next
'        For Each ctl In Me.Controls
'            If UCase(ctl.Container.Name) = UCase("discount") Then
'                If TypeOf ctl Is textbox Or TypeOf ctl Is ComboBox Or TypeOf ctl Is CheckBox Then
'                    ctl.Enabled = True
'                End If
'            End If
'        Next
  '  End If
    
    
    
'If SSTab1.Tab = 5 Then
    '/**  deactivate other tabs**/
        
       
        
'        For i = 0 To 5
'            If i <> SSTab1.Tab Then
'                SSTab1.TabEnabled(i) = False
'            End If
'        Next
        For Each ctl In Me.Controls
            If UCase(ctl.Container.Name) = UCase("cashend") Then
                If TypeOf ctl Is TextBox Or TypeOf ctl Is ComboBox Or TypeOf ctl Is CheckBox Then
                    ctl.Enabled = True
                End If
            End If
        Next
    
       'cashCombocnepcontragenledgerdesc.SetFocus
      'frmGLedger.cashTextInvePrintOrder.SetFocus
      'frmGLedger.cashTextcneprate.Text = "0"
   ' End If
    

    
    
    
    
    
    Commandmasteradd.Enabled = False
    Commandmasteredit.Enabled = False
    CommandmasterPrint.Enabled = False
    Commandmastersave.Enabled = True
    Commandmasterabandon.Enabled = True
    CommandmasterReturn.Enabled = True
    Commandmastersearch.Enabled = True
    
    ComboSPECIALCATEGORY.SetFocus
    
End Sub
    
Private Sub Commandmasterdelete_Click()

If MsgBox("Want To Delete ?", vbInformation + vbYesNo) = vbYes Then
  con.Execute "delete from GLEDGER where (Category='" & ComboSPECIALCATEGORY.Text & "' and gledger='" & Textglgeneralledgerdiscription.Text & "') and " & "" & stringyear
  Commandmasteradd_Click
  
  Commandmasterdelete.Enabled = False
  Commandmastersave.Enabled = False
  
  fillGrid
End If


End Sub
Public Sub Commandmasteredit_Click()
    
    editing = True
    addmaster = False
    Me.Commandmasteradd.Enabled = False
    Me.Commandmasteredit.Enabled = False
    Me.Commandmasterabandon.Enabled = True
    Me.Commandmasteredit.Enabled = True
    
    
'    If SSTab1.Tab = 0 Then
       Textglyearopeningbalance.Enabled = True
       Textglgeneralledgerdiscription.Enabled = True
       Commandmastersave.Enabled = True
       Commandmasterdelete.Enabled = True
       frmGLedger.GMASTERPL.Enabled = True
       frmGLedger.GMASTERBS.Enabled = True
       frmGLedger.Enabled = True
       frmGLedger.GMASTERSL.Enabled = True
       frmGLedger.Cashbankbook.Enabled = True
       frmGLedger.ComboSPECIALCATEGORY.Enabled = True
       frmGLedger.Textglgeneralledgerdiscription.Enabled = False
       frmGLedger.Textglyearopeningbalance.Enabled = True
       Me.Textglgeneralledgerdiscription.Enabled = True
       Textglgeneralledgerdiscription.Enabled = True
       
 '   End If
    
''    If SStab1.Tab = 1 Then
''        Me.Comboslgenledgerdiscription.Enabled = True
''        Me.Textslsubledgerdiscription.Enabled = True
''        Textslfindgl.Enabled = True
''        TextFINDSUBLEADGER.Enabled = True
''        Textsldiscriptionforinvoice.Enabled = True
''        Textsladdress1.Enabled = True
''        Textsladdress2.Enabled = True
''        Textsladdress3.Enabled = True
''        Textslyearopeningbalance.Enabled = True
''        Combosldiscountcategory.Enabled = True
''        Combosldistrictcode.Enabled = True
''    End If
''
'''    If SSTab1.Tab = 2 Then
'''        INVEVar = Val(TextInvePrintOrder.Text)
'''    End If
'''    If SSTab1.Tab = 4 Then
'''        INVEVar = Val(TextcnePrintOrder.Text)
'''    End If
'''
'''     If SSTab1.Tab = 5 Then
'''        INVEVar = Val(TextOInvePrintOrder.Text)
'''    End If
'       frmGLedger.Commandmastersave.Enabled = True
'       frmGLedger.Commandmasteredit.Enabled = False
'       frmGLedger.Commandmasteradd.Enabled = False
'       frmGLedger.Commandmasterdelete.Enabled = True
'       frmGLedger.Commandmasterabandon.Enabled = True

End Sub
Private Sub CommandmasterPrint_Click()
'    If SStab1.Tab = 0 Then
'    '    Genledgerprinting.Show
'    ElseIf SStab1.Tab = 1 Then
'    'MainMenu.cr1.Connect = constr
'        'MainMenu.cr1.ReportFileName = strrptpath & "\rEPORTS\subledgerlist.RPT"
'        'MainMenu.cr1.SelectionFormula = "{SLEDGER.fyear}='" & main.session & "' and {SLEDGER.setupid}=" & main.setupid & IIf(Trim(Comboslgenledgerdiscription) <> "", " AND {SLEDGER.GENLEDGER}='" & Comboslgenledgerdiscription & "'", "")
'        'MainMenu.cr1.WindowShowPrintBtn = True
'        'MainMenu.cr1.WindowShowPrintSetupBtn = True
'        'MainMenu.cr1.WindowState = crptMaximized
'        'MainMenu.cr1.Action = 1
'    End If
    
End Sub
Private Sub CommandmasterReturn_Click()

'''''MainMenu.Toolbar1.Visible = True
    Unload Me
End Sub
Private Sub Commandmastersave_Click()
Dim SAVED As Boolean
SAVED = False
'/////////////////*************
'   saving gen ledger
'/////////////////*************
      
'If SStab1.Tab = 0 Then
  Commandmasteradd.Enabled = True
  Commandmasteredit.Enabled = True
  'gledger.Enabled = False  '************ for frame unlock
  Commandmasteradd.Enabled = True
  Commandmasteredit.Enabled = True
  Commandmastersave.Enabled = False
  If ComboSPECIALCATEGORY.Text <> "" And Textglgeneralledgerdiscription <> "" Then
       If RS.State = 1 Then RS.Close
        RS.Open "select * from gledger where " & stringyear, con, adOpenKeyset, adLockOptimistic, adCmdText
        If addmaster = True Then
            RS.Find "gledger='" + Trim(Me.Textglgeneralledgerdiscription.Text) + "'"
            If Not RS.EOF Then
                MsgBox "Record Already exist... "
            Else
                For I = 0 To UBound(arycname)
                RS.AddNew
                RS!gledger = Trim(UCase(Textglgeneralledgerdiscription.Text))
                RS!category = ComboSPECIALCATEGORY.Text
                If GMASTERPL.value = 0 Then
                    RS!PLC = False
                Else
                    RS!PLC = True
                End If
                If GMASTERBS.value = 0 Then
                    RS!BSC = False
                Else
                    RS!BSC = True
                End If
                If GMASTERSL.value = 0 Then
                    RS!SLF = False
                Else
                    RS!SLF = True
                End If
                RS!YEAROPENING = Val(Textglyearopeningbalance.Text)
                If Cashbankbook.value = 0 Then
                    RS!Cashbankbook = False
                Else
                    RS!Cashbankbook = True
                End If
                
                If DebitFromRepSale.value = 0 Then
                   RS!DebitFromRepSale = False
                Else
                   RS!DebitFromRepSale = True
                End If
                
                
                RS!fyear = main.session
                RS!setupid = Val(Left(arycname(I), InStr(1, arycname(I), " (")))
                'RS!createdby = main.username
                'RS!createdon = Now
                RS.update
                Next
                SAVED = True
            End If
        Else
            If Not RS.BOF Then
                RS.MoveFirst
            End If
            RS.Find "gledger='" + Trim(Me.Textfindgl.Text) + "'"
            If RS.EOF Then
                MsgBox "Not Found!.."
            Else
                
                RS!gledger = Trim(UCase(Textglgeneralledgerdiscription.Text))
                RS!category = ComboSPECIALCATEGORY.Text
                If GMASTERPL.value = 0 Then
                    RS!PLC = False
                Else
                    RS!PLC = True
                End If
                If GMASTERBS.value = 0 Then
                    RS!BSC = False
                Else
                    RS!BSC = True
                End If
                If Cashbankbook.value = 0 Then
                    RS!Cashbankbook = False
                Else
                    RS!Cashbankbook = True
                End If
                If GMASTERSL.value = 0 Then
                    RS!SLF = False
                Else
                    RS!SLF = True
                End If
                RS!YEAROPENING = Val(Textglyearopeningbalance.Text)
                'rs!updatedby = main.username
                'rs!updatedon = Now
                
                If DebitFromRepSale.value = 0 Then
                   RS!DebitFromRepSale = False
                Else
                   RS!DebitFromRepSale = True
                End If
                
                RS.update
                SAVED = True
                
                '------------------------------------------------
                
                
                'updateGledger Trim(UCase(Textglgeneralledgerdiscription.Text)), vs.TextMatrix(vs.RowSel, 1)
                con.Execute "exec UpdateLedger 'GLEDGER','" & vs.TextMatrix(vs.RowSel, 1) & "','" & Trim(UCase(Textglgeneralledgerdiscription.Text)) & "','" & session & "','" & main.setupid & "'"
                
                
                MsgBox "Data Saved ...", vbInformation
                
            
            End If
        End If
    Else
        MsgBox "ONE OF THE REQUIRED DATA IS NULL. PLEASE CHECK.."
    End If
    If RS.State = 1 Then
        RS.Close
    End If
    Commandmasteredit.Enabled = False
    Commandmasterdelete.Enabled = False
    


Commandmasterdelete.Enabled = False
Commandmastersave.Enabled = False
  



Commandmastersave.Enabled = False
Commandmasteredit.Enabled = False
fillGrid

End Sub
Private Sub Commandmastersearch_Click()
  Me.Enabled = False
  Call searchscreen.tempr(SSTab1.Tab, Me.Name)
  SetButton Commandmasteradd, Commandmasteredit, Commandmastersave, Commandmasterdelete
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
End Sub
Sub fillGrid()
    
    Dim f As New ADODB.Recordset
    If f.State = 1 Then f.Close
    If ComboSPECIALCATEGORY.Text = "" Then
      s = "SELECT [Category],[gledger],[PLC],[BSC],[SLF],[YEAROPENING],[cashbankbook],DebitFromRepSale,Auto FROM [GLEDGER] WHERE " & stringyear & " ORDER BY [Category],[gledger]"
    Else
      s = "SELECT [Category],[gledger],[PLC],[BSC],[SLF],[YEAROPENING],[cashbankbook],DebitFromRepSale,Auto FROM [GLEDGER] where " & stringyear & " AND Category='" & Trim(ComboSPECIALCATEGORY.Text) & "' ORDER BY [Category],[gledger]"
    End If
    
    f.Open s, con
    Set vs.DataSource = f
    
    
    vs.ColWidth(0) = 1000
    vs.ColWidth(1) = 1900
    vs.ColWidth(2) = 800
    vs.ColWidth(3) = 800
    vs.ColWidth(4) = 800
    vs.ColWidth(5) = 1300
    vs.ColWidth(6) = 1200
    vs.ColWidth(7) = 1500
    vs.ColWidth(8) = 0
      
    
End Sub
Private Sub Form_Load()



' /****      FRAMEINI      ****/
    fillGrid
    Me.Top = 20
    Me.Left = 100
    Me.Height = 9120
    Me.Width = 10800
    
    
    Dim TMPA As Control
    editing = False
    INVEVar = 0
    
    Commandmastersearch.Enabled = True
    CommandmasterPrint.Enabled = True
    SetButton Commandmasteradd, Commandmasteredit, Commandmastersave, Commandmasterdelete
    Commandmasteredit.Enabled = False
    Commandmastersave.Enabled = False
    Commandmasterdelete.Enabled = False
        
      ComboSPECIALCATEGORY.AddItem "Assets"
    ComboSPECIALCATEGORY.AddItem "Liability"
    ComboSPECIALCATEGORY.AddItem "Income"
    ComboSPECIALCATEGORY.AddItem "Expences"
      
        
        
    'BackColorFrom Me
        
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
'refreshme
End Sub

Private Sub Textcbgenledger_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Dim RS As New ADODB.Recordset
   If RS.State = 1 Then RS.Close
   RS.Open "select * from SLEDGER where GLEDGER ='" & Textcbgenledger.Text & "' and " & stringyear, con, adOpenKeyset, adLockOptimistic, adCmdText
   If Not RS.EOF Then
        Do While Not RS.EOF
            Textcbsubledger.AddItem RS(1)
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.Close
End If
End Sub

Private Sub Textcbgenledger_LostFocus()
   
       If Textcbgenledger.Text <> "" Then
            Dim rs4 As New ADODB.Recordset
            rs4.Open "Select* from gledger where GLEDGER = '" + Trim(Textcbgenledger.Text) + "' and " & stringyear, con, adOpenStatic
            If rs4.RecordCount <= 0 Then
                 MsgBox "No valid Gen.Ledger"
                 Textcbgenledger.SetFocus
            End If
        End If
   

End Sub

Private Sub Textcbsubledger_LostFocus()
If Textcbsubledger.Text <> "" And Textcbgenledger.Text <> "" Then
   Dim rs4 As New ADODB.Recordset
   rs4.Open "Select* from sledger where GLEDGER='" + Trim(Textcbgenledger.Text) + "' and SubLedger='" + Trim(Textcbsubledger.Text) + "' and " & stringyear, con, adOpenStatic
   If rs4.RecordCount <= 0 Then
      MsgBox "No valid Sub Ledger"
      Textcbsubledger.SetFocus
   End If
End If
If Textcbsubledger.ListCount > 0 And Textcbsubledger.Text = "" Then
'    Textcbsubledger.SetFocus
  
 End If

End Sub

Private Sub Textcnep20chartext_LostFocus()
 Textcnep20chartext.Text = UCase(Textcnep20chartext.Text)
End Sub

Private Sub Textcneprate_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Then
    Else
        If KeyAscii <> 46 Then
            If KeyAscii <> 8 And KeyAscii <> 45 Then
                KeyAscii = 0
            End If
        End If
    End If
End Sub
Private Sub Textdcdiscountcategorycode_LostFocus()
    Textdcdiscountcategorycode.Text = UCase(Textdcdiscountcategorycode.Text)
End Sub
Private Sub Textdcdiscountrate_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Then
 Else
        If KeyAscii <> 46 Then
            If KeyAscii <> 8 And KeyAscii <> 45 Then
                KeyAscii = 0
            End If
        End If
    End If
End Sub

Private Sub Textdcdiscountrate_LostFocus()
   Textdcdiscountrate.Text = Format(Textdcdiscountrate.Text, "0.00")
End Sub

Private Sub Form_Resize()
panel.Left = (Me.ScaleWidth - panel.Width) / 2
panel.Top = (Me.ScaleHeight - panel.Height) / 2

End Sub

Private Sub Textglgeneralledgerdiscription_LostFocus()
    Textglgeneralledgerdiscription.Text = UCase(Textglgeneralledgerdiscription.Text)
        Set RS = New ADODB.Recordset
        RS.Open "select * from gledger where " & stringyear, con, adOpenKeyset, adLockReadOnly, adCmdText
        RS.Find "gledger='" + Trim(Textglgeneralledgerdiscription.Text) + "'"
        If Not RS.EOF Then
            Textglgeneralledgerdiscription.Text = RS!gledger
            Textfindgl.Text = RS!gledger
            ComboSPECIALCATEGORY.Text = RS!category
            If RS!PLC = False Then
                GMASTERPL.value = 0
            Else
                GMASTERPL.value = 1
            End If
            If RS!BSC = False Then
                GMASTERBS.value = 0
            Else
                GMASTERBS.value = 1
            End If
            If RS!Cashbankbook = False Then
                Cashbankbook.value = 0
            Else
                Cashbankbook.value = 1
            End If
            If RS!SLF = False Then
                GMASTERSL.value = 0
            Else
                GMASTERSL.value = 1
            End If
            Textglyearopeningbalance.Text = RS!YEAROPENING
            ComboSPECIALCATEGORY.Text = RS!category
        End If
        RS.Close
End Sub

Private Sub Textglyearopeningbalance_GotFocus()
Textglyearopeningbalance.Text = Format(Textglyearopeningbalance.Text, "0.00")
End Sub

Private Sub Textglyearopeningbalance_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 Then
    Else
        If KeyAscii <> 46 Then
            If KeyAscii <> 8 And KeyAscii <> 45 Then
                KeyAscii = 0
            End If
        End If
    End If
End Sub



Private Sub Textinvep20chartext_LostFocus()
Textinvep20chartext.Text = UCase(Textinvep20chartext.Text)
End Sub

Private Sub Textinveprate_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 Then
    Else
        If KeyAscii <> 46 Then
            If KeyAscii <> 8 And KeyAscii <> 45 Then
                KeyAscii = 0
            End If
        End If
    End If
End Sub

Private Sub TextInvePrintOrder_LostFocus()


If IsNumeric(TextInvePrintOrder.Text) = False Then
    MsgBox "Please Enter Any No..."
    TextInvePrintOrder.SetFocus
End If


End Sub

Private Sub Textsladdress1_LostFocus()
    Textsladdress1.Text = UCase(Textsladdress1.Text)
End Sub
Private Sub Textsladdress2_LostFocus()
    Textsladdress2.Text = UCase(Textsladdress2.Text)
End Sub
Private Sub Textsladdress3_LostFocus()
    Textsladdress3.Text = UCase(Textsladdress3.Text)
End Sub
Private Sub Textsldiscriptionforinvoice_LostFocus()
    Textsldiscriptionforinvoice.Text = UCase(Textsldiscriptionforinvoice.Text)
End Sub

Private Sub Textslsubledgerdiscription_GotFocus()
If Trim(Comboslgenledgerdiscription.Text) = "" Then
    MsgBox "Please Select Gen. Ledger ."
    Comboslgenledgerdiscription.SetFocus
End If
End Sub

Private Sub Textslsubledgerdiscription_LostFocus()
If Trim(Comboslgenledgerdiscription.Text) <> "" Then
    Textslsubledgerdiscription.Text = UCase(Textslsubledgerdiscription.Text)
    Set RS = New ADODB.Recordset
        RS.Open "select * from SLEDGER where gledger='" + Trim(Comboslgenledgerdiscription.Text) + "' and subledger='" + IIf(Combosldistrictcode.Text <> "", Combosldistrictcode.Text & "-" & TXTCUSTCODE, "") & Trim(Textslsubledgerdiscription.Text) + "' and " & stringyear, con, adOpenKeyset, adLockReadOnly, adCmdText
        If Not RS.EOF Then
            Textslfindgl.Text = RS!gledger
            TextFINDSUBLEADGER.Text = RS!subledger
            
             
            If IsNull(RS!DESCFORINVOICE) Then
               Textsldiscriptionforinvoice.Text = ""
            Else
                Textsldiscriptionforinvoice.Text = RS!DESCFORINVOICE

            End If
            
            If IsNull(RS!address1) Then
               Textsladdress1.Text = ""
            Else
               Textsladdress1.Text = RS!address1
            End If
            
            If IsNull(RS!address2) Then
                    Textsladdress2.Text = ""
               Else
                  Textsladdress2.Text = RS!address2

            End If
            
            If IsNull(RS!address3) Then
               Textsladdress3.Text = ""
            Else
               Textsladdress3.Text = RS!address3
            End If
            'Textsladdress2.Text = rs!ADDRESS2
            'Textsladdress3.Text = rs!ADDRESS3
            Textslyearopeningbalance.Text = Format(RS!YEAROPENING, "0.00")
            If IsNull(RS!DISCATEGORY) Then
                 Combosldiscountcategory.Text = ""
            Else
                Combosldiscountcategory.Text = RS!DISCATEGORY
                
            End If
            
            If IsNull(RS!distcode) Then
                Combosldistrictcode.Text = ""
               
            ElseIf RS!distcode <> "" Then
                Combosldistrictcode.Text = RS!distcode
                Else
                Combosldistrictcode.ListIndex = 0
            End If

            
            If editing Then
                'editing = False
                'Me.Comboslgenledgerdiscription.Enabled = False
            End If
        End If
        RS.Close
End If
If Trim(Textslsubledgerdiscription) = "" Then
    'Me.Comboslgenledgerdiscription.Enabled = True
End If
End Sub
Private Sub Textslyearopeningbalance_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 Then
    Else
        If KeyAscii <> 46 Then
            If KeyAscii <> 8 And KeyAscii <> 45 Then
                KeyAscii = 0
            End If
        End If
    End If
End Sub

Function refreshme()
' /****      FRAMEINI      ****/
    Me.Top = 20
    Me.Left = 200
    Dim TMPA As Control
    editing = False
    For Each TMPA In Me.Controls
        If TypeOf TMPA Is VB.Frame Then
            TMPA.Top = 1200
            TMPA.Left = 800
            TMPA.Width = 7515
            TMPA.Height = 4005
        End If
        If TypeOf TMPA Is TextBox Then
            TMPA.Enabled = False
        End If
        If TypeOf TMPA Is CheckBox Then
            TMPA.Enabled = False
        End If
        If TypeOf TMPA Is ComboBox Then
            TMPA.Enabled = False
        End If
    Next
    ' ComboSPECIALCATEGORY INI
    ComboSPECIALCATEGORY.AddItem "Assets"
    ComboSPECIALCATEGORY.AddItem "Liability"
    ComboSPECIALCATEGORY.AddItem "Income"
    ComboSPECIALCATEGORY.AddItem "Expences"
    
    Set RS = New ADODB.Recordset
    RS.Open "select * from gledger where slf=1 and " & stringyear, con, adOpenKeyset, adLockOptimistic, adCmdText
    If Not RS.EOF Then
        Do While Not RS.EOF
            Comboslgenledgerdiscription.AddItem RS!gledger
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.Close
    
    RS.Close
    RS.Open "select gledger from gledger where slf=1 and  " & stringyear & "   order by gledger", con, adOpenKeyset, adLockReadOnly, adCmdText
    If Not RS.EOF Then
        Do While Not RS.EOF
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.Close
    RS.Open "select subledger from sledger where  " & stringyear & "   order by subledger", con, adOpenKeyset, adLockReadOnly, adCmdText
    If Not RS.EOF Then
        Do While Not RS.EOF
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.Close
    Commandmastersearch.Enabled = True
    CommandmasterPrint.Enabled = True
End Function

Private Sub Textslyearopeningbalance_LostFocus()
Textslyearopeningbalance.Text = Format(Textslyearopeningbalance.Text, "0.00")
End Sub
Private Sub vs_Click()
                
                
Set RS = New ADODB.Recordset
RS.Open "select * from GLEDGER where (" & stringyear & " and Category='" & vs.TextMatrix(vs.RowSel, 0) & "' and gledger='" & vs.TextMatrix(vs.RowSel, 1) & "')", con
If RS.EOF = False Then
    
    Commandmasteredit.Enabled = True
    
    Textglgeneralledgerdiscription.Text = RS!gledger
    Textfindgl.Text = RS!gledger
    ComboSPECIALCATEGORY.Text = RS!category
    If RS!PLC = False Then
       GMASTERPL.value = 0
    Else
       GMASTERPL.value = 1
    End If
    
    If RS!BSC = False Then
       GMASTERBS.value = 0
    Else
        GMASTERBS.value = 1
    End If
    
    
    If RS!SLF = False Then
       GMASTERSL.value = 0
    Else
       GMASTERSL.value = 1
    End If
    
    Textglyearopeningbalance.Text = RS!YEAROPENING
    
    If RS!Cashbankbook = False Then
       Cashbankbook.value = 0
    Else
       Cashbankbook.value = 1
    End If
    
    If (RS!DebitFromRepSale = False Or IsNull(RS!DebitFromRepSale)) Then
        DebitFromRepSale.value = 0
      Else
        DebitFromRepSale.value = 1
    End If
    
    

End If

End Sub
