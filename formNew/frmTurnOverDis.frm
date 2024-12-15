VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTurnOverDis 
   ClientHeight    =   9900
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   14976
   Icon            =   "frmTurnOverDis.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9900
   ScaleWidth      =   14976
   Begin VB.TextBox txtOSAmt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   9225
      TabIndex        =   37
      Top             =   7830
      Width           =   870
   End
   Begin VB.ListBox cboschool 
      Appearance      =   0  'Flat
      ForeColor       =   &H000000FF&
      Height          =   3864
      Left            =   10215
      TabIndex        =   35
      Top             =   1230
      Width           =   4365
   End
   Begin VB.TextBox txtTotal10 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   5370
      TabIndex        =   34
      Top             =   7830
      Width           =   975
   End
   Begin VB.CommandButton cmdOK_ 
      Caption         =   "&OK"
      Height          =   495
      Left            =   11955
      TabIndex        =   33
      Top             =   420
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox Check1_Edit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Date Where The Data Used"
      Height          =   435
      Left            =   10155
      TabIndex        =   31
      Top             =   420
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.TextBox txtTODDiS 
      Height          =   315
      Left            =   13485
      TabIndex        =   28
      Top             =   5265
      Width           =   735
   End
   Begin VB.TextBox txtTotal8 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   11850
      TabIndex        =   27
      Top             =   7830
      Width           =   855
   End
   Begin VB.TextBox txtTotal7 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   10995
      TabIndex        =   26
      Top             =   7830
      Width           =   825
   End
   Begin VB.TextBox txtTotal9 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   13380
      TabIndex        =   25
      Top             =   7830
      Width           =   1035
   End
   Begin VB.TextBox txtTotal6 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   10125
      TabIndex        =   24
      Top             =   7830
      Width           =   855
   End
   Begin VB.CommandButton cmdView 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&View"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6615
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      Width           =   1275
   End
   Begin VB.TextBox txtTotal1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3390
      TabIndex        =   16
      Top             =   7830
      Width           =   975
   End
   Begin VB.TextBox txtSponsorshipNo 
      Height          =   315
      Left            =   1275
      TabIndex        =   15
      Top             =   60
      Width           =   855
   End
   Begin VB.TextBox txtRemarks 
      Height          =   315
      Left            =   15
      MaxLength       =   200
      TabIndex        =   14
      Top             =   8520
      Width           =   6675
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   930
      Left            =   6900
      ScaleHeight     =   936
      ScaleWidth      =   6048
      TabIndex        =   6
      Top             =   8505
      Width           =   6045
      Begin VB.CommandButton cmdPrint_7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Print"
         Height          =   720
         Left            =   3900
         Picture         =   "frmTurnOverDis.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   60
         Width           =   1005
      End
      Begin VB.CommandButton cmdEdit_4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   1920
         Picture         =   "frmTurnOverDis.frx":0BF0
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   60
         Width           =   945
      End
      Begin VB.CommandButton close 
         BackColor       =   &H00FFFFFF&
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   720
         Left            =   4905
         Picture         =   "frmTurnOverDis.frx":0FFD
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   60
         Width           =   1065
      End
      Begin VB.CommandButton REPORTCD 
         Caption         =   "&Print"
         Height          =   585
         Left            =   4860
         TabIndex        =   10
         Top             =   1020
         Width           =   1245
      End
      Begin VB.CommandButton Abandon 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Refresh"
         Height          =   720
         Left            =   0
         Picture         =   "frmTurnOverDis.frx":1BE1
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   60
         Width           =   945
      End
      Begin VB.CommandButton Del 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         Height          =   720
         Left            =   2880
         Picture         =   "frmTurnOverDis.frx":27C5
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   60
         Width           =   1005
      End
      Begin VB.CommandButton save 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         Height          =   720
         Left            =   960
         Picture         =   "frmTurnOverDis.frx":33A9
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   60
         Width           =   945
      End
   End
   Begin VB.TextBox txtTotal3 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   6330
      TabIndex        =   5
      Top             =   7830
      Width           =   930
   End
   Begin VB.TextBox txtTotal2 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4365
      TabIndex        =   4
      Top             =   7830
      Width           =   975
   End
   Begin VB.ComboBox cmbAgentName 
      Appearance      =   0  'Flat
      Height          =   288
      ItemData        =   "frmTurnOverDis.frx":3F8D
      Left            =   1275
      List            =   "frmTurnOverDis.frx":3F8F
      Style           =   1  'Simple Combo
      TabIndex        =   2
      Top             =   660
      Width           =   5160
   End
   Begin VB.TextBox txtTotal4 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   7290
      TabIndex        =   1
      Top             =   7830
      Width           =   915
   End
   Begin VB.TextBox txtTotal5 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   8190
      TabIndex        =   0
      Top             =   7830
      Width           =   960
   End
   Begin Crystal.CrystalReport cr 
      Left            =   13395
      Top             =   8760
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   2115
      Left            =   15
      TabIndex        =   17
      Top             =   5640
      Width           =   14565
      _cx             =   25691
      _cy             =   3731
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
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
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmTurnOverDis.frx":3F91
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
   Begin MSComCtl2.DTPicker txtDates 
      Height          =   330
      Left            =   2295
      TabIndex        =   18
      Top             =   60
      Width           =   1305
      _ExtentX        =   2307
      _ExtentY        =   593
      _Version        =   393216
      Format          =   182452225
      CurrentDate     =   39795
   End
   Begin MSComCtl2.DTPicker txtMaxDate 
      Height          =   330
      Left            =   11955
      TabIndex        =   30
      Top             =   60
      Visible         =   0   'False
      Width           =   1305
      _ExtentX        =   2307
      _ExtentY        =   593
      _Version        =   393216
      Format          =   182452225
      CurrentDate     =   39795
   End
   Begin VSFlex7Ctl.VSFlexGrid vs1 
      Height          =   4035
      Left            =   15
      TabIndex        =   32
      Top             =   1200
      Width           =   10095
      _cx             =   17806
      _cy             =   7117
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmTurnOverDis.frx":40CA
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
   Begin VB.Label lblsc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "School Name :"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   10260
      TabIndex        =   36
      Top             =   990
      Width           =   1050
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "TOD Dis."
      Height          =   315
      Left            =   12765
      TabIndex        =   29
      Top             =   5265
      Width           =   735
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Press F4 Key To Delete A Grid Item"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   15
      TabIndex        =   23
      Top             =   7920
      Width           =   2955
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Entry No :"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   75
      TabIndex        =   22
      Top             =   180
      Width           =   705
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks :"
      Height          =   285
      Left            =   15
      TabIndex        =   21
      Top             =   8340
      Width           =   1515
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0078CFE9&
      BorderWidth     =   4
      Height          =   1020
      Left            =   6855
      Top             =   8460
      Width           =   6165
   End
   Begin VB.Label Label19 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Party Name :"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   75
      TabIndex        =   20
      Top             =   720
      Width           =   915
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Press F2 For Search"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   1275
      TabIndex        =   19
      Top             =   360
      Width           =   2715
   End
End
Attribute VB_Name = "frmTurnOverDis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dt_from As Date
Dim dt_to As Date
Dim dt_str As String
Dim bb_2 As Boolean
Dim bb1 As Boolean
Dim Edit As Boolean
Dim Add As Boolean
Dim bal_ As Double
Dim CON_next As ADODB.Connection
Dim salecurrent, saleRetcurrent, saleRetnext As String
Dim ss_, saleCol_curr, saleRetCol_curr, saleRetCol_Next As String
Dim salenext As String
Dim nextYrs As String
Function fatchDate(fyear_ As String, type_ As String, inv As Integer, rows_) As String

   '''================================================================
   '''Checking
   '''================================================================
   Select Case fyear_
       
   Case session
    If type_ = "I" Then
        If rs1.State = 1 Then rs1.close
        rs1.Open "select top 1 INVOICEDATE from PartyWiseItemWiseQtySales where INVOICENO='" & inv & "'", con
        If rs1.EOF = False Then
            fatchDate = "PartyWiseItemWiseQtySales"
            current_next = "current"
            
        End If
    ElseIf type_ = "C" Then
        If rs1.State = 1 Then rs1.close
        rs1.Open "select top 1 INVOICEDATE from PartyWiseItemWiseQtySales_return where INVOICENO='" & inv & "'", con
        If rs1.EOF = False Then
           current_next = "current"
           fatchDate = "PartyWiseItemWiseQtySales_return"
        End If
    End If
    
    
    Case session_next
    
      If type_ = "I" Then
        If rs1.State = 1 Then rs1.close
        rs1.Open "select top 1 INVOICEDATE from PartyWiseItemWiseQtySales where INVOICENO='" & inv & "'", CON_next
        If rs1.EOF = False Then
           'fatchDate = rs1!INVOICEDATE
           current_next = "next"
           fatchDate = "PartyWiseItemWiseQtySales"
        End If
    ElseIf type_ = "C" Then
        If rs1.State = 1 Then rs1.close
        rs1.Open "select top 1 INVOICEDATE from PartyWiseItemWiseQtySales_return where INVOICENO='" & inv & "'", CON_next
        If rs1.EOF = False Then
           'fatchDate = rs1!INVOICEDATE
           fatchDate = "PartyWiseItemWiseQtySales_return"
           current_next = "next"
        End If
    End If
    
    
    End Select


End Function
Private Sub ABANDON_Click()

refresh_
max_sp

con.Execute "delete from tmpTurnOver"

cmdEdit_4.Enabled = False
Del.Enabled = False


End Sub

Private Sub cboPayment_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtWhomTobeGiven.SetFocus
End Sub

Private Sub cboSer_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   txtAdPer.SetFocus
End If

End Sub



Private Sub cboSponse_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtReturnAdj.SetFocus
End Sub

Private Sub cboTOD_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cmdView_Click
End Sub
Sub showData()

Dim contr As New ADODB.Connection
Dim fillVs As New ADODB.Recordset
Dim dr, cr As Double

Dim bb As Boolean

Dim op, drcr
Dim rs1 As New ADODB.Recordset
con.Execute "delete from templedger4"

con.Execute "INSERT INTO tempLedger4 (dates,Billtype,bill,des,dr,cr,Party)  SELECT INVOICEDATE,'I',INVOICENO,'Invoice Sales',netamount,BAA,SUBLEDGER from INVOICEA WHERE (((netamount-BAA)>0 or (netamount-BAA)<0) and SUBLEDGER='" & cmbAgentName.text & "')"
con.Execute "INSERT INTO tempLedger4 (dates,Billtype,bill,des,dr,cr,Party) Select INVOICEDATE,'CI',INVOICENO,'Credit Note Item',baa,netamount,SUBLEDGER from CREDITA WHERE (((netamount-BAA)>0 or (netamount-BAA)<0) and SUBLEDGER='" & cmbAgentName.text & "')"
con.Execute "INSERT INTO tempLedger4 (dates,Billtype,bill,des,dr,cr,Party) Select INVOICEDATE,'C/M',INVOICENO,'Cash Memo',netamount,baa,SUBLEDGER from CASHA where (((netamount-BAA)>0 or (netamount-BAA)<0) and SUBLEDGER='" & cmbAgentName.text & "')"
con.Execute "INSERT INTO tempLedger4 (dates,Billtype,bill,des,dr,cr,Party) select dnd,'DN',dnn,'Debit Note',na,'0',PSLD from dnfa where " & stringyear & " and PSLD='" & cmbAgentName.text & "'"
con.Execute "INSERT INTO tempLedger4 (dates,Billtype,bill,des,cr,dr,Party) Select cnd,'CN',cnn,'Credit Note',na,'0',psld from Cnf1a where PSLD='" & cmbAgentName.text & "'"
con.Execute "INSERT INTO tempLedger4 (dates,Billtype,bill,des,dr,cr,Party) Select dates,'J',Recno,Particullar,Dr,CR,PartyName from ReceiveIssueParty where PartyName='" & cmbAgentName.text & "'"

DoEvents
DoEvents
DoEvents
DoEvents

'==================

If fillVs.State = 1 Then fillVs.close
fillVs.Open "select DISTCODE as City,SUBLEDGER as Party,op,drcr from SLEDGER where gledger='SUNDRY DEBTORS'  and subledger='" & cmbAgentName.text & "' order by SUBLEDGER", con, adOpenDynamic, adLockOptimistic
If fillVs.EOF = False Then
  op = IIf(IsNull(fillVs(2)), 0, fillVs(2))
  If RS.State = 1 Then RS.close
  RS.Open "select sum(dr),sum(cr) from tempLedger4  where party='" & fillVs(1) & "'", con, adOpenDynamic, adLockOptimistic
  If Not IsNull(RS(0)) Then
     dr = RS(0)
  End If
  
  
  If Not IsNull(RS(1)) Then
     cr = RS(1)
  End If
  
  If fillVs(3) = "Cr" Then
    op = (-1 * fillVs(2))
  End If
  
  drcr = Round((op + (dr - cr)), 2)
  If Val(drcr) >= 0 Then
     bal_ = Round((op + (dr - cr)), 2)
  End If
 
End If
 
End Sub
Private Sub cboschool_Click()

For J = 1 To vs1.rows - 1
For k1 = 0 To 5
vs1.Cell(flexcpBackColor, J, k1) = vbWhite
DoEvents
Next
Next



If rs1.State = 1 Then rs1.close
rs1.Open "select  billno as invoiceno from tmpTurnOver where scname='" & cboschool.text & "'", con
For I = 1 To rs1.RecordCount
  For J = 1 To vs1.rows - 1
         If rs1!invoiceNo = vs1.TextMatrix(J, 0) Then
           For k1 = 0 To 6
               vs1.Cell(flexcpBackColor, J, k1) = vbGreen
               DoEvents
            Next
         End If
  Next
  rs1.MoveNext
Next


End Sub

Private Sub Check1_Edit_Click()

If Check1_edit.value = 1 Then
   txtMaxDate.Visible = True
   cmdOK_.Visible = True
Else
   txtMaxDate.Visible = False
   cmdOK_.Visible = False
End If

End Sub

Private Sub close_Click()
Unload Me
End Sub
Sub createData()

On Error GoTo err1

Screen.MousePointer = vbHourglass

Dim rs1_ As New ADODB.Recordset
con.Execute "delete FROM tmpTurnOver"

DebitFromRepSale
 
If cmbAgentName = "" Then Exit Sub

If rs1.State = 1 Then rs1.close
rs1.Open "select fromDate,toDate,fyear,Current_Next,DataBase,fromDateSRet,toDateSRet from turnOverDis where NotCreated='y' order by fyear", CCON
While rs1.EOF = False

 Del.Enabled = False
 cmdEdit_4.Enabled = False
 save.Enabled = True


If rs1!current_next = "current" Then

If debitForAgn <> "" Then
    
    con.Execute "insert into tmpTurnOver(EntryNo,dates,INVOICEDATE,SUBLEDGER,NetAmt,RepName,Category,Fyear,billno) " & _
   " SELECT " & txtSponsorshipNo & " as EntryNo ,'" & Format(txtDates, "MM/dd/yyyy") & "' as dates,INVOICEDATE,SUBLEDGER,Amount,repname,'CN' as Category,'" & rs1!fyear & "',cnn FROM PartyCreditRegisternew where (INVOICEDATE>=Convert(smalldatetime,'" & rs1!FromDate & "', 103) and INVOICEDATE<=Convert(smalldatetime,'" & rs1!toDate & "', 103)) and subledger='" & cmbAgentName & "' and cncategory='Adjustment' and todid is null"

End If
    con.Execute "insert into tmpTurnOver(EntryNo,dates,INVOICEDATE,SUBLEDGER,NetAmt,RepName,Category,Fyear,billno,scname) " & _
    " SELECT " & txtSponsorshipNo & " as EntryNo ,'" & Format(txtDates, "MM/dd/yyyy") & "' as dates,INVOICEDATE,SUBLEDGER,NetAmount,AgentName as RepName,'I' as Category,'" & rs1!fyear & "',INVOICENO,scname   FROM INVOICEA where (INVOICEDATE>=Convert(smalldatetime,'" & rs1!FromDate & "', 103) and INVOICEDATE<=Convert(smalldatetime,'" & rs1!toDate & "', 103)) and subledger='" & cmbAgentName & "' and todid is null"
    
    con.Execute "insert into tmpTurnOver(EntryNo,dates,INVOICEDATE,SUBLEDGER,NetAmt,RepName,Category,Fyear,billno,scname) " & _
    " SELECT " & txtSponsorshipNo & " as EntryNo ,'" & Format(txtDates, "MM/dd/yyyy") & "' as dates,INVOICEDATE,SUBLEDGER,NetAmount,AgentName,'CI' as Category,'" & rs1!fyear & "',INVOICENO,scname   FROM creditA where (INVOICEDATE>=Convert(smalldatetime,'" & rs1!fromDateSRet & "', 103) and INVOICEDATE<=Convert(smalldatetime,'" & rs1!toDateSRet & "', 103)) and subledger='" & cmbAgentName & "' and todid is null"
Else
    
    
    '------Fatch Data From Next Session-------------'
    Set CON_next = New ADODB.Connection
    If LCase(server_) = "server" Then
       CON_next.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverName_ & "; DATABASE=" & rs1!Database & "; UID=" & sql_user & "; PWD=" & sql_pass
       CON_next.Open
    Else
       CON_next.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverName_ & "; DATABASE=" & rs1!Database & "; UID=; PWD=;"
       CON_next.Open
    End If

    
    
    
    
    If RS.State = 1 Then RS.close
    RS.Open "SELECT INVOICEDATE,SUBLEDGER,NetAmount,AgentName as RepName,INVOICENO,scname   FROM INVOICEA where (INVOICEDATE>=Convert(smalldatetime,'" & rs1!FromDate & "', 103) and INVOICEDATE<=Convert(smalldatetime,'" & rs1!toDate & "', 103)) and subledger='" & cmbAgentName & "' and todid is null", CON_next
    While RS.EOF = False
       con.Execute "insert into tmpTurnOver(EntryNo,dates,INVOICEDATE,SUBLEDGER,NetAmt,RepName,Category,Fyear,billno,scname) " & _
       " values(" & txtSponsorshipNo & " ,'" & Format(txtDates, "MM/dd/yyyy") & "','" & Format(RS!invoiceDate, "MM/dd/yyyy") & "','" & RS!subledger & "'," & RS!netamount & ",'" & RS!RepName & "','I','" & rs1!fyear & "','" & RS!invoiceNo & "','" & RS!scname & "')"
    RS.MoveNext
    Wend
        
       
    
    If RS.State = 1 Then RS.close
    RS.Open "select INVOICEDATE,SUBLEDGER,Amount,repname,'CreditNote' as Category,'" & rs1!fyear & "',cnn,netamount FROM PartyCreditRegisternew where (INVOICEDATE>=Convert(smalldatetime,'" & rs1!fromDateSRet & "', 103) and INVOICEDATE<=Convert(smalldatetime,'" & rs1!toDateSRet & "', 103)) and subledger='" & cmbAgentName & "' and cncategory='Adjustment' and todid is null", CON_next
    While RS.EOF = False
        con.Execute "insert into tmpTurnOver(EntryNo,dates,INVOICEDATE,SUBLEDGER,NetAmt,RepName,Category,Fyear,billno) " & _
        " values(" & txtSponsorshipNo & " ,'" & Format(txtDates, "MM/dd/yyyy") & "','" & Format(RS!invoiceDate, "MM/dd/yyyy") & "','" & RS!subledger & "'," & RS!amount & ",'" & RS!RepName & "','CN','" & rs1!fyear & "','" & RS!cnn & "')"
        RS.MoveNext
    Wend
    
   
   
    
    If RS.State = 1 Then RS.close
    RS.Open "SELECT INVOICEDATE,SUBLEDGER,NetAmount,AgentName,'Sale' as Category,INVOICENO,scname  FROM creditA where (INVOICEDATE>=Convert(smalldatetime,'" & rs1!fromDateSRet & "', 103) and INVOICEDATE<=Convert(smalldatetime,'" & rs1!toDateSRet & "', 103)) and subledger='" & cmbAgentName & "' and todid is null", CON_next
    While RS.EOF = False
        con.Execute "insert into tmpTurnOver(EntryNo,dates,INVOICEDATE,SUBLEDGER,NetAmt,RepName,Category,Fyear,billno,scname) " & _
        " values(" & txtSponsorshipNo & " ,'" & Format(txtDates, "MM/dd/yyyy") & "','" & Format(RS!invoiceDate, "MM/dd/yyyy") & "','" & RS!subledger & "'," & RS!netamount & ",'" & RS!agentname & "','CI','" & rs1!fyear & "','" & RS!invoiceNo & "','" & RS!scname & "')"
        RS.MoveNext
    Wend
  
    
End If
 
 rs1.MoveNext
Wend


debitForAgn = ""


Screen.MousePointer = vbDefault



dinesh:
Exit Sub
err1:
Screen.MousePointer = vbDefault
MsgBox "" & err.DESCRIPTION


End Sub
Sub createData_SerWise()

On Error GoTo err1

Screen.MousePointer = vbHourglass

Dim rs1_ As New ADODB.Recordset
con.Execute "delete FROM tmpTurnOver"

DebitFromRepSale
 
If cmbAgentName = "" Then Exit Sub

If rs1.State = 1 Then rs1.close
rs1.Open "select fromDate,toDate,fyear,Current_Next,DataBase,fromDateSRet,toDateSRet from turnOverDis where NotCreated='y' order by fyear", CCON
While rs1.EOF = False

 Del.Enabled = False
 cmdEdit_4.Enabled = False
 save.Enabled = True


If rs1!current_next = "current" Then

If debitForAgn <> "" Then
    
    con.Execute "insert into tmpTurnOver(EntryNo,dates,INVOICEDATE,SUBLEDGER,NetAmt,RepName,Category,Fyear,billno) " & _
   " SELECT " & txtSponsorshipNo & " as EntryNo ,'" & Format(txtDates, "MM/dd/yyyy") & "' as dates,INVOICEDATE,SUBLEDGER,Amount,repname,'CN' as Category,'" & rs1!fyear & "',cnn FROM PartyCreditRegisternew where (INVOICEDATE>=Convert(smalldatetime,'" & rs1!FromDate & "', 103) and INVOICEDATE<=Convert(smalldatetime,'" & rs1!toDate & "', 103)) and subledger='" & cmbAgentName & "' and cncategory='Adjustment' and todid is null"

End If

    con.Execute "insert into tmpTurnOver(EntryNo,dates,INVOICEDATE,SUBLEDGER,NetAmt,RepName,Category,Fyear,billno,scname,sername) " & _
    " SELECT  " & txtSponsorshipNo & " as EntryNo ,'" & Format(txtDates, "MM/dd/yyyy") & "' as dates,INVOICEDATE,SUBLEDGER,sum(NetAmount),AgentName as RepName,'I' as Category,'" & rs1!fyear & "',INVOICENO,scname,sername   FROM invoiceBQry where (INVOICEDATE>=Convert(smalldatetime,'" & rs1!FromDate & "', 103) and INVOICEDATE<=Convert(smalldatetime,'" & rs1!toDate & "', 103)) and subledger='" & cmbAgentName & "' and todid is null group by INVOICEDATE,SUBLEDGER,AgentName,INVOICENO,scname,sername"
    
    con.Execute "insert into tmpTurnOver(EntryNo,dates,INVOICEDATE,SUBLEDGER,NetAmt,RepName,Category,Fyear,billno,scname,sername) " & _
    " SELECT " & txtSponsorshipNo & " as EntryNo ,'" & Format(txtDates, "MM/dd/yyyy") & "' as dates,INVOICEDATE,SUBLEDGER,sum(NetAmount),AgentName,'CI' as Category,'" & rs1!fyear & "',INVOICENO,scname,sername   FROM CreditbQry where (INVOICEDATE>=Convert(smalldatetime,'" & rs1!fromDateSRet & "', 103) and INVOICEDATE<=Convert(smalldatetime,'" & rs1!toDateSRet & "', 103)) and subledger='" & cmbAgentName & "' and todid is null group by INVOICEDATE,SUBLEDGER,AgentName,INVOICENO,scname,sername"
    
    '    con.Execute "insert into tmpTurnOver(EntryNo,dates,INVOICEDATE,SUBLEDGER,NetAmt,RepName,Category,Fyear,billno,scname) " & _
    '" SELECT " & txtSponsorshipNo & " as EntryNo ,'" & Format(txtDates, "MM/dd/yyyy") & "' as dates,INVOICEDATE,SUBLEDGER,NetAmount,AgentName,'CI' as Category,'" & rs1!fyear & "',INVOICENO,scname   FROM creditA where (INVOICEDATE>=Convert(smalldatetime,'" & rs1!fromDateSRet & "', 103) and INVOICEDATE<=Convert(smalldatetime,'" & rs1!toDateSRet & "', 103)) and subledger='" & cmbAgentName & "' and todid is null"

Else
    
    
    '------Fatch Data From Next Session-------------'
    Set CON_next = New ADODB.Connection
    If LCase(server_) = "server" Then
       CON_next.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverName_ & "; DATABASE=" & rs1!Database & "; UID=" & sql_user & "; PWD=" & sql_pass
       CON_next.Open
    Else
       CON_next.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverName_ & "; DATABASE=" & rs1!Database & "; UID=; PWD=;"
       CON_next.Open
    End If

    
    
    
    
    
    If RS.State = 1 Then RS.close
    RS.Open "SELECT INVOICEDATE,SUBLEDGER,sum(NetAmount),AgentName as RepName,INVOICENO,scname,sername   FROM invoiceBQry where (INVOICEDATE>=Convert(smalldatetime,'" & rs1!FromDate & "', 103) and INVOICEDATE<=Convert(smalldatetime,'" & rs1!toDate & "', 103)) and subledger='" & cmbAgentName & "' and todid is null group by INVOICEDATE,SUBLEDGER,AgentName,INVOICENO,scname,sername", CON_next
    While RS.EOF = False
       con.Execute "insert into tmpTurnOver(EntryNo,dates,INVOICEDATE,SUBLEDGER,NetAmt,RepName,Category,Fyear,billno,scname,sername) " & _
       " values(" & txtSponsorshipNo & " ,'" & Format(txtDates, "MM/dd/yyyy") & "','" & Format(RS!invoiceDate, "MM/dd/yyyy") & "','" & RS!subledger & "'," & RS(2) & ",'" & RS!RepName & "','I','" & rs1!fyear & "','" & RS!invoiceNo & "','" & RS!scname & "','" & RS!sername & "')"
    RS.MoveNext
    Wend
        
       
    
    If RS.State = 1 Then RS.close
    RS.Open "select INVOICEDATE,SUBLEDGER,Amount,repname,'CreditNote' as Category,'" & rs1!fyear & "',cnn,netamount FROM PartyCreditRegisternew where (INVOICEDATE>=Convert(smalldatetime,'" & rs1!fromDateSRet & "', 103) and INVOICEDATE<=Convert(smalldatetime,'" & rs1!toDateSRet & "', 103)) and subledger='" & cmbAgentName & "' and cncategory='Adjustment' and todid is null", CON_next
    While RS.EOF = False
        con.Execute "insert into tmpTurnOver(EntryNo,dates,INVOICEDATE,SUBLEDGER,NetAmt,RepName,Category,Fyear,billno) " & _
        " values(" & txtSponsorshipNo & " ,'" & Format(txtDates, "MM/dd/yyyy") & "','" & Format(RS!invoiceDate, "MM/dd/yyyy") & "','" & RS!subledger & "'," & RS!amount & ",'" & RS!RepName & "','CN','" & rs1!fyear & "','" & RS!cnn & "')"
        RS.MoveNext
    Wend
    
   
    
    If RS.State = 1 Then RS.close
    RS.Open "SELECT INVOICEDATE,SUBLEDGER,sum(NetAmount),AgentName,'Sale' as Category,INVOICENO,scname,sername  FROM CreditbQry where (INVOICEDATE>=Convert(smalldatetime,'" & rs1!fromDateSRet & "', 103) and INVOICEDATE<=Convert(smalldatetime,'" & rs1!toDateSRet & "', 103)) and subledger='" & cmbAgentName & "' and todid is null group by INVOICEDATE,SUBLEDGER,AgentName,INVOICENO,scname,sername", CON_next
    While RS.EOF = False
        con.Execute "insert into tmpTurnOver(EntryNo,dates,INVOICEDATE,SUBLEDGER,NetAmt,RepName,Category,Fyear,billno,scname,sername) " & _
        " values(" & txtSponsorshipNo & " ,'" & Format(txtDates, "MM/dd/yyyy") & "','" & Format(RS!invoiceDate, "MM/dd/yyyy") & "','" & RS!subledger & "'," & RS(2) & ",'" & RS!agentname & "','CI','" & rs1!fyear & "','" & RS!invoiceNo & "','" & RS!scname & "','" & RS!sername & "')"
        RS.MoveNext
    Wend
    
'    If RS.State = 1 Then RS.close
'    RS.Open "SELECT INVOICEDATE,SUBLEDGER,NetAmount,AgentName,'Sale' as Category,INVOICENO,scname  FROM creditA where (INVOICEDATE>=Convert(smalldatetime,'" & rs1!fromDateSRet & "', 103) and INVOICEDATE<=Convert(smalldatetime,'" & rs1!toDateSRet & "', 103)) and subledger='" & cmbAgentName & "' and todid is null", CON_next
'    While RS.EOF = False
'        con.Execute "insert into tmpTurnOver(EntryNo,dates,INVOICEDATE,SUBLEDGER,NetAmt,RepName,Category,Fyear,billno,scname) " & _
'        " values(" & txtSponsorshipNo & " ,'" & Format(txtDates, "MM/dd/yyyy") & "','" & Format(RS!invoiceDate, "MM/dd/yyyy") & "','" & RS!SUBLEDGER & "'," & RS!netamount & ",'" & RS!agentname & "','CI','" & rs1!fyear & "','" & RS!invoiceNo & "','" & RS!scname & "')"
'        RS.MoveNext
'    Wend

    
    
    
  
    
End If
 
 rs1.MoveNext
Wend


debitForAgn = ""


Screen.MousePointer = vbDefault



dinesh:
Exit Sub
err1:
Screen.MousePointer = vbDefault
MsgBox "" & err.DESCRIPTION


End Sub

Private Sub cmbAgentName_GotFocus()

If PopUpValue1 <> "" Then
   cmbAgentName = PopUpValue3
   
   If Right(session, 2) >= 20 Then
      createData_SerWise
   Else
      createData
   End If
   
   con.Execute "update tmpTurnOver set tod='Yes'"

   PopUpValue1 = ""
   PopUpValue2 = ""
   PopUpValue3 = ""
End If

End Sub

Private Sub cmbAgentName_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
    cmdView_Click
End If

If KeyCode = 113 Then
    searchType = "party"
    lblCr = "dr"
    value = "select distinct(Party),Code,subledger from SLEDGER where gledger='SUNDRY DEBTORS' and " & stringyear & "  order by party"
    popuplist_client value, CCON
End If

End Sub
Sub max_sp()
Set RS = New ADODB.Recordset
RS.Open "select max(entryNo) from TurnOver", con
If Not IsNull(RS(0)) Then
   txtSponsorshipNo = RS(0) + 1
Else
   txtSponsorshipNo = 1
End If

End Sub


Private Sub cmdEdit_4_Click()


Del.Enabled = True
save.Enabled = True
Edit = True


vs.Enabled = True
End Sub
Private Sub cmdok_Click()

Screen.MousePointer = vbHourglass

txtAdjAmt = 0
txtDiffNet = 0
For I = 1 To vs.rows - 1
If vs.TextMatrix(I, 0) <> "" Then
    
If RS.State = 1 Then RS.close
RS.Open "select bookcode from books where sername='" & cboser & "' order by bookcode", con
While RS.EOF = False
If vs.TextMatrix(I, 4) = RS(0) Then
    vs.TextMatrix(I, 11) = Val(txtAdPer)
    vs.TextMatrix(I, 12) = Round(Val(vs.TextMatrix(I, 8)) - (Val(vs.TextMatrix(I, 8)) * Val(txtAdPer) / 100), 0)
    vs.TextMatrix(I, 13) = (Val(vs.TextMatrix(I, 10)) - Val(vs.TextMatrix(I, 12)))
End If
RS.MoveNext
Wend
End If
Next


txtAdjAmt = 0
txtDiffNet = 0

For k1 = 1 To vs.rows - 1
If vs.TextMatrix(k1, 0) <> "" Then
If vs.TextMatrix(k1, 1) = "I" Then
   txtAdjAmt = Val(txtAdjAmt) + Val(vs.TextMatrix(k1, 12))
   txtDiffNet = Val(txtDiffNet) + Val(vs.TextMatrix(k1, 13))
Else
   txtAdjAmt = Val(txtAdjAmt) - Val(vs.TextMatrix(k1, 12))
   txtDiffNet = Val(txtDiffNet) - Val(vs.TextMatrix(k1, 13))
End If
End If
Next








Screen.MousePointer = vbDefault
           
End Sub

Private Sub cmdOK__Click()

  Set RS = New ADODB.Recordset
 RS.Open "select * from turnOverDis where fyear='" & session_next & "'", CCON, adOpenDynamic, adLockOptimistic
 If RS.EOF = False Then
    RS!toDate = txtMaxDate.value
    RS!toDateSRet = txtMaxDate.value
    RS.update
    MsgBox "Date have Changed...", vbInformation
 End If
 
End Sub
Private Sub cmdPrint_7_Click()

DSNNew


cr.Reset
cr.ReportFileName = rptPath & "/TurnOverDis.rpt"
cr.ReplaceSelectionFormula "{TurnOver.entryNo}=" & txtSponsorshipNo & ""
cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
cr.Formulas(0) = "sale_current='" & vs.TextMatrix(0, 1) & "'"
cr.Formulas(1) = "saleRet_current='" & vs.TextMatrix(0, 2) & "'"
cr.Formulas(2) = "sale_next='" & vs.TextMatrix(0, 3) & "'"
cr.Formulas(3) = "saleRet_next='" & vs.TextMatrix(0, 4) & "'"


cr.WindowShowPrintSetupBtn = True
cr.WindowShowPrintBtn = True
cr.WindowShowExportBtn = True
cr.WindowState = crptMaximized
cr.Action = 1


End Sub






Private Sub Combo1_Change()

End Sub
Private Sub cmdView_Click()
'Fill Grid=================================================================================

DoEvents
DoEvents
DoEvents
DoEvents
DoEvents


'------------------------------
showData
'------------------------------


 If Right(session, 2) >= 20 Then
    vs1.Cols = 7
 Else
    vs1.Cols = 6
 End If

Dim sale_current, saleRet_current, saleRet_next, salenext_, aduj As Double
Dim rs_ As New ADODB.Recordset
Dim k9 As Integer



If RS.State = 1 Then RS.close
RS.Open "SELECT fyear,Current_Next FROM turnOverDis order by fyear", CCON, adOpenDynamic, adLockReadOnly
If RS.EOF = False Then
  salecurrent = "Sale:" & RS!fyear
  saleRetcurrent = "SaleRet:" & RS!fyear
  RS.MoveNext
  saleRetnext = "SaleRet:" & RS!fyear
  salenext = "Sale:" & RS!fyear
End If



If RS.State = 1 Then RS.close
RS.Open "select top 1 EntryNo from TurnOver where subledger='" & cmbAgentName & "'", con
If RS.EOF = False Then
   MsgBox "TOD is already created for This Party", vbCritical
   Exit Sub
End If

'''con.Execute "UPDATE a SET a.tod = b.tod  FROM tmpTurnOver AS a " & _
'''" INNER JOIN AppForm AS b ON (b.School_PartyName = a.scname or b.PName = a.scname)  where (b.id = '" & Mid(cmbAgentName, 1, 5) & "' or b.code = '" & Mid(cmbAgentName, 1, 5) & "') "

'=======================================================
'=======================================================

   cboschool.Clear
   
   If RS.State = 1 Then RS.close
   RS.Open "select  distinct scname,tod from tmpTurnOver where scname is not null", con
   While RS.EOF = False
   
   If RS!tod = "Yes" Then
   cboschool.AddItem RS(0)
   End If
   
   RS.MoveNext
   Wend
   '==============================



vs.Clear
vs.rows = 1
k9 = 1
'If cboTOD.Text = "" Then Exit Sub




If rs_.State = 1 Then rs_.close
rs_.Open "SELECT fyear,Current_Next FROM turnOverDis", CCON, adOpenDynamic, adLockReadOnly


Dim k1 As Integer
k1 = 1

vs1.Cols = 7
vs1.rows = 1

If rs1.State = 1 Then rs1.close
rs1.Open "SELECT RepName FROM tmpTurnOver where tod='Yes' group by RepName", con
While rs1.EOF = False
    
    vs.rows = vs.rows + 1
    sale_current = 0
    saleRet_current = 0
    saleRet_next = 0
    aduj = 0
    salenext_ = 0
    
    If RS.State = 1 Then RS.close
    RS.Open "SELECT SUBLEDGER,NetAmt,Category,Fyear,BillNO,Dates,RepName,invoicedate,sername  FROM tmpTurnOver where RepName='" & rs1!RepName & "' and tod='Yes' order by category desc,convert(int,billno) asc", con
    
    While RS.EOF = False
    
     vs1.rows = vs1.rows + 1
     vs1.TextMatrix(k9, 0) = RS!billno
     vs1.TextMatrix(k9, 1) = RS!invoiceDate
     vs1.TextMatrix(k9, 2) = RS!RepName
     vs1.TextMatrix(k9, 3) = RS!category
     vs1.TextMatrix(k9, 4) = RS!fyear
     vs1.TextMatrix(k9, 5) = RS!sername & ""

     vs1.TextMatrix(k9, 6) = RS!netamt
     k9 = k9 + 1
     
     rs_.MoveFirst
     rs_.Find "fyear='" & RS!fyear & "'"
     If rs_.EOF = False Then
           
        
        
        If rs_!current_next = "current" Then
           
           saleRetCol_curr = "Sale Ret:" & rs_!fyear
           saleCol_curr = "Sale:" & rs_!fyear

           
           
           If RS!category = "I" Then
              sale_current = sale_current + RS!netamt
              
           ElseIf RS!category = "CI" Then
              saleRet_current = saleRet_current + RS!netamt
           ElseIf RS!category = "CN" Then
              aduj = aduj + RS!netamt
           End If
           
        Else
          
          If RS!category = "I" Then
              salenext_ = salenext_ + RS!netamt
          ElseIf RS!category = "CI" Then
             saleRet_next = saleRet_next + RS!netamt
          ElseIf RS!category = "CN" Then   'ElseIf RS!category = "CreditNote" Then
              aduj = aduj + RS!netamt
          End If
          
          saleRetCol_Next = "Sale Ret:" & rs_!fyear
          
        End If
           
           
     End If
    
    
    RS.MoveNext
    Wend
    
    vs.TextMatrix(k1, 0) = rs1!RepName
    vs.TextMatrix(k1, 1) = Round(sale_current, 0)
    vs.TextMatrix(k1, 2) = Round(saleRet_current, 0)
    vs.TextMatrix(k1, 3) = Round(salenext_, 0)
    
    vs.TextMatrix(k1, 4) = Round(saleRet_next, 0)
    
    vs.TextMatrix(k1, 5) = ((Val(vs.TextMatrix(k1, 1)) + Val(vs.TextMatrix(k1, 3))) - (Val(vs.TextMatrix(k1, 2)) + Val(vs.TextMatrix(k1, 4))))
    
    
    vs.TextMatrix(k1, 6) = Round(aduj, 0)
   vs.TextMatrix(k1, 7) = Round(bal_, 0)
    
     
     '=======Changing is working------------------------------------------------
     
    If Val(vs.TextMatrix(k1, 5)) <> 0 Then
       vs.TextMatrix(k1, 8) = Round((Val(vs.TextMatrix(k1, 5)) - Val(vs.TextMatrix(k1, 6))) - Val(vs.TextMatrix(k1, 7)), 0)
    End If
    
    
    '=======End Code    -----------------------------------------------------------
    
    
    k1 = k1 + 1
    
 rs1.MoveNext
Wend


 
 
 
' If Right(session, 2) >= 20 Then
 
 vs1.Cols = 7
 
 vs1.TextMatrix(0, 0) = "Bill No"
 vs1.TextMatrix(0, 1) = "Date"
 vs1.TextMatrix(0, 2) = "Rep.Name"
 vs1.TextMatrix(0, 3) = "VType"
 vs1.TextMatrix(0, 4) = "Session"
 vs1.TextMatrix(0, 5) = "Series"
 vs1.TextMatrix(0, 6) = "Amount"
 
 vs1.ColWidth(0) = 900
 vs1.ColWidth(1) = 1000
 vs1.ColWidth(2) = 2500
 vs1.ColWidth(3) = 1400
 vs1.ColWidth(4) = 1400
 vs1.ColWidth(5) = 1200
 vs1.ColWidth(6) = 1200
 
 
' Else
'
' vs1.Cols = 6
'
' vs1.TextMatrix(0, 0) = "Bill No"
' vs1.TextMatrix(0, 1) = "Date"
' vs1.TextMatrix(0, 2) = "Rep.Name"
' vs1.TextMatrix(0, 3) = "VType"
' vs1.TextMatrix(0, 4) = "Session"
' vs1.TextMatrix(0, 5) = "Amount"
'
' vs1.ColWidth(0) = 900
' vs1.ColWidth(1) = 1000
' vs1.ColWidth(2) = 3000
' vs1.ColWidth(3) = 1500
' vs1.ColWidth(4) = 1500
' vs1.ColWidth(5) = 1200
'
'
' End If
'
 vs.Cols = 13


 vs.TextMatrix(0, 0) = "Party Name"
 vs.TextMatrix(0, 1) = "" & salecurrent
 vs.TextMatrix(0, 2) = "" & saleRetcurrent
 vs.TextMatrix(0, 3) = "" & salenext
 vs.TextMatrix(0, 4) = "" & saleRetnext
 vs.TextMatrix(0, 5) = "Total Sale"
 vs.TextMatrix(0, 6) = "Credit Amt"
 vs.TextMatrix(0, 7) = "OS Amt"
 vs.TextMatrix(0, 8) = "Final Amt"
 vs.TextMatrix(0, 9) = "TOD(%)"
 vs.TextMatrix(0, 10) = "TOD Amt"
 vs.TextMatrix(0, 11) = "Adj.Amt"
 vs.TextMatrix(0, 12) = "Net Sale"

 
 
 
 

Total




End Sub
Sub searchData()

On Error Resume Next

Dim k1 As Integer


vs.Clear
vs.Cols = 13
vs.rows = 1

'If cboTOD.Text = "" Then Exit Sub

If RS.State = 1 Then RS.close
RS.Open "SELECT fyear,Current_Next FROM turnOverDis order by fyear", CCON, adOpenDynamic, adLockReadOnly
If RS.EOF = False Then
  salecurrent = "Sale:" & RS!fyear
  saleRetcurrent = "SaleRet:" & RS!fyear
  
  RS.MoveNext
  saleRetnext = "SaleRet:" & RS!fyear
  salenext = "Sale:" & RS!fyear
End If


k1 = 1


If rs1.State = 1 Then rs1.close
rs1.Open "SELECT * FROM TurnOver where entryno=" & txtSponsorshipNo & " order by SUBLEDGER", con
While rs1.EOF = False
    
    txtDates.value = rs1!dates
    txtRemarks.text = rs1!remarks & ""
    cmbAgentName.text = rs1!RepName & ""
    
    save.Enabled = False
    cmdEdit_4.Enabled = True
    Del.Enabled = False
    
    vs.rows = vs.rows + 1
    
    vs.TextMatrix(k1, 0) = rs1!subledger
    vs.TextMatrix(k1, 1) = rs1!sale_current
    vs.TextMatrix(k1, 2) = rs1!saleRet_current
    
    vs.TextMatrix(k1, 3) = rs1!sale_next & ""
    'new
    
    vs.TextMatrix(k1, 4) = rs1!saleRet_next
    vs.TextMatrix(k1, 5) = rs1!netsale
    vs.TextMatrix(k1, 6) = rs1!AdjustAmt
    
    vs.TextMatrix(k1, 7) = rs1!OSAmt & ""
    
    vs.TextMatrix(k1, 8) = rs1!finalAmt
    vs.TextMatrix(k1, 9) = rs1!TODDis & ""
    vs.TextMatrix(k1, 10) = rs1!TODAmt & ""
    vs.TextMatrix(k1, 11) = rs1!AdjAmt & ""
    vs.TextMatrix(k1, 12) = rs1!FinalAmount & ""
    
    k1 = k1 + 1
    
 rs1.MoveNext
Wend


'==========================================================
vs1.rows = 1
vs1.Cols = 7

If RS.State = 1 Then RS.close
RS.Open "SELECT * FROM TurnBillDet where entryno=" & txtSponsorshipNo & "", con
For k9 = 1 To RS.RecordCount
    vs1.rows = vs1.rows + 1
    vs1.TextMatrix(k9, 0) = RS!billno
    vs1.TextMatrix(k9, 1) = RS!dates
    vs1.TextMatrix(k9, 2) = RS!RepName
    vs1.TextMatrix(k9, 3) = RS!category
    vs1.TextMatrix(k9, 4) = RS!fyear
    vs1.TextMatrix(k9, 5) = RS!sername & ""
    vs1.TextMatrix(k9, 6) = RS!netamt
    RS.MoveNext
Next
'==========================================================



vs1.TextMatrix(0, 0) = "Bill No"
vs1.TextMatrix(0, 1) = "Date"
vs1.TextMatrix(0, 2) = "Rep.Name"
vs1.TextMatrix(0, 3) = "VType"
vs1.TextMatrix(0, 4) = "Session"
vs1.TextMatrix(0, 5) = "SerName"
vs1.TextMatrix(0, 6) = "Amount"
 
vs1.ColWidth(0) = 900
vs1.ColWidth(1) = 1000
vs1.ColWidth(2) = 2800
vs1.ColWidth(3) = 1100
vs1.ColWidth(4) = 1100
vs1.ColWidth(5) = 1400
vs1.ColWidth(6) = 1200






vs.Cols = 13

 vs.TextMatrix(0, 0) = "Party Name"
 vs.TextMatrix(0, 1) = "" & salecurrent
 vs.TextMatrix(0, 2) = "" & saleRetcurrent
 
 vs.TextMatrix(0, 3) = "" & salenext
 vs.TextMatrix(0, 4) = "" & saleRetnext
 
 vs.TextMatrix(0, 5) = "Total Sale"
 vs.TextMatrix(0, 6) = "Credit Amt"
 
 vs.TextMatrix(0, 7) = "OS. Amt"
 
 vs.TextMatrix(0, 8) = "Final Amt"
 vs.TextMatrix(0, 9) = "TOD(%)"
 vs.TextMatrix(0, 10) = "TOD Amt"
 vs.TextMatrix(0, 11) = "Adj.Amt"
 vs.TextMatrix(0, 12) = "Net Sale"

Total

End Sub
Sub SearchDataNew()

On Error Resume Next

Dim k1 As Integer

'=======================================================================================

Dim conadj As ADODB.Connection
Dim rs_adj As ADODB.Recordset

Set conadj = New ADODB.Connection
Set rs_adj = New ADODB.Recordset
rs_adj.Open "select LastDatabase from data", CCON
If rs_adj.EOF = False Then
    Set con_don = New ADODB.Connection
    If LCase(server_) = "server" Then
       conadj.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverName_ & "; DATABASE=" & rs_adj!LastDatabase & "; UID=" & sql_user & "; PWD=" & sql_pass
       conadj.Open
    Else
       conadj.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=HP\SQL2008; DATABASE=" & rs_adj!LastDatabase & "; UID=; PWD=;"
       conadj.Open
    End If
End If
'=================================================================================

Dim rs_1 As New ADODB.Recordset
Dim lastData_ As Boolean

lastData_ = False
txtSponsorshipNo.text = inviceNo

str11 = "SELECT top 1 BillNO FROM  TurnBillDet where fyear='" & session & "' and BillNO='" & inv_ledger & "'"
ss1_ = ""
If rs_1.State = 1 Then rs_1.close
rs_1.Open str11, conadj
If rs_1.EOF = False Then
   lastData_ = True
   pname_ = ""
End If





'========================================================================================

vs.Clear
vs.Cols = 13
vs.rows = 1



If RS.State = 1 Then RS.close
RS.Open "SELECT fyear,Current_Next FROM turnOverDis order by fyear", CCON, adOpenDynamic, adLockReadOnly
If RS.EOF = False Then
  salecurrent = "Sale:" & RS!fyear
  saleRetcurrent = "SaleRet:" & RS!fyear
  
  RS.MoveNext
  saleRetnext = "SaleRet:" & RS!fyear
  salenext = "Sale:" & RS!fyear
End If


k1 = 1


If rs1.State = 1 Then rs1.close
If lastData_ = True Then
   rs1.Open "SELECT * FROM TurnOver where entryno=" & txtSponsorshipNo & " order by SUBLEDGER", conadj
 Else
   rs1.Open "SELECT * FROM TurnOver where entryno=" & txtSponsorshipNo & " order by SUBLEDGER", con
End If
While rs1.EOF = False
    
    txtDates.value = rs1!dates
    txtRemarks.text = rs1!remarks & ""
    cmbAgentName.text = rs1!RepName & ""
    
    save.Enabled = False
    cmdEdit_4.Enabled = True
    Del.Enabled = False
    
    vs.rows = vs.rows + 1
    
    vs.TextMatrix(k1, 0) = rs1!subledger
    vs.TextMatrix(k1, 1) = rs1!sale_current
    vs.TextMatrix(k1, 2) = rs1!saleRet_current
    
    vs.TextMatrix(k1, 3) = rs1!sale_next & ""
    'new
    
    vs.TextMatrix(k1, 4) = rs1!saleRet_next
    vs.TextMatrix(k1, 5) = rs1!netsale
    vs.TextMatrix(k1, 6) = rs1!AdjustAmt
    
    vs.TextMatrix(k1, 7) = rs1!OSAmt & ""
    
    vs.TextMatrix(k1, 8) = rs1!finalAmt
    vs.TextMatrix(k1, 9) = rs1!TODDis & ""
    vs.TextMatrix(k1, 10) = rs1!TODAmt & ""
    vs.TextMatrix(k1, 11) = rs1!AdjAmt & ""
    vs.TextMatrix(k1, 12) = rs1!FinalAmount & ""
    
    k1 = k1 + 1
    
 rs1.MoveNext
Wend


'==========================================================
vs1.rows = 1

If RS.State = 1 Then RS.close
If lastData_ = True Then
RS.Open "SELECT * FROM TurnBillDet where entryno=" & txtSponsorshipNo & "", conadj
Else
RS.Open "SELECT * FROM TurnBillDet where entryno=" & txtSponsorshipNo & "", con
End If
For k9 = 1 To RS.RecordCount
    vs1.rows = vs1.rows + 1
    vs1.TextMatrix(k9, 0) = RS!billno
    vs1.TextMatrix(k9, 1) = RS!dates
    vs1.TextMatrix(k9, 2) = RS!RepName
    vs1.TextMatrix(k9, 3) = RS!category
    vs1.TextMatrix(k9, 4) = RS!fyear
    vs1.TextMatrix(k9, 5) = RS!netamt
    RS.MoveNext
Next
'==========================================================



vs1.TextMatrix(0, 0) = "Bill No"
vs1.TextMatrix(0, 1) = "Date"
vs1.TextMatrix(0, 2) = "Rep.Name"
vs1.TextMatrix(0, 3) = "VType"
vs1.TextMatrix(0, 4) = "Session"
vs1.TextMatrix(0, 5) = "Amount"
 
vs1.ColWidth(0) = 900
vs1.ColWidth(1) = 1000
vs1.ColWidth(2) = 3000
vs1.ColWidth(3) = 1500
vs1.ColWidth(4) = 1500
vs1.ColWidth(5) = 1200






vs.Cols = 13

 vs.TextMatrix(0, 0) = "Party Name"
 vs.TextMatrix(0, 1) = "" & salecurrent
 vs.TextMatrix(0, 2) = "" & saleRetcurrent
 
 vs.TextMatrix(0, 3) = "" & salenext
 vs.TextMatrix(0, 4) = "" & saleRetnext
 
 vs.TextMatrix(0, 5) = "Total Sale"
 vs.TextMatrix(0, 6) = "Credit Amt"
 
 vs.TextMatrix(0, 7) = "OS. Amt"
 
 vs.TextMatrix(0, 8) = "Final Amt"
 vs.TextMatrix(0, 9) = "TOD(%)"
 vs.TextMatrix(0, 10) = "TOD Amt"
 vs.TextMatrix(0, 11) = "Adj.Amt"
 vs.TextMatrix(0, 12) = "Net Sale"

Total

End Sub

Sub Total()
 
vs.Cols = 13

 
 
 vs.ColWidth(0) = 2600
 vs.ColWidth(1) = 1100
 vs.ColWidth(2) = 1300
 vs.ColWidth(3) = 1300
 vs.ColWidth(4) = 1000
 vs.ColWidth(5) = 900
 vs.ColWidth(6) = 1000
 vs.ColWidth(7) = 800    ''' add
 
 vs.ColWidth(8) = 800
 
 vs.ColWidth(9) = 800
 vs.ColWidth(10) = 800
 vs.ColWidth(11) = 1000
 
  
 
 
 
 txtTotal1 = 0
 txtTotal2 = 0
 txtTotal3 = 0
 txtTotal4 = 0
 txtTotal5 = 0
 txtTotal6 = 0
 txtTotal7 = 0
 txtTotal8 = 0
 txtTotal9 = 0
 txtTotal10 = 0
 txtOSAmt = 0
 
 
 For J = 1 To vs.rows - 1
 If vs.TextMatrix(J, 0) <> "" Then
    txtTotal1 = Val(txtTotal1) + Val(vs.TextMatrix(J, 1))
    txtTotal2 = Val(txtTotal2) + Val(vs.TextMatrix(J, 2))
    
    txtTotal10 = Val(txtTotal10) + Val(vs.TextMatrix(J, 3))
    
    txtTotal3 = Val(txtTotal3) + Val(vs.TextMatrix(J, 4))
    
    txtTotal4 = Val(txtTotal4) + Val(vs.TextMatrix(J, 5))
    txtTotal5 = Val(txtTotal5) + Val(vs.TextMatrix(J, 6))
    
    txtOSAmt = Val(txtOSAmt) + Val(vs.TextMatrix(J, 7))
    
    txtTotal6 = Val(txtTotal6) + Val(vs.TextMatrix(J, 8))
    txtTotal7 = Val(txtTotal7) + Val(vs.TextMatrix(J, 9))
    txtTotal8 = Val(txtTotal8) + Val(vs.TextMatrix(J, 10))
    txtTotal9 = Val(txtTotal9) + Val(vs.TextMatrix(J, 12))
    
    
 End If
 Next

End Sub

Private Sub Del_Click()
If MsgBox("want to delete ?", vbQuestion + vbYesNo) = vbYes Then
   
   con.Execute "delete from TurnOver where entryno=" & txtSponsorshipNo & ""
   con.Execute "delete from TurnBillDet where entryno=" & txtSponsorshipNo & ""
   
   
   
   
'========================================================================================================
'========================================================================================================
For J = 1 To vs1.rows - 1

If vs1.TextMatrix(J, 0) <> "" Then
    fyear_ = vs1.TextMatrix(J, 4)
    If rs1.State = 1 Then rs1.close
    rs1.Open "select fromDate,toDate,fyear,Current_Next,DataBase,fromDateSRet,toDateSRet from turnOverDis where fyear='" & fyear_ & "'", CCON
'    While rs1.EOF = False
    If rs1.EOF = False Then
        If rs1!current_next = "current" Then

           If vs1.TextMatrix(J, 3) = "I" Then
              con.Execute "update INVOICEA set todid=null,toddate=null where invoiceno='" & vs1.TextMatrix(J, 0) & "'"
           ElseIf vs1.TextMatrix(J, 3) = "CI" Then
              con.Execute "update creditA set todid=null,toddate=null where invoiceno='" & vs1.TextMatrix(J, 0) & "'"
           ElseIf vs1.TextMatrix(J, 3) = "CN" Then
              con.Execute "update CNF1A set todid=null,toddate=null where cnn='" & vs1.TextMatrix(J, 0) & "'"
           End If


        Else

            '------Fatch Data From Next Session-------------'
            Set CON_next = New ADODB.Connection
            If LCase(server_) = "server" Then
               CON_next.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverName_ & "; DATABASE=" & rs1!Database & "; UID=" & sql_user & "; PWD=" & sql_pass
               CON_next.Open
            Else
               CON_next.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverName_ & "; DATABASE=" & rs1!Database & "; UID=; PWD=;"
               CON_next.Open
            End If


           If vs1.TextMatrix(J, 3) = "I" Then
              CON_next.Execute "update INVOICEA set todid=null,toddate=null where invoiceno='" & vs1.TextMatrix(J, 0) & "'"
           ElseIf vs1.TextMatrix(J, 3) = "CI" Then
              CON_next.Execute "update creditA set todid=null,toddate=null where invoiceno='" & vs1.TextMatrix(J, 0) & "'"
           ElseIf vs1.TextMatrix(J, 3) = "CN" Then
              CON_next.Execute "update CNF1A set todid=null,toddate=null where cnn='" & vs1.TextMatrix(J, 0) & "'"
           End If

        End If

    'rs1.MoveNext
    End If

End If

Next


'===========================================================================================
   
   
   
   
   refresh_
End If
End Sub
Sub refresh_()
vs.Enabled = True
Add = True
Edit = False
vs.Clear
cboschool.Clear

txtOSAmt = 0
txtNetBal = ""
txtWhomToBeGivenMob = ""
txtSponsorshipNo.text = ""
txtDates.value = Format(Date, "dd/MM/yyyy")
cmbAgentName.text = ""

txtRemarks = ""

txtTotal1 = ""
txtTotal2 = ""
txtTotal3 = ""
txtTotal4 = ""
txtTotal5 = ""
txtTotal6 = ""
txtTotal7 = ""
txtTotal8 = ""
txtTotal9 = ""
txtSponsorshipNo.SetFocus


 vs1.Clear
 
vs1.Cols = 7
 
vs1.TextMatrix(0, 0) = "Bill No"
vs1.TextMatrix(0, 1) = "Date"
vs1.TextMatrix(0, 2) = "Rep.Name"
vs1.TextMatrix(0, 3) = "VType"
vs1.TextMatrix(0, 4) = "Session"
vs1.TextMatrix(0, 5) = "SerName"
vs1.TextMatrix(0, 6) = "Amount"
 
vs1.ColWidth(0) = 900
vs1.ColWidth(1) = 1000
vs1.ColWidth(2) = 2800
vs1.ColWidth(3) = 1100
vs1.ColWidth(4) = 1100
vs1.ColWidth(5) = 1400
vs1.ColWidth(6) = 1200


save.Enabled = True
End Sub
Private Sub Form_Load()

On Error GoTo ee


Screen.MousePointer = vbHourglass

txtDates.value = Format(Date, "dd/MM/yyyy")

max_sp



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
RS.Open "select toDate,fyear from turnOverDis order by fyear", CCON
If RS.EOF = False Then
   RS.MoveNext
   txtMaxDate.value = RS!toDate
   nextYrs = RS!fyear
End If

'If cboTOD.ListCount > 0 Then
'cboTOD.ListIndex = 0
'End If


Picture3.Enabled = True

If Len(inviceNo) > 0 Then
   
   txtSponsorshipNo = inviceNo
   'txtDates.value = PopUpValue2
   'cmbAgentName.Text = PopUpValue3
   
   SearchDataNew
   PopUpValue1 = ""
   PopUpValue2 = ""
   PopUpValue3 = ""
   inviceNo = ""
   Picture3.Enabled = False
End If




Screen.MousePointer = vbDefault
'------------------
Me.top = 0
Me.Left = 0
Me.Width = 14700
Me.Height = 10350
BackColorFrom Me






Exit Sub
ee:

Screen.MousePointer = vbDefault
'------------------
Me.top = 0
Me.Left = 0
Me.Width = 14400
Me.Height = 10435
BackColorFrom Me


End Sub

Private Sub save_Click()

'On Error GoTo aa10



If Edit = True Then
    con.Execute "delete from TurnOver where (EntryNo=" & txtSponsorshipNo & ")"
    Edit = False
Else
    max_sp
End If


If RS.State = 1 Then RS.close
RS.Open "select * from TurnOver where EntryNo=" & txtSponsorshipNo & "", con, adOpenDynamic, adLockOptimistic
If RS.EOF = True Then

For J = 1 To vs.rows - 1
    
If vs.TextMatrix(J, 0) <> "" Then
    RS.AddNew
    RS!entryNo = txtSponsorshipNo.text
    RS!dates = txtDates.value
    RS!RepName = cmbAgentName.text
    RS!subledger = vs.TextMatrix(J, 0)
    RS!sale_current = Val(vs.TextMatrix(J, 1))
    RS!saleRet_current = Val(vs.TextMatrix(J, 2))
    RS!sale_next = Val(vs.TextMatrix(J, 3))
    
    RS!saleRet_next = Val(vs.TextMatrix(J, 4))
    RS!netsale = Val(vs.TextMatrix(J, 5))
    RS!AdjustAmt = Val(vs.TextMatrix(J, 6))
    
    RS!OSAmt = Val(vs.TextMatrix(J, 7))
    
    RS!finalAmt = Val(vs.TextMatrix(J, 8))
    RS!TODDis = Val(vs.TextMatrix(J, 9))
    RS!TODAmt = Val(vs.TextMatrix(J, 10))
    RS!AdjAmt = Val(vs.TextMatrix(J, 11))
    RS!FinalAmount = Val(vs.TextMatrix(J, 12))
    
    RS!UserName = UserName
    RS!remarks = Trim(txtRemarks.text)
    
    
    RS.update
End If

Next

End If






'================================================================================================================
'Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.close
RS.Open "SELECT * FROM TurnBillDet where entryno=" & txtSponsorshipNo & "", con, adOpenDynamic, adLockOptimistic
If RS.EOF = False Then
   con.Execute "delete from TurnBillDet where entryno=" & txtSponsorshipNo & ""
End If
'================================================================================================================
'================================================================================================================
For J = 1 To vs1.rows - 1

If vs1.TextMatrix(J, 0) <> "" Then
    
    
    RS.AddNew
    RS!entryNo = txtSponsorshipNo.text
    RS!billno = vs1.TextMatrix(J, 0)
    RS!dates = vs1.TextMatrix(J, 1)
    RS!RepName = vs1.TextMatrix(J, 2)
    RS!category = vs1.TextMatrix(J, 3)
    RS!fyear = vs1.TextMatrix(J, 4)
    RS!sername = vs1.TextMatrix(J, 5)
    RS!netamt = vs1.TextMatrix(J, 6)
    RS.update

    
    fyear_ = vs1.TextMatrix(J, 4)
    If rs1.State = 1 Then rs1.close
    rs1.Open "select fromDate,toDate,fyear,Current_Next,DataBase,fromDateSRet,toDateSRet from turnOverDis where fyear='" & fyear_ & "'", CCON
    If rs1.EOF = False Then

        If rs1!current_next = "current" Then

           If vs1.TextMatrix(J, 3) = "I" Then
              con.Execute "update INVOICEA set todid='" & txtSponsorshipNo & "',toddate='" & Format(txtDates.value, "MM/dd/yyyy") & "' where invoiceno='" & vs1.TextMatrix(J, 0) & "' and  SUBLEDGER='" & cmbAgentName.text & "'"
           ElseIf vs1.TextMatrix(J, 3) = "CI" Then
              con.Execute "update creditA set todid='" & txtSponsorshipNo & "',toddate='" & Format(txtDates.value, "MM/dd/yyyy") & "' where invoiceno='" & vs1.TextMatrix(J, 0) & "' and  SUBLEDGER='" & cmbAgentName.text & "'"
           ElseIf vs1.TextMatrix(J, 3) = "CN" Then
              con.Execute "update CNF1A set todid='" & txtSponsorshipNo & "',toddate='" & Format(txtDates.value, "MM/dd/yyyy") & "' where cnn='" & vs1.TextMatrix(J, 0) & "' and  PSLD='" & cmbAgentName.text & "'"
           End If


        Else
            '------Fatch Data From Next Session-------------'
            Set CON_next = New ADODB.Connection
            If LCase(server_) = "server" Then
               CON_next.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverName_ & "; DATABASE=" & rs1!Database & "; UID=" & sql_user & "; PWD=" & sql_pass
               CON_next.Open
            Else
               CON_next.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverName_ & "; DATABASE=" & rs1!Database & "; UID=; PWD=;"
               CON_next.Open
            End If


           If vs1.TextMatrix(J, 3) = "I" Then
              CON_next.Execute "update INVOICEA set todid='" & txtSponsorshipNo & "',toddate='" & Format(txtDates.value, "MM/dd/yyyy") & "' where invoiceno='" & vs1.TextMatrix(J, 0) & "' and  SUBLEDGER='" & cmbAgentName.text & "'"
           ElseIf vs1.TextMatrix(J, 3) = "CI" Then
              CON_next.Execute "update creditA set todid='" & txtSponsorshipNo & "',toddate='" & Format(txtDates.value, "MM/dd/yyyy") & "' where invoiceno='" & vs1.TextMatrix(J, 0) & "' and  SUBLEDGER='" & cmbAgentName.text & "'"
           ElseIf vs1.TextMatrix(J, 3) = "CN" Then
              CON_next.Execute "update CNF1A set todid='" & txtSponsorshipNo & "',toddate='" & Format(txtDates.value, "MM/dd/yyyy") & "' where cnn='" & vs1.TextMatrix(J, 0) & "' and  PSLD='" & cmbAgentName.text & "'"
           End If


          End If

     End If
     
     
     

End If

Next


'===========================================================================================

save.Enabled = False
cmdEdit_4.Enabled = True
Add = False
Edit = False

MsgBox "Data Saved...", vbInformation


Exit Sub
aa10:
MsgBox "" & err.DESCRIPTION




End Sub


Private Sub txtDates_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   cmbAgentName.SetFocus
End If
End Sub

Private Sub txtMaxDate_LostFocus()
''  If MsgBox("want to edit", vbQuestion + vbYesNo) = vbYes Then
''     CCON.Execute "update turnOverDis set toDate='" & txtMaxDate.value & "' where fyear='" & nextYrs & "'"
''  End If
  
End Sub

Private Sub txtRemarks_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cboSponse.SetFocus
End Sub
Private Sub txtRemarks_LostFocus()
txtRemarks = UCase(txtRemarks)
End Sub
Private Sub txtSponsorshipNo_GotFocus()
  If PopUpValue1 <> "" Then
     
     txtSponsorshipNo = PopUpValue1
     txtDates.value = PopUpValue2
     cmbAgentName.text = PopUpValue3
     
     searchData
     PopUpValue1 = ""
     PopUpValue2 = ""
     PopUpValue3 = ""
     
  End If
End Sub
Private Sub txtSponsorshipNo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then
    searchType = "inv"
    popuplist10 "select distinct EntryNo,Dates,RepName,TypeOfTOD from TurnOver order by Entryno", con
End If

If KeyCode = 13 Then
   searchData
   txtDates.SetFocus
End If
End Sub

Private Sub txtWhomTobeGiven_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtWhomToBeGivenMob.SetFocus
End Sub

Private Sub txtWhomTobeGiven_LostFocus()
txtWhomTobeGiven = UCase(txtWhomTobeGiven)
End Sub

Private Sub txtTODDiS_Change()

For k1 = 1 To vs.rows - 1
  If vs.TextMatrix(k1, 0) <> "" Then
     
     vs.TextMatrix(k1, 9) = Val(txtTODDiS.text)
     
     If vs.TextMatrix(k1, 9) <> "" Then
        vs.TextMatrix(k1, 10) = Round(Val(vs.TextMatrix(k1, 8)) * Val(vs.TextMatrix(k1, 9)) / 100, 0)
        vs.TextMatrix(k1, 12) = (Val(vs.TextMatrix(k1, 8)) - (Val(vs.TextMatrix(k1, 10)) + Val(vs.TextMatrix(k1, 11))))
     End If
     
  End If
Next


Total

End Sub

Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = 115 Then
   'con.Execute "delete from TurnOver where entryNo=" & txtSponsorshipNo & " and SUBLEDGER='" & vs.TextMatrix(vs.RowSel, 0) & "'"
   vs.RemoveItem vs.RowSel
   Total
End If
End Sub
Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)
  
On Error Resume Next
  
If KeyCode = 13 Then
    If vs.Col = 7 Then
       vs.TextMatrix(vs.RowSel, 8) = Round((Val(vs.TextMatrix(vs.RowSel, 5)) - Val(vs.TextMatrix(vs.RowSel, 6))) - Val(vs.TextMatrix(vs.RowSel, 7)), 0)
       sendkeys ("{right}")
       sendkeys ("{right}")
    ElseIf vs.Col = 11 Then
      vs.TextMatrix(vs.RowSel, 12) = (Val(vs.TextMatrix(vs.RowSel, 8)) - (Val(vs.TextMatrix(vs.RowSel, 10)) + Val(vs.TextMatrix(vs.RowSel, 11))))
      sendkeys "{down}"
      Total
     ElseIf vs.Col = 9 Then
       vs.TextMatrix(vs.RowSel, 10) = Round(Val(vs.TextMatrix(vs.RowSel, 8)) * Val(vs.TextMatrix(vs.RowSel, 9)) / 100, 0)
       vs.TextMatrix(vs.RowSel, 12) = (Val(vs.TextMatrix(vs.RowSel, 8)) - (Val(vs.TextMatrix(vs.RowSel, 10)) + Val(vs.TextMatrix(vs.RowSel, 11))))
       sendkeys "{down}"
       
       Total
   End If
End If
  
End Sub

Private Sub vs1_Click()
vv = vs1.Cell(flexcpBackColor, vs1.RowSel, 1)

If vv = 16777215 Then
   vs1.SelectionMode = flexSelectionByRow
Else
   vs1.SelectionMode = flexSelectionFree
End If

End Sub

Private Sub vs1_KeyDown(KeyCode As Integer, Shift As Integer)

'On Error Resume Next
If KeyCode = 115 Then
   
   If vs1.TextMatrix(vs1.RowSel, 3) = "I" Then
      If (vs1.TextMatrix(vs1.RowSel, 5) = "") Then
         con.Execute "delete from tmpTurnOver where (entryNo=" & txtSponsorshipNo & " and billno='" & vs1.TextMatrix(vs1.RowSel, 0) & "')"
      Else
         con.Execute "delete from tmpTurnOver where (entryNo=" & txtSponsorshipNo & " and billno='" & vs1.TextMatrix(vs1.RowSel, 0) & "' and sername='" & vs1.TextMatrix(vs1.RowSel, 5) & "')"
      End If
   Else
      
      If (vs1.TextMatrix(vs1.RowSel, 5) = "") Then
         con.Execute "delete from tmpTurnOver where entryNo=" & txtSponsorshipNo & " and billno='" & vs1.TextMatrix(vs1.RowSel, 0) & "'"
      Else
         con.Execute "delete from tmpTurnOver where (entryNo=" & txtSponsorshipNo & " and billno='" & vs1.TextMatrix(vs1.RowSel, 0) & "' and sername='" & vs1.TextMatrix(vs1.RowSel, 5) & "')"
      End If

   End If
   
   vs1.RemoveItem vs1.RowSel
End If

End Sub
Private Sub vs1_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      If vs1.Col < 4 Then
         sendkeys "{right}"
      ElseIf vs1.Col = 4 Then
         vs1.rows = vs1.rows + 1
         sendkeys "{down}"
         sendkeys "{home}"
      End If
   End If
End Sub
