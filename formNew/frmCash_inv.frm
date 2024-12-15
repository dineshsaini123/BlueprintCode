VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCash_inv 
   Caption         =   "Cash"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   9480
   WindowState     =   2  'Maximized
   Begin VB.ComboBox customercode 
      Height          =   2325
      Left            =   4410
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   20
      Top             =   750
      Visible         =   0   'False
      Width           =   3165
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   0
      ScaleHeight     =   390
      ScaleWidth      =   9480
      TabIndex        =   8
      Top             =   6405
      Width           =   9480
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   345
         Left            =   4215
         TabIndex        =   12
         Top             =   30
         Width           =   1125
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "Update"
         Height          =   360
         Left            =   5460
         TabIndex        =   24
         Top             =   30
         Width           =   1320
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Show"
         Height          =   345
         Left            =   3165
         TabIndex        =   23
         Top             =   30
         Width           =   945
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Party"
         Height          =   345
         Left            =   2130
         TabIndex        =   19
         Top             =   30
         Width           =   1005
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Print"
         Height          =   285
         Left            =   4230
         TabIndex        =   18
         Top             =   30
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.OptionButton Optionall 
         Caption         =   "ALL"
         Height          =   255
         Left            =   10860
         TabIndex        =   17
         Top             =   30
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.OptionButton Optiongrno 
         Caption         =   "GR No"
         Height          =   255
         Left            =   9945
         TabIndex        =   16
         Top             =   45
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.OptionButton Optionrrno 
         Caption         =   "RR No"
         Height          =   255
         Left            =   9045
         TabIndex        =   15
         Top             =   60
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   300
         Left            =   480
         TabIndex        =   14
         Top             =   45
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   345
         TabIndex        =   13
         Top             =   45
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   405
         TabIndex        =   11
         Top             =   30
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   300
         Left            =   360
         TabIndex        =   10
         Top             =   30
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   300
         Left            =   405
         TabIndex        =   9
         Top             =   30
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.PictureBox picStatBox 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   9480
      TabIndex        =   2
      Top             =   6105
      Width           =   9480
      Begin VB.CommandButton cmdLast 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3165
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   15
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2835
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   15
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   705
         TabIndex        =   7
         Top             =   15
         Width           =   2130
      End
   End
   Begin VB.ComboBox cbostation 
      Height          =   2325
      Left            =   7635
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   1
      Top             =   765
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.ComboBox cboTrans 
      Height          =   2325
      Left            =   9030
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   0
      Top             =   780
      Visible         =   0   'False
      Width           =   1320
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   2415
      Top             =   4200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
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
      Height          =   315
      Left            =   360
      TabIndex        =   21
      Top             =   3375
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSDataGridLib.DataGrid grdDataGrid 
      Height          =   3315
      Left            =   75
      TabIndex        =   22
      Top             =   15
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   5847
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   17
      TabAction       =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCash_inv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents adoPrimaryRS As Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Dim flg1 As Boolean
Private Sub cbostation_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If grdDataGrid.Col = 3 Then
   Me.grdDataGrid.Text = Me.cboStation.Text
   Me.grdDataGrid.SetFocus
   Me.grdDataGrid.SetFocus
   cboStation.Visible = False
  End If
End If
End Sub
Private Sub cboTrans_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
If grdDataGrid.Col = 5 Then
   Me.grdDataGrid.Text = Me.cboTrans.Text
   Me.grdDataGrid.SetFocus
   Me.grdDataGrid.SetFocus
   Me.cboTrans.Visible = False
  End If
End If
End Sub

Private Sub cmdSales_Click()

End Sub

Private Sub cmdImport_Click()
On Error GoTo aa1:

con.BeginTrans
con.Execute "delete from CashRegister"
con.Execute "insert into  CashRegister(SNO,[Date],Particulars,BDL,WT,Freight,RR,GR,GR_DT,Freight_Paid,CMNo) SELECT INVOICENO, INVOICEDATE,SUBLEDGER,STATION,BUNDLES,transportname,MARKA,val(BILTYNO),BILTYDATE,FREIGHT,INVOICENO FROM casha where " & stringyear & " and len(STATION)>0"
con.CommitTrans

ss = ""
K = 1
If RS.State = 1 Then RS.close
RS.Open "select station,invoiceno from casha where " & stringyear, con, adOpenKeyset, adLockReadOnly
While RS.EOF = False
 
s = InStr(1, RS(0), " ")
If s <> 0 Then
   ss = Trim(Mid(RS(0), 1, s))
Else
   ss = RS(0)
End If
  
con.Execute "update CashRegister set bdl='" & ss & "',sno=" & K & " where " & stringyear & " and CMNo=" & RS(1) & ""
  
  
  K = K + 1
  RS.MoveNext
Wend

MsgBox "Updated ..", vbInformation

Exit Sub
aa1:
MsgBox "" & err.DESCRIPTION

End Sub

Private Sub Command1_Click()

DSNNew

'CON1.Execute "Delete * from CashRegister"
' S1 = App.Path & "\" & directory & "\tchitra.mdb"
s1 = App.Path & "\2003-04\tchitra.mdb"
'con.Execute "insert into  CashRegister  IN '" & S1 & "' select * from CashRegister"
DoEvents
cr1.ReportFileName = rptPath & "\CashRegister.rpt"
If customercode <> "" Then
cr1.ReplaceSelectionFormula ("{BILTYRETURNREGISTER.Particulars} = '" + customercode + "'")
Else
If Optionrrno = True Then
'cr1.ReplaceSelectionFormula ("{CashRegister.Recd_dt} < Date (2000, 01, 01) and {CashRegister.RR} <> ''")
cr1.ReplaceSelectionFormula ("totext({BILTYRETURNREGISTER.Recd_dt}) = '' and {CashRegister.RR} <> ''")
Else
If Optiongrno = True Then
cr1.ReplaceSelectionFormula ("totext({BILTYRETURNREGISTER.Recd_dt}) = '' and {CashRegister.GR} <>  0")
Else
'cr1.ReplaceSelectionFormula ("totext({CashRegister.Recd_dt}) = '' ")
End If
End If
End If
cr1.Action = 1
customercode.Visible = False
customercode = ""
End Sub
Private Sub Command2_Click()
customercode.Top = 3000
customercode.Left = 1500
customercode.Visible = True
End Sub

Private Sub Command3_Click()
addData
End Sub

Private Sub customercode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If grdDataGrid.Col = 2 Then
   Me.grdDataGrid.Text = Me.customercode.Text
   Me.grdDataGrid.SetFocus
   Me.grdDataGrid.SetFocus
   customercode.Visible = False
  End If
  
End If
End Sub
Sub addData()
  If customercode.Text = "" Then Exit Sub
  flg1 = True
  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "select SNO,Date,Particulars,BDL as Station,WT as Bundle,Freight as Transport,RR as PM,CMNo,GR as Bilty,GR_DT,Freight_Paid AS Fright,Remarks from CashRegister where " & stringyear & " and Particulars='" & customercode.Text & "' Order by SNO", con, adOpenDynamic, adLockOptimistic
  Set grdDataGrid.DataSource = adoPrimaryRS
  grdDataGrid.Columns(0).Width = 400
  grdDataGrid.Columns(1).Width = 1000
  grdDataGrid.Columns(2).Width = 2500
  grdDataGrid.Columns(3).Width = 1500
  grdDataGrid.Columns(4).Width = 1000
  grdDataGrid.Columns(5).Width = 1500
  grdDataGrid.Columns(6).Width = 600
  grdDataGrid.Columns(7).Width = 600
  grdDataGrid.Columns(8).Width = 600
  grdDataGrid.Columns(9).Width = 1000
  grdDataGrid.Columns(10).Width = 1000
  grdDataGrid.Columns(11).Width = 1000
  mbDataChanged = False
  customercode.Visible = False
End Sub
Private Sub Form_Load()
  
  
If main.UserName = "v" Then
   Me.cmdImport.Visible = True
Else
   Me.cmdImport.Visible = False
End If

  
  
  flg1 = True
  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "select CMNo,Date,Particulars,BDL as Station,Freight as Transport,WT as PM,GR as Bilty,GR_DT,Freight_Paid AS Fright from CashRegister where " & stringyear & " Order by CMNo", con, adOpenDynamic, adLockOptimistic
  Set grdDataGrid.DataSource = adoPrimaryRS
  ',Remarks
  'WT as Bundle,
  grdDataGrid.Columns(0).Width = 600
  grdDataGrid.Columns(1).Width = 1000
  grdDataGrid.Columns(2).Width = 2500
  grdDataGrid.Columns(3).Width = 1500
  grdDataGrid.Columns(4).Width = 1000
  grdDataGrid.Columns(5).Width = 1500
  grdDataGrid.Columns(6).Width = 1000
  grdDataGrid.Columns(7).Width = 1000
  grdDataGrid.Columns(8).Width = 1000
  'grdDataGrid.Columns(9).Width = 1000
  'grdDataGrid.Columns(10).Width = 1000
  'grdDataGrid.Columns(11).Width = 500
  'grdDataGrid.Columns(12).Width = 500
  'grdDataGrid.Columns(13).Width = 2000
  mbDataChanged = False

  Dim rs1 As New ADODB.Recordset
   at = "SUNDRY DEBTORS"
    rs1.Open "select * from sledger where " & stringyear & " and gledger='" + Trim(at) + "'", con, adOpenDynamic, adLockReadOnly, adCmdText
    If Not rs1.BOF Then
        Do While Not rs1.EOF
            Me.customercode.AddItem rs1(1)
            If Not rs1.EOF Then
                rs1.MoveNext
            End If
        Loop
    End If
    rs1.close
    customercode.Visible = False
    
    rs1.Open "select distinct(STATION) from INVOICEA where " & stringyear, con, adOpenDynamic, adLockReadOnly, adCmdText
    If Not rs1.BOF Then
        Do While Not rs1.EOF
            If rs1(0) <> "" Then
            Me.cboStation.AddItem rs1(0)
            End If
            If Not rs1.EOF Then
                rs1.MoveNext
            End If
        Loop
    End If
    rs1.close
    cboStation.Visible = False
    
    rs1.Open "select distinct(Transportname) from transportmaster where " & stringyear, con, adOpenDynamic, adLockReadOnly, adCmdText
    If Not rs1.BOF Then
        Do While Not rs1.EOF
           If rs1(0) <> "" Then
            Me.cboTrans.AddItem rs1(0)
           End If
           If Not rs1.EOF Then
                rs1.MoveNext
           End If
        Loop
    End If
   rs1.close
   cboTrans.Visible = False
  
   SetButton cmdAdd, cmdUpdate, cmdCancel, cmdDelete
   
   BackColorFrom Me
   
End Sub
Private Sub Form_Resize()
  On Error Resume Next
  'This will resize the grid when the form is resized
  grdDataGrid.Height = ((Me.ScaleHeight - 100 - picButtons.Height - picStatBox.Height) - 100)
  lblStatus.Width = Me.Width - 1500
  cmdNext.Left = lblStatus.Width + 700
  cmdLast.Left = cmdNext.Left + 340
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If mbEditFlag Or mbAddNewFlag Then Exit Sub
  Select Case KeyCode
    Case vbKeyEscape
      cmdClose_Click
    Case vbKeyEnd
     ' cmdLast_Click
    Case vbKeyHome
      'cmdFirst_Click
    Case vbKeyUp, vbKeyPageUp
      If Shift = vbCtrlMask Then
        'cmdFirst_Click
      Else
        'cmdPrevious_Click
      End If
    Case vbKeyDown, vbKeyPageDown
      If Shift = vbCtrlMask Then
       'cmdLast_Click
      Else
        'cmdNext_Click
      End If
  End Select
  
End Sub
Private Sub Form_Unload(cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub
Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, Adstatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This will display the current record position for this recordset
  lblStatus.Caption = "Record: " & CStr(adoPrimaryRS.AbsolutePosition)
End Sub
Private Sub adoPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, Adstatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This is where you put validation code
  'This event gets called when the following actions occur
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then Adstatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  adoPrimaryRS.MoveLast
  grdDataGrid.SetFocus
  adoPrimaryRS.AddNew
  grdDataGrid_OnAddNew
  grdDataGrid.SetFocus
  Exit Sub
AddErr:
  MsgBox err.DESCRIPTION
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  With adoPrimaryRS
    .delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
DeleteErr:
  MsgBox err.DESCRIPTION
End Sub
Private Sub cmdedit_Click()
  On Error GoTo EditErr

  lblStatus.Caption = "Edit record"
  mbEditFlag = True
  SetButtons False
  Exit Sub

EditErr:
  MsgBox err.DESCRIPTION
End Sub
Private Sub cmdCancel_Click()
  On Error Resume Next

  SetButtons True
  mbEditFlag = False
  mbAddNewFlag = False
  adoPrimaryRS.CancelUpdate
  If mvBookMark > 0 Then
    adoPrimaryRS.Bookmark = mvBookMark
  Else
    adoPrimaryRS.MoveFirst
  End If
  mbDataChanged = False

End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr
  adoPrimaryRS.UpdateBatch adAffectAll
  If mbAddNewFlag Then
    adoPrimaryRS.MoveLast              'move to the new record
  End If
  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False
  Exit Sub
UpdateErr:
  MsgBox err.DESCRIPTION
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub
Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError

  adoPrimaryRS.MoveFirst
  mbDataChanged = False

  Exit Sub

GoFirstError:
  MsgBox err.DESCRIPTION
End Sub
Private Sub cmdLast_Click()
  On Error GoTo GoLastError

  adoPrimaryRS.MoveLast
  mbDataChanged = False

  Exit Sub

GoLastError:
  MsgBox err.DESCRIPTION
End Sub
Private Sub cmdNext_Click()
On Error GoTo GoNextError
If Not adoPrimaryRS.EOF Then adoPrimaryRS.MoveNext
If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
Beep
 'moved off the end so go back
adoPrimaryRS.MoveLast
End If
'show the current record
mbDataChanged = False
Exit Sub
GoNextError:
MsgBox err.DESCRIPTION
End Sub
Private Sub cmdPrevious_Click()
On Error GoTo GoPrevError
If Not adoPrimaryRS.BOF Then adoPrimaryRS.MovePrevious
If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then
Beep
'moved off the end so go back
adoPrimaryRS.MoveFirst
End If
'show the current record
mbDataChanged = False
Exit Sub
GoPrevError:
MsgBox err.DESCRIPTION
End Sub
Private Sub SetButtons(bVal As Boolean)
cmdAdd.Visible = bVal
cmdedit.Visible = bVal
cmdUpdate.Visible = Not bVal
cmdCancel.Visible = Not bVal
cmdDelete.Visible = bVal
cmdClose.Visible = bVal
cmdNext.Enabled = bVal
cmdFirst.Enabled = bVal
cmdLast.Enabled = bVal
cmdPrevious.Enabled = bVal
End Sub
Private Sub grdDataGrid_BeforeDelete(cancel As Integer)
If MsgBox("Are you sure....", vbOKCancel) = vbCancel Then
cancel = -1
End If
End Sub
Private Sub grdDataGrid_RowColChange(lastrow As Variant, ByVal lastcol As Integer)
 If grdDataGrid.Col = 1 Then
'''        i_dt.Visible = True
'''        If grdDataGrid.Text <> "" Then
'''            i_dt.Text = grdDataGrid.Text
'''        End If
'''        i_dt.Left = grdDataGrid.Columns(grdDataGrid.col).Left + 80
'''        i_dt.Width = grdDataGrid.Columns(grdDataGrid.col).Width
'''        i_dt.Top = grdDataGrid.RowTop(grdDataGrid.row) + 100
'''        i_dt.SetFocus
   
Else
 If grdDataGrid.Col = 4 Then
'        i_dt.Visible = True
'        'If grdDataGrid.Text <> "" Then
'        '    i_dt.Text = grdDataGrid.Text
'        'End If
'        i_dt.Left = grdDataGrid.Columns(grdDataGrid.col).Left
'        i_dt.Width = grdDataGrid.Columns(grdDataGrid.col).Width
'        i_dt.Top = grdDataGrid.RowTop(grdDataGrid.row)
'        i_dt.SetFocus
'         grdDataGrid.col
Else
If grdDataGrid.Col = 9 Then
        'i_dt.Visible = True
'        'i_dt.Enabled = False
'        If grdDataGrid.Text <> "" Then
'        '    i_dt.Text = grdDataGrid.Text
'        End If
'        i_dt.Left = grdDataGrid.Columns(grdDataGrid.col).Left + 80
'        i_dt.Width = grdDataGrid.Columns(grdDataGrid.col).Width
'        i_dt.Top = grdDataGrid.RowTop(grdDataGrid.row) + 100
        'i_dt.SetFocus
Else
   i_dt.Visible = False
End If
End If
End If
End Sub
Private Sub i_dt_GotFocus()
i_dt.SetFocus
End Sub
Private Sub i_dt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
' If i_dt.Text = "__/__/____" Then
'      MsgBox "Please enter valid date...."
'      i_dt.SetFocus
'      Exit Sub
'   End If
   
   If grdDataGrid.Col = 1 Then
        
        grdDataGrid.Text = ""
        
        'If IsDate(i_dt.Text) = False Then
        '    MsgBox "Please enter valid date...."
        '    i_dt.SetFocus
        '    Exit Sub
        'End If
        grdDataGrid.Text = i_dt.Text
        'i_dt.Visible = False
        a = grdDataGrid.Columns(2).Text
        grdDataGrid.Col = 2
        grdDataGrid.Text = a
        'i_dt.Text = "__/__/____"
        grdDataGrid.SetFocus
        End If
        
    If grdDataGrid.Col = 5 Then
        grdDataGrid.Text = ""
       ' If IsDate(i_dt.Text) = False Then
       '     MsgBox "Please enter valid date...."
       '     i_dt.SetFocus
       '     Exit Sub
       ' End If
       'If i_dt.Text <> "__/__/____" Then
       ' grdDataGrid.Text = i_dt.Text
       ' End If
       ' i_dt.Visible = False
        a = grdDataGrid.Columns(6).Text
        grdDataGrid.Col = 6
        grdDataGrid.Text = a
       ' i_dt.Text = "__/__/____"
        
   End If
   If grdDataGrid.Col = 9 Then
        grdDataGrid.Text = ""
       ' If IsDate(i_dt.Text) = False Then
       '     MsgBox "Please enter valid date...."
       '     i_dt.SetFocus
       '     Exit Sub
       ' End If
       If i_dt.Text <> "__/__/____" Then
       If IsDate(i_dt.Text) Then
        grdDataGrid.Text = i_dt.Text
       End If
       
       End If
        i_dt.Visible = False
        a = grdDataGrid.Columns(10).Text
        grdDataGrid.Col = 10
        grdDataGrid.Text = a
        i_dt.Text = "__/__/____"
   End If
End If
End Sub
Private Sub grdDataGrid_HeadClick(ByVal ColIndex As Integer)

On Error GoTo abc:

h1 = grdDataGrid.Columns(ColIndex).DataField
If h1 = "Station" Then
h1 = "BDL"
End If
If h1 = "Bundle" Then
h1 = "GR"
End If


If h1 = "Transport" Then
h1 = "Freight"
End If


If h1 <> "" Then
    Set adoPrimaryRS = New ADODB.Recordset
    If adoPrimaryRS.State = 1 Then adoPrimaryRS.close
    If flg1 = True Then
       'adoPrimaryRS.Open "select SNO,Date ,Particulars,RR,GR,GR_dt,BDL,WT,Freight,Recd_dt,Freight_Paid,CNO,RNO,Remarks from CashRegister Order by  " & h1, con, adOpenDynamic, adLockOptimistic
       adoPrimaryRS.Open "select SNO,Date,Particulars,BDL as Station,WT as Bundle,Freight as Transport,RR as PM,GR as Bilty,GR_DT,Freight_Paid AS Fright,Remarks from CashRegister where " & stringyear & " Order by  " & h1, con, adOpenDynamic, adLockOptimistic
       flg1 = False
    Else
       'adoPrimaryRS.Open "select SNO,Date ,Particulars,RR,GR,GR_dt,BDL,WT,Freight,Recd_dt,Freight_Paid,CNO,RNO,Remarks from CashRegister Order by " & h1 & " desc", con, adOpenDynamic, adLockOptimistic
       adoPrimaryRS.Open "select SNO,Date,Particulars,BDL as Station,WT as Bundle,Freight as Transport,RR as PM,GR as Bilty,GR_DT,Freight_Paid AS Fright,Remarks from CashRegister where " & stringyear & " Order by " & h1 & " desc", con, adOpenDynamic, adLockOptimistic
       flg1 = True
    End If
    
    Set grdDataGrid.DataSource = adoPrimaryRS
  End If
  grdDataGrid.Columns(0).Width = 400
  grdDataGrid.Columns(1).Width = 1000
  grdDataGrid.Columns(2).Width = 2500
  grdDataGrid.Columns(3).Width = 700
  grdDataGrid.Columns(4).Width = 700
  grdDataGrid.Columns(5).Width = 1000
  grdDataGrid.Columns(6).Width = 400
  grdDataGrid.Columns(7).Width = 500
  grdDataGrid.Columns(8).Width = 1200
  grdDataGrid.Columns(9).Width = 1000
  grdDataGrid.Columns(10).Width = 1000
  'grdDataGrid.Columns(11).Width = 500
  'grdDataGrid.Columns(12).Width = 500
  'grdDataGrid.Columns(13).Width = 1500

Exit Sub
abc:
MsgBox "" & err.DESCRIPTION

End Sub
Private Sub grdDataGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rs1 As New ADODB.Recordset
If rs1.State = 1 Then rs1.close
If grdDataGrid.Col = 2 Then
    If KeyCode = 113 Then
         customercode.Visible = True
         customercode.SetFocus
         customercode.Text = grdDataGrid.Text
         customercode.Left = grdDataGrid.Columns(grdDataGrid.Col).Left + 50
         customercode.Width = grdDataGrid.Columns(grdDataGrid.Col).Width
         customercode.Top = grdDataGrid.RowTop(grdDataGrid.Row) + 100
    End If
    Exit Sub
Else
   customercode.Visible = False
End If

If grdDataGrid.Col = 3 Then
    'If KeyCode = 113 Then
    '     cboStation.Visible = True
    '     cboStation.SetFocus
    '     cboStation.Text = grdDataGrid.Text
    '     cboStation.Left = grdDataGrid.Columns(grdDataGrid.col).Left + 50
    '     cboStation.Width = grdDataGrid.Columns(grdDataGrid.col).Width + 1000
    '     cboStation.Top = grdDataGrid.RowTop(grdDataGrid.row) + 100
    'End If
Else
    '     cboStation.Visible = False
End If

If grdDataGrid.Col = 5 Then
    If KeyCode = 113 Then
       cboTrans.Visible = True
       cboTrans.SetFocus
       cboTrans.Text = grdDataGrid.Text
       cboTrans.Left = grdDataGrid.Columns(grdDataGrid.Col).Left + 50
       cboTrans.Width = grdDataGrid.Columns(grdDataGrid.Col).Width + 1000
       cboTrans.Top = grdDataGrid.RowTop(grdDataGrid.Row) + 100
    End If
Else
       cboTrans.Visible = False
End If


''If KeyCode = 13 Then
''If grdDataGrid.col = 10 Then
''   SendKeys "{Home}"
''   SendKeys "{down}"
''   grdDataGrid.AllowAddNew = True
''End If
''End If
End Sub

Private Sub grdDataGrid_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If grdDataGrid.Col <> 1 Then
   SendKeys "{Tab}"
   If grdDataGrid.Col = 10 Then
        SendKeys "{Home}"
        SendKeys "{DOWN}"
   End If
  Else
  i_dt.Text = "__/__/____"
  'i_dt.Visible = True
  'i_dt.SetFocus
  End If
  
End If
End Sub

Private Sub grdDataGrid_OnAddNew()
Dim rs1 As New ADODB.Recordset
rs1.Open "Select max(Sno)as mno from CashRegister", con, adOpenDynamic
If rs1.RecordCount > 0 Then
  grdDataGrid.Columns(0).Text = IIf(IsNull(rs1!Mno), 1, rs1!Mno + 1)
   grdDataGrid_KeyPress (65)
Else
  grdDataGrid.Columns(0).Text = 1
End If
End Sub

