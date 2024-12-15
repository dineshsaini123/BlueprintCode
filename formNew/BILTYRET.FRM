VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form BILTYRETURNREGISTER 
   Caption         =   "BILTYRETURNREGISTER"
   ClientHeight    =   8688
   ClientLeft      =   1116
   ClientTop       =   348
   ClientWidth     =   15252
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8688
   ScaleWidth      =   15252
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport cr1 
      Left            =   2280
      Top             =   4080
      _ExtentX        =   593
      _ExtentY        =   593
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
      Left            =   840
      TabIndex        =   15
      Top             =   3360
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1820
      _ExtentY        =   550
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VB.ComboBox customercode 
      Height          =   2256
      Left            =   2850
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   14
      Top             =   600
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   15252
      TabIndex        =   7
      Top             =   7725
      Width           =   15252
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Party"
         Height          =   435
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   0
         Width           =   960
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Show"
         Height          =   435
         Left            =   3510
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   0
         Width           =   1005
      End
      Begin VB.OptionButton Optionall 
         Caption         =   "C/N Empty List"
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
         Left            =   8076
         TabIndex        =   18
         Top             =   90
         Value           =   -1  'True
         Width           =   1944
      End
      Begin VB.OptionButton Optiongrno 
         Caption         =   "GR No"
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
         Left            =   6885
         TabIndex        =   17
         Top             =   90
         Width           =   975
      End
      Begin VB.OptionButton Optionrrno 
         Caption         =   "RR No"
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
         Left            =   5865
         TabIndex        =   16
         Top             =   90
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cancel"
         Height          =   435
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   450
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Update"
         Height          =   435
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   405
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Close"
         Height          =   435
         Left            =   4530
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         Width           =   1140
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         Height          =   435
         Left            =   1395
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   300
         Left            =   5400
         TabIndex        =   9
         Top             =   630
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add"
         Height          =   435
         Left            =   290
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
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
      ScaleWidth      =   15252
      TabIndex        =   1
      Top             =   8385
      Width           =   15252
      Begin VB.CommandButton cmdLast 
         Caption         =   ">>"
         Height          =   300
         Left            =   4410
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         Height          =   300
         Left            =   4065
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<"
         Height          =   300
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "<<"
         Height          =   300
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   6
         Top             =   0
         Width           =   3360
      End
   End
   Begin MSDataGridLib.DataGrid grdDataGrid 
      Height          =   9192
      Left            =   0
      TabIndex        =   0
      Top             =   48
      Width           =   15096
      _ExtentX        =   26628
      _ExtentY        =   16214
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   17
      TabAction       =   2
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
Attribute VB_Name = "BILTYRETURNREGISTER"
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
Private Sub Command1_Click()

DSNNew


con.Execute "update BILTYRETURNREGISTER set setupid=0 where rno is null"

cr1.Reset
cr1.ReportFileName = rptPath & "\biltyreturnregister.rpt"
cr1.Connect = constr
If customercode <> "" Then
    cr1.ReplaceSelectionFormula ("{BILTYRETURNREGISTER.Particulars} = '" + customercode + "'")
ElseIf Optionrrno = True Then
    cr1.ReplaceSelectionFormula ("totext({BILTYRETURNREGISTER.Recd_dt}) = '' and {BILTYRETURNREGISTER.RR} <> ''")
ElseIf Optiongrno = True Then
    cr1.ReplaceSelectionFormula ("totext({BILTYRETURNREGISTER.Recd_dt}) = '' and {BILTYRETURNREGISTER.GR} <>  0")
ElseIf Optionall = True Then
    cr1.ReplaceSelectionFormula "{BILTYRETURNREGISTER.setupid}=0"

End If



cr1.WindowState = crptMaximized
cr1.Action = 1
customercode.Visible = False
customercode = ""
End Sub

Private Sub Command2_Click()
customercode.top = 3000
customercode.Left = 1720
customercode.Visible = True
End Sub

Private Sub customercode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If grdDataGrid.Col = 2 Then
   Me.grdDataGrid.text = Me.customercode.text
   Me.grdDataGrid.SetFocus
   Me.grdDataGrid.SetFocus
   customercode.Visible = False
  End If
End If
End Sub
Private Sub Form_Load()

Screen.MousePointer = vbHourglass

  flg1 = True
  Set adoPrimaryRS = New Recordset
  'adoPrimaryRS.Open "select SNO,Date ,Particulars,RR,GR,GR_dt,BDL,WT,Freight,Recd_dt,Freight_Paid,CNO,RNO,Remarks from BILTYRETURNREGISTER where " & stringyear & " Order by SNO", CON, adOpenDynamic, adLockOptimistic
  adoPrimaryRS.Open "select SNO,Date ,Particulars,RR,GR,GR_dt,BDL,WT,Freight,Recd_dt,Freight_Paid,CNO as RNO ,RNO as CNO,Remarks from BILTYRETURNREGISTER  Order by SNO", con, adOpenDynamic, adLockOptimistic
  Set grdDataGrid.DataSource = adoPrimaryRS
  grdDataGrid.Columns(0).Width = 700
  grdDataGrid.Columns(1).Width = 1000
  grdDataGrid.Columns(2).Width = 3000
  grdDataGrid.Columns(3).Width = 700
  grdDataGrid.Columns(4).Width = 700
  grdDataGrid.Columns(5).Width = 1000
  grdDataGrid.Columns(6).Width = 400
  grdDataGrid.Columns(7).Width = 500
  grdDataGrid.Columns(8).Width = 1200
  grdDataGrid.Columns(9).Width = 1000
  grdDataGrid.Columns(10).Width = 1000
  grdDataGrid.Columns(11).Width = 500
  grdDataGrid.Columns(12).Width = 1000
  grdDataGrid.Columns(13).Width = 1800

  mbDataChanged = False
  
Dim rs1 As New ADODB.Recordset
at = "SUNDRY DEBTORS"
If rs1.State = 1 Then rs1.close
'rs1.Open "select SUBLEDGER from sledger where " & stringyear & " and (gledger in('SUNDRY DEBTORS','SUNDRY CREDITORS'))", con, adOpenDynamic, adLockReadOnly, adCmdText
Set rs1 = con.Execute("exec AllDebtors")
If Not rs1.BOF Then
    Do While Not rs1.EOF
       
        Me.customercode.AddItem rs1(0)
        If Not rs1.EOF Then
            rs1.MoveNext
        End If
    Loop
End If
rs1.close
customercode.Visible = False
   
SetButton cmdAdd, cmdUpdate, cmdCancel, cmdDelete
 
   
BackColorFrom Me
   
Screen.MousePointer = vbDefault
   
End Sub
Private Sub Form_Resize()
  On Error Resume Next
  'This will resize the grid when the form is resized
  grdDataGrid.Height = Me.ScaleHeight - 30 - picButtons.Height - picStatBox.Height
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
        
        i_dt.Visible = True
        If grdDataGrid.text <> "" Then
            i_dt.text = grdDataGrid.text
        End If
        i_dt.Left = grdDataGrid.Columns(grdDataGrid.Col).Left
        i_dt.Width = grdDataGrid.Columns(grdDataGrid.Col).Width
        i_dt.top = grdDataGrid.RowTop(grdDataGrid.Row)
        i_dt.SetFocus
   
Else
 If grdDataGrid.Col = 5 Then
        i_dt.Visible = True
        If grdDataGrid.text <> "" Then
            i_dt.text = grdDataGrid.text
        End If
        i_dt.Left = grdDataGrid.Columns(grdDataGrid.Col).Left
        i_dt.Width = grdDataGrid.Columns(grdDataGrid.Col).Width
        i_dt.top = grdDataGrid.RowTop(grdDataGrid.Row)
        i_dt.SetFocus
Else
If grdDataGrid.Col = 9 Then
        i_dt.Visible = True
        If grdDataGrid.text <> "" Then
            i_dt.text = grdDataGrid.text
        End If
        i_dt.Left = grdDataGrid.Columns(grdDataGrid.Col).Left
        i_dt.Width = grdDataGrid.Columns(grdDataGrid.Col).Width
        i_dt.top = grdDataGrid.RowTop(grdDataGrid.Row)
        i_dt.SetFocus
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


   If grdDataGrid.Col = 1 Then
   
        grdDataGrid.text = ""
        If IsDate(i_dt.text) = False Then
            MsgBox "Please enter valid date...."
            i_dt.SetFocus
            Exit Sub
        End If
        grdDataGrid.text = i_dt.text
        i_dt.Visible = False
        a = grdDataGrid.Columns(2).text
        grdDataGrid.Col = 2
        grdDataGrid.text = a
        i_dt.text = "__/__/____"
        grdDataGrid.SetFocus
        End If
        
    If grdDataGrid.Col = 5 Then
        
        grdDataGrid.text = ""
        If i_dt.text <> "__/__/____" Then
        If IsDate(i_dt.text) = False Then
            MsgBox "Please enter valid date...."
            i_dt.SetFocus
            Exit Sub
        End If
       End If
       
       
       If i_dt.text <> "__/__/____" Then
        grdDataGrid.text = i_dt.text
       End If
        
        
        i_dt.Visible = False
        a = grdDataGrid.Columns(6).text
        grdDataGrid.Col = 6
        grdDataGrid.text = a
        i_dt.text = "__/__/____"
        
   End If
   If grdDataGrid.Col = 9 Then
        grdDataGrid.text = ""
       ' If IsDate(i_dt.Text) = False Then
       '     MsgBox "Please enter valid date...."
       '     i_dt.SetFocus
       '     Exit Sub
       ' End If
       
        If i_dt.text <> "__/__/____" Then
        If IsDate(i_dt.text) = False Then
            MsgBox "Please enter valid date...."
            i_dt.SetFocus
            Exit Sub
        End If
     End If
       
       
       
        If i_dt.text <> "__/__/____" Then
        grdDataGrid.text = i_dt.text
        End If
        i_dt.Visible = False
        a = grdDataGrid.Columns(10).text
        grdDataGrid.Col = 10
        grdDataGrid.text = a
        i_dt.text = "__/__/____"
   End If
End If
End Sub

Private Sub grdDataGrid_HeadClick(ByVal ColIndex As Integer)


h1 = grdDataGrid.Columns(ColIndex).DataField
If h1 <> "" Then
    If adoPrimaryRS.State = 1 Then adoPrimaryRS.close
    If flg1 = True Then
       'adoPrimaryRS.Open "select SNO,Date ,Particulars,RR,GR,GR_dt,BDL,WT,Freight,Recd_dt,Freight_Paid,CNO,RNO,Remarks from BILTYRETURNREGISTER where " & stringyear & " Order by  " & h1, CON, adOpenDynamic, adLockOptimistic
       adoPrimaryRS.Open "select SNO,Date ,Particulars,RR,GR,GR_dt,BDL,WT,Freight,Recd_dt,Freight_Paid,CNO as RNO,RNO as CNO,Remarks from BILTYRETURNREGISTER Order by  " & h1, con, adOpenDynamic, adLockOptimistic
       flg1 = False
    Else
       'adoPrimaryRS.Open "select SNO,Date ,Particulars,RR,GR,GR_dt,BDL,WT,Freight,Recd_dt,Freight_Paid,CNO,RNO,Remarks from BILTYRETURNREGISTER where " & stringyear & " Order by " & h1 & " desc", CON, adOpenDynamic, adLockOptimistic
       adoPrimaryRS.Open "select SNO,Date ,Particulars,RR,GR,GR_dt,BDL,WT,Freight,Recd_dt,Freight_Paid,CNO as RNO,RNO as CNO,Remarks from BILTYRETURNREGISTER Order by " & h1 & " desc", con, adOpenDynamic, adLockOptimistic
       flg1 = True
    End If
    
    Set grdDataGrid.DataSource = adoPrimaryRS
    
End If

  grdDataGrid.Columns(0).Width = 700
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
  grdDataGrid.Columns(11).Width = 500
  grdDataGrid.Columns(12).Width = 500
  grdDataGrid.Columns(13).Width = 1500

End Sub
Private Sub grdDataGrid_KeyDown(KeyCode As Integer, Shift As Integer)
If grdDataGrid.Col = 2 Then
    If KeyCode = 113 Then
         customercode.Visible = True
         customercode.SetFocus
         customercode.text = grdDataGrid.text
         customercode.Left = grdDataGrid.Columns(grdDataGrid.Col).Left
         customercode.Width = grdDataGrid.Columns(grdDataGrid.Col).Width
         customercode.top = grdDataGrid.RowTop(grdDataGrid.Row)
    End If
Else
   customercode.Visible = False
End If
End Sub

Private Sub grdDataGrid_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If grdDataGrid.Col <> 1 Then
   sendkeys "{Tab}"
   If grdDataGrid.Col = 13 Then
        sendkeys "{Home}"
        sendkeys "{DOWN}"
   End If
  Else
  i_dt.text = "__/__/____"
  i_dt.Visible = True
  i_dt.SetFocus
  End If
  
  
If (grdDataGrid.Col = 3) Then
    kkk = 0
    grdDataGrid.Col = 4
    If (grdDataGrid.text <> "") Then
       kkk = 1
    End If
    grdDataGrid.Col = 3
    If (grdDataGrid.text <> "") Then
       kkk = kkk + 1
    End If
    
    If kkk >= 2 Then
       MsgBox "RR/GR Only 1 Field required ", vbCritical
       Exit Sub
    End If
    
 End If
  
  
If (grdDataGrid.Col = 4) Then
    kkk = 0
    grdDataGrid.Col = 3
    If (grdDataGrid.text <> "") Then
       kkk = 1
    End If
    grdDataGrid.Col = 4
    If (grdDataGrid.text <> "") Then
       kkk = kkk + 1
    End If
    
    If kkk >= 2 Then
       MsgBox "RR/GR Only 1 Field required ", vbCritical
       Exit Sub
    End If
    
 End If
  
  
End If
End Sub

Private Sub grdDataGrid_OnAddNew()
Dim rs1 As New ADODB.Recordset
'rs1.Open "Select max(Sno)as mno from BILTYRETURNREGISTER where " & stringyear, CON, adOpenDynamic
rs1.Open "Select max(Sno)as mno from BILTYRETURNREGISTER", con, adOpenDynamic
If rs1.RecordCount > 0 Then
  grdDataGrid.Columns(0).text = IIf(IsNull(rs1!Mno), 1, rs1!Mno + 1)
  grdDataGrid_KeyPress (65)
Else
  grdDataGrid.Columns(0).text = 1
End If

End Sub
Private Sub i_dt_LostFocus()
'If grdDataGrid.AddNewMode = dbgAddNewCurrent Then
'            grdDataGrid.Text = "12/12/2003"
'             grdDataGrid.Text = i_dt.Text
'             adoPrimaryRS.MovePrevious
'             adoPrimaryRS.MoveLast
'   i_dt.Text = "__/__/____"
             
 '        End If

'i_dt.Text = "__/__/____"
'grdDataGrid.col = 2
End Sub
