VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmPermission 
   BackColor       =   &H00E0E0E0&
   Caption         =   "User Permission"
   ClientHeight    =   8556
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   13572
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8556
   ScaleWidth      =   13572
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFC0&
      Caption         =   "&Save"
      Height          =   555
      Left            =   5445
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   7290
      Width           =   1395
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Module List"
      Height          =   2580
      Left            =   10125
      TabIndex        =   15
      Top             =   90
      Visible         =   0   'False
      Width           =   3315
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   2160
         Left            =   45
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   16
         Top             =   270
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Height          =   645
      Left            =   120
      TabIndex        =   11
      Top             =   15
      Width           =   4920
      Begin VB.OptionButton Option3_report 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Reports"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3465
         TabIndex        =   14
         Top             =   180
         Width           =   1230
      End
      Begin VB.OptionButton Option2_trans 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Transaction"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1500
         TabIndex        =   13
         Top             =   180
         Width           =   1680
      End
      Begin VB.OptionButton Option1_master 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Master"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   90
         TabIndex        =   12
         Top             =   180
         Value           =   -1  'True
         Width           =   1230
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Exit"
      Height          =   555
      Left            =   8130
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1965
      Width           =   1335
   End
   Begin VB.CommandButton cmdRef 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Refresh"
      Height          =   555
      Left            =   5235
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1965
      Width           =   1350
   End
   Begin VB.TextBox txtPass 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   6180
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   600
      Width           =   2220
   End
   Begin VB.ComboBox cboUser 
      Height          =   315
      Left            =   6180
      TabIndex        =   5
      Top             =   180
      Width           =   2235
   End
   Begin VB.CheckBox Check_Edit 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Edit"
      Height          =   315
      Left            =   7395
      TabIndex        =   4
      Top             =   1320
      Width           =   1035
   End
   Begin VB.CheckBox Check_Delete 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Delete"
      Height          =   315
      Left            =   6285
      TabIndex        =   3
      Top             =   1320
      Width           =   1035
   End
   Begin VB.CheckBox Check_save 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Save"
      Height          =   315
      Left            =   5280
      TabIndex        =   2
      Top             =   1320
      Width           =   1035
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&OK"
      Height          =   555
      Left            =   6660
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1980
      Width           =   1395
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   7176
      Left            =   135
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   675
      Width           =   4890
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   3495
      Left            =   5400
      TabIndex        =   17
      Top             =   3645
      Width           =   7635
      _cx             =   13467
      _cy             =   6165
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
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPermission.frx":0000
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
   Begin VB.Shape Shape1 
      Height          =   4965
      Left            =   5265
      Top             =   3105
      Width           =   7800
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Form Wise Permission :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5400
      TabIndex        =   18
      Top             =   3240
      Width           =   3435
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   255
      Left            =   5280
      TabIndex        =   8
      Top             =   600
      Width           =   915
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      Height          =   255
      Left            =   5280
      TabIndex        =   6
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmPermission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim types As String
Sub types_()
    
    If Option1_master.value = True Then
       types = "Master"
    ElseIf Option2_trans.value = True Then
       types = "Transaction"
    ElseIf Option3_report.value = True Then
       types = "Report"
    End If
    
End Sub
Private Sub cboUser_Click()
 
 types_
 
 Check_Delete.value = 0
 Check_save.value = 0
 Check_Edit.value = 0
 
 
 
 
For I = 1 To List1.ListCount - 1
   List1.Selected(I) = False
Next
 
For I = 1 To List2.ListCount - 1
   List2.Selected(I) = False
Next
 
 
If RS.State = 1 Then RS.close
RS.Open "select * from UsrePermission where (username='" & cboUser & "' and type='" & types & "') order by taskname", coninfo, adOpenKeyset, adLockReadOnly
If RS.EOF = False Then
  
If RS![delete] = "y" Then
 Check_Delete.value = 1
Else
 Check_Delete.value = 0
End If


If RS![save] = "y" Then
 Check_save.value = 1
Else
 Check_save.value = 0
End If

If RS![Edit] = "y" Then
 Check_Edit.value = 1
Else
 Check_Edit.value = 0
End If

End If
  
  


'---------------------------------------
If RS.State = 1 Then RS.close
RS.Open "select distinct TaskName from UsrePermission where [module]='" & module_ & "' and username='" & cboUser & "' and type='" & types & "' order by TaskName", coninfo, adOpenKeyset, adLockReadOnly
For J = 0 To RS.RecordCount - 1
    
  For I = 0 To List1.ListCount - 1
    If List1.List(I) = RS(0) Then
      List1.Selected(I) = True
    End If
  Next
  
  RS.MoveNext
Next
 
 
'===============================================

If RS.State = 1 Then RS.close
RS.Open "select * from UsrePermission where (username='" & cboUser & "') order by taskname", coninfo, adOpenKeyset, adLockReadOnly
If RS.EOF = False Then

txtpass = RS!Password & ""

For J = 0 To List2.ListCount

If RS.EOF = False Then
If List2.List(J) = RS!TaskName Then
   List2.Selected(J) = True
   RS.MoveNext
End If
End If

Next
  
End If
 
 
 
 
'=================================================================
vs.Clear

If RS.State = 1 Then RS.close
RS.Open "select TaskName,[Save],[Delete],[Edit]  from UsrePermission where UserName='" & cboUser.Text & "' and type='Transaction' and formWiseP='y'"
For I = 1 To RS.RecordCount
  
  vs.TextMatrix(I, 0) = RS!TaskName
  vs.TextMatrix(I, 1) = RS(1)
  vs.TextMatrix(I, 2) = RS(2)
  vs.TextMatrix(I, 3) = RS(3)

  RS.MoveNext
Next


vs.FormatString = "Form Name|Save|Delete|Edit"
vs.ColWidth(0) = 3500
vs.ColWidth(1) = 1100
vs.ColWidth(2) = 1100
vs.ColWidth(3) = 1100

'==================================================================
 
 
 
End Sub

Private Sub cboUser_LostFocus()
'=================================================================
vs.Clear

If RS.State = 1 Then RS.close
RS.Open "select TaskName,[Save],[Delete],[Edit]  from UsrePermission where UserName='" & cboUser.Text & "' and type='Transaction' and formWiseP='y'"
For I = 1 To RS.RecordCount
  
  vs.TextMatrix(I, 0) = RS!TaskName
  
  vs.TextMatrix(I, 1) = RS(1)
  vs.TextMatrix(I, 2) = RS(2)
  vs.TextMatrix(I, 3) = RS(3)

  RS.MoveNext
Next


vs.FormatString = "Form Name|Save|Delete|Edit"
vs.ColWidth(0) = 3500
vs.ColWidth(1) = 1100
vs.ColWidth(2) = 1100
vs.ColWidth(3) = 1100

'==================================================================
 

End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub
Sub addUser()

cboUser.Clear
If RS.State = 1 Then RS.close
RS.Open "select distinct username from UsrePermission where username<>'Admin' order by username", coninfo, adOpenKeyset, adLockReadOnly
While RS.EOF = False
If Not IsNull(RS(0)) Then
cboUser.AddItem RS(0)
End If
RS.MoveNext
Wend

End Sub
Private Sub cmdok_Click()
   
Dim save, delete, Edit, taskType As String
Dim rs_11 As New ADODB.Recordset



If Check_Delete.value = 1 Then
   delete = "y"
Else
   delete = "n"
End If

If Check_save.value = 1 Then
   save = "y"
Else
   save = "n"
End If

If Check_Edit.value = 1 Then
   Edit = "y"
Else
   Edit = "n"
End If

J = 0

types_


coninfo.Execute "delete from UsrePermission where [module]='" & module_ & "' and userName='" & cboUser & "' and Type='" & types & "'"

coninfo.BeginTrans

For I = 1 To List1.ListCount
   
If List1.Selected(J) = True Then

If RS.State = 1 Then RS.close
RS.Open "select tasktype from UsrePermission where [module]='" & module_ & "' and username='" & Trim(cboUser) & "' and taskname='" & List1.List(J) & "'", coninfo, adOpenKeyset, adLockReadOnly
If RS.EOF = True Then

If rs_11.State = 1 Then rs_11.close
rs_11.Open "select Type,tasktype,order_by,[Module] from UsrePermission where  [module]='" & module_ & "' and username='Admin' and taskname='" & List1.List(J) & "'", coninfo, adOpenKeyset, adLockReadOnly
If rs_11.EOF = False Then
    coninfo.Execute "insert into [UsrePermission](type,taskname,permission,[save],[Delete],[Edit],TaskType,[UserName],[password],[module],order_by) " & _
    " values('" & rs_11!Type & "','" & List1.List(J) & "','y','" & [save] & "','" & delete & "','" & Edit & "','" & rs_11!taskType & "','" & cboUser & "','" & txtpass & "','" & rs_11.Fields("Module").value & "','" & rs_11!order_by & "')"
End If

End If

End If

J = J + 1
Next

coninfo.CommitTrans
coninfo.Execute "update UsrePermission set [Save]='" & save & "',[delete]='" & delete & "',[edit]='" & Edit & "' where [module]='" & module_ & "' and userName='" & cboUser & "'"

addModule_permission

MsgBox "Data Saved ...", vbInformation

End Sub
Sub addModule_permission()

coninfo.Execute "delete from UsrePermission where [module]='Module' and userName='" & cboUser & "'"


For I = 1 To List1.ListCount - 1
   
If List1.Selected(I) = True Then

If RS.State = 1 Then RS.close
RS.Open "select tasktype from UsrePermission where [module]='Module' and username='" & Trim(cboUser) & "' and taskname='" & List2.List(J) & "'", coninfo, adOpenKeyset, adLockReadOnly
If RS.EOF = True Then

If rs1.State = 1 Then rs1.close
rs1.Open "select tasktype,order_by from UsrePermission where [module]='Module' and username='Admin' and taskname='" & List2.List(J) & "'", coninfo, adOpenKeyset, adLockReadOnly
If rs1.EOF = False Then
    con.Execute "insert into UsrePermission(type,taskname,permission,TaskType,[UserName],[password],[module],order_by) " & _
    " values('" & types & "','" & List2.List(J) & "','y','" & rs1!taskType & "','" & cboUser & "','" & txtpass & "','Module','" & rs1!order_by & "')"
End If

End If

End If

J = J + 1
Next



End Sub

Private Sub cmdref_Click()
addUser
cmdok.Enabled = True
End Sub

Private Sub cmdSave_Click()


For I = 1 To vs.rows - 1
If vs.TextMatrix(I, 0) <> "" Then
   coninfo.Execute "update UsrePermission set [Save]='" & vs.TextMatrix(I, 1) & "',[Delete]='" & vs.TextMatrix(I, 2) & "',Edit='" & vs.TextMatrix(I, 3) & "',formWiseP='y' where (UserName='" & cboUser.Text & "' and TaskName='" & vs.TextMatrix(I, 0) & "' and [Type]='Transaction') "
End If
Next

MsgBox "Updated .....", vbInformation

End Sub

Private Sub Form_Load()

 If (UCase(UserName) = UCase("admin") Or UCase(UserName) = UCase("v")) Then
    cmdok.Enabled = True
 Else
    cmdok.Enabled = False
 End If



types_
If RS.State = 1 Then RS.close
RS.Open "select distinct TaskName from UsrePermission where [module]='" & module_ & "' and username='Admin' and type='" & types & "' order by TaskName", coninfo, adOpenKeyset, adLockReadOnly
While RS.EOF = False
List1.AddItem RS(0)
RS.MoveNext
Wend

If RS.State = 1 Then RS.close
RS.Open "select distinct [module] from UsrePermission where [username]='Admin' order by [module]", coninfo, adOpenKeyset, adLockReadOnly
While RS.EOF = False
List2.AddItem RS(0)
RS.MoveNext
Wend


'=================================================================
If RS.State = 1 Then RS.close
RS.Open "select TaskName,[Save],[Delete],[Edit]  from UsrePermission where type='Transaction' and formWiseP='y'"
For I = 1 To RS.RecordCount
  
  vs.TextMatrix(I, 0) = RS!TaskName
  
  vs.TextMatrix(I, 1) = RS(1)
  vs.TextMatrix(I, 2) = RS(2)
  vs.TextMatrix(I, 3) = RS(3)

  RS.MoveNext
Next


''vs.FormatString = "Form Name|Save|Delete|Edit"
''
''vs.ColWidth(0) = 3500
''vs.ColWidth(1) = 1100
''vs.ColWidth(2) = 1100
''vs.ColWidth(3) = 1100


'==================================================================


Dim ss_ As String

ss_ = ""

If RS.State = 1 Then RS.close
RS.Open "select distinct TaskName from UsrePermission where [type]='Transaction' order by TaskName", coninfo, adOpenKeyset, adLockReadOnly
While RS.EOF = False

If ss_ = "" Then
   ss_ = RS(0)
Else
   ss_ = ss_ & "|" & RS(0)
End If

RS.MoveNext
Wend

ss_ = ss_ & "|" & "Sub Ledger Master"

vs.ColComboList(0) = ss_
vs.FormatString = "Form Name|Save|Delete|Edit"

vs.ColWidth(0) = 3500
vs.ColWidth(1) = 1100
vs.ColWidth(2) = 1100
vs.ColWidth(3) = 1100


addUser

cboUser.ListIndex = 0

Option1_master_Click

End Sub
Private Sub Option1_master_Click()

types_
List1.Clear

If RS.State = 1 Then RS.close
RS.Open "select distinct TaskName from UsrePermission where [module]='" & module_ & "' and username='Admin' and type='" & types & "' order by TaskName", coninfo, adOpenKeyset, adLockReadOnly
While RS.EOF = False
List1.AddItem RS(0)
RS.MoveNext
Wend

For J = 0 To List1.ListCount - 1
List1.Selected(J) = False
Next

'---------------------------------------
If RS.State = 1 Then RS.close
RS.Open "select distinct TaskName from UsrePermission where [module]='" & module_ & "' and username='" & cboUser & "' and type='" & types & "' order by TaskName", coninfo, adOpenKeyset, adLockReadOnly
For J = 0 To RS.RecordCount - 1
    
  For I = 0 To List1.ListCount - 1
    If List1.List(I) = RS(0) Then
      List1.Selected(I) = True
    End If
  Next
  
  RS.MoveNext
Next

End Sub

Private Sub Option2_trans_Click()
types_
List1.Clear

If RS.State = 1 Then RS.close
RS.Open "select distinct TaskName from UsrePermission where [module]='" & module_ & "' and username='Admin' and type='" & types & "' order by TaskName", coninfo, adOpenKeyset, adLockReadOnly
While RS.EOF = False
List1.AddItem RS(0)
RS.MoveNext
Wend

For J = 0 To List1.ListCount - 1
List1.Selected(J) = False
Next

'---------------------------------------
If RS.State = 1 Then RS.close
RS.Open "select distinct TaskName from UsrePermission where [module]='" & module_ & "' and username='" & cboUser & "' and type='" & types & "' order by TaskName", coninfo, adOpenKeyset, adLockReadOnly
For J = 0 To RS.RecordCount - 1
    
  For I = 0 To List1.ListCount - 1
    If List1.List(I) = RS(0) Then
      List1.Selected(I) = True
    End If
  Next
  
  RS.MoveNext
Next

End Sub
Private Sub Option3_report_Click()
types_
List1.Clear

If RS.State = 1 Then RS.close
RS.Open "select distinct TaskName from UsrePermission where [module]='" & module_ & "' and username='Admin' and type='" & types & "' order by TaskName", coninfo, adOpenKeyset, adLockReadOnly
While RS.EOF = False
List1.AddItem RS(0)
RS.MoveNext
Wend

For J = 0 To List1.ListCount - 1
List1.Selected(J) = False
Next

'---------------------------------------
If RS.State = 1 Then RS.close
RS.Open "select distinct TaskName from UsrePermission where [module]='" & module_ & "' and username='" & cboUser & "' and type='" & types & "' order by TaskName", coninfo, adOpenKeyset, adLockReadOnly
For J = 0 To RS.RecordCount - 1
  
  For I = 0 To List1.ListCount - 1
    If List1.List(I) = RS(0) Then
      List1.Selected(I) = True
    End If
  Next
  
  RS.MoveNext
Next


End Sub

Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then
       
       If (vs.Col = 0 Or vs.Col <= 20) Then
          SendKeys "{right}"
       Else
          SendKeys "{home}"
          SendKeys "{down}"
       End If
       
    End If
    
End Sub
