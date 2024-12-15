VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form popuplist 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Item"
   ClientHeight    =   5736
   ClientLeft      =   5436
   ClientTop       =   2772
   ClientWidth     =   10584
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.6
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   5736
   ScaleWidth      =   10584
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   10500
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5010
      Left            =   45
      TabIndex        =   1
      Top             =   405
      Width           =   10515
      _ExtentX        =   18542
      _ExtentY        =   8827
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      NumItems        =   0
   End
   Begin VB.Label lblRaw 
      BackColor       =   &H0078CFE9&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   60
      TabIndex        =   2
      Top             =   5400
      Width           =   10500
   End
End
Attribute VB_Name = "popuplist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sortmode As Boolean
Public itemname As String
Private Sub Form_Activate()
    sortmode = True
    Text1.text = ""
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
End Sub
Private Sub Form_Load()
  Me.top = 2000
  Me.Left = 2000
End Sub

''Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
''    If KeyCode = 27 Then
''        If check = 5 Then
''            frmBlood_Requist.cmdDelete.Enabled = False
''            frmBlood_Requist.cmdModify.Enabled = False
''            frmBlood_Requist.cmdSave.Enabled = True
''        ElseIf check = 3 Then
''            frmdonar.cmdDelete.Enabled = False
''            frmdonar.cmdModify.Enabled = False
''            frmdonar.cmdSave.Enabled = True
''        ElseIf check = 2 Then
''            frmbloodSupply.cmdDelete.Enabled = False
''            frmbloodSupply.cmdModify.Enabled = False
''            frmbloodSupply.cmdSave.Enabled = True
''        ElseIf check = 6 Then
''            frmDonarTest.cmdDelete.Enabled = False
''            frmDonarTest.cmdModify.Enabled = False
''            frmDonarTest.cmdSave.Enabled = True
''        End If
''        Unload Me
''    End If
''End Sub
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.Sorted = True
        If sortmode = True Then
            ListView1.SortOrder = lvwAscending
            sortmode = False
        Else
            ListView1.SortOrder = lvwDescending
            sortmode = True
        End If
End Sub
Private Sub ListView1_DblClick()
    ListView1_KeyPress 13
    Unload Me
End Sub

Sub ListView1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ct = ListView1.ColumnHeaders.Count
        If ct >= 1 Then PopUpValue1 = ListView1.SelectedItem.text
        If ct >= 2 Then PopUpValue2 = ListView1.SelectedItem.SubItems(1)
        If ct >= 3 Then PopUpValue3 = ListView1.SelectedItem.SubItems(2)
        If ct >= 4 Then popupvalue4 = ListView1.SelectedItem.SubItems(3)
        If ct >= 5 Then popupvalue5 = ListView1.SelectedItem.SubItems(4)
        itemname = Text1.text
        'Text1.Text = ""
        popuplist.Hide
    End If
    If KeyAscii = 27 Then
        Unload Me
    End If
        KeyAscii = 0
End Sub
Public Sub Text1_Change()
    
    Dim ITMFOUND As ListItem
    Set ITMFOUND = ListView1.FindItem(Text1.text, 0, , 1)
        If ITMFOUND Is Nothing Then
            
        Else
            ITMFOUND.EnsureVisible
            ITMFOUND.Selected = True
            ListView1.SetFocus
            Text1.SetFocus
        End If
        
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        sendkeys "{TAB}"
    End If
End Sub