VERSION 5.00
Begin VB.Form frmRestoredata 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Restore Data"
   ClientHeight    =   6585
   ClientLeft      =   3705
   ClientTop       =   960
   ClientWidth     =   4500
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   Begin VB.FileListBox File1 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   360
      TabIndex        =   9
      Top             =   4200
      Width           =   3795
   End
   Begin VB.DriveListBox DriveR 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   330
      TabIndex        =   0
      Top             =   795
      Width           =   3840
   End
   Begin VB.DirListBox DirR 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2790
      Left            =   345
      TabIndex        =   1
      Top             =   1170
      Width           =   3825
   End
   Begin VB.PictureBox P1 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      FillColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   2100
      ScaleHeight     =   345
      ScaleWidth      =   1140
      TabIndex        =   2
      Top             =   5970
      Width           =   1170
      Begin VB.Label L1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Restore"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   60
         TabIndex        =   5
         Top             =   75
         Width           =   990
      End
   End
   Begin VB.PictureBox P1 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      FillColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   1
      Left            =   3270
      ScaleHeight     =   345
      ScaleWidth      =   960
      TabIndex        =   3
      Top             =   5970
      Width           =   990
      Begin VB.Label L1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   225
         TabIndex        =   4
         Top             =   45
         Width           =   450
      End
   End
   Begin VB.Label lblRestore 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   195
      TabIndex        =   8
      Top             =   6045
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      X1              =   120
      X2              =   4380
      Y1              =   4095
      Y2              =   4080
   End
   Begin VB.Label LblCaption 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C29221&
      Caption         =   "DATA RESTORE"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   105
      Width           =   4245
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      X1              =   105
      X2              =   4350
      Y1              =   450
      Y2              =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please Select Path ..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   165
      TabIndex        =   6
      Top             =   525
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   6465
      Left            =   60
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "frmRestoredata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MDB_Path As String
Private Sub DirR_Change()
On Error Resume Next
File1.Path = DirR.Path

End Sub
Private Sub DriveR_Change()

On Error GoTo last
    DirR.Path = DriveR.Drive
    Exit Sub
last:
    MsgBox "Drive not Accessable", vbCritical
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Screen.ActiveControl.TabIndex > -1 Then
        SendKeys "{TAB}"
    End If
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
Me.Move 6800, 1000

End Sub

Private Sub L1_Click(Index As Integer)

sss = File1.FileName

If File1.FileName = "" Then
   MsgBox "Select File Name ..", vbCritical
   Exit Sub
End If

ReStoreData

End Sub
Private Sub p1_GotFocus(Index As Integer)
    P1(Index).BackColor = &HC0C0C0
    L1(Index).ForeColor = vbBlack
End Sub
Private Sub P1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

If KeyCode = 39 And Screen.ActiveControl.TabIndex > -1 Then
    SendKeys "{TAB}"
End If
If KeyCode = 37 And Screen.ActiveControl.TabIndex > -1 Then
    SendKeys "+{TAB}"
End If
If KeyCode = 13 Then
    Select Case Index
    Case 0
        ReStoreData
    Case 1
        Unload Me
    End Select
End If

End Sub
Private Sub P1_LostFocus(Index As Integer)
    P1(Index).BackColor = vbBlack
    L1(Index).ForeColor = vbWhite '&HC0C0C0
End Sub
Sub ReStoreData()

If MsgBox("This action will remove the present data." & vbCrLf & "Do you want to continue 'Yes' OR 'No' ?", vbQuestion + vbYesNo, "Warning") = vbYes Then
        
      ses_ = "data_2013-14"
      s_path1 = DirR.Path & "\" & File1.FileName
      s_path2 = "ExportData to " & App.Path & "\" & ses_ & "\ExportData.mdf"
      s_path3 = "ExportData to " & App.Path & "\" & ses_ & "\ExportData_Log.ldf"
      CON.Execute "exec DataRestore '" & s_path1 & "','" & s_path2 & "','" & s_path3 & "'"
        
      lblRestore.Visible = False
      MsgBox "Restore completed !" & vbCrLf & "Please Restart Software !", vbInformation
End
        Exit Sub
End If


End Sub
