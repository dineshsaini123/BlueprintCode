VERSION 5.00
Begin VB.Form frmDataBK 
   Caption         =   "Data Backup"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   516
   ClientWidth     =   5448
   Icon            =   "frmDataBK.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   5448
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1_datafile 
      Height          =   2208
      ItemData        =   "frmDataBK.frx":000C
      Left            =   75
      List            =   "frmDataBK.frx":002B
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   525
      Width           =   5280
   End
   Begin VB.TextBox txtdes 
      Height          =   1545
      Left            =   45
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   3705
      Width           =   5235
   End
   Begin VB.CommandButton cmdCon 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Data Base Backup"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2760
      Width           =   5265
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Database File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   45
      TabIndex        =   3
      Top             =   90
      Width           =   2580
   End
End
Attribute VB_Name = "frmDataBK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCon_Click()
    
    Dim filename, day_ As String
    Dim status_ As String
    Dim path_ As String
    
    path_ = "\\192.168.0.140\blueprintsales\databk"
    day_ = DatePart("d", Date) & "" & DatePart("m", Date) & "" & Year(Date)
    
Screen.MousePointer = vbHourglass

 

On Error GoTo aa1
    
     
    
 For k1 = 0 To List1_datafile.ListCount - 1
 
 If List1_datafile.Selected(k1) = True Then
 
    filename = List1_datafile.List(k1) + "_BK_" & day_
    con.Execute "BACKUP DATABASE " & List1_datafile.List(k1) & " TO DISK='" & path_ & "\" & filename & ".bak'"
    status_ = status_ & List1_datafile.List(k1) & "BACKUP DATABASE successfully"
    
    DoEvents
    DoEvents
    txtdes.Text = status_
    
 End If
    
    
 Next
    
    
    
    Screen.MousePointer = vbDefault

Exit Sub
aa1:
   
    Screen.MousePointer = vbDefault
    MsgBox "" & err.DESCRIPTION
    
    
    
End Sub
