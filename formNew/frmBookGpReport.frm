VERSION 5.00
Begin VB.Form frmBookGpReport 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   8625
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame buttonFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   945
      Left            =   135
      TabIndex        =   18
      Top             =   3915
      Width           =   2985
      Begin VB.CommandButton cmdAdd_1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   45
         Picture         =   "frmBookGpReport.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1125
         Width           =   1230
      End
      Begin VB.CommandButton cmdSave_2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   75
         Picture         =   "frmBookGpReport.frx":0BE4
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   45
         Width           =   1320
      End
      Begin VB.CommandButton cmdDelete_3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   2880
         Picture         =   "frmBookGpReport.frx":17C8
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1215
         Width           =   1275
      End
      Begin VB.CommandButton cmdEdit_4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   4155
         Picture         =   "frmBookGpReport.frx":23AC
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1215
         Width           =   1275
      End
      Begin VB.CommandButton cmdExit_12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   1470
         Picture         =   "frmBookGpReport.frx":27B9
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   75
         Width           =   1410
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   5445
         Picture         =   "frmBookGpReport.frx":339D
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1215
         Width           =   1365
      End
   End
   Begin VB.ListBox List1 
      Height          =   3180
      ItemData        =   "frmBookGpReport.frx":3F81
      Left            =   90
      List            =   "frmBookGpReport.frx":3F83
      Sorted          =   -1  'True
      TabIndex        =   17
      Top             =   495
      Width           =   1695
   End
   Begin VB.ListBox List3 
      Height          =   3180
      ItemData        =   "frmBookGpReport.frx":3F85
      Left            =   3840
      List            =   "frmBookGpReport.frx":3F87
      Sorted          =   -1  'True
      TabIndex        =   11
      Top             =   450
      Width           =   1035
   End
   Begin VB.CommandButton cmdm4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3420
      Picture         =   "frmBookGpReport.frx":3F89
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1920
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdm3 
      Height          =   540
      Left            =   3420
      Picture         =   "frmBookGpReport.frx":42CB
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1200
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.ListBox List2 
      Height          =   3180
      ItemData        =   "frmBookGpReport.frx":460D
      Left            =   2340
      List            =   "frmBookGpReport.frx":460F
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   450
      Width           =   1035
   End
   Begin VB.CommandButton cmdm2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1860
      Picture         =   "frmBookGpReport.frx":4611
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1920
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.CommandButton cmdm1 
      Height          =   540
      Left            =   1860
      Picture         =   "frmBookGpReport.frx":4953
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1200
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.ListBox List4 
      Height          =   3180
      ItemData        =   "frmBookGpReport.frx":4C95
      Left            =   5355
      List            =   "frmBookGpReport.frx":4C97
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   450
      Width           =   1035
   End
   Begin VB.CommandButton cmdm5 
      Height          =   540
      Left            =   4905
      Picture         =   "frmBookGpReport.frx":4C99
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmd6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4890
      Picture         =   "frmBookGpReport.frx":4FDB
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.ListBox List5 
      Height          =   3180
      ItemData        =   "frmBookGpReport.frx":531D
      Left            =   6975
      List            =   "frmBookGpReport.frx":5324
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   450
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   6510
      Picture         =   "frmBookGpReport.frx":532F
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Height          =   540
      Left            =   6525
      Picture         =   "frmBookGpReport.frx":5671
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0078CFE9&
      BorderWidth     =   4
      Height          =   1050
      Left            =   90
      Top             =   3870
      Width           =   3120
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Group3"
      Height          =   345
      Left            =   5535
      TabIndex        =   16
      Top             =   135
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Group2 "
      Height          =   285
      Left            =   3915
      TabIndex        =   15
      Top             =   135
      Width           =   840
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Group1"
      Height          =   270
      Left            =   2385
      TabIndex        =   14
      Top             =   135
      Width           =   960
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Group Name"
      Height          =   315
      Left            =   135
      TabIndex        =   13
      Top             =   135
      Width           =   1635
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Group4"
      Height          =   300
      Left            =   7155
      TabIndex        =   12
      Top             =   135
      Width           =   735
   End
End
Attribute VB_Name = "frmBookGpReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd6_Click()
If List4.ListCount > 0 Then
    Dim I
    Dim delitem As Integer
    For I = 0 To List4.ListCount - 1
        If List4.Selected(I) Then
                List1.AddItem List4.List(I)
                delitem = I
         End If
    Next
    List4.RemoveItem delitem
End If
End Sub

Private Sub cmdExit_12_Click()
Unload Me
End Sub

Private Sub cmdm1_Click()
If List1.ListCount > 0 Then
    Dim I
    Dim delitem As Integer
    For I = 0 To List1.ListCount - 1
        If List1.Selected(I) Then
                List2.AddItem List1.List(I)
                delitem = I
         End If
    Next I
    List1.RemoveItem delitem
End If
End Sub
Sub addgp4()
   
    
 List5.Clear
 If RS.State = 1 Then RS.close
 RS.Open "select groupcode from groups where " & stringyear & " and Group4=1", CON
 While RS.EOF = False
    List5.AddItem RS(0)
 RS.MoveNext
 Wend
 RS.close
 
List1.Clear
Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset
rs1.Open "Select groupcode from groups where " & stringyear & " and group1= 0 and group2=0 and  group3=0 and group4=0", CON, adOpenStatic, adLockOptimistic

If rs1.RecordCount > 0 Then
rs1.MoveFirst
While Not rs1.EOF
    List1.AddItem rs1!groupcode
    rs1.MoveNext
Wend
End If



End Sub
Sub addData()

List2.Clear
List3.Clear
List4.Clear
List5.Clear


If rs1.State = 1 Then rs1.close
rs1.Open "Select groupcode from groups where " & stringyear & " and group1= 1", CON, adOpenStatic, adLockOptimistic
If rs1.RecordCount > 0 Then
       rs1.MoveFirst
       While Not rs1.EOF
           List2.AddItem rs1!groupcode
           rs1.MoveNext
       Wend
End If

If rs1.State = 1 Then rs1.close
rs1.Open "Select groupcode from groups where " & stringyear & " and  group2=1", CON, adOpenStatic, adLockOptimistic
If rs1.RecordCount > 0 Then
       rs1.MoveFirst
       While Not rs1.EOF
           List3.AddItem rs1!groupcode
           rs1.MoveNext
       Wend
End If


If rs1.State = 1 Then rs1.close
rs1.Open "Select groupcode from groups where " & stringyear & " and group3=1", CON, adOpenStatic, adLockOptimistic
If rs1.RecordCount > 0 Then
       rs1.MoveFirst
       While Not rs1.EOF
           List4.AddItem rs1!groupcode
           rs1.MoveNext
       Wend
End If


If RS.State = 1 Then RS.close
RS.Open "select groupcode from groups where  " & stringyear & " and Group4=1", CON
While RS.EOF = False
List5.AddItem RS(0)
RS.MoveNext
Wend
RS.close




End Sub
Private Sub cmdm2_Click()
If List2.ListCount > 0 Then
    Dim I
    Dim delitem As Integer
    For I = 0 To List2.ListCount - 1
        If List2.Selected(I) Then
                List1.AddItem List2.List(I)
                delitem = I
         End If
    Next
    List2.RemoveItem delitem
End If
End Sub

Private Sub cmdm3_Click()
If List1.ListCount > 0 Then
    Dim I
    Dim delitem As Integer
    For I = 0 To List1.ListCount - 1
        If List1.Selected(I) Then
                List3.AddItem List1.List(I)
                delitem = I
         End If
    Next
    List1.RemoveItem delitem
End If

End Sub

Private Sub cmdm4_Click()
If List3.ListCount > 0 Then
    Dim I
    Dim delitem As Integer
    For I = 0 To List3.ListCount - 1
        If List3.Selected(I) Then
                List1.AddItem List3.List(I)
                delitem = I
         End If
    Next
    List3.RemoveItem delitem
End If
End Sub

Private Sub cmdm5_Click()
If List1.ListCount > 0 Then
    Dim I
    Dim delitem As Integer
    For I = 0 To List1.ListCount - 1
        If List1.Selected(I) Then
                List4.AddItem List1.List(I)
                delitem = I
         End If
    Next
    List1.RemoveItem delitem
End If
End Sub
Private Sub cmdSave_2_Click()

CON.Execute "update GROUPS set group1=0,group2=0,group3=0,group4=0 where " & stringyear & ""

For I = 0 To List2.ListCount - 1
    CON.Execute "update GROUPS set group1=1 where " & stringyear & " and groupcode='" & List2.List(I) & "'"
Next

For I = 0 To List3.ListCount - 1
    CON.Execute "update GROUPS set group2=1 where " & stringyear & " and groupcode='" & List3.List(I) & "'"
Next


For I = 0 To List4.ListCount - 1
    CON.Execute "update GROUPS set group3=1 where " & stringyear & " and groupcode='" & List4.List(I) & "'"
Next


For I = 0 To List5.ListCount - 1
    CON.Execute "update GROUPS set group4=1 where " & stringyear & " and groupcode='" & List5.List(I) & "'"
Next



End Sub
Private Sub Command1_Click()

If List5.ListCount > 0 Then
    Dim I
    Dim delitem As Integer
    For I = 0 To List5.ListCount - 1
        If List5.Selected(I) Then
                List1.AddItem List5.List(I)
                delitem = I
         End If
    Next
    List5.RemoveItem delitem
End If

End Sub

Private Sub Command2_Click()

If List1.ListCount > 0 Then
    Dim I
    Dim delitem As Integer
    For I = 0 To List1.ListCount - 1
        If List1.Selected(I) Then
                List5.AddItem List1.List(I)
                delitem = I
         End If
    Next
    List1.RemoveItem delitem
End If

End Sub

Private Sub Form_Load()
    
    Me.Top = 2000
    Me.Left = 1000
   
    addgp4
    
    addData
    
    BackColorFrom Me

End Sub
