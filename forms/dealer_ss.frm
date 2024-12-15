VERSION 5.00
Begin VB.Form dealer_ss 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dealer SS Discount"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnExit 
      Caption         =   "Exit"
      Height          =   555
      Left            =   3240
      TabIndex        =   10
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   555
      Left            =   1560
      TabIndex        =   9
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox txtd_dis2 
      Height          =   315
      Left            =   3060
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtd_dis3 
      Height          =   315
      Left            =   4560
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtss_dis1 
      Height          =   315
      Left            =   1560
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtss_dis2 
      Height          =   315
      Left            =   3060
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtss_dis3 
      Height          =   315
      Left            =   4560
      TabIndex        =   7
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtss_dis4 
      Height          =   315
      Left            =   6000
      TabIndex        =   8
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtd_dis1 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label SSlabel 
      AutoSize        =   -1  'True
      Caption         =   "S.S."
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   300
      TabIndex        =   4
      Top             =   1020
      Width           =   660
   End
   Begin VB.Label Dealerlable 
      AutoSize        =   -1  'True
      Caption         =   "Delaer"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   300
      TabIndex        =   0
      Top             =   540
      Width           =   810
   End
End
Attribute VB_Name = "dealer_ss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnExit_Click()
Unload Me
End Sub

Private Sub btnOK_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from dealer_ss", CON, adOpenKeyset, adLockOptimistic


If Not rs.EOF Then rs.Delete adAffectCurrent

rs.AddNew
rs(0) = txtd_dis1
rs(1) = txtd_dis2
rs(2) = txtd_dis3
rs(3) = txtss_dis1
rs(4) = txtss_dis2
rs(5) = txtss_dis3
rs(6) = txtss_dis4
rs.Update

MsgBox "Data saved", vbInformation, "Done"


End Sub

Private Sub Form_Load()

Me.TOP = 1500
Me.Left = 1000

If rs.State = 1 Then rs.Close
rs.Open "select * from dealer_ss", CON, adOpenKeyset, adLockOptimistic
If Not rs.EOF Then
txtd_dis1 = rs(0)
txtd_dis2 = rs(1)
txtd_dis3 = rs(2)
txtss_dis1 = rs(3)
txtss_dis2 = rs(4)
txtss_dis3 = rs(5)
txtss_dis4 = rs(6)
End If


End Sub
