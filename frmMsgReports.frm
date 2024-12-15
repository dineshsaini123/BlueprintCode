VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmMsgReports 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3000
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4830
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   4830
   Begin Crystal.CrystalReport cr 
      Left            =   225
      Top             =   2475
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Height          =   915
      Left            =   765
      TabIndex        =   3
      Top             =   1665
      Width           =   3135
      Begin VB.CommandButton cmdexit 
         Caption         =   "&Exit"
         Height          =   465
         Left            =   1440
         TabIndex        =   4
         Top             =   225
         Width           =   1290
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Print"
         Height          =   465
         Left            =   135
         TabIndex        =   1
         Top             =   225
         Width           =   1215
      End
   End
   Begin VB.TextBox txtRawCode 
      Height          =   315
      Left            =   900
      TabIndex        =   0
      Top             =   1095
      Width           =   1740
   End
   Begin VB.Label Label4 
      Caption         =   "Raw Code"
      Height          =   240
      Left            =   915
      TabIndex        =   2
      Top             =   795
      Width           =   1140
   End
End
Attribute VB_Name = "frmMsgReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim K As Integer
Dim b As Boolean
Sub setwidth()
vs.FormatString = "Raw Code|Finish Code|Finish Qty"
vs.ColWidth(0) = 2000
vs.ColWidth(1) = 2000
vs.ColWidth(2) = 2000
vs.Rows = 2
End Sub


Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdref_Click()
vs.Clear
setwidth

txtPno.Text = ""
txtRawCode.Text = ""
txtFCode.Text = ""
txtqty.Text = ""
txtPno.SetFocus


End Sub


Private Sub pdates_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Me.txtRawCode.SetFocus
End Sub

Private Sub txtFCode_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then
   popuplist10 "select BOOKcode as ItemCode,BOOKNAME + ': ' + convert(varchar,size1) + ' ' + unit1 + ' ' + convert(varchar,size2) + ' ' + unit2 + ': ' + quality as Item from books where " & stringyear & "  and upper(GROUPCODE)='YES'", CON
End If
End Sub
Private Sub txtPno_GotFocus()




End Sub

Private Sub txtPno_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then


End If


If KeyCode = 113 Then
   popuplist10 "select Pno,Dates from mfgtable where " & stringyear & " group by Pno,Dates", CON
End If

End Sub
Private Sub txtqty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
End If
End Sub

Private Sub CmdSave_Click()
cr.Reset
cr.Connect = constr
cr.ReportFileName = strrptpath & "\rEPORTS\MfgReports.rpt"
cr.ReplaceSelectionFormula "{mfgitem.rcode}='" & Me.txtRawCode.Text & "'"
cr.WindowShowPrintSetupBtn = True
cr.WindowShowRefreshBtn = True
cr.WindowState = crptMaximized
cr.Action = 1
End Sub

Private Sub txtRawCode_GotFocus()
If PopUpValue1 <> "" Then
Me.txtRawCode.Text = PopUpValue1
PopUpValue1 = ""
End If
End Sub
Private Sub txtRawCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
   popuplist10 "select BOOKCODE as ItemCode,BOOKNAME + ': ' + convert(varchar,size1) + ' ' + unit1 + ' ' + convert(varchar,size2) + ' ' + unit2 + ': ' + quality as Item from books where " & stringyear & " and upper(GROUPCODE)='NO'", CON
End If
End Sub

Private Sub VS_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
vs.RemoveItem (vs.RowSel)
End If
End Sub
