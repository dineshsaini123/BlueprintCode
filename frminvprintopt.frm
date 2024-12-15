VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frminvprintopt 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   6075
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cbocopy 
      Height          =   315
      ItemData        =   "frminvprintopt.frx":0000
      Left            =   2040
      List            =   "frminvprintopt.frx":0013
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   1140
      Width           =   3675
   End
   Begin VB.TextBox txtinvoiceno 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Top             =   480
      Width           =   1200
   End
   Begin VB.CommandButton cmdprintchallan 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   2760
      Width           =   1635
   End
   Begin VB.CommandButton cmdviewchallan 
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2070
      TabIndex        =   6
      Top             =   2760
      Width           =   1635
   End
   Begin ComCtl2.UpDown UpDown1 
      Height          =   285
      Left            =   2490
      TabIndex        =   11
      Top             =   1830
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   503
      _Version        =   327681
      Value           =   1
      AutoBuddy       =   -1  'True
      BuddyControl    =   "txtnoofcopies"
      BuddyDispid     =   196613
      OrigLeft        =   2520
      OrigTop         =   1350
      OrigRight       =   2775
      OrigBottom      =   1665
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtnoofcopies 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "1"
      Top             =   1830
      Width           =   450
   End
   Begin VB.CommandButton cmdview 
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2070
      TabIndex        =   4
      Top             =   2280
      Width           =   1635
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   2280
      Width           =   1635
   End
   Begin VB.CheckBox chkagentprint 
      Caption         =   "Yes"
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   1530
      Width           =   855
   End
   Begin VB.ComboBox cbomanufacturedby 
      Height          =   315
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   780
      Width           =   3675
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CHALLAN :"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   2760
      Width           =   1755
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "INVOICE :"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   2280
      Width           =   1755
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000002&
      BorderWidth     =   2
      Height          =   1875
      Left            =   -30
      Top             =   360
      Width           =   6195
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Original/Duplicate : "
      Height          =   255
      Left            =   270
      TabIndex        =   15
      Top             =   1170
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Agent Name on Print :"
      Height          =   255
      Left            =   270
      TabIndex        =   13
      Top             =   1500
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Invoice No. : "
      Height          =   255
      Left            =   270
      TabIndex        =   12
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No of Copies :"
      Height          =   255
      Left            =   270
      TabIndex        =   10
      Top             =   1830
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Invoice Print Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   -30
      TabIndex        =   9
      Top             =   30
      Width           =   6135
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Manufactured By : "
      Height          =   255
      Left            =   270
      TabIndex        =   8
      Top             =   810
      Width           =   1695
   End
End
Attribute VB_Name = "frminvprintopt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbomanufacturedby_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{tab}"
End If

End Sub

Private Sub chkagentprint_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{tab}"
End If

End Sub

Private Sub cmdprint_Click()
invoiceopot True
End Sub

Private Sub cmdprintchallan_Click()
challanopt True
End Sub
Sub invoiceopot(printinv As Boolean)
If validinv = False Then
MsgBox "Invalid Invoice no."
txtinvoiceno.SetFocus
Exit Sub
End If

INVOICE.cr1.SelectionFormula = ""
    INVOICE.cr1.Connect = constr
    INVOICE.cr1.SelectionFormula = "{invoicea.invoiceno} = " & txtinvoiceno.Text & " AND {invoicea.setupid} = " & main.setupid & " AND {invoicea.fyear} = '" & main.session & "'"
    INVOICE.cr1.ReportFileName = strrptpath & "\reports\" & strinvrpt
    INVOICE.cr1.Formulas(0) = "agentprint=" & IIf(chkagentprint.Value = 1, "'True'", "'False'")
    INVOICE.cr1.Formulas(1) = "Manufacturedby=" & IIf(cbomanufacturedby.Text <> "", "'Manufactured By : " & cbomanufacturedby.Text & "'", "''")
    INVOICE.cr1.Formulas(2) = "invoicecopy='" & cbocopy.Text & "'"
    'INVOICE.cr1.Formulas(3) = "toword='" & toword(CDbl(INVOICE.mna.Caption)) & "'"
    
    If printinv = True Then
    INVOICE.cr1.Destination = crptToPrinter
    INVOICE.cr1.CopiesToPrinter = Val(txtnoofcopies.Text)
    Else
    INVOICE.cr1.WindowShowPrintBtn = True
    INVOICE.cr1.WindowShowPrintSetupBtn = True
    INVOICE.cr1.Destination = crptToWindow
    End If
    INVOICE.cr1.Action = 1


End Sub
Private Sub cmdview_Click()
invoiceopot False
End Sub

Private Sub cmdviewchallan_Click()
challanopt False
End Sub
Sub challanopt(printchallan As Boolean)
If validinv = False Then
MsgBox "Invalid Invoice no."
txtinvoiceno.SetFocus
Exit Sub
End If
INVOICE.cr1.SelectionFormula = ""
    INVOICE.cr1.Connect = constr
    INVOICE.cr1.SelectionFormula = "{invoicea.invoiceno} = " & txtinvoiceno.Text & " AND {invoicea.setupid} = " & main.setupid & " AND {invoicea.fyear} = '" & main.session & "'"
    INVOICE.cr1.ReportFileName = strrptpath & "\reports\invchallan.rpt"
    INVOICE.cr1.Formulas(0) = "agentprint='True'"
    INVOICE.cr1.Formulas(1) = "Manufacturedby=''"
    If printchallan = True Then
    INVOICE.cr1.WindowState = crptMaximized
    INVOICE.cr1.Destination = crptToPrinter
    INVOICE.cr1.CopiesToPrinter = Val(txtnoofcopies.Text)
    Else
    INVOICE.cr1.Destination = crptToWindow
    INVOICE.cr1.WindowShowPrintBtn = True
    INVOICE.cr1.WindowShowPrintSetupBtn = True
    End If
    INVOICE.cr1.Action = 1
    

End Sub
Private Sub Form_Load()

    If rs.State = adStateOpen Then rs.Close
    rs.Open "select  * from manufacture where " & stridnyear & " order by mname ", CON, adOpenKeyset, adLockReadOnly, adCmdText
    cbomanufacturedby.Clear
    cbomanufacturedby.AddItem ""
    If Not rs.EOF Then
       Do While Not rs.EOF
          If IsNull(rs(0)) = False Then
            Me.cbomanufacturedby.AddItem rs!mname
          End If
          If Not rs.EOF Then rs.MoveNext
        Loop
    End If
    rs.Close
End Sub

Function validinv() As Boolean
txtinvoiceno.Text = Val(Trim(txtinvoiceno.Text))
If rs.State = 1 Then rs.Close
rs.Open "select * from invoicea where invoiceno=" & Val(Trim(txtinvoiceno.Text)) & " and fyear='" & main.session & "' and setupid=" & main.setupid, CON, adOpenKeyset, adLockOptimistic
If rs.RecordCount <= 0 Then
validinv = False
Else
rs!amttoword = toword(rs!netamount)
rs.Update
validinv = True
End If
rs.Close
End Function

Private Sub txtinvoiceno_KeyPress(KeyAscii As Integer)
txtinvoiceno.Text = Val(txtinvoiceno.Text)
If KeyAscii = 13 Then
SendKeys "{tab}"
End If

End Sub

Private Sub txtnoofcopies_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{tab}"
End If

End Sub
