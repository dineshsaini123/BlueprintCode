VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTotBookQty_Amt 
   Caption         =   "Total Book Qty/Amount"
   ClientHeight    =   2964
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   5016
   Icon            =   "frmq.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   2964
   ScaleWidth      =   5016
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Print (Total Book Amt.)"
      Height          =   900
      Left            =   1740
      Picture         =   "frmq.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1680
      Width           =   1185
   End
   Begin VB.CommandButton close 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&Exit"
      Height          =   900
      Left            =   2940
      Picture         =   "frmq.frx":0BF0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1680
      Width           =   1185
   End
   Begin VB.CommandButton cmdPrint_7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Print (Total Book Qty.)"
      Height          =   900
      Left            =   360
      Picture         =   "frmq.frx":17D4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1680
      Width           =   1365
   End
   Begin MSComCtl2.DTPicker txtDateTo 
      Height          =   315
      Left            =   2280
      TabIndex        =   0
      Top             =   840
      Width           =   1575
      _ExtentX        =   2773
      _ExtentY        =   550
      _Version        =   393216
      Format          =   155582465
      CurrentDate     =   42409
   End
   Begin MSComCtl2.DTPicker txtFrom 
      Height          =   315
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   1575
      _ExtentX        =   2773
      _ExtentY        =   550
      _Version        =   393216
      Format          =   155582465
      CurrentDate     =   42409
   End
   Begin VB.Label Label4 
      Caption         =   "To"
      Height          =   315
      Left            =   1980
      TabIndex        =   2
      Top             =   900
      Width           =   195
   End
End
Attribute VB_Name = "frmTotBookQty_Amt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub close_Click()
Unload Me
End Sub
Private Sub cmdPrint_7_Click()

DSNNew

con.Execute "delete from tmpTotalBookIssue"
con.Execute "insert into tmpTotalBookIssue select bookcode,bookname,sum(Issue),sum(netamount),GROUPCODE from TotalBookIssue where (invoicedate >= convert(smalldatetime,'" & txtFrom.value & "',103) and invoicedate <= convert(smalldatetime,'" & txtDateTo.value & "',103)) group by bookcode,bookname,GROUPCODE"

con.Execute "delete from tmpTotalBookRec"
con.Execute "insert into tmpTotalBookRec select bookcode,sum(Qty),sum(amount) from TotalBookRec where (invoicedate >= convert(smalldatetime,'" & txtFrom.value & "',103) and invoicedate <= convert(smalldatetime,'" & txtDateTo.value & "',103)) group by bookcode"


MsgBox " View ? ", vbQuestion

frmAgentLadger.cr.Reset
    frmAgentLadger.cr.ReportFileName = rptPath & "/TotalBookLedger.rpt"
        frmAgentLadger.cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
        frmAgentLadger.cr.Formulas(0) = "fdate='" & txtFrom.value & "'"
        frmAgentLadger.cr.Formulas(1) = "tdate='" & txtDateTo.value & "'"
            frmAgentLadger.cr.WindowShowPrintBtn = True
        frmAgentLadger.cr.WindowShowPrintSetupBtn = True
    frmAgentLadger.cr.WindowShowSearchBtn = True
    
frmAgentLadger.cr.WindowState = crptMaximized
frmAgentLadger.cr.Action = 1


End Sub
Private Sub Command1_Click()

DSNNew

con.Execute "delete from tmpTotalBookIssue"
con.Execute "insert into tmpTotalBookIssue select bookcode,bookname,sum(Issue),sum(netamount),GROUPCODE from TotalBookIssue where (invoicedate >= convert(smalldatetime,'" & txtFrom.value & "',103) and invoicedate <= convert(smalldatetime,'" & txtDateTo.value & "',103)) group by bookcode,bookname,GROUPCODE"

con.Execute "delete from tmpTotalBookRec"
con.Execute "insert into tmpTotalBookRec select bookcode,sum(Qty),sum(amount) from TotalBookRec where (invoicedate >= convert(smalldatetime,'" & txtFrom.value & "',103) and invoicedate <= convert(smalldatetime,'" & txtDateTo.value & "',103)) group by bookcode"


MsgBox " View ? ", vbQuestion


frmAgentLadger.cr.Reset
  frmAgentLadger.cr.ReportFileName = rptPath & "/TotalBookLedgerTotal.rpt"
   frmAgentLadger.cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
           frmAgentLadger.cr.Formulas(0) = "fdate='" & txtFrom.value & "'"
        frmAgentLadger.cr.Formulas(1) = "tdate='" & txtDateTo.value & "'"

    frmAgentLadger.cr.WindowShowPrintBtn = True
    frmAgentLadger.cr.WindowShowPrintSetupBtn = True
   frmAgentLadger.cr.WindowShowSearchBtn = True
  frmAgentLadger.cr.WindowState = crptMaximized
frmAgentLadger.cr.Action = 1

End Sub

Private Sub Form_Load()
 Me.txtFrom.value = from_date
 Me.txtDateTo.value = to_date
End Sub
