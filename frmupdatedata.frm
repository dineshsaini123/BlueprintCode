VERSION 5.00
Begin VB.Form frmupdatedata 
   Caption         =   "Form1"
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   5505
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmupdatedata.frx":0000
      Left            =   1110
      List            =   "frmupdatedata.frx":0025
      TabIndex        =   1
      Top             =   600
      Width           =   3225
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update"
      Height          =   675
      Left            =   1530
      TabIndex        =   0
      Top             =   1140
      Width           =   2535
   End
End
Attribute VB_Name = "frmupdatedata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset

Private Sub Command1_Click()
On Error Resume Next

CON.Execute "alter table agentmaster alter column agentname nvarchar(30) not null"
CON.Execute "alter table books alter column bookname nvarchar(40) not null"
CON.Execute "alter table books alter column bookcode nvarsmalldatetime not null"
CON.Execute "alter table cashend alter column cgenledger nvarchar(40) not null"
CON.Execute "alter table cashend alter column csubledger nvarchar(40) not null"
CON.Execute "alter table cashend alter column genledger nvarchar(40) not null"
CON.Execute "alter table cashend alter column subledger nvarchar(40) not null"
CON.Execute "alter table cashend alter column text nvarchar(20) not null"
CON.Execute "alter table challanend alter column cgenledger nvarchar(40) not null"
CON.Execute "alter table challanend alter column csubledger nvarchar(40) not null"
CON.Execute "alter table challanend alter column genledger nvarchar(40) not null"
CON.Execute "alter table challanend alter column subledger nvarchar(40) not null"
CON.Execute "alter table challanend alter column text nvarchar(20) not null"
CON.Execute "alter table creditend alter column cgenledger nvarchar(40) not null"
CON.Execute "alter table creditend alter column csubledger nvarchar(40) not null"
CON.Execute "alter table creditend alter column genledger nvarchar(40) not null"
CON.Execute "alter table creditend alter column subledger nvarchar(40) not null"
CON.Execute "alter table creditend alter column text nvarchar(20) not null"
CON.Execute "alter table districts alter column districtname nvarchar(30) not null"
CON.Execute "alter table districts alter column distcode nvarsmalldatetime not null"
CON.Execute "alter table gledger alter column category nvarchar(11) not null"
CON.Execute "alter table gledger alter column gledger nvarchar(40) not null"
CON.Execute "alter table groups alter column groupname nvarchar(50) not null"
CON.Execute "alter table groups alter column groupcode nvarchar(2) not null"
CON.Execute "alter table invoiceend alter column cgenledger nvarchar(40) not null"
CON.Execute "alter table invoiceend alter column csubledger nvarchar(40) not null"
CON.Execute "alter table invoiceend alter column genledger nvarchar(40) not null"
CON.Execute "alter table invoiceend alter column subledger nvarchar(40) not null"
CON.Execute "alter table invoiceend alter column text nvarchar(20) not null"
CON.Execute "alter table purchaseend alter column cgenledger nvarchar(40) not null"
CON.Execute "alter table purchaseend alter column csubledger nvarchar(40) not null"
CON.Execute "alter table purchaseend alter column genledger nvarchar(40) not null"
CON.Execute "alter table purchaseend alter column subledger nvarchar(40) not null"
CON.Execute "alter table purchaseend alter column text nvarchar(20) not null"
CON.Execute "alter table sledger alter column gledger nvarchar(40) not null"
CON.Execute "alter table sledger alter column subledger nvarchar(50) not null"

CON.Execute "alter table " & Combo1.Text & " alter column fyear nvarsmalldatetime not null"
CON.Execute "alter table " & Combo1.Text & " alter column setupid tinyint not null"

CON.Execute "alter table agentmaster add CONSTRAINT agentkey primary key (agentname,fyear,setupid)"
CON.Execute "alter table books add CONSTRAINT bookkey primary key (bookcode,bookname,fyear,setupid)"
CON.Execute "alter table cashend add CONSTRAINT cahskey primary key (cgenledger,csubledger,genledger,subledger,text,fyear,setupid)"
CON.Execute "alter table challanend add CONSTRAINT challankey primary key (cgenledger,csubledger,genledger,subledger,text,fyear,setupid)"
CON.Execute "alter table creditend add CONSTRAINT creditey primary key (cgenledger,csubledger,genledger,subledger,text,fyear,setupid)"
CON.Execute "alter table districts add CONSTRAINT distkey primary key (districtname,distcode,fyear,setupid)"
CON.Execute "alter table gledger add CONSTRAINT gledkey primary key (category,gledger,fyear,setupid)"
CON.Execute "alter table groups add CONSTRAINT groupkey primary key (groupname,groupcode,fyear,setupid)"
CON.Execute "alter table invoiceend add CONSTRAINT invkey primary key (cgenledger,csubledger,genledger,subledger,text,fyear,setupid)"
CON.Execute "alter table purchaseend add CONSTRAINT purkey primary key (cgenledger,csubledger,genledger,subledger,text,fyear,setupid)"
CON.Execute "alter table sledger add CONSTRAINT slkey primary key (gledger,subledger,fyear,setupid)"

Set rs1 = Nothing
rs1.Open "select * from " & Combo1.Text, CON, adOpenKeyset, adLockOptimistic
Dim arycname2() As String
arycname2 = arycname
For I = 0 To UBound(arycname)
    If rs.State = adStateOpen Then rs.Close
    rs.Open "select * from " & Combo1.Text & " where setupid=" & Left(arycname(I), InStr(1, arycname(I), " (")), CON, adOpenKeyset, adLockOptimistic
    While rs.EOF = False
    K = 0
    For K = 0 To UBound(arycname2)
    If arycname(I) <> arycname2(K) Then
        rs1.AddNew
        For J = 0 To rs.Fields.Count - 1
        If UCase(rs.Fields(J).Name) = UCase("Setupid") Then
        rs1(J) = Left(arycname2(K), InStr(1, arycname2(K), " ("))
        Else
        rs1(J) = rs(J)
        End If
        
        Next
        rs1.Update
    End If
    Next
    rs.MoveNext
    Wend

Next

MsgBox "Updation Done"
End Sub

