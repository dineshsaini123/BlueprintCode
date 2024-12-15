VERSION 5.00
Begin VB.Form frmLoginDonn 
   Caption         =   "Login"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5145
   Icon            =   "frmLoginDonn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   5145
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   450
      Left            =   2640
      TabIndex        =   10
      Top             =   1590
      Width           =   1140
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmLoginDonn.frx":000C
      Left            =   1500
      List            =   "frmLoginDonn.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   30
      Width           =   1515
   End
   Begin VB.TextBox txtUserName 
      Height          =   315
      Left            =   1500
      TabIndex        =   8
      Top             =   540
      Width           =   3015
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1500
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   1050
      Width           =   1485
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   465
      Left            =   1470
      TabIndex        =   6
      Top             =   1590
      Width           =   1155
   End
   Begin VB.ComboBox cbocname 
      Height          =   315
      ItemData        =   "frmLoginDonn.frx":002C
      Left            =   5760
      List            =   "frmLoginDonn.frx":002E
      TabIndex        =   5
      Text            =   "cbocname"
      Top             =   -60
      Width           =   975
   End
   Begin VB.CheckBox chkdefault 
      Caption         =   "Set as Default"
      Height          =   285
      Left            =   3120
      TabIndex        =   4
      Top             =   90
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.ComboBox cboModule 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmLoginDonn.frx":0030
      Left            =   -120
      List            =   "frmLoginDonn.frx":0043
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1725
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton cmdCon 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Connect Data Base Server"
      Height          =   585
      Left            =   420
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2445
      Width           =   4500
   End
   Begin VB.CheckBox Check1_other 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Click  for other option connect to server"
      Height          =   315
      Left            =   540
      TabIndex        =   1
      Top             =   3165
      Visible         =   0   'False
      Width           =   3675
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DSN"
      Height          =   375
      Left            =   4380
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3165
      Width           =   495
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Session :"
      Height          =   285
      Index           =   2
      Left            =   315
      TabIndex        =   15
      Top             =   60
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name :"
      Height          =   285
      Index           =   0
      Left            =   315
      TabIndex        =   14
      Top             =   570
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Password :"
      Height          =   285
      Index           =   1
      Left            =   315
      TabIndex        =   13
      Top             =   1110
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Module Name :"
      Height          =   195
      Index           =   3
      Left            =   -120
      TabIndex        =   12
      Top             =   1785
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   300
      TabIndex        =   11
      Top             =   3885
      Visible         =   0   'False
      Width           =   4695
   End
End
Attribute VB_Name = "frmLoginDonn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public LoginSucceeded As Boolean
Dim mydsn_ As Boolean
'Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
     (ByVal lpBuffer As String, nSize As Long) As Long
Sub clearExeFromProcess()

''''On Error GoTo ErrHandler
''''Dim oWMI
''''Dim ret
''''Dim sService
''''Dim oWMIServices
''''Dim oWMIService
''''Dim oServices
''''Dim oService
''''Dim servicename
''''Set oWMI = GetObject("winmgmts:")
''''Set oServices = oWMI.InstancesOf("win32_process")
''''
''''For Each oService In oServices
''''                 servicename = LCase(Trim(CStr(oService.Name) & ""))
''''                 If InStr(1, servicename, LCase("bluePrintSale.exe"), vbTextCompare) > 0 Then
''''                    ret = oService.Terminate
''''                 End If
''''Next
''''
''''Set oServices = Nothing
''''Set oWMI = Nothing
''''
''''ErrHandler:
''''err.Clear


End Sub
Sub Loadrights()
  'MainMenu.Caption = "Publication Software System __ " & cbocname.Text & "__ " + Combo1
   MainMenu.Show
   
   
End Sub
''Sub DSN()
''Dim FSO As FileSystemObject
''Dim f As File
''Dim txt As TextStream
''Dim matter As String
''Dim Total As String
''Dim s(1, 2) As String
''Set FSO = New FileSystemObject
''Dim ss As String
''If win7 = "x" Then
''   dstrpath = "C:\Progra~1\Common~1\ODBC\DataSo~1\chitradsn.dsn"
''   If FSO.FolderExists(dstrpath) = True Then
''      Set f = FSO.DeleteFile("" & dstrpath)
''
''   End If
''Else
''   dstrpath = "C:\Users\" & sys_user & "\Documents\chitradsn.dsn"
''   Set f = FSO.DeleteFile("" & dstrpath)
''   f.delete
''End If
''
''                Set txt = FSO.CreateTextFile("" & dstrpath)
''                matter = matter & "[ODBC]" & vbNewLine
''                matter = matter & "DRIVER=Microsoft Access Driver (*.mdb)" & vbNewLine
''                matter = matter & "UID = admin" & vbNewLine
''                matter = matter & "UserCommitSync = Yes" & vbNewLine
''                matter = matter & "Threads = 3" & vbNewLine
''                matter = matter & "afeTransactions = 0" & vbNewLine
''                matter = matter & "PageTimeout = 5" & vbNewLine
''                matter = matter & "MaxScanRows = 8" & vbNewLine
''                matter = matter & "MaxBufferSize = 2048" & vbNewLine
''                matter = matter & "FIL=MS Access" & vbNewLine
''                matter = matter & "DriverId = 25" & vbNewLine
''                matter = matter & "DefaultDir=" & App.Path & vbNewLine
''                matter = matter & "DBQ=" & App.Path & "\" & main.directory & "\" & "Data.mdb"
''
''txt.Write matter
''txt.close
''
''End Sub
Sub fillsession()

   If rs.State <> adStateClosed Then rs.close
   'Combo1.Clear
   rs.Open "select setupid from setup1 where cname='" & cbocname.Text & "'", con, adOpenStatic, adLockReadOnly, adCmdText
   If Not rs.EOF Then
        main.setupid = Val(rs!setupid)
        rs.close
        rs.Open "select fyear from financialyear where setupid=" & main.setupid & " order by fyear desc", con, adOpenStatic, adLockReadOnly, adCmdText
        If Not rs.EOF Then
            Do While Not rs.EOF
                'Combo1.AddItem RS!fyear
                If Not rs.EOF Then
                   rs.MoveNext
                End If
            Loop
            Combo1.ListIndex = 0
        Else
        MsgBox "Session not Declared for this Compnay."
        End If
   Else
        MsgBox "Compnay Name Not Found."
   End If
    
End Sub
Private Sub cboCName_Click()
 fillsession
End Sub

Private Sub cboModule_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then SendKeys "{tab}"
End Sub

Private Sub cmdAdd_Click()

'''----------------------------------
''Dim dr, cr
''
''If rs1.State = 1 Then rs1.close
''rs1.Open "select * from INVOICEa order by invoiceno", CON
''While rs1.EOF = False
''
''dr = 0
''cr = 0
''
''If RS.State = 1 Then RS.close
''RS.Open "select * from INVOICEC where invoiceno=" & rs1!invoiceno & "", CON
''While RS.EOF = False
''
''If RS!DEBITORCREDIT = "Credit" Then
''cr = cr + RS!amount
''Else
''dr = dr + RS!amount
''End If
''
''
''RS.MoveNext
''Wend
''
''
''dr = Val(dr - cr)
''
''If dr > 0 Then
''   dr = -1 * dr
''Else
''   dr = Abs(dr)
''End If
''
''
''CON.Execute "update INVOICEC set saleType='" & dr & "' where invoiceno=" & rs1!invoiceno & ""
''
''rs1.MoveNext
''Wend
''
''Exit Sub

'----------------------------------


GetCom
Unload Me
frmLogout.Show 1
End Sub

Private Sub cmdCancel_Click()
    If rs.State = 1 Then
        rs.close
    End If
    LoginSucceeded = False
    Me.Hide
    End
End Sub
Sub Fill_Module()

If rs.State = 1 Then rs.close
rs.Open "select * from UsrePermission where (module='Module' and username='" & UserName & "') order by order_by", con, adOpenKeyset, adLockReadOnly
If rs.EOF = False Then
   module_ = rs!taskname
End If

End Sub
Function BlueDSN() As Boolean

On Error GoTo conErr


Set CON_blue = New ADODB.Connection
If CON_blue.State = 1 Then CON_blue.close
mydsn_ = False

If server_ = "server" Then

  CON_blue.ConnectionString = "filedsn=MyDSN;PWD=dinesh.123;"
  'constr = "filedsn=MyDSN;PWD=dinesh.123;"
  'CON_blue.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=NNSERVER\BLUEPRINT; DATABASE=mydata_new; UID=sa; PWD=dinesh.123;"
Else
  DoEvents
  DoEvents
  DoEvents
  CON_blue.ConnectionString = "filedsn=MyDSN;PWD=sidc;"
  'constr = "filedsn=MyDSN;PWD=sidc;"
  'CON_blue.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=COMPAQ\SQL2008NEW; DATABASE=mydata_new; UID=; PWD=;"
End If

CON_blue.CursorLocation = adUseClient
  DoEvents
  DoEvents
  DoEvents

CON_blue.Open
  DoEvents
  DoEvents
  DoEvents


BlueDSN = True
Exit Function
BlueDSN = False
conErr:

If err.Number = "-2147467259 " Then
   mydsn_ = True
   Screen.MousePointer = vbDefault
End If


End Function
Public Function GetCom() As String
Dim sResult As String * 255
    GetComputerNameA sResult, 255
    GetCom = Left$(sResult, InStr(sResult, Chr$(0)) - 1)
    com_name = Left$(sResult, InStr(sResult, Chr$(0)) - 1)
    
    
    Dim lpBuff As String * 25
    Dim ret As Long, UserName As String

     ' Get the user name minus any trailing spaces found in the name.
     ret = GetUserName(lpBuff, 25)
     com_user = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)

End Function
Sub DSN()


Dim FSO As FileSystemObject
Dim f As File
Dim txt As TextStream
Dim matter As String
Dim Total As String
Dim s(1, 2) As String
Set FSO = New FileSystemObject
Dim ss, database_ As String

matter = ""

Dim op_system, Dusername, dstrpath As String

GetCom

If rs.State = 1 Then rs.close
rs.Open "select * from UserDSN where UserName='" & com_name & "' and usrename_='" & com_user & "'", CCON
If rs.EOF = False Then

  dstrpath = rs!Path '& "\chitradsn.dsn"
  Set txt = FSO.CreateTextFile("" & dstrpath)
  Label1.Caption = dstrpath
   
End If

If server_ = "client" Then
  ss = "Server=COMPAQ\SQL2008NEW"
Else
  ss = "Server=BPSERVER"
End If

If Combo1.Text = "2016-17" Then
database_ = "Database=chitraData_1617"
ElseIf Combo1.Text = "2015-16" Then
database_ = "Database = chitraData"
End If


If server_ = "client" Then

    matter = matter & "[ODBC]" & vbNewLine
    matter = matter & "DRIVER=SQL Server" & vbNewLine
    matter = matter & "UId=dinesh" & vbNewLine
    matter = matter & "Trusted_Connection=Yes" & vbNewLine
    matter = matter & "Network=DBMSLPCN" & vbNewLine
    matter = matter & database_ & "" & vbNewLine
    matter = matter & "WSID=COMPAQ" & vbNewLine
    matter = matter & "APP=Microsoft Data Access Components" & vbNewLine
    matter = matter & ss & "" & vbNewLine
    txt.Write matter
    txt.close

    'MsgBox "" & database_
    
Else


  
    matter = matter & "[ODBC]" & vbNewLine
    matter = matter & "DRIVER=SQL Server" & vbNewLine
    matter = matter & "UId=sa" & vbNewLine
    matter = matter & database_ & "" & vbNewLine
    matter = matter & "WSID=COMPAQ" & vbNewLine
    matter = matter & "APP=Microsoft Data Access Components" & vbNewLine
    matter = matter & ss & "" & vbNewLine
    txt.Write matter
    txt.close



End If



End Sub
Private Sub cmdCon_Click()

On Error GoTo conErr



Screen.MousePointer = vbHourglass


If mydsn_ = False Then
    
    DoEvents
    DoEvents
    GetCom
    
    Combo1.Enabled = True
    
    'for client
    
    DSN
    
    'for server
    'sql_pass = "dinesh.123"
    'server_ = "server"
    
    
    Set con = New ADODB.Connection
    
    constr = "filedsn=chitradsn;uid=sa;pwd=" & sql_pass
    con.ConnectionString = constr
    con.CursorLocation = adUseClient
    
    
    If Check1_other.value = 0 Then
    Else
    'NNSERVER\BLUEPRINT
    If server_ = "server" Then
     '  CON.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=NNSERVER\BLUEPRINT; DATABASE=chitradata; UID=sa; PWD=" & sql_pass
    Else
     '  CON.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=COMPAQ\SQL2008NEW; DATABASE=chitradata; UID=; PWD=;"
    End If
    End If
    
    DoEvents
    DoEvents

    con.Open
    
    DoEvents
    DoEvents


End If


'Blue Print-----------------------------------------------------------
'===========================================
If BlueDSN = False Then
   MsgBox "Create DSN...", vbCritical
   Exit Sub
End If

'===========================================

'=====================================================================

updateCon

lblLabels(0).Enabled = True
lblLabels(2).Enabled = True
lblLabels(1).Enabled = True
Combo1.Enabled = True
txtUserName.Enabled = True
txtPassword.Enabled = True
cmdOK.Enabled = True
cmdCancel.Enabled = True


'============================================================
Screen.MousePointer = vbDefault
Me.Caption = "Login Connection is established successfully..."
Check1_other.Visible = False
txtUserName.SetFocus
Exit Sub

conErr:

'MsgBox "" & err.DESCRIPTION

Me.Caption = "Login Connection is not established successfully ..."
Check1_other.Visible = True
Check1_other.value = 1
'clearExeFromProcess
Screen.MousePointer = vbDefault

End Sub
Private Sub cmdOk_Click()

Dim a1, a2 As String


Screen.MousePointer = vbHourglass


If chkdefault.value = 1 Then
CCON.Execute "DELETE FROM MYCHOICE"
CCON.Execute "INSERT INTO MYCHOICE VALUES('" & cbocname.Text & "')"
End If


If Not Trim(Combo1.Text) <> "" Then
MsgBox "Select a Session"
Combo1.SetFocus
Exit Sub
End If

  
   

    If rs.State = 1 Then
        rs.close
    End If
    'rs.Open "select * from UsrePermission where fyear='" & Trim(Combo1.Text) & "' and setupid=" & main.setupid, CON, adOpenKeyset, adLockReadOnly, adCmdText
    rs.Open "select * from UsrePermission where (fyear='" & Trim(Combo1.Text) & "' and USERNAME='" + Trim(Me.txtUserName) + "')", con, adOpenKeyset, adLockReadOnly
    'rs.Find "USERNAME='" + Trim(Me.txtUserName) + "'"
    If Not rs.EOF Then
        If Trim(txtPassword.Text) = Trim(rs!Password) Then
                UId = rs!UserId
                rs.close
                'RS.Open "Select * from setup1 where fyear='" & Trim(Combo1.Text) & "' and setupid=" & main.setupid, CON, adOpenKeyset, adLockReadOnly
                rs.Open "Select * from setup1 where fyear='" & Trim(Combo1.Text) & "' and setupid=2", con, adOpenKeyset, adLockReadOnly
                If rs.EOF = False Then
                
                 
                cname_1 = rs!cname
                cname_2 = rs!add1
                main.setupid = 2
                module_ = Trim(cboModule.Text)
                LoginSucceeded = True
                main.UserName = Trim(Me.txtUserName.Text)
                main.session = Trim(Combo1.Text)
                stringyear = " fyear='" & Trim(Combo1.Text) & "' and setupid=" & main.setupid & " "
                stringyear = " fyear='" & Trim(Combo1.Text) & "' "
                stringyear = stringyear & " and setupid=" & main.setupid & " "
                
                Dim apppath_
                
                Open App.Path + "\client.txt" For Input As #1
                Line Input #1, apppath_
                Close #1

                
                'If main.session = "2015-16" Then
                   rptPath = apppath_ & "\reports"
                'Else
                '   rptPath = App.Path & "\reports"
                'End If
                
                a1 = Left(session, 4) + 1
                a2 = Right(session, 2) + 1
                session_next = a1 & "-" & a2
 
                
                Fill_Module
                
                Loadrights
                CNSetup
                
                ''''''''''''''''' load master on the Client-----------------------
                ''addmaster
                
                'FatchData from Noida=============================================
                If rs1.State = 1 Then rs1.close
                rs1.Open "select FatchOrderFromNoida from setup1 where " & stringyear, con
                If rs1.EOF = False Then
                If (Not IsNull(rs1!FatchOrderFromNoida) And rs1!FatchOrderFromNoida = "y") Then
                    Set CON_noida = New ADODB.Connection
                    constr = "filedsn=blueacc;"
                    CON_noida.ConnectionString = constr
                    CON_noida.CursorLocation = adUseClient
                    CON_noida.Open
                End If
                
                End If
               '=============================================================
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Screen.MousePointer = vbDefault
                
                Unload Me
                Else
                MsgBox "Please contact to the Vendor."
                End If
        Else
            MsgBox "Invalid Password, try again!"
            txtPassword.SetFocus
            HIT
        End If
    Else
        MsgBox "USER  NOT FOUND"
        Me.txtUserName.SetFocus
        HIT
    End If
    
Screen.MousePointer = vbDefault
    
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Me.txtUserName.Enabled = True Then
      Me.txtUserName.SetFocus
    Else
      cmdCon.SetFocus
    End If
    KeyAscii = 0
End If
End Sub
Private Sub Combo1_LostFocus()
'==============
ConOpen      ' dataconfigure File

If rs.State = 1 Then rs.close
rs.Open "select rptpath from data", CCON
If rs.EOF = False Then
   server_ = rs(0)
End If


If rs.State = 1 Then rs.close
rs.Open "select WSID from data", CCON
If rs.EOF = False Then
   sql_pass = rs(0)
End If

'--------------------
End Sub

Private Sub Command1_Click()
GetCom
End Sub

 Private Sub Form_Load()




'If RS.State = 1 Then RS.close
'RS.Open "select fyear from financialyear order by fyear desc", CCON
'While RS.EOF = False
'   Combo1.AddItem RS(0)
'RS.MoveNext
'Wend

Combo1.ListIndex = 1


Me.Height = Me.Height - 650



BackColorFrom Me


lblLabels(0).Enabled = False
lblLabels(2).Enabled = False
lblLabels(1).Enabled = False
Combo1.Enabled = False
txtUserName.Enabled = False
txtPassword.Enabled = False
cmdOK.Enabled = False
cmdCancel.Enabled = False

Combo1.Enabled = True

End Sub
Sub updateCon()





hindi = "kruti dev 010"
english = "Times New Roman"

'-----------------------------------------
On Error Resume Next

con.Execute "ALTER TABLE tempLedger1 add UserId nvarchar(50)"
con.Execute "ALTER TABLE ordera add Godown nvarchar(50)"
con.Execute "alter table tempLedger1 add rptid int"
con.Execute "alter table tempLedger1 add rptype nvarchar(50)"
con.Execute "alter table CREDITB add BookDesc nvarchar(300)"
con.Execute "alter table ORDERA add sale_sp nvarchar(10)"

con.Execute "ALTER TABLE DonnationMain add AdvAmt float"
con.Execute "ALTER TABLE DonnationMain add RoundOfAAmt float"




'CON.Execute "ALTER TABLE ORDERB add DISCOUNT float"
'CON.Execute "ALTER TABLE ORDERB add billNo nvarchar(20)"
'CON.Execute "ALTER TABLE ORDERa add partyname nvarchar(50)"

''CON.Execute "ALTER TABLE sledger add Transport nvarchar(30)"
''CON.Execute "ALTER TABLE TmpBook ALTER column BName nvarchar(100)"
''CON.Execute "ALTER TABLE TmpBook add Area nvarchar(500)"
''CON.Execute "ALTER TABLE invoicea add Remarks nvarchar(100)"
''CON.Execute "ALTER TABLE BOOKS_KIT add Apply nvarchar(1)"
''CON.Execute "ALTER TABLE PackinkSlipA add OrderNo int"
''CON.Execute "ALTER TABLE sledger add pin nvarchar(25)"
''CON.Execute "ALTER TABLE ORDERA  add party_state nvarchar(50)"
''CON.Execute "ALTER TABLE ORDERA add party_dist nvarchar(100)"
''CON.Execute "ALTER TABLE ORDERA add shipto_dist nvarchar(100)"
''CON.Execute "ALTER TABLE BOOKS_KIT add Qty nvarchar(5)"
''CON.Execute "ALTER TABLE ordera add Shipto_CityId nvarchar(7)"
''CON.Execute "ALTER TABLE ordera add Shipto nvarchar(100)"
''CON.Execute "ALTER TABLE ordera add Shipto_Add1 nvarchar(60)"
''CON.Execute "ALTER TABLE ordera add Shipto_Add2 nvarchar(60)"
''CON.Execute "ALTER TABLE ordera add Shipto_City nvarchar(50)"
''CON.Execute "ALTER TABLE ordera add Shipto_district nvarchar(50)"
''CON.Execute "ALTER TABLE ordera add Shipto_States nvarchar(50)"
''CON.Execute "ALTER TABLE ordera add RepName nvarchar(50)"
''CON.Execute "ALTER TABLE ordera add ScName nvarchar(200)"
''CON.Execute "ALTER TABLE ordera add ScID nvarchar(6)"
''CON.Execute "ALTER TABLE orderb add Bquantity nvarchar(5)"
''CON.Execute "ALTER TABLE invoiceA add Shipto_Scholl nvarchar(1)"
''CON.Execute "ALTER TABLE CNFA ALTER column N nvarchar(150)"
''CON.Execute "ALTER TABLE DNFA ALTER column N nvarchar(150)"
''CON.Execute "ALTER TABLE DNFA add Amtwords nvarchar(100)"
''CON.Execute "ALTER TABLE SLEDGER add PAN nvarchar(12)"
''CON.Execute "ALTER TABLE BookStock add BAuthorized bit"
''CON.Execute "ALTER TABLE INVOICEA add Shipto nvarchar(100)"
''CON.Execute "ALTER TABLE INVOICEA add ShiptoAdd nvarchar(200)"
''CON.Execute "ALTER TABLE INVOICEA add ScName nvarchar(200)"
''CON.Execute "ALTER TABLE INVOICEA add ScID nvarchar(6)"
''CON.Execute "ALTER TABLE INVOICEC_sp add UserName nvarchar(30)"
''CON.Execute "ALTER TABLE INVOICECtmp_sp add UserName nvarchar(30)"
''CON.Execute "ALTER TABLE INVOICEC add UserName nvarchar(30)"
''CON.Execute "ALTER TABLE INVOICECtmp add UserName nvarchar(30)"
''CON.Execute "ALTER TABLE BookStock add BookDesc nvarchar(300)"



On Error GoTo 0
'-----------------------------------------


cboModule.ListIndex = 0
cbocname.Clear

Set rs = New ADODB.Recordset
rs.Open "select * from setup1 order by cname asc", con, adOpenStatic, adLockReadOnly, adCmdText
If rs.EOF = False Then
    Do While Not rs.EOF
        cbocname.AddItem rs!cname
        If Not rs.EOF Then
            rs.MoveNext
        End If
    Loop
End If

If CCON.State = adStateClosed Then CCON.Open
If rs.State <> adStateClosed Then rs.close
rs.Open "SELECT * FROM MYCHOICE", CCON, adOpenKeyset, adLockOptimistic
If rs.EOF = True Then
   cbocname.ListIndex = 0
Else
   cbocname.Text = rs!cname
End If
''fillsession

End Sub
Private Sub txtPassword_GotFocus()
txtPassword.SelLength = 10
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdOK.SetFocus
        KeyAscii = 0
    End If
End Sub
Private Sub txtUserName_GotFocus()
txtUserName.SelLength = 15
End Sub
Private Sub txtUserName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtPassword.SetFocus
        KeyAscii = 0
    End If
End Sub



