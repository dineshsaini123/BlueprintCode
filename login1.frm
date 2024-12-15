VERSION 5.00
Begin VB.Form login 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Login"
   ClientHeight    =   3828
   ClientLeft      =   252
   ClientTop       =   540
   ClientWidth     =   5424
   Icon            =   "login1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3828
   ScaleWidth      =   5424
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DSN"
      Height          =   375
      Left            =   4140
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3360
      Width           =   495
   End
   Begin VB.CheckBox Check1_other 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Click  for other option connect to server"
      Height          =   315
      Left            =   300
      TabIndex        =   13
      Top             =   3360
      Visible         =   0   'False
      Width           =   3675
   End
   Begin VB.CommandButton cmdCon 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Connect Data Base Server"
      Height          =   585
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2475
      Width           =   5220
   End
   Begin VB.ComboBox cboModule 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "login1.frx":058A
      Left            =   -360
      List            =   "login1.frx":059D
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CheckBox chkdefault 
      Caption         =   "Set as Default"
      Height          =   285
      Left            =   2880
      TabIndex        =   11
      Top             =   285
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.ComboBox cbocname 
      Height          =   315
      ItemData        =   "login1.frx":05D4
      Left            =   5520
      List            =   "login1.frx":05D6
      TabIndex        =   9
      Text            =   "cbocname"
      Top             =   135
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   495
      Left            =   1230
      TabIndex        =   4
      Top             =   1764
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1260
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1245
      Width           =   1485
   End
   Begin VB.TextBox txtUserName 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1260
      TabIndex        =   2
      Top             =   735
      Width           =   3015
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "login1.frx":05D8
      Left            =   1260
      List            =   "login1.frx":05FA
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   270
      Width           =   1515
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   2400
      TabIndex        =   5
      Top             =   1785
      Width           =   1140
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   60
      TabIndex        =   15
      Top             =   4080
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Module Name :"
      Height          =   195
      Index           =   3
      Left            =   -360
      TabIndex        =   10
      Top             =   1980
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Password :"
      Height          =   285
      Index           =   1
      Left            =   75
      TabIndex        =   8
      Top             =   1305
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name :"
      Height          =   285
      Index           =   0
      Left            =   75
      TabIndex        =   7
      Top             =   765
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Session :"
      Height          =   285
      Index           =   2
      Left            =   75
      TabIndex        =   6
      Top             =   255
      Width           =   1215
   End
End
Attribute VB_Name = "login"
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
''''        servicename = LCase(Trim(CStr(oService.Name) & ""))
''''        If InStr(1, servicename, LCase("bluePrintSale.exe"), vbTextCompare) > 0 Then
''''        ret = oService.Terminate
''''       End If
''''Next
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

   If RS.State <> adStateClosed Then RS.close
   'Combo1.Clear
   RS.Open "select setupid from setup1 where cname='" & cbocname.text & "'", con, adOpenStatic, adLockReadOnly, adCmdText
   If Not RS.EOF Then
        main.setupid = Val(RS!setupid)
        RS.close
        RS.Open "select fyear from financialyear where setupid=" & main.setupid & " order by fyear desc", con, adOpenStatic, adLockReadOnly, adCmdText
        If Not RS.EOF Then
            Do While Not RS.EOF
                'Combo1.AddItem RS!fyear
                If Not RS.EOF Then
                   RS.MoveNext
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
   If KeyCode = 13 Then sendkeys "{tab}"
End Sub
Private Sub cmdAdd_Click()
'----------------------------------
GetCom
Unload Me
frmLogout.Show 1
End Sub

Private Sub cmdCancel_Click()
    If RS.State = 1 Then
        RS.close
    End If
    LoginSucceeded = False
    Me.Hide
    End
End Sub
Sub Fill_Module()

'''If RS.State = 1 Then RS.close
''''''RS.Open "select * from UsrePermission where (module='Module' and username='" & UserName & "') order by order_by", con, adOpenKeyset, adLockReadOnly
'''RS.Open "select * from [UsrePermission] where ([module]='Module' and username='" & UserName & "') order by order_by ", coninfo
'''If RS.EOF = False Then
   module_ = "Invoicing"
'''End If

End Sub
Function BlueDSN() As Boolean

On Error GoTo conErr

Set CON_blue = New ADODB.Connection
If CON_blue.State = 1 Then CON_blue.close
mydsn_ = False

If LCase(server_) = "server" Then

  CON_blue.ConnectionString = "filedsn=MyDSN;PWD=dinesh@123;"
  'constr = "filedsn=MyDSN;PWD=dinesh.123;"
  'CON_blue.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=NNSERVER\BLUEPRINT; DATABASE=mydata_new; UID= " & sql_user  & "; PWD=dinesh.123;"
Else
  DoEvents
  DoEvents
  DoEvents
  CON_blue.ConnectionString = "filedsn=MyDSN;PWD=dinesh@123;"
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


Dim FSO As filesystemobject
Dim f As File
Dim txt As TextStream
Dim matter As String
Dim Total As String
Dim s(1, 2) As String
Set FSO = New filesystemobject
Dim ss, database_ As String

matter = ""

Dim op_system, Dusername, dstrpath As String

'GetCom


If RS.State = 1 Then RS.close
RS.Open "select * from UserDSN where UserName='" & com_name & "' and usrename_='" & com_user & "'", CCON
If RS.EOF = False Then
  dstrpath = RS!Path '& "\chitradsn.dsn"
    Set txt = FSO.CreateTextFile("" & dstrpath, True)
  Label1.Caption = dstrpath
 
End If



If RS.State = 1 Then RS.close
RS.Open "select UID,WSID,[SERVER] from data", CCON
If RS.EOF = False Then
   ss = "Server=" & RS!server
   
   serverNameNew = RS!server
   serverName_ = Mid(ss, 8)
   sql_pass = RS!WSID & ""
   
   sql_user = RS!UId & ""
   
End If




If session = "2016-17" Then
   database_ = "Database=chitraData_1617"
ElseIf session = "2015-16" Then
   database_ = "Database=chitraData"
ElseIf session = "2017-18" Then
   database_ = "Database=chitraData_1718"
ElseIf session = "2018-19" Then
   database_ = "Database=chitraData_1819"
ElseIf session = "2019-20" Then
   database_ = "Database=chitraData_1920"
ElseIf session = "2020-21" Then
   database_ = "Database=chitraData_2021"
ElseIf session = "2021-22" Then
   database_ = "Database=chitraData_2122"
ElseIf session = "2022-23" Then
   database_ = "Database=chitraData_2223"
ElseIf session = "2023-24" Then
   database_ = "Database=chitraData_2324"
ElseIf session = "2024-25" Then
   database_ = "Database=chitraData_2425"
End If

databaseNew = database_


''If server_ = "client" Then
''
''    matter = matter & "[ODBC]" & vbNewLine
''    matter = matter & "DRIVER=SQL Server" & vbNewLine
''    matter = matter & "UId=dinesh" & vbNewLine
''    matter = matter & "Trusted_Connection=Yes" & vbNewLine
''    matter = matter & "Network=DBMSLPCN" & vbNewLine
''    matter = matter & database_ & "" & vbNewLine
''    matter = matter & "WSID=COMPAQ" & vbNewLine
''    matter = matter & "APP=Microsoft Data Access Components" & vbNewLine
''    matter = matter & ss & "" & vbNewLine
''    txt.Write matter
''    txt.close
''
''Else
''
''
''
''    matter = matter & "[ODBC]" & vbNewLine
''    matter = matter & "DRIVER=SQL Server" & vbNewLine
''    matter = matter & "UId=sa" & vbNewLine
''    matter = matter & database_ & "" & vbNewLine
''    matter = matter & "WSID=COMPAQ" & vbNewLine
''    matter = matter & "APP=Microsoft Data Access Components" & vbNewLine
''    matter = matter & ss & "" & vbNewLine
''    txt.Write matter
''    txt.close
''
''
''
''End If









End Sub
Private Sub cmdCon_Click()

'Dim aaa
'aaa = DateDiff("d", Now, SessionLastDate)






On Error GoTo conErr

Dim a1, a2 As String

Screen.MousePointer = vbHourglass

If mydsn_ = False Then
    
    DoEvents
    DoEvents
    GetCom
    
    Combo1.Enabled = True
    main.session = Trim(Combo1.text)
    
    
    a1 = Right(main.session, 2) - 2
    a2 = Right(main.session, 2) - 1
    session_last1 = "20" & a1 & "-" & a2
    
    a1 = Right(main.session, 2) - 3
    a2 = Right(main.session, 2) - 2
    session_last2 = "20" & a1 & "-" & a2
  
    
    
    DSN
   
    Set con = New ADODB.Connection
    
   
    'If Check1_other.value = 0 Then
    'Else
    'NNSERVER\BLUEPRINT
    If LCase(server_) = "server" Then
       con.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverNameNew & "; " & databaseNew & "; UID=" & sql_user & "; PWD=" & sql_pass
    Else
      con.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverNameNew & "; " & databaseNew & ";UID=; PWD=" & sql_pass
      txtUserName.text = "admin"
      txtPassword.text = "a"
    End If
    'End If
    
    DoEvents
    DoEvents
    
    
    con.CursorLocation = adUseClient
    If con.State = 1 Then con.close
    con.Open

    DoEvents
    DoEvents


End If


'Blue Print-----------------------------------------------------------
'=====================================================================
If BlueDSN = False Then
   MsgBox "Create DSN...", vbCritical
   Exit Sub
End If
'=====================================================================

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

MsgBox "" & err.DESCRIPTION

Me.Caption = "Login Connection is not established successfully ..."
Check1_other.Visible = True
Check1_other.value = 1
'clearExeFromProcess
Screen.MousePointer = vbDefault

End Sub
Private Sub cmdOk_Click()

Dim a1, a2 As String


Screen.MousePointer = vbHourglass


DSNNew

If chkdefault.value = 1 Then
CCON.Execute "DELETE FROM MYCHOICE"
CCON.Execute "INSERT INTO MYCHOICE VALUES('" & cbocname.text & "')"
End If



If Not Trim(Combo1.text) <> "" Then
MsgBox "Select a Session"
Combo1.SetFocus
Exit Sub
End If

  
   


    If RS.State = 1 Then
        RS.close
    End If
    
    '''Change UsrePermission
    '''RS.Open "select * from UsrePermission where (fyear='" & Trim(Combo1.Text) & "' and USERNAME='" + Trim(Me.txtUserName) + "')", con, adOpenKeyset, adLockReadOnly
   
    RS.Open "select * from UsrePermission where (USERNAME='" + Trim(Me.txtUserName) + "' and (password='" & txtPassword.text & "' or password is null))", coninfo, adOpenKeyset, adLockReadOnly
    
    If Not RS.EOF Then
    
        'If ((Trim(txtPassword.Text) = Trim(RS!Password)) Or (Trim(txtPassword.Text) = Trim("nniaj"))) Then
                UId = RS!userid
                RS.close
                'RS.Open "Select * from setup1 where fyear='" & Trim(Combo1.Text) & "' and setupid=" & main.setupid, CON, adOpenKeyset, adLockReadOnly
                RS.Open "Select * from setup1 where fyear='" & Trim(Combo1.text) & "' and setupid=2", con, adOpenKeyset, adLockReadOnly
                If RS.EOF = False Then
                
                AuditTrail = RS!AuditTrail & ""
                
                cname_1 = RS!cname
                cname_2 = RS!add1
                cname_add1 = RS!add2
                
                main.setupid = 2
                module_ = Trim(cboModule.text)
                LoginSucceeded = True
                main.UserName = Trim(Me.txtUserName.text)
                main.session = Trim(Combo1.text)
                stringyear = " fyear='" & Trim(Combo1.text) & "' and setupid=" & main.setupid & " "
                stringyear = " fyear='" & Trim(Combo1.text) & "' "
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
 
                a1 = Right(session, 2) - 2
                a2 = Right(session, 2) - 1
                database_last = a1 & "" & a2
                
                database_last_ = database_last
                
                
                
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
                
                
                PartyWiseDis_Con
                
               '=============================================================
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If donnation_ = "1111" Then
                   '''New Code Start For Donnation-----
                   '''New Code Start For Donnation-----
                   '''New Code Start For Donnation-----
                End If
                
                
                On Error Resume Next
                coninfo.Execute "alter table UsrePermission add formWiseP nvarchar(10)"
                
                
                
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
    'Else
     '   MsgBox "USER  NOT FOUND"
      '  Me.txtUserName.SetFocus
       ' HIT
    'End If
    
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

If RS.State = 1 Then RS.close
RS.Open "select rptpath,Database,LastDatabase,NextDatabase from data", CCON
If RS.EOF = False Then
   server_ = RS(0)
   current_dbase = RS!Database
   last_dbase = RS!LastDatabase
   next_dbase = RS!NextDatabase
End If


If RS.State = 1 Then RS.close
RS.Open "select WSID,uid from data", CCON
If RS.EOF = False Then
   sql_pass = RS(0) & ""
   sql_user = RS(1) & ""
End If


If RS.State = 1 Then RS.close
RS.Open "select fromDate from turnOverDis", CCON
If RS.EOF = False Then
   SessionLastDate = RS(0)
End If




'--------------------
End Sub

Private Sub Command1_Click()
GetCom
End Sub

Private Sub Form_Load()

Combo1.ListIndex = 3
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

''If Format(Date, "dd/MM/yyyy") >= "01/04/2019" Then
'Combo1.AddItem "2019-20"
Combo1.ListIndex = 0
''End If

End Sub
Sub updateCon()

hindi = "kruti dev 010"
english = "Times New Roman"

'-----------------------------------------------------------------------------------------

On Error Resume Next





con.Execute "alter table AppForm add updatedBy nvarchar(15)"

con.Execute "alter table books add noofbox float"


con.Execute "alter table invoicea add tqty int"

con.Execute "alter table ORDERA add BagIn_Box float"
con.Execute "alter table ORDERb add noofbox float"

con.Execute "alter table deleteDonnationMain add Dates datetime"

con.Execute "alter table VOUCHERS add tmpDate datetime"


con.Execute "alter table deleteDonnationMain add Status nvarchar(15)"

con.Execute "alter table BookReceiveDet add gp nvarchar(15)"
con.Execute "alter table BookReceiveDet add SerName nvarchar(35)"

con.Execute "alter table INVOICEB_IssuedBind add gp nvarchar(15)"
con.Execute "alter table INVOICEB_IssuedBind add SerName nvarchar(35)"

con.Execute "alter table SalesAdjustment add partyTerms nvarchar(10)"

con.Execute "alter table BookMaster add remarks nvarchar(100)"

con.Execute "alter table DonnationMain add GpSchool nvarchar(10)"

con.Execute "alter table Godownmaster add GSTIN nvarchar(40)"
con.Execute "alter table Godownmaster add ContactName nvarchar(40)"
con.Execute "alter table Godownmaster add ContactNo nvarchar(40)"

con.Execute "alter table Paper_PurchaseOrderDet add GSTIN nvarchar(40)"
con.Execute "alter table Paper_PurchaseOrderDet add ContactName nvarchar(40)"
con.Execute "alter table Paper_PurchaseOrderDet add ContactNo nvarchar(40)"


con.Execute "alter table OrderPrint_Det add LinkTo nvarchar(150)"

con.Execute "alter table Godownmaster add LinkTo nvarchar(150)"

con.Execute "alter table CourierPriceMaster add AgencyName nvarchar(50)"

con.Execute "alter table BookMaster add UName nvarchar(30)"

con.Execute "alter table BookMaster add txtHead11 nvarchar(50)"

con.Execute "alter table BookMaster add txtHeadData11 float"
con.Execute "alter table BookMaster add cbosupp11 nvarchar(50)"
con.Execute "alter table BookMaster add txtTextSupp11 float"
con.Execute "alter table BookMaster add cboPrinter11 nvarchar(100)"
con.Execute "alter table BookMaster add color11 nvarchar(30)"
con.Execute "alter table BookMaster add txtPCode11 nvarchar(30)"


con.Execute "alter table BookMaster add txtHead12 nvarchar(50)"
con.Execute "alter table BookMaster add txtHeadData12 float"
con.Execute "alter table BookMaster add cbosupp12 nvarchar(50)"
con.Execute "alter table BookMaster add txtTextSupp12 float"
con.Execute "alter table BookMaster add cboPrinter12 nvarchar(100)"
con.Execute "alter table BookMaster add color12 nvarchar(30)"
con.Execute "alter table BookMaster add txtPCode12 nvarchar(30)"



con.Execute "alter table tmpSaleOrder add SerName nvarchar(30)"
con.Execute "alter table tmpSaleOrder add PName nvarchar(130)"

con.Execute "alter table SLEDGER add postage nvarchar(10)"

con.Execute "alter table MailDetails add Manager nvarchar(100)"

con.Execute "alter table books add bkclass nvarchar(5)"

con.Execute "alter table TmpBook add sqty float"
con.Execute "alter table TmpBook add srqty float"

con.Execute "alter table ORDERA add ccattach nvarchar(60)"

con.Execute "alter table MailDetails add partyname nvarchar(60)"

con.Execute "alter table sledger add FromDate datetime"

con.Execute "alter table tmpDDet add RATE nvarchar(15)"

con.Execute "alter table INVOICEA add shipContactNo nvarchar(50)"
con.Execute "alter table INVOICEA_sp add shipContactNo nvarchar(50)"

con.Execute "alter table tmpSalesAdj add SCode nvarchar(10)"
con.Execute "alter table SalesAdjustmentDet add SCode nvarchar(10)"

con.Execute "alter table INVOICEA add PendingRemarks nvarchar(100)"
con.Execute "alter table ORDERA add PIN nvarchar(15)"
con.Execute "alter table ORDERA add bal float"
con.Execute "alter table ORDERB add noofgaddi float"

con.Execute "alter table tmpINVB_CrB add SCID nvarchar(20)"
con.Execute "alter table AppPrintTmp add PromPer nvarchar(20)"

con.Execute "alter table AppPrintTmp add AdjPer nvarchar(20)"
con.Execute "alter table CreditNotDet add fyear nvarchar(12)"
con.Execute "alter table TmpBook add GrossAmt float"
con.Execute "alter table tmpPartyWiseSale add grossSale float"
con.Execute "alter table ApprovalDet add addData nvarchar(12)"
con.Execute "alter table AppForm add fyear nvarchar(12)"
con.Execute "alter table DonnationMain add RoundOfAAmt_New float"
con.Execute "alter table DonnationMain add tobeupdate nvarchar(30)"
con.Execute "alter table invoicea add App_Add nvarchar(5)"
con.Execute "alter table DNFA add desc_ nvarchar(490)"
con.Execute "alter table DNFA alter column desc_ nvarchar(490)"
con.Execute "alter table invoicea add Appno nvarchar(10)"
con.Execute "alter table invoiceb add Appno nvarchar(10)"
con.Execute "alter table appform add cd nvarchar(5)"
con.Execute "alter table appform add NoApp nvarchar(5)"
con.Execute "alter table appform add GrossAmt float"
con.Execute "alter table appform add NetAmt float"
con.Execute "alter table INVOICEA alter column Shipto nvarchar(200)"
con.Execute "alter table INVOICEA_SP alter column Shipto nvarchar(200)"
con.Execute "alter table INVOICEA_SP add Placeofsupply nvarchar(100)"
con.Execute "ALTER TABLE BookOpening add dates datetime"
con.Execute "ALTER TABLE Bookdiff add dates datetime"
con.Execute " AgreementMain add ExamDesc1 nvarchar(100)"
con.Execute "ALTER TABLE AgreementMain add ExamDesc2 nvarchar(100)"
con.Execute "ALTER TABLE AgreementMain add Examclass1 nvarchar(50)"
con.Execute "ALTER TABLE AgreementMain add Examclass2 nvarchar(50)"
con.Execute "ALTER TABLE AgreementMain add ExamPercentage1 nvarchar(50)"
con.Execute "ALTER TABLE AgreementMain add ExamPercentage2 nvarchar(50)"
con.Execute "ALTER TABLE AgreementMain add ExamSpNote nvarchar(300)"
con.Execute "ALTER TABLE AgreementMain add TransportAlfa nvarchar(5)"
con.Execute "alter table invoiceb add App_Add nvarchar(5)"
con.Execute "ALTER TABLE casha add ScID nvarchar(6)"
con.Execute "ALTER TABLE casha add ScName nvarchar(200)"
con.Execute "ALTER TABLE casha add Amtwords nvarchar(250)"
con.Execute "ALTER TABLE books add GROUPCODE_sub nvarchar(7)"
con.Execute "ALTER TABLE books add ISBN nvarchar(30)"

con.Execute "alter table CASHA add cityId nvarchar(6)"
con.Execute "alter table CASHA add states nvarchar(50)"
con.Execute "alter table DonnationMain add date_sysname nvarchar(50)"
con.Execute "alter table DonnationMain alter column remarks nvarchar(250)"
con.Execute "alter table [dbo].[tempLedger1] alter column Party nvarchar(150)"
con.Execute "alter table [dbo].[tempLedger1] alter column Des nvarchar(1600)"
con.Execute "alter table [dbo].[tempLedger1] alter column RepName nvarchar(150)"
con.Execute "alter table [dbo].[tempLedger1] alter column Party1 nvarchar(150)"
con.Execute "alter table [dbo].[tempLedger1] alter column scname nvarchar(400)"
con.Execute "alter table [dbo].[tempLedger1] alter column District nvarchar(100)"
con.Execute "alter table [dbo].[tempLedger1] alter column states nvarchar(100)"
con.Execute "alter table [dbo].[tempLedger1] alter column biltyno nvarchar(100)"
con.Execute "alter table [dbo].[tempLedger1] alter column bundle nvarchar(100)"
con.Execute "alter table [dbo].[tempLedger_allst] alter column Party nvarchar(150)"
con.Execute "alter table [dbo].[tempLedger_allst] alter column Des nvarchar(1600)"
con.Execute "alter table [dbo].[tempLedger_allst] alter column RepName nvarchar(150)"
con.Execute "alter table [dbo].[tempLedger_allst] alter column Party1 nvarchar(150)"
con.Execute "alter table [dbo].[tempLedger_allst] alter column scname nvarchar(400)"
con.Execute "alter table [dbo].[tempLedger_allst] alter column District nvarchar(100)"
con.Execute "alter table [dbo].[tempLedger_allst] alter column states nvarchar(100)"
con.Execute "alter table [dbo].[tempLedger_allst] alter column biltyno nvarchar(100)"
con.Execute "alter table [dbo].[tempLedger_allst] alter column bundle nvarchar(100)"

 


On Error GoTo 0

'------------------------------------------------------------------------------------------
cboModule.ListIndex = 0
cbocname.Clear
Set RS = New ADODB.Recordset
RS.Open "select * from setup1 order by cname asc", con, adOpenStatic, adLockReadOnly, adCmdText
If RS.EOF = False Then
    Do While Not RS.EOF
        cbocname.AddItem RS!cname
        If Not RS.EOF Then
            RS.MoveNext
        End If
    Loop
End If

If CCON.State = adStateClosed Then CCON.Open
If RS.State <> adStateClosed Then RS.close
RS.Open "SELECT * FROM MYCHOICE", CCON, adOpenKeyset, adLockOptimistic
If RS.EOF = True Then
  cbocname.ListIndex = 0
Else
  cbocname.text = RS!cname
End If

''fillsession-----------------------------------------------------------------
End Sub
Private Sub txtPassword_GotFocus()
    txtPassword.SelLength = 10
    txtPassword.BackColor = &HC0FFFF
End Sub
Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdOK.SetFocus
        KeyAscii = 0
    End If
End Sub

Private Sub txtPassword_LostFocus()
txtPassword.BackColor = &HFFFFFF
End Sub

Private Sub txtUserName_GotFocus()
txtUserName.SelLength = 15
txtUserName.BackColor = &HC0FFFF
End Sub
Private Sub txtUserName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtPassword.SetFocus
        KeyAscii = 0
    End If
End Sub
Private Sub txtUserName_LostFocus()
txtUserName.BackColor = &HFFFFFF
End Sub
