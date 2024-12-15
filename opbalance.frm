VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmopbalance 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6495
   ClientLeft      =   2040
   ClientTop       =   2040
   ClientWidth     =   8940
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "opbalance.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optsub 
      Caption         =   "Subledger"
      Height          =   285
      Left            =   1440
      TabIndex        =   8
      Top             =   600
      Width           =   1365
   End
   Begin VB.OptionButton optgen 
      Caption         =   "Genledger"
      Height          =   285
      Left            =   60
      TabIndex        =   7
      Top             =   600
      Value           =   -1  'True
      Width           =   1365
   End
   Begin VB.ComboBox Combogenledgerdesc 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1440
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   930
      Width           =   3165
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   345
      Left            =   7020
      TabIndex        =   2
      Top             =   1170
      Width           =   1785
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   5610
      Top             =   1710
      Visible         =   0   'False
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   582
      ConnectMode     =   1
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSMask.MaskEdBox textsearch 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   1290
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      PromptChar      =   "_"
   End
   Begin MSDataGridLib.DataGrid Grid1 
      Bindings        =   "opbalance.frx":000C
      Height          =   4845
      Left            =   60
      Negotiate       =   -1  'True
      TabIndex        =   3
      Top             =   1620
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   8546
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAcrossSplits =   -1  'True
      TabAction       =   2
      WrapCellPointer =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label17 
      Caption         =   "Gen.ledger Desc."
      Height          =   225
      Left            =   60
      TabIndex        =   6
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Year Opening Balance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   90
      TabIndex        =   4
      Top             =   90
      Width           =   8775
   End
   Begin VB.Label Label1 
      Caption         =   "Search String"
      Height          =   225
      Left            =   60
      TabIndex        =   1
      Top             =   1320
      Width           =   1305
   End
End
Attribute VB_Name = "frmopbalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim CON As ADODB.Connection
Dim rs As ADODB.Recordset
Dim tablabel As Integer
Dim masterlabel As String
Dim maxrow As Integer
Dim ctl As Control
Dim intselcol As Integer
Dim blnsortacr As Boolean
Dim strsortorder As String

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub Combogenledgerdesc_Click()
If optsub.Value = True Then
Me.tempr 0, "subledger"
End If
End Sub



Private Sub Form_Load()
    Set rs = New ADODB.Recordset
    DATA1.ConnectionString = CON
    Me.Left = 145
    Me.TOP = 600
    intselcol = 0
    blnsortacr = True
    strsortorder = " asc"
    If rs.State = adStateOpen Then rs.Close
    rs.Open "select * from gledger where slf=1 and " & stringyear, CON, adOpenDynamic, adLockOptimistic, adCmdText
    If Not rs.EOF Then
        Do While Not rs.EOF
            Combogenledgerdesc.AddItem rs!gledger
            If Not rs.EOF Then
                rs.MoveNext
            End If
        Loop
    End If
    rs.Close
    Combogenledgerdesc.ListIndex = 0
    tempr 0, "genledger"
    End Sub

Private Sub Grid1_GotFocus()
On Error Resume Next
'For I = 0 To DATA1.Recordset.Fields.Count - 1
      Grid1.Columns(0).Locked = True
'Next I
End Sub

Function tempr(tb As Integer, master As String)
    tablabel = tb
    masterlabel = master
    If tb = 0 And Trim(UCase(master)) = Trim(UCase("genledger")) Then
         DATA1.RecordSource = "select gledger,yearopening from gledger where  " & stringyear & " and slf=0 order by gledger"
         DATA1.Refresh
         Grid1.ReBind
         Grid1.Columns(1).NumberFormat = "0.00"
         Grid1.Columns(0).Width = 5000
         Grid1.Columns(1).Width = 2000
    ElseIf tb = 0 And Trim(UCase(master)) = Trim(UCase("subledger")) Then
         DATA1.RecordSource = "select subledger,yearopening from sledger where  " & stringyear & " and gledger='" & Combogenledgerdesc.Text & "' order by subledger"
         DATA1.Refresh
         Grid1.ReBind
         Grid1.Columns(1).NumberFormat = "0.00"
         Grid1.Columns(0).Width = 5000
         Grid1.Columns(1).Width = 2000
    End If
    'Me.Show
    'textsearch.SetFocus
End Function

Private Sub PgCntFootBeg1_GotFocus()

End Sub



Private Sub optgen_Click()
If optgen.Value = True Then
Combogenledgerdesc.Enabled = False
Me.tempr 0, "genledger"
End If
End Sub

Private Sub optsub_Click()
If optsub.Value = True Then
Combogenledgerdesc_Click
Combogenledgerdesc.Enabled = True
End If
End Sub

Private Sub textsearch_Change()
    'Data1.Refresh
    If tablabel = 0 And Trim(UCase(masterlabel)) = UCase("") Then
        Grid1.col = 1
        sets = True
    End If
End Sub
Private Sub textsearch_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        DATA1.Refresh
        Grid1.Columns(0).Width = 5000
        Grid1.Columns(1).Width = 2000
        Grid1.col = 0
        sets = True
        DATA1.Recordset.Sort = DATA1.Recordset.Fields(Grid1.col).Name
        
       Dim strsearch As String
       If textsearch.Text <> "" Then
           If DATA1.Recordset.RecordCount > 0 Then
              Select Case DATA1.Recordset.Fields(Grid1.col).Type
              Case 3, 5, 17: 'number
                If IsNumeric(textsearch.Text) = False Then
                MsgBox "Invalid Number."
                textsearch.SetFocus
                Exit Sub
                End If
                strsearch = " like " + Trim(textsearch.Text) + ""
              Case 202: 'text
                strsearch = " like '" + Trim(textsearch.Text) + "*'"
              Case 135: 'date
                If IsDate(textsearch.Text) = False Then
                MsgBox "Invalid Date."
                textsearch.SetFocus
                Exit Sub
                End If
                strsearch = " like '" + Trim(textsearch.Text) + "'"
              Case 11: 'boolean
              strsearch = " = " + Trim(textsearch.Text) + ""
              Case Else:
              strsearch = " = '" + Trim(textsearch.Text) + "'"
              End Select
              
              If IsNumeric(Trim(textsearch.Text)) Then
                  DATA1.Recordset.Find Trim(DATA1.Recordset.Fields(Grid1.col).Name) + strsearch
              Else
                  DATA1.Recordset.Find Trim(DATA1.Recordset.Fields(Grid1.col).Name) + strsearch
              End If
              
              If DATA1.Recordset.AbsolutePosition > 0 Then
                  a = DATA1.Recordset.AbsolutePosition
                  DATA1.Recordset.AbsolutePosition = a
              End If
           End If
              Grid1.SetFocus
              'For I = 0 To DATA1.Recordset.Fields.Count - 1
                   Grid1.Columns(0).Locked = True
              'Next I
              'Grid1.Columns(1).Locked = True
              'SendKeys ("{LEFT}")
              Grid1.SetFocus
              
           End If
    End If
End Sub

