VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form GRIDpreview 
   Caption         =   "GRIDpreview"
   ClientHeight    =   6792
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   9480
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.6
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6792
   ScaleWidth      =   9480
   WindowState     =   2  'Maximized
   Begin VB.CommandButton export 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2430
      Picture         =   "GRview.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6420
      Visible         =   0   'False
      Width           =   345
   End
   Begin MSComDlg.CommonDialog d1 
      Left            =   810
      Top             =   5250
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton return1 
      Caption         =   "&Return"
      Height          =   375
      Left            =   4590
      TabIndex        =   3
      Top             =   6390
      Width           =   1155
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      ItemData        =   "GRview.frx":0451
      Left            =   3480
      List            =   "GRview.frx":0464
      TabIndex        =   1
      Text            =   "100 %"
      Top             =   6390
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2820
      Picture         =   "GRview.frx":048A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6360
      Width           =   585
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid R1 
      Height          =   6195
      Left            =   30
      TabIndex        =   2
      Top             =   120
      Width           =   11775
      _ExtentX        =   20765
      _ExtentY        =   10922
      _Version        =   393216
      BackColor       =   16777215
      BackColorFixed  =   16777215
      BackColorSel    =   16777215
      BackColorBkg    =   16777215
      BackColorUnpopulated=   16777215
      GridColorFixed  =   16777215
      GridColorUnpopulated=   16777215
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      FillStyle       =   1
      GridLinesFixed  =   0
      MergeCells      =   1
      RowSizingMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   0
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "GRIDpreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQLSTRING As String
'Dim CON As ADODB.Connection
Dim RS As ADODB.Recordset
Private Sub Combo1_Change()

If Trim(Combo1.text) = "75 %" Then
    r1.Font.Size = 8
End If

If Trim(Combo1.text) = "100 %" Then
    r1.Font.Size = 10
End If

If Trim(Combo1.text) = "200 %" Then
    r1.Font.Size = 18
End If

If Trim(Combo1.text) = "125 %" Then
    r1.Font.Size = 12
End If

If Trim(Combo1.text) = "150 %" Then
    r1.Font.Size = 14
End If


End Sub
Private Sub Combo1_Click()

r1.Row = 1
    
If Trim(Combo1.text) = "75 %" Then
    r1.Font.Size = 8
End If
    If Trim(Combo1.text) = "100 %" Then
        r1.Font.Size = 10
    End If
If Trim(Combo1.text) = "200 %" Then
    r1.Font.Size = 18
    For I = 0 To 5
        r1.Col = I
        'r1.ColWidth(i) = r1
        'r1.ColWidth (i) * 1.5
    Next
End If

If Trim(Combo1.text) = "125 %" Then
    r1.Font.Size = 12
End If
If Trim(Combo1.text) = "150 %" Then
    r1.Font.Size = 14
End If

End Sub

Private Sub Command1_Click()
    Printdlg.Show
End Sub
Private Sub export_Click()

d1.ShowPrinter
MsgBox "copies =" + Str(d1.copies)
'd1.Copies
'Printer.PaperSize

End Sub
Public Function printnow()

Dim X As Long
Dim p As Printer
For I = 0 To Printers.Count - 1
    Set p = Printers(I)
    If Trim(p.DeviceName) = Trim(Printdlg.Comboprinters.text) Then
        Exit For
    End If
Next
For I = 1 To (Printdlg.UpDown1.value)
    X = Shell("" + App.Path + "\ppp.bat " + App.Path + "\vipin.txt " & Trim(p.Port), vbHide)
Next

Printdlg.UpDown1.value = 1
Printdlg.Text1.text = "1"
    
End Function
Private Sub Form_Load()
    Command1.top = r1.top + r1.Height + 30
    Combo1.top = r1.top + r1.Height + 30
    
End Sub

Public Function SQL(SQLSTR As String)
    Dim rs1 As ADODB.Recordset
    SQLSTRING = SQLSTR
    
    'Set CON = New ADODB.Connection
    'With CON
    '.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + App.Path + "\" + Trim(main.directory) + "\data.mdb"
    '.Open
    'End With
    
    Set RS = New ADODB.Recordset
    Set rs1 = New ADODB.Recordset
    RS.Open SQLSTRING, con, adOpenStatic, adLockReadOnly, adCmdText
    If Not RS.EOF Then
        r1.Cols = 0
        r1.rows = 0
        r1.rows = 2
        r1.Cols = RS.Fields.Count
        r1.Col = 0
        r1.Row = 0
        '*************************
        rs1.Open "printing", con, adOpenKeyset, adLockReadOnly, adCmdText
        For I = 0 To RS.Fields.Count - 1
            r1.Col = I
            rs1.Find "fieldname='" + Trim(RS(I).Name) + "'"
            If Not rs1.EOF Then
                r1.CellFontBold = True
                r1.text = rs1(1)
                r1.ColWidth(I) = Len(Trim(rs1(1))) * 125
                
                rs1.MoveFirst
            Else
                r1.text = " "
            End If
        Next
        rs1.close
        r1.Row = 1
        For I = 0 To RS.Fields.Count - 1
            r1.Col = I
            r1.text = "--------------------------------------------------------------------------------------------------------------------"
        Next
        '*********************
        r1.rows = 3
        'r1.Cols = RS.Fields.Count
        r1.Row = 2
        r1.Col = 0
        RS.MoveFirst
        Do While Not RS.EOF
            For I = 0 To RS.Fields.Count - 1
                r1.Col = I
                '   r1.Text = RS(i)
                '   r1.ColAlignment(i) = 2
                    If r1.ColWidth(I) < (RS(I).DefinedSize * 120) Then
                        r1.ColWidth(I) = RS(I).DefinedSize * 120
                        If RS(I).DefinedSize <= 2 Then
                            r1.ColWidth(I) = 5 * 120
                        End If
                    End If
                'Else
                    r1.text = (RS(I))
               ' End If
                    
            Next
            r1.rows = r1.rows + 1
            r1.Row = r1.Row + 1
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
        RS.close
    Me.Show
 End If
 
End Function
Private Sub Form_Resize()
If Me.Width > 350 And Me.Height > 1500 Then
    r1.Width = Me.Width - 250
    r1.Height = Me.Height - 1000
    Command1.top = r1.top + r1.Height + 30
    Combo1.top = r1.top + r1.Height + 30
    return1.top = Combo1.top + 100
'    setup.Top = Command1.Top
    export.top = Command1.top
End If
End Sub
Private Sub return1_Click()
    Unload Me
End Sub

