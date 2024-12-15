VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmDQry 
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13080
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8265
   ScaleWidth      =   13080
   WindowState     =   2  'Maximized
   Begin VB.Frame CONTAINER 
      Height          =   7845
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   12690
      Begin VB.CommandButton cmdExit 
         Caption         =   "&Exit"
         Height          =   645
         Left            =   10935
         TabIndex        =   31
         Top             =   180
         Width           =   1635
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   645
         Left            =   9270
         TabIndex        =   30
         Top             =   180
         Width           =   1635
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "&Report"
         Height          =   645
         Left            =   7470
         TabIndex        =   29
         Top             =   180
         Width           =   1770
      End
      Begin VB.CommandButton cmdReport 
         Caption         =   "&Report Head"
         Height          =   645
         Left            =   5445
         TabIndex        =   28
         Top             =   180
         Width           =   1995
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Abandon"
         Height          =   645
         Left            =   3555
         TabIndex        =   27
         Top             =   180
         Width           =   1860
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save Query"
         Height          =   645
         Left            =   1755
         TabIndex        =   26
         Top             =   180
         Width           =   1770
      End
      Begin VB.CommandButton cmdViewQry 
         Caption         =   "&View Query"
         Height          =   645
         Left            =   90
         TabIndex        =   25
         Top             =   180
         Width           =   1635
      End
      Begin VB.ListBox List1 
         Height          =   6105
         Left            =   7560
         TabIndex        =   20
         Top             =   900
         Width           =   4920
      End
      Begin VB.Frame FrameSTUDENTRECORDS 
         Height          =   3915
         Left            =   30
         TabIndex        =   5
         Top             =   3255
         Width           =   7365
         Begin VB.CheckBox Checksno 
            Caption         =   "S.No"
            Height          =   240
            Left            =   3105
            TabIndex        =   19
            Top             =   180
            Width           =   735
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Remove"
            Height          =   345
            Left            =   3030
            TabIndex        =   18
            Top             =   1110
            Width           =   945
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Add"
            Height          =   345
            Left            =   3030
            TabIndex        =   17
            Top             =   750
            Width           =   945
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Sorting"
            ForeColor       =   &H00FF0000&
            Height          =   795
            Left            =   2925
            TabIndex        =   14
            Top             =   3030
            Width           =   1455
            Begin VB.OptionButton Option5 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Ascending"
               ForeColor       =   &H00FF0000&
               Height          =   225
               Left            =   135
               TabIndex        =   16
               Top             =   255
               Value           =   -1  'True
               Width           =   1065
            End
            Begin VB.OptionButton Option6 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Descending"
               ForeColor       =   &H00FF0000&
               Height          =   225
               Left            =   120
               TabIndex        =   15
               Top             =   525
               Width           =   1215
            End
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Add"
            Height          =   345
            Left            =   3060
            TabIndex        =   13
            Top             =   2190
            Width           =   945
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Remove"
            Height          =   345
            Left            =   3060
            TabIndex        =   12
            Top             =   2550
            Width           =   945
         End
         Begin VB.ListBox List3 
            ForeColor       =   &H00FF0000&
            Height          =   1425
            ItemData        =   "frmDQry.frx":0000
            Left            =   60
            List            =   "frmDQry.frx":0002
            TabIndex        =   11
            Top             =   2160
            Width           =   2475
         End
         Begin VB.ListBox List2 
            ForeColor       =   &H00FF0000&
            Height          =   1425
            ItemData        =   "frmDQry.frx":0004
            Left            =   4500
            List            =   "frmDQry.frx":0006
            TabIndex        =   10
            Top             =   2160
            Width           =   2145
         End
         Begin VB.CommandButton Command2 
            Caption         =   "down"
            Height          =   255
            Left            =   6660
            TabIndex        =   9
            Top             =   1140
            Width           =   600
         End
         Begin VB.CommandButton Command1 
            Caption         =   "up"
            Height          =   285
            Left            =   6660
            TabIndex        =   8
            Top             =   840
            Width           =   600
         End
         Begin VB.ListBox SelectAdmfieldname 
            Height          =   1815
            ItemData        =   "frmDQry.frx":0008
            Left            =   4500
            List            =   "frmDQry.frx":000A
            TabIndex        =   7
            Top             =   180
            Width           =   2145
         End
         Begin VB.ListBox admFieldname 
            Height          =   1815
            ItemData        =   "frmDQry.frx":000C
            Left            =   30
            List            =   "frmDQry.frx":000E
            Sorted          =   -1  'True
            TabIndex        =   6
            Top             =   120
            Width           =   2500
         End
      End
      Begin VB.TextBox Textan3 
         Height          =   675
         Left            =   2640
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   7395
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.TextBox Textan1 
         Height          =   1050
         Left            =   7740
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   7575
         Visible         =   0   'False
         Width           =   3480
      End
      Begin VB.Frame frame7 
         Height          =   2430
         Left            =   60
         TabIndex        =   2
         Top             =   825
         Width           =   7365
         Begin VSFlex7Ctl.VSFlexGrid vs 
            Height          =   2175
            Left            =   45
            TabIndex        =   24
            Top             =   180
            Width           =   7260
            _cx             =   12806
            _cy             =   3836
            _ConvInfo       =   1
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   0
            ShowComboButton =   -1  'True
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
         End
      End
      Begin VB.TextBox Textan2 
         Height          =   600
         Left            =   8700
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   8655
         Width           =   4710
      End
      Begin Crystal.CrystalReport CrystalReport2 
         Bindings        =   "frmDQry.frx":0010
         Left            =   450
         Top             =   8955
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileType   =   19
         ReportSource    =   3
         PrintFileLinesPerPage=   60
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   1530
         Top             =   8865
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         ReportFileName  =   "c:\scope\scope31.rpt"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   270
         TabIndex        =   23
         Top             =   3855
         Width           =   1095
      End
      Begin VB.Label LABELUSERNAME 
         BackStyle       =   0  'Transparent
         Height          =   330
         Left            =   60
         TabIndex        =   22
         Top             =   9375
         Width           =   7785
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   180
         TabIndex        =   21
         Top             =   3735
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmDQry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim maintablename As String
Private Sub cmdClear_Click()
searchframEini
'resultgrid1.Visible = False
List1.Visible = False

End Sub
Sub searchframEini()
Dim I As Integer
    Dim MyTableDef As Recordset
    searchfieldname1.Clear
    admFieldname.Clear
    Set MyTableDef = Db.OpenRecordset(maintablename)
    For I = 0 To MyTableDef.Fields.Count - 1
        searchfieldname1.AddItem MyTableDef.Fields(I).Name
        admFieldname.AddItem MyTableDef.Fields(I).Name
    Next I
    logicaloperator1.Clear
logicaloperator1.AddItem "And"
logicaloperator1.AddItem "Or"
logicaloperator1.AddItem "Not"
simpleoperator1.Clear
simpleoperator1.AddItem "="
simpleoperator1.AddItem ">"
simpleoperator1.AddItem "<"
simpleoperator1.AddItem ">="
simpleoperator1.AddItem "<="
simpleoperator1.AddItem "<>"
simpleoperator1.AddItem "Like"
searchgrid1.Col = 0
searchgrid1.Row = 0
searchgrid1.Width = 7200
searchgrid1.ColWidth(0) = 2500
searchgrid1.Text = "Search Field"
searchgrid1.Col = 1
searchgrid1.Row = 0
searchgrid1.ColWidth(1) = 600
searchgrid1.Text = ""
searchgrid1.Col = 2
searchgrid1.Row = 0
searchgrid1.ColWidth(2) = 2500
searchgrid1.Text = "Search Value"
searchgrid1.Col = 3
searchgrid1.Row = 0
searchgrid1.ColWidth(3) = 1000
searchgrid1.Text = "Opr "
searchfieldname1.Visible = False
searchfieldname1.Text = ""
simpleoperator1.Visible = False
simpleoperator1.Text = ""
searchvalue1.Visible = False
searchvalue1.Text = ""
logicaloperator1.Visible = False
logicaloperator1.Text = ""
For I = 1 To 8
searchgrid1.Row = I
searchgrid1.Col = 0
searchgrid1.Text = ""
searchgrid1.Col = 1
searchgrid1.Text = ""
searchgrid1.Col = 2
searchgrid1.Text = ""
searchgrid1.Col = 3
searchgrid1.Text = ""
Next
End Sub

Private Sub cmdDelete_Click()

If Len(List1) > 0 Then
Dim X As Integer
                X = MsgBox("Are you sure Y/N", vbQuestion + vbYesNo)
    If X = 6 Then

Dim ss As String
ss = "DESCRIPTION = '" + List1 + "'"
ANLCRI.FindFirst ss
If Not ANLCRI.NoMatch Then
    Dim ANLCRITRN As Recordset
   ss = "delete * from ANLCRITARIATRANS WHERE " + ss
   Db.Execute ss
   ss = "DESCRIPTION = '" + List1 + "'"
   ss = "delete * from ANLCRITARIAFIELDS WHERE " + ss
   Db.Execute ss
  
    ss = "DESCRIPTION = '" + List1 + "'"
   ss = "delete * from  ANLCRITARIA WHERE " + ss
   Db.Execute ss
   
   List1.RemoveItem List1.ListIndex
 End If
End If
End If

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub CmdSave_Click()

Dim ANLCRITRN As Recordset
Dim anlcrifld As Recordset
Dim ANLCRI As Recordset

Set ANLCRITRN = Db.OpenRecordset("ANLCRITARIATRANS")
If REFCHECK = "Y" Then
    Dim LLL As Integer
    Dim rtmp As String
    LLL = Len(Textan2)
    rtmp = Mid(Trim(Textan2), LLL - 3, 3)
    If rtmp = "AND" Then
        Textan2 = Left(Textan2, (LLL - 4))
    End If
    REFCHECK = "N"
End If
Dim savefilnam As String
savefilnam = InputBox("Enter Search Description ", "Search Description ", "")
Dim ss As String
ss = "Description = '" + savefilnam + "'"
Set ANLCRI = Db.OpenRecordset("SELECT * FROM ANLCRITARIA WHERE " + ss + " and  SEARCHSTATUS = '" + SEARCHSTATUS + "'")
If ANLCRI.RecordCount <= 0 Then
    ANLCRI.AddNew
    ANLCRI!DESCRIPTION = savefilnam
    ANLCRI!CRITARIA = Textan2
    ANLCRI!SEARCHSTATUS = SEARCHSTATUS
    ANLCRI.update
Else
    Dim X As Integer
    X = MsgBox("Query already exists. Overwrite it ?", 4, "Confirmation")
    If X = 6 Then
        'ANLCRI.Edit
        ANLCRI!CRITARIA = Textan2
        ANLCRI!SEARCHSTATUS = SEARCHSTATUS
        ANLCRI.update
        Db.Execute "DELETE * FROM ANLCRITARIAtrans WHERE DESCRIPTION = '" + Trim(savefilnam) + "'"
    End If
End If
Dim I As Integer
For I = 1 To 8
    searchgrid1.Row = I
    searchgrid1.Col = 0
    If searchgrid1.Text <> "" Then
        ANLCRITRN.AddNew
        ANLCRITRN!DESCRIPTION = savefilnam
        ANLCRITRN!searchfieldname1 = searchgrid1.Text
        searchgrid1.Col = 1
        ANLCRITRN!simpleoperator1 = searchgrid1.Text
        searchgrid1.Col = 2
        ANLCRITRN!searchvalue1 = searchgrid1.Text
        searchgrid1.Col = 3
        ANLCRITRN!logicaloperator1 = searchgrid1.Text
        
                ANLCRITRN!SEARCHSTATUS = SEARCHSTATUS

        ANLCRITRN.update
    End If
Next
Set anlcrifld = Db.OpenRecordset("ANLCRITARIAFIELDS")
Db.Execute "DELETE * FROM ANLCRITARIAFIELDS WHERE DESCRIPTION = '" + Trim(savefilnam) + "'"
For I = 0 To SelectAdmfieldname.ListCount - 1
SelectAdmfieldname.ListIndex = I
anlcrifld.AddNew
anlcrifld!admFieldname = SelectAdmfieldname
anlcrifld!DESCRIPTION = savefilnam
anlcrifld!admsortorder = I
Dim MyTableDef As Recordset
Dim J, K As Integer

        Set MyTableDef = Db.OpenRecordset(maintablename)
        For J = 0 To MyTableDef.Fields.Count - 1
            If MyTableDef.Fields(J).Name = SelectAdmfieldname Then anlcrifld!admFieldtype = MyTableDef.Fields(J).Type: Exit For
          
        Next J
        MyTableDef.close
          For K = 0 To List2.ListCount - 1
             List2.ListIndex = K
             If SelectAdmfieldname = List2 Then
             anlcrifld!admFieldSORT = K
             If Option6.value = True Then
             anlcrifld!ASCDEC = "D"
             Else
             anlcrifld!ASCDEC = "A"
             End If
             End If
             Next K
        
        
        
                anlcrifld!SEARCHSTATUS = SEARCHSTATUS

anlcrifld.update
Next


List1.AddItem savefilnam



End Sub
Private Sub BUILTSEARCHQUERY()
Dim SEARCHFIELDNAME As String
Dim SEARCHFIELDTYPE As String
Dim simpleopr As String
Dim logicalopr As String
Dim searchtext As String
Dim sq As String
Dim SQ1 As String
Dim I As Integer

SQ1 = "Select * from " + maintablename + " where "
For I = 1 To 8
    
    SEARCHFIELDNAME = ""
    simpleopr = ""
    logicalopr = ""
    searchtext = ""
    
    
    
    If vs.TextMatrix(I, 0) <> "" Then
        SEARCHFIELDNAME = Trim(vs.TextMatrix(I, 0))
        SQ1 = SQ1 + Trim(Trim(vs.TextMatrix(I, 0))) + " "
    End If
    
    If Trim(vs.TextMatrix(I, 1)) <> "" Then
        simpleopr = Trim(vs.TextMatrix(I, 1))
        SQ1 = SQ1 + Trim(vs.TextMatrix(I, 1)) + " "
    End If

       
    
    If Trim(vs.TextMatrix(I, 2)) <> "" Then
        searchtext = Trim(vs.TextMatrix(I, 2))
        Dim J As Integer
        Dim MyTableDef As Recordset
        Set MyTableDef = Db.OpenRecordset(maintablename)
        For J = 0 To MyTableDef.Fields.Count - 1
            If MyTableDef.Fields(J).Name = SEARCHFIELDNAME Then SEARCHFIELDTYPE = MyTableDef.Fields(J).Type: Exit For
        Next J
        Select Case SEARCHFIELDTYPE
        Case 10 ' TEXT
         SQ1 = SQ1 + " '" + Trim(searchgrid1) + "' "
        Case 1, 3, 4
        SQ1 = SQ1 + Trim(searchgrid1) + " "
        Case 8
        SQ1 = SQ1 + " cdate('" + Trim(searchgrid1) + "') "
        End Select
    End If
    searchgrid1.Col = 3
    If Trim(searchgrid1) <> "" Then
        logicalopr = Trim(searchgrid1)
        SQ1 = SQ1 + Trim(searchgrid1) + " "
    End If
Next
sq = SQ1
'MsgBox sq
Textan2 = sq
End Sub

Private Sub cmdshow_Click()
Dim LLL As Integer
Dim rtmp As String
Dim sql1, mysql As String
Dim PATPR As Recordset

'On Error Resume Next
'query built for demographical data

BUILTSEARCHQUERY
LLL = Len(Textan2)
rtmp = Mid(Trim(Textan2), LLL - 3, 3)
If rtmp = "AND" Then
    Textan2 = Left(Textan2, (LLL - 4))
End If
 'insertin the heading
If UCase(Textan2) <> "SELECT * FROM " + maintablename + " WHERE " Then
    sql1 = Textan2
Else
    sql1 = "SELECT * FROM " + maintablename
End If
'mysql = "delete * from " + reporttablename
'db.Execute mysql
'mysql = " INSERT INTO " + reporttablename + " " + sql1
'db.Execute mysql
'reference search end
Dim resultgrid1set As Recordset
Dim sq As String
'sq = "select * from " + reporttablename + " "
sq = sql1
If List2.List(0) <> "" Then
    sq = sq + "order by " + List2.List(0) + " "
    If Option5.value = False Then sq = sq + "desc"
End If
If List2.List(1) <> "" Then
    sq = sq + ", " + List2.List(1) + " "
    If Option5.value = False Then sq = sq + "desc"
End If
If List2.List(2) <> "" Then
    sq = sq + ", " + List2.List(2) + " "
    If Option5.value = False Then sq = sq + "desc"
End If
Dim tmpsetup As Recordset
Set resultgrid1set = Db.OpenRecordset(sq)
Set tmpsetup = Db.OpenRecordset("Setup")
' DBGrid start
'Data1.DatabaseName = db.Name
'Set Data1.Recordset = resultgrid1set
'db grid end
' excel format conversion start
   ' Declare object variables for Microsoft Excel,
   ' application workbook, and worksheet objects.
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim I As Integer
Dim xl As Excel.Application
If xl Is Nothing Then
   Set xl = CreateObject("Excel.Application")
End If
xl.Visible = True
Set xlBook = xl.Workbooks.Add
Set xlSheet = xlBook.Worksheets.Add
Dim c, r As Long
c = 1
r = 1
Dim noofcolumn As Integer
Dim lastcolumn1, lastcolumn2 As String
noofcolumn = SelectAdmfieldname.ListCount
Select Case noofcolumn
Case 1
lastcolumn1 = "a1"
lastcolumn2 = "a2"
Case 2
lastcolumn1 = "b1"
lastcolumn2 = "b2"
Case 3
lastcolumn1 = "c1"
lastcolumn2 = "c2"
Case 4
lastcolumn1 = "d1"
lastcolumn2 = "d2"
Case 5
lastcolumn1 = "e1"
lastcolumn2 = "e2"
Case 6
lastcolumn1 = "f1"
lastcolumn2 = "f2"
Case 7
lastcolumn1 = "g1"
lastcolumn2 = "g2"
Case 8
lastcolumn1 = "h1"
lastcolumn2 = "h2"
Case 9
lastcolumn1 = "i1"
lastcolumn2 = "i2"
Case Is >= 10
lastcolumn1 = "j1"
lastcolumn2 = "j2"
End Select
xl.Range("a1", lastcolumn1).Merge
xl.Range("a1").value = tmpsetup!sName
xl.Range("a1").Font.Bold = True
xl.Range("a1").Font.Size = 14
xl.Range("a2", lastcolumn2).Merge
xl.Range("a2").value = tmpsetup!city
xl.Range("a2").Font.Bold = False
xl.Range("a2").Font.Size = 12
r = r + 3
Dim MyTableDef As Recordset
Dim J As Integer
Set MyTableDef = Db.OpenRecordset(maintablename)
Dim admfieldtypearray(300)
Dim tbladmmast As Recordset
  If Checksno.value = 1 Then
                xlSheet.Cells(r, c + I).value = "S.No."
                xlSheet.Cells(r, c + I).Font.Bold = True
                c = c + 1
                End If
For I = 0 To SelectAdmfieldname.ListCount - 1
    SelectAdmfieldname.ListIndex = I
  
    For J = 0 To MyTableDef.Fields.Count - 1
        If MyTableDef.Fields(J).Name = SelectAdmfieldname Then
                
            admfieldtypearray(I) = MyTableDef.Fields(J).Type
            Set tbladmmast = Db.OpenRecordset("select * from anladmmast where admfieldname = '" + SelectAdmfieldname + "'")
            If tbladmmast.RecordCount > 0 Then
                 xlSheet.Cells(r, c + I).value = tbladmmast!admreporthead
                 xlSheet.Cells(r, c + I).Font.Bold = True
            End If
            Exit For
         End If
     Next J
Next I
MyTableDef.close
r = r + 1
Dim sno As Integer
sno = 1
Do While Not resultgrid1set.EOF
   
   ' Assign the values entered in the text boxes to
   ' Microsoft Excel cells.
   Dim xx As String
    xlSheet.Columns.AutoFit
   If Checksno.value = 1 Then
                xlSheet.Cells(r, 1).value = sno
            '    C = C + 1
                End If
   For I = 0 To SelectAdmfieldname.ListCount - 1
 
        SelectAdmfieldname.ListIndex = I
        If admfieldtypearray(I) = 8 Then
            If resultgrid1set.Fields(SelectAdmfieldname) <> "" Then
                xlSheet.Cells(r, c + I).value = Format(CStr(resultgrid1set.Fields(SelectAdmfieldname)), "mm/dd/yyyy")
            End If
        Else
            xlSheet.Cells(r, c + I).value = resultgrid1set.Fields(SelectAdmfieldname)
        End If
    Next
    If Not resultgrid1set.EOF Then
        resultgrid1set.MoveNext
        
        r = r + 1
        sno = sno + 1
        
    End If
Loop
' Close the Workbook
'   xlBook.Close
'   xl.Quit
'resultgrid1.Visible = True
Exit Sub
errortrap:
errorcall

End Sub

Private Sub cmdViewQry_Click()
List1.Visible = True
End Sub
Private Sub Form_Load()

'frmBack Me, Me.CONTAINER
'db.Execute "Update ANLCRITARIA set SEARCHSTATUS = 'STUDENT' where SEARCHSTATUS = NULL"
'db.Execute "Update ANLADMMAST set SEARCHSTATUS = 'STUDENT' where SEARCHSTATUS = NULL"
'db.Execute "Update ANLCRITARIAFIELDS set SEARCHSTATUS = 'STUDENT' where SEARCHSTATUS = NULL"
'db.Execute "Update ANLCRITARIATRANS set SEARCHSTATUS = 'STUDENT' where SEARCHSTATUS = NULL"

''
''Me.Caption = SEARCHSTATUS + " Search"
''
''REFCHECK = "N"
''
''List1.Visible = False
''If SEARCHSTATUS = "STAFF" Then
''maintablename = "staffhistory"
''Textan2 = "SELECT * FROM STAFFHISTORY WHERE "
''Set ANLCRI = db.OpenRecordset("SELECT * FROM ANLCRITARIA WHERE SEARCHSTATUS = 'STAFF'")
''Else
''maintablename = "admform"
''Textan2 = "SELECT * FROM " + maintablename + " WHERE "
''Set ANLCRI = db.OpenRecordset("SELECT * FROM ANLCRITARIA WHERE SEARCHSTATUS = '          ' OR SEARCHSTATUS = 'STUDENT'")
''End If
''searchframEini

maintablename = "admform"

If rs.State = 1 Then rs.close
rs.Open "select * from DQurey", CON, adOpenKeyset, adLockReadOnly
If rs.RecordCount > 0 Then
List1.Clear

Do While Not rs.EOF
   List1.AddItem rs!DESCRIPTION
   If Not rs.EOF Then
      rs.MoveNext
    End If
Loop
End If


End Sub
Private Sub List1_Click()

Dim rsS As New ADODB.Recordset

If rs.State = 1 Then rs.close
rs.Open "SELECT * FROM DQurey WHERE DESCRIPTION='" & List1.Text & "'"
If rs.EOF = False Then
       
    Textan2 = rs!CRITARIA
    
    If rsS.State = 1 Then rsS.close
    rsS.Open "SELECT * FROM DQurey_critaria WHERE DESCRIPTION='" & List1.Text & "'", CON

    Dim I As Integer
    I = 1
    Do While Not rsS.EOF
        
        vs.TextMatrix(I, 0) = rsS!searchfieldname1
        vs.TextMatrix(I, 1) = rsS!simpleoperator1
        vs.TextMatrix(I, 2) = rsS!searchvalue1
        vs.TextMatrix(I, 3) = rsS!logicaloperator1
        If Not rsS.EOF Then
            rsS.MoveNext
            I = I + 1
        End If
    Loop
    
    Dim anlcrifld As Recordset
    If rs.State = 1 Then rs.close
    rs.Open "SELECT * FROM DQurey_fields WHERE DESCRIPTION='" & List1.Text & "'", CON, adOpenKeyset, adLockReadOnly
 SelectAdmfieldname.Clear
 List3.Clear
 
If rs.RecordCount > 0 Then
    Do While Not rs.EOF
 SelectAdmfieldname.AddItem rs!admFieldname
  List3.AddItem rs!admFieldname
  If rs!ASCDEC = "A" Then
  Option5.value = True
  Else
  Option6.value = True
  End If
  
  
  
  If Not rs.EOF Then
         rs.MoveNext
  End If
Loop
End If

If rs.State = 1 Then rs.close
rs.Open "SELECT * FROM DQurey_fields where DESCRIPTION='" & List1.Text & "' AND admFieldSORT <> NULL  order by admFIELDSort", CON
 List2.Clear
 If rs.RecordCount > 0 Then
 Do While Not rs.EOF
 If IsNumeric(Val(rs!admFieldSORT)) Then
  List2.AddItem rs!admFieldname
  End If
  If Not rs.EOF Then
         rs.MoveNext
End If
Loop
End If
  
    
    End If
    
    
    
    
    
    
    


End Sub
Private Sub List1_DblClick()
 
 List1_Click
 cmdshow_Click
 List1.Visible = False

End Sub
