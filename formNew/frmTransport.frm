VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmTransport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transport Master"
   ClientHeight    =   7980
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   9372
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   9372
   ShowInTaskbar   =   0   'False
   Begin VB.Frame buttonFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   840
      Left            =   225
      TabIndex        =   1
      Top             =   6840
      Width           =   6585
      Begin VB.CommandButton cmdAdd_1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   45
         Picture         =   "frmTransport.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   45
         Width           =   1065
      End
      Begin VB.CommandButton cmdSave_2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   1140
         Picture         =   "frmTransport.frx":0BE4
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   30
         Width           =   1065
      End
      Begin VB.CommandButton cmdDelete_3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   2220
         Picture         =   "frmTransport.frx":17C8
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   45
         Width           =   1065
      End
      Begin VB.CommandButton cmdEdit_4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   3315
         Picture         =   "frmTransport.frx":23AC
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   30
         Width           =   1065
      End
      Begin VB.CommandButton cmdExit_12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   5490
         Picture         =   "frmTransport.frx":27B9
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   45
         Width           =   1065
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "S&earch"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   4395
         Picture         =   "frmTransport.frx":339D
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   30
         Width           =   1065
      End
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   765
      MaxLength       =   50
      TabIndex        =   0
      Top             =   180
      Width           =   5055
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   5970
      Left            =   135
      TabIndex        =   10
      Top             =   675
      Width           =   6630
      _cx             =   11695
      _cy             =   10530
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   11162880
      BackColorFixed  =   7917545
      ForeColorFixed  =   8388608
      BackColorSel    =   16777153
      ForeColorSel    =   4210752
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
      ExplorerBar     =   1
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
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
   Begin VB.Label lblName1 
      Height          =   285
      Left            =   5940
      TabIndex        =   9
      Top             =   180
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0078CFE9&
      BorderWidth     =   4
      Height          =   960
      Left            =   180
      Top             =   6795
      Width           =   6720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
      Height          =   240
      Left            =   180
      TabIndex        =   8
      Top             =   225
      Width           =   555
   End
End
Attribute VB_Name = "frmTransport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim edit1 As Boolean
Private Sub cmdAdd_1_Click()


txtName = ""
lblName1 = ""
vs.Clear
cmdSave_2.Enabled = True
cmdEdit_4.Enabled = False
cmdDelete_3.Enabled = False
edit1 = False

vs.FormatString = "Id|Placeofsupply|GeneralRate|DoordeliveryRate"
vs.ColWidth(0) = 800
vs.ColWidth(1) = 3000
vs.ColWidth(2) = 1200
vs.ColWidth(3) = 1200

txtName.SetFocus

   
   
End Sub

Private Sub cmdDelete_3_Click()

If MsgBox("want To delete ?", vbInformation + vbYesNo) = vbYes Then
 con.Execute "delete from  transportmaster where Transportname='" & txtName & "' and " & stringyear
End If


cmdAdd_1_Click

End Sub

Private Sub cmdEdit_4_Click()

edit1 = True
cmdEdit_4.Enabled = False
cmdSave_2.Enabled = True
cmdDelete_3.Enabled = True
cmdSave_2.SetFocus

End Sub

Private Sub cmdExit_12_Click()
Unload Me
End Sub



Private Sub cmdPrint_Click()

End Sub

Private Sub cmdSave_2_Click()

On Error GoTo aa:


If txtName = "" Then
   MsgBox "Enter Transport Name. ...", vbInformation
   txtName.SetFocus
   Exit Sub
End If

If MsgBox("Want to Save ?", vbQuestion + vbYesNo) = vbNo Then
   Exit Sub
End If
   




If edit1 = True Then
   
   con.Execute "update [transportmaster] set Transportname='" & UCase(txtName) & "' where " & stringyear & " and Transportname='" & lblName1 & "'"

Else

con.Execute "INSERT INTO  [transportmaster]" & _
           "([Transportname]" & _
           ",[Fyear]" & _
           ",[setupid]" & _
     ") Values" & _
           "('" & UCase(txtName) & "'" & _
           ",'" & main.session & "'" & _
           ",'" & main.setupid & "')"
        
        

End If

Screen.MousePointer = vbHourglass

For J = 1 To vs.rows - 1

If vs.TextMatrix(J, 0) <> "" Then
    v1 = IIf(vs.TextMatrix(J, 2) = "", 0, vs.TextMatrix(J, 2))
    v2 = IIf(vs.TextMatrix(J, 3) = "", 0, vs.TextMatrix(J, 3))
    con.Execute "update [TransportDet] set Placeofsupply='" & vs.TextMatrix(J, 1) & "',GeneralRate=" & v1 & ",Doordelivery=" & v2 & " where id=" & vs.TextMatrix(J, 0) & ""
ElseIf vs.TextMatrix(J, 0) = "" Then
  If vs.TextMatrix(J, 1) <> "" Then
    v1 = IIf(vs.TextMatrix(J, 2) = "", 0, vs.TextMatrix(J, 2))
    v2 = IIf(vs.TextMatrix(J, 3) = "", 0, vs.TextMatrix(J, 3))
    con.Execute "insert into [TransportDet](TransportName,Placeofsupply,GeneralRate,Doordelivery) values('" & UCase(txtName) & "','" & UCase(vs.TextMatrix(J, 1)) & "'," & v1 & "," & v2 & ")"
  End If
End If

Next

Screen.MousePointer = vbDefault

cmdSave_2.Enabled = False
cmdEdit_4.Enabled = False
cmdDelete_3.Enabled = False



txtName.SetFocus

   
Exit Sub

aa:

Screen.MousePointer = vbDefault
MsgBox "" & err.DESCRIPTION

End Sub

Private Sub cmdSearch_Click()

   value = "select Transportname from [transportmaster] order by Transportname"
   popuplistModel10 value, con
 
   cmdSave_2.Enabled = False
   cmdEdit_4.Enabled = True

End Sub
Private Sub cmdSearch_GotFocus()
  

If PopUpValue1 <> "" Then
   
   
   
   
   txtName = PopUpValue1
   lblName1 = PopUpValue1
    
   vs.Clear
   vs.rows = 2
   
   Set RS = New ADODB.Recordset
   RS.Open "select Id,Placeofsupply,GeneralRate,Doordelivery from TransportDet where TransportName='" & PopUpValue1 & "' order by Id", con
   For n1 = 1 To RS.RecordCount
     DoEvents
     vs.TextMatrix(n1, 0) = RS!id
     vs.TextMatrix(n1, 1) = RS!Placeofsupply
     If RS!GeneralRate > 0 Then
        vs.TextMatrix(n1, 2) = RS!GeneralRate
     End If
     
     If RS!Doordelivery > 0 Then
        vs.TextMatrix(n1, 3) = RS!Doordelivery
     End If
     
     DoEvents
     DoEvents
     RS.MoveNext
     vs.rows = vs.rows + 1
   Next
   
   vs.FormatString = "Id|Placeofsupply|GeneralRate|DoordeliveryRate"
   vs.ColWidth(0) = 800
   vs.ColWidth(1) = 3000
   vs.ColWidth(2) = 1200
   vs.ColWidth(3) = 1200
   
End If
 
PopUpValue1 = ""

cmdEdit_4.Enabled = True
  

End Sub


Private Sub Form_Activate()
cmdAdd_1_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()
 Me.top = 500
 Me.Left = 500
 
 BackColorFrom Me
 
End Sub
Private Sub txtname_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      cmdSave_2.SetFocus
   End If
End Sub
Private Sub txtName_LostFocus()
   txtName = UCase(txtName)
End Sub
Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 115 Then
     If MsgBox("Want to Delete ?", vbQuestion + vbYesNo) = vbYes Then
        con.Execute "delete from TransportDet  where id=" & vs.TextMatrix(vs.RowSel, 0) & ""
        vs.RemoveItem (vs.RowSel)
     End If
  End If
End Sub
Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      If vs.Col = 2 Then
         SendKeys "{right}"
      ElseIf vs.Col = 3 Then
         SendKeys "{down}"
         vs.Col = 2
      End If
   End If
End Sub
