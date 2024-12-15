VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmBookRetailers 
   Caption         =   "Book Retailers Details"
   ClientHeight    =   7452
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   12048
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.6
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBookRetailers.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7452
   ScaleWidth      =   12048
   Begin VB.CommandButton cmdAddFile 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add &File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9810
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   6390
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.CommandButton cmdImportExcel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "U&pdate Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   10845
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   6390
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.CheckBox Check1_import 
      Caption         =   "Import Excel File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2925
      TabIndex        =   29
      Top             =   6435
      Width           =   1770
   End
   Begin VB.TextBox txtpath 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4815
      TabIndex        =   28
      Top             =   6435
      Visible         =   0   'False
      Width           =   4965
   End
   Begin Crystal.CrystalReport cr 
      Left            =   7515
      Top             =   4590
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame buttonFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00B8E4F1&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   630
      TabIndex        =   25
      Top             =   5340
      Width           =   6525
      Begin VB.CommandButton cmdExit_12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   5310
         Picture         =   "frmBookRetailers.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   90
         Width           =   1095
      End
      Begin VB.CommandButton cmdPrint_7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Print Label"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   4245
         Picture         =   "frmBookRetailers.frx":0BF0
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   90
         Width           =   1005
      End
      Begin VB.CommandButton cmdEdit_4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   3165
         Picture         =   "frmBookRetailers.frx":17D4
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   75
         Width           =   1005
      End
      Begin VB.CommandButton cmdDelete_3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   2100
         Picture         =   "frmBookRetailers.frx":1C16
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   75
         Width           =   1005
      End
      Begin VB.CommandButton cmdSave_2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   1050
         Picture         =   "frmBookRetailers.frx":27FA
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   75
         Width           =   1005
      End
      Begin VB.CommandButton cmdAdd_1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   -15
         Picture         =   "frmBookRetailers.frx":33DE
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   75
         Width           =   1005
      End
   End
   Begin VB.ComboBox cbogp 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "frmBookRetailers.frx":3FC2
      Left            =   2925
      List            =   "frmBookRetailers.frx":3FC9
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   270
      Width           =   4110
   End
   Begin VB.ComboBox cbostate 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2925
      TabIndex        =   7
      Top             =   4635
      Width           =   4110
   End
   Begin VB.ComboBox cboCity 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2925
      TabIndex        =   6
      Top             =   4140
      Width           =   4110
   End
   Begin VB.TextBox txtPin 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2925
      TabIndex        =   5
      Top             =   3645
      Width           =   2760
   End
   Begin VB.TextBox txtadd 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2925
      TabIndex        =   4
      Top             =   3195
      Width           =   8790
   End
   Begin VB.TextBox txtMobile2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2925
      TabIndex        =   3
      Top             =   2745
      Width           =   2760
   End
   Begin VB.TextBox txtMobile1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2925
      TabIndex        =   2
      Top             =   2250
      Width           =   3390
   End
   Begin VB.TextBox txtContact 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2925
      TabIndex        =   1
      Top             =   1800
      Width           =   5190
   End
   Begin VB.TextBox txtNameShop 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2925
      TabIndex        =   0
      Top             =   1305
      Width           =   8790
   End
   Begin VB.TextBox txtsno 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2925
      TabIndex        =   23
      Top             =   720
      Width           =   1140
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   8010
      Top             =   4545
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "F2 For Search"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2970
      TabIndex        =   27
      Top             =   1080
      Width           =   1500
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0078CFE9&
      BorderWidth     =   4
      Height          =   960
      Left            =   495
      Top             =   5265
      Width           =   6765
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Group :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   9
      Left            =   405
      TabIndex        =   24
      Top             =   270
      Width           =   1230
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "State  :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   8
      Left            =   405
      TabIndex        =   22
      Top             =   4680
      Width           =   1995
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "City  :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   7
      Left            =   405
      TabIndex        =   21
      Top             =   4185
      Width           =   1995
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PIN Code  :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   6
      Left            =   405
      TabIndex        =   20
      Top             =   3645
      Width           =   1995
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Address  :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   5
      Left            =   405
      TabIndex        =   19
      Top             =   3195
      Width           =   1995
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile No -2  :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   4
      Left            =   405
      TabIndex        =   18
      Top             =   2745
      Width           =   1995
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile No -1  :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   405
      TabIndex        =   17
      Top             =   2250
      Width           =   1995
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Number :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   405
      TabIndex        =   16
      Top             =   1800
      Width           =   1995
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name Of The Shop :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   405
      TabIndex        =   15
      Top             =   1350
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "S.No :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   405
      TabIndex        =   14
      Top             =   720
      Width           =   1230
   End
End
Attribute VB_Name = "frmBookRetailers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()

End Sub

Private Sub Check1_import_Click()

If Check1_import.value = 1 Then
   cmdAddFile.Visible = True
   cmdImportExcel.Visible = True
   txtpath.Visible = True
Else
   cmdAddFile.Visible = False
   cmdImportExcel.Visible = False
   txtpath.Visible = False

End If

End Sub

Private Sub cmdAdd_1_Click()
clearForm
End Sub

Private Sub cmdAddFile_Click()
cd.ShowOpen
txtpath.Text = cd.filename

End Sub

Private Sub cmdEdit_4_Click()

cmdEdit_4.Enabled = False
cmdSave_2.Enabled = True
cmdDelete_3.Enabled = True

cmdSave_2.SetFocus


End Sub

Private Sub cmdExit_12_Click()
Unload Me
End Sub
Private Sub cmdImportExcel_Click()

Exit Sub

Dim sconn As String
Dim I As Integer

sFile = Me.txtpath.Text
sconn = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=" & sFile

txtTotQty = 0
txtBillQty = 0

Dim rs_fatch As New ADODB.Recordset
Dim rs_em As New ADODB.Recordset


I = 0
k1 = 1


If RS.State = 1 Then RS.close
RS.Open "select * from BookRetailer where sno=" & txtsno.Text & "", con, adOpenDynamic, adLockOptimistic
If RS.EOF = True Then
   RS.AddNew
End If
   
RS!gp = cbogp.Text
RS!sno = txtsno.Text
RS!NameOfShop = txtNameShop.Text
RS!ContactNo = txtContact.Text
RS!MobileNo1 = txtMobile1.Text
RS!MobileNo2 = txtMobile2.Text
RS!Address = txtadd.Text
RS!PinCode = txtPin.Text
RS!city = cboCity.Text
RS!State = cbostate.Text
RS.update


If RS.State = 1 Then RS.close
RS.Open "SELECT * FROM [sheet1$]", sconn
While RS.EOF = False

  




v1 = IIf(IsNull(RS(2)), 0, RS(2))
v2 = IIf(IsNull(RS(3)), 0, RS(3))

 If (v1 > 0 Or v2 > 0) Then
    
    
'========================================
If Option1_sale.value = True Then
          
If InStr(txtParty, "(SUNDRY)") = 0 Then

    rs_em.MoveFirst
    rs_em.Find "bookcode='" & UCase(RS(0)) & "'"
    If rs_em.EOF = False Then
             
    If party_type = "EM" Then
        If party_type <> rs_em(0) Then
             MsgBox "Book List is not valid for this customer ...", vbCritical
             Exit Sub
             ''GoTo abc:
        End If
    Else
        If rs_em(0) = "EM" Then
             MsgBox "Book List is not valid for this customer ...", vbCritical
             Exit Sub
             ''GoTo abc:
        End If
    End If
 End If

End If
End If
'=======================================
    
    
    
    
    vs.TextMatrix(k1, 0) = k1
    vs.TextMatrix(k1, 1) = UCase(RS(0))
    vs.TextMatrix(k1, 2) = UCase(RS(1))
    If Option1_sale.value = False Then
       vs.TextMatrix(k1, 4) = v2
       txtTotQty = Val(txtTotQty) + v2
    Else
       
       vs.TextMatrix(k1, 3) = v1
       vs.TextMatrix(k1, 4) = v2
       txtTotQty = Val(txtTotQty) + v1
       txtBillQty = Val(txtBillQty) + v2
    End If
    '------------------------------------------------------------
      If rs1.State = 1 Then rs1.close
      rs1.Open "select Bookcode,bookname,rate,DISCOUNT from BOOKS where bookcode='" & vs.TextMatrix(k1, 1) & "'", con
      If rs1.EOF = False Then
         vs.TextMatrix(k1, 6) = rs1!rate
         
         qty = 0
         qty = IIf(vs.TextMatrix(k1, 3) = "", 0, vs.TextMatrix(k1, 3))
         qty = Val(qty) + Val(IIf(vs.TextMatrix(k1, 4) = "", 0, vs.TextMatrix(k1, 4)))

         
         net = (qty * Val(vs.TextMatrix(k1, 6)))
         
         vs.TextMatrix(k1, 9) = rs1!discount
         
         dis = rs1!discount
         If Option1_sale.value = True Then
            rs_fatch.MoveFirst
            rs_fatch.Find "bookcode='" & RS(0) & "'"
            If rs_fatch.EOF = False Then
               vs.TextMatrix(k1, 9) = rs_fatch!discount
               dis = rs_fatch!discount
            End If
         End If
         
         vs.TextMatrix(k1, 7) = net - Format(Round(net * (dis / 100), 2), "0.00")
         
         vs.TextMatrix(k1, 11) = net
      End If
          
    '------------------------------------------------------------


    k1 = k1 + 1
 End If



RS.MoveNext
Wend



MsgBox "Data import Successfully", vbInformation

End Sub

Private Sub cmdPrint_7_Click()
Screen.MousePointer = vbHourglass

DSNNew

If (cboCity.Text <> "") Then

    cr.Reset
    cr.ReportFileName = rptPath & "/BOOKRETAILS.rpt"
    cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    cr.ReplaceSelectionFormula "{BookRetailer.city}='" & cboCity.Text & "'"
    cr.WindowShowPrintSetupBtn = True
    cr.WindowState = crptMaximized
    cr.Action = 1
    
ElseIf (cbostate.Text <> "") Then

    cr.Reset
    cr.ReportFileName = rptPath & "/BOOKRETAILS.rpt"
    cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    cr.ReplaceSelectionFormula "{BookRetailer.state}='" & cbostate.Text & "'"
    cr.WindowShowPrintSetupBtn = True
    cr.WindowState = crptMaximized
    cr.Action = 1
    
Else
    cr.Reset
    cr.ReportFileName = rptPath & "/BOOKRETAILS.rpt"
    cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    cr.WindowShowPrintSetupBtn = True
    cr.WindowState = crptMaximized
    cr.Action = 1

End If

Screen.MousePointer = vbDefault

End Sub

Private Sub cmdSave_2_Click()
 
 
If RS.State = 1 Then RS.close
RS.Open "select * from BookRetailer where sno=" & txtsno.Text & "", con, adOpenDynamic, adLockOptimistic
If RS.EOF = True Then
   RS.AddNew
End If
   
   RS!gp = cbogp.Text
   RS!sno = txtsno.Text
   RS!NameOfShop = txtNameShop.Text
   RS!ContactNo = txtContact.Text
   RS!MobileNo1 = txtMobile1.Text
   RS!MobileNo2 = txtMobile2.Text
   RS!Address = txtadd.Text
   RS!PinCode = txtPin.Text
   RS!city = cboCity.Text
   RS!State = cbostate.Text
   RS.update


MsgBox "Data Updated...", vbInformation



 
 
 
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdUpLoad_Click()

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub Form_Load()

Me.Height = 7770
Me.Width = 12470

cbogp.ListIndex = 0

addData
Max_Sno
   
BackColorFrom Me
   
End Sub
Sub Max_Sno()

Screen.MousePointer = vbHourglass

If rs1.State = 1 Then rs1.close
rs1.Open "select max(sno) from BookRetailer"
If IsNull(rs1(0)) Then
   txtsno.Text = 1
Else
   txtsno.Text = rs1(0) + 1
End If


Screen.MousePointer = vbDefault

End Sub

Sub addData()

Screen.MousePointer = vbHourglass

If rs1.State = 1 Then rs1.close
rs1.Open "select City from BookRetailer group by City order by City"
While rs1.EOF = False

cboCity.AddItem rs1(0)
rs1.MoveNext
Wend

If rs1.State = 1 Then rs1.close
rs1.Open "select [state] from BookRetailer group by [state] order by [state]"
While rs1.EOF = False

cbostate.AddItem rs1(0)
rs1.MoveNext
Wend


Screen.MousePointer = vbDefault

End Sub
Sub clearForm()
    

    
'cbogp.ListIndex = -1
txtsno.Text = ""
txtNameShop.Text = ""
txtContact.Text = ""
txtMobile1.Text = ""
txtMobile2.Text = ""

txtadd.Text = ""
txtPin.Text = ""

cboCity.Text = ""
cbostate.Text = ""

cmdEdit_4.Enabled = False
cmdDelete_3.Enabled = False

Max_Sno

txtNameShop.SetFocus

End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtNameShop_GotFocus()
  
  If PopUpValue1 <> "" Then
     
     txtsno.Text = popupvalue4
     
     If RS.State = 1 Then RS.close
     RS.Open "select * from BookRetailer where sno=" & txtsno.Text & "", con
     If RS.EOF = False Then
        
        cmdEdit_4.Enabled = True
        cmdSave_2.Enabled = False
        cmdDelete_3.Enabled = False

        
        cbogp.Text = RS!gp
        txtsno.Text = RS!sno
        txtNameShop.Text = RS!NameOfShop & ""
        txtContact.Text = RS!ContactNo & ""
        txtMobile1.Text = RS!MobileNo1 & ""
        txtMobile2.Text = RS!MobileNo2 & ""
        
        txtadd.Text = RS!Address & ""
        txtPin.Text = RS!PinCode & ""
        
        cboCity.Text = RS!city
        cbostate.Text = RS!State
        
        PopUpValue1 = ""
        PopUpValue2 = ""
        PopUpValue3 = ""
        popupvalue4 = ""
        
     End If
     
     
  End If
  
End Sub
Private Sub txtNameShop_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then
    
    searchType = "bookretailer"
    popuplistFast "select NameOfShop,City,State,SNO from BookRetailer order by NameOfShop,City,State", con, , , "bookretailer"

End If

End Sub

