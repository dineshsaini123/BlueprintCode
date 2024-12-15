VERSION 5.00
Begin VB.Form frmPriceMasterCourier 
   Caption         =   "Price Master Courier..."
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   7548
   Icon            =   "frmPriceMasterCourier.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4500
   ScaleWidth      =   7548
   Begin VB.ComboBox cboAName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      ItemData        =   "frmPriceMasterCourier.frx":000C
      Left            =   2520
      List            =   "frmPriceMasterCourier.frx":000E
      TabIndex        =   14
      Top             =   756
      Width           =   4344
   End
   Begin VB.CommandButton Command13 
      Appearance      =   0  'Flat
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   6930
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1224
      Width           =   495
   End
   Begin VB.CommandButton cmdExit_12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "E&xit"
      Height          =   660
      Left            =   4230
      Picture         =   "frmPriceMasterCourier.frx":0010
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3444
      Width           =   1005
   End
   Begin VB.CommandButton Commandsave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sa&ve"
      Height          =   645
      Left            =   1152
      Picture         =   "frmPriceMasterCourier.frx":0BF4
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3444
      Width           =   990
   End
   Begin VB.CommandButton Commanddelete 
      BackColor       =   &H00FFFFFF&
      Caption         =   "De&lete"
      Height          =   645
      Left            =   2190
      Picture         =   "frmPriceMasterCourier.frx":17D8
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3444
      Width           =   990
   End
   Begin VB.CommandButton Commandsearch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Search"
      Height          =   645
      Left            =   3210
      Picture         =   "frmPriceMasterCourier.frx":23BC
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3444
      Width           =   990
   End
   Begin VB.CommandButton Commandadd 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Add"
      Height          =   645
      Left            =   135
      Picture         =   "frmPriceMasterCourier.frx":2FA0
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3444
      Width           =   990
   End
   Begin VB.TextBox txtkgCharge 
      Height          =   375
      Left            =   2520
      MaxLength       =   4
      TabIndex        =   3
      Top             =   2436
      Width           =   1140
   End
   Begin VB.TextBox txtMinCharge 
      Height          =   375
      Left            =   2520
      MaxLength       =   4
      TabIndex        =   2
      Top             =   1860
      Width           =   1140
   End
   Begin VB.ComboBox cboPlace 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      ItemData        =   "frmPriceMasterCourier.frx":3B84
      Left            =   2520
      List            =   "frmPriceMasterCourier.frx":3B86
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1320
      Width           =   4335
   End
   Begin VB.ComboBox cboCourier 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      ItemData        =   "frmPriceMasterCourier.frx":3B88
      Left            =   2520
      List            =   "frmPriceMasterCourier.frx":3B8A
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   270
      Width           =   4335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Agency Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   108
      TabIndex        =   15
      Top             =   756
      Width           =   1860
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Per kg/gm Charges:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   132
      TabIndex        =   12
      Top             =   2484
      Width           =   2580
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Min. Charges:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   132
      TabIndex        =   11
      Top             =   1944
      Width           =   1860
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Place of Supply :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   132
      TabIndex        =   10
      Top             =   1320
      Width           =   2088
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Courier Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   135
      TabIndex        =   9
      Top             =   270
      Width           =   1860
   End
End
Attribute VB_Name = "frmPriceMasterCourier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cboAName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   cboPlace.SetFocus
End If
End Sub

Private Sub cboCourier_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   cboAName.SetFocus
End If
End Sub


Private Sub cboPlace_Click()
 'search
End Sub

Private Sub cboPlace_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
txtMinCharge.SetFocus
End If
End Sub

Private Sub cboPlace_LostFocus()
search_Data
End Sub

Private Sub cmdExit_12_Click()
Unload Me
End Sub

Private Sub Command13_Click()
HeadTbl = "cplace"
frmMasters.Show 1
End Sub

Private Sub Commandadd_Click()

cboCourier.ListIndex = -1
cboPlace.ListIndex = -1
txtMinCharge.text = ""
txtkgCharge.text = ""
cboAName.ListIndex = -1
cboCourier.SetFocus

Commandsave.Enabled = True
Commanddelete.Enabled = False

End Sub

Private Sub Commanddelete_Click()
 If MsgBox("Want to Delete ?", vbQuestion + vbYesNo) = vbYes Then
    con.Execute "delete from CourierPriceMaster where (CourierMaster='" & cboCourier & "' and PlaceOfSupp='" & cboPlace & "')"
    Commandadd_Click
 End If
End Sub
Sub search_Data()

If RS.State = 1 Then RS.close
RS.Open "select  * from CourierPriceMaster where (CourierMaster='" & cboCourier & "' and PlaceOfSupp='" & cboPlace & "' and AgencyName='" & cboAName.text & "')", con, adOpenDynamic, adLockOptimistic
If RS.EOF = False Then

cboCourier.text = RS!CourierMaster & ""
cboPlace.text = RS!PlaceOfSupp & ""
txtMinCharge.text = RS!ChargePerMin
txtkgCharge.text = RS!ChargePerKG
cboAName.text = RS!AgencyName & ""

Else

'cboCourier.ListIndex = -1
'cboPlace.ListIndex = -1
txtMinCharge.text = ""
txtkgCharge.text = ""
'cboAName.text = ""

End If


  
End Sub



Private Sub Commandsave_Click()

If RS.State = 1 Then RS.close
RS.Open "select  * from CourierPriceMaster where (CourierMaster='" & cboCourier & "' and PlaceOfSupp='" & cboPlace & "')", con, adOpenDynamic, adLockOptimistic
If RS.EOF = True Then
RS.AddNew
End If
RS!CourierMaster = cboCourier.text
RS!PlaceOfSupp = cboPlace.text
RS!ChargePerMin = IIf(txtMinCharge.text = "", 0, txtMinCharge.text)
RS!ChargePerKG = IIf(txtkgCharge.text = "", 0, txtkgCharge.text)
RS!AgencyName = cboAName.text
RS.update

Commandsave.Enabled = False
Commandadd.SetFocus
  
End Sub

Private Sub Commandsearch_Click()
popuplist1 "Select distinct CourierMaster,PlaceOfSupp,ChargePerMin,ChargePerKG,AgencyName from CourierPriceMaster order by CourierMaster,PlaceOfSupp", con
End Sub
Private Sub Commandsearch_GotFocus()

If PopUpValue1 <> "" Then
    cboCourier.text = PopUpValue1
    cboPlace.text = PopUpValue2
    txtMinCharge.text = PopUpValue3
    txtkgCharge.text = popupvalue4
    cboAName.text = popupvalue5
    
    Commanddelete.Enabled = True
    
    PopUpValue1 = ""
    PopUpValue2 = ""
    PopUpValue3 = ""
    popupvalue4 = ""
    popupvalue5 = ""
End If

End Sub

Private Sub Form_Load()

Me.top = 500
Me.Width = 7905
Me.Height = 5475

filecombo
BackColorFrom Me

End Sub
Sub filecombo()
cboCourier.Clear
cboPlace.Clear
cboAName.Clear

If RS.State = 1 Then RS.close
RS.Open "select  name from MasterTbl where category='aname'", con
While RS.EOF = False
    cboCourier.AddItem RS(0)
    RS.MoveNext
Wend

If RS.State = 1 Then RS.close
RS.Open "select  name from MasterTbl where category='aname'", con
While RS.EOF = False
    cboAName.AddItem RS(0)
    RS.MoveNext
Wend

If RS.State = 1 Then RS.close
RS.Open "select  name from MasterTbl where category='cplace'", con
While RS.EOF = False
    cboPlace.AddItem RS(0)
    RS.MoveNext
Wend
End Sub

Private Sub txtkgCharge_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
   Commandsave_Click
End If

End Sub

Private Sub txtMinCharge_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   txtkgCharge.SetFocus
End If
End Sub
