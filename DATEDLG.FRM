VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form datedlg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Date"
   ClientHeight    =   2490
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   Begin MSACAL.Calendar Calendar1 
      Height          =   2475
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4485
      _Version        =   524288
      _ExtentX        =   7911
      _ExtentY        =   4366
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2001
      Month           =   5
      Day             =   30
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "datedlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FNAME As String
Dim TEXT As String
Option Explicit

Private Sub Calendar1_Click()
    If UCase(Trim(FNAME)) = UCase(Trim("Voucherform")) Then
        Voucherform.vdate.TEXT = Calendar1.Value
        Voucherform.Enabled = True
        Voucherform.Grid1.SetFocus
        Call Voucherform.getval
    End If
    If UCase(Trim(FNAME)) = UCase(Trim("INVOICE")) And UCase(Trim(TEXT)) = UCase(Trim("I_DT")) Then
        INVOICE.i_dt.TEXT = Calendar1.Value
        INVOICE.Enabled = True
        INVOICE.I_OB.SetFocus
    End If
    Unload Me
    
    
End Sub

Private Sub Form_GotFocus()
'    Calendar1.SetFocus
End Sub

Public Function GETDATE(FORMNAME As String, TEXTNAME As String)
        FNAME = FORMNAME
        TEXT = TEXTNAME
End Function

Private Sub Form_Load()
    Me.Hide
End Sub
