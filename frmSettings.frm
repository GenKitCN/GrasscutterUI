VERSION 5.00
Begin VB.Form frmSettings 
   BackColor       =   &H80000005&
   Caption         =   "���ݻ� - ����"
   ClientHeight    =   2835
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3165
   BeginProperty Font 
      Name            =   "΢���ź�"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2835
   ScaleWidth      =   3165
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox txtLogCount 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "��־��������:"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
txtLogCount.Text = CStr(LogCount)
End Sub

Private Sub txtLogCount_Change()
LogCount = CInt(txtLogCount.Text)
WriteIni "GCUI", "LogCount", txtLogCount.Text, App.Path & "\Config.ini"
End Sub
