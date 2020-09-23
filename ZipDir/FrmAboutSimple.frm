VERSION 5.00
Begin VB.Form FrmAboutSimple 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sobre..."
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.PictureBox PicLogo 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      ScaleHeight     =   615
      ScaleWidth      =   735
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Label LSite 
      AutoSize        =   -1  'True
      Caption         =   "http://www.mcunha98.cjb.net"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   2145
   End
   Begin VB.Label LApp 
      Caption         =   "Label1"
      Height          =   675
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   4080
   End
End
Attribute VB_Name = "FrmAboutSimple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const CorUrl = vbBlue

Private Sub Command1_Click()
 Unload Me
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 LSite.ForeColor = vbButtonText
End Sub

Private Sub Form_Load()
 LApp.Caption = App.Title & vbCrLf & "Versão: " & App.Major & "." & Format(App.Revision, "00")
 PicLogo.Picture = FrmMenu.Icon
 Me.Icon = FrmMenu.Icon
 MyLanguage = GetSetting(App.EXEName, "last", "language", 0)
 Me.Caption = Replace(LoadResString(MyLanguage + 157), "&", "")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 LSite.ForeColor = vbButtonText
End Sub

Private Sub LApp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 LSite.ForeColor = vbButtonText
End Sub

Private Sub LSite_Click()
 Call ShellExecute(Me.hWnd, "Open", LSite.Caption, "", CurDir, 0)
End Sub

Private Sub LSite_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 LSite.ForeColor = CorUrl
End Sub

Private Sub PicLogo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 LSite.ForeColor = vbButtonText
End Sub
