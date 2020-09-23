VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "BuzMail Sample App"
   ClientHeight    =   1320
   ClientLeft      =   2010
   ClientTop       =   3585
   ClientWidth     =   4095
   LinkTopic       =   "Form2"
   ScaleHeight     =   1320
   ScaleWidth      =   4095
   Begin VB.CommandButton btnInbox 
      Caption         =   "&View Inbox"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   300
      Width           =   1215
   End
   Begin VB.CommandButton btnSend 
      Caption         =   "&Send Mail"
      Height          =   495
      Left            =   300
      TabIndex        =   0
      Top             =   300
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Success As Boolean

Private Sub Command1_Click()

End Sub


Private Sub btnInbox_Click()
Load frmInbox
frmInbox.Show 1

End Sub

Private Sub btnSend_Click()
Load Form1
Form1.Show 1

End Sub


Private Sub Form_Load()
Dim UserName As String
UserName = InputBox("Enter the Windows messaging profile name you want to use:", "Logon", "")
If UserName = "" Then
    Unload Me
    End
End If

Success = MailLogon(UserName, "")
End Sub


Private Sub Form_Unload(Cancel As Integer)
Success = MailLogoff
End Sub


