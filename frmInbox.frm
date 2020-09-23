VERSION 5.00
Begin VB.Form frmInbox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inbox Sample"
   ClientHeight    =   2625
   ClientLeft      =   1710
   ClientTop       =   2955
   ClientWidth     =   7740
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnView 
      Caption         =   "&View"
      Height          =   495
      Left            =   6420
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.ListBox lstInbox 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7515
   End
End
Attribute VB_Name = "frmInbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnView_Click()
Dim MsgNum As Long
Dim tmpMail As Mail
Dim Dummy As Integer

If lstInbox.Text = "" Then Exit Sub

MsgNum = lstInbox.ItemData(lstInbox.ListIndex)

tmpMail = MailGet(MsgNum)

'display message
Load Form1
Form1.lblRecips.Caption = "From:"
Form1.btnBook.Visible = False
Form1.btnSend.Visible = False
Form1.Caption = "View Sample"

Form1.txtSubject.Text = tmpMail.Subject
Form1.txtText.Text = tmpMail.Text

Form1.lstRecips.AddItem tmpMail.From.Name

Form1.Show 1

End Sub

Private Sub Form_Load()
Dim NumOfMsgs As Long
Dim Dummy As Long
'temporary mail object
Dim tmpMail As Mail

'get the number of messages in mailbox
NumOfMsgs = MailCount(False)

lstInbox.Clear
If NumOfMsgs > 0 Then
    'read each message (in reverse order to get most recent first)
    For Dummy = NumOfMsgs To 1 Step -1
        tmpMail = MailGet(Dummy)
        lstInbox.AddItem tmpMail.From.Name + vbTab + vbTab + tmpMail.Subject
        'add the message number to the itemdata property for quick access later
        lstInbox.ItemData(lstInbox.NewIndex) = Dummy
    Next
End If


End Sub


