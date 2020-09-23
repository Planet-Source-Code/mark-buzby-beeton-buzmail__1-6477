VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sample Send"
   ClientHeight    =   3915
   ClientLeft      =   2925
   ClientTop       =   2520
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstRecips 
      Height          =   840
      Left            =   1080
      TabIndex        =   7
      Top             =   120
      Width           =   3075
   End
   Begin VB.CommandButton btnSend 
      Caption         =   "&Send"
      Height          =   495
      Left            =   5340
      TabIndex        =   6
      Top             =   3300
      Width           =   1215
   End
   Begin VB.TextBox txtText 
      Height          =   1665
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1560
      Width           =   5475
   End
   Begin VB.TextBox txtSubject 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   1200
      Width           =   3075
   End
   Begin VB.CommandButton btnBook 
      Caption         =   "&Add Recips"
      Height          =   315
      Left            =   4260
      TabIndex        =   1
      Top             =   120
      Width           =   1275
   End
   Begin VB.Label Label3 
      Caption         =   "Text:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Subject:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblRecips 
      Caption         =   "Recipients:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'set up mail 'object' for this form
Dim MyMail As Mail

'couple of return variables
Dim Success As Boolean
Dim SuccessInt As Integer

Private Sub btnBook_Click()
Dim Dummy As Integer

MailAddress MyMail
lstRecips.Clear

If MyMail.RecipCount > 0 Then
    For Dummy = 1 To MyMail.RecipCount
        lstRecips.AddItem MyMail.Recips(Dummy).Name
    Next
End If
End Sub


Private Sub btnSend_Click()
'add the subject and text
MailContents MyMail, txtSubject.Text, txtText.Text

'send the mail
SuccessInt = MailSend(MyMail)

Select Case SuccessInt
Case 0
    MsgBox "Mail send succesfully."
    Unload Me
Case Is > 0
    MsgBox "Can't resolve recipient " + MyMail.Recips(SuccessInt).Name
Case -1
    MsgBox "No recipients"
Case -2
    MsgBox "Not logged on"
End Select
    
End Sub


