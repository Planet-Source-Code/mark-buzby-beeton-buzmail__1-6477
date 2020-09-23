VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmBuzMail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BuzMail"
   ClientHeight    =   570
   ClientLeft      =   5160
   ClientTop       =   2130
   ClientWidth     =   2355
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmBuzMail.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   570
   ScaleWidth      =   2355
   Visible         =   0   'False
   Begin MSMAPI.MAPISession BuzSession 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSMAPI.MAPIMessages BuzMessages 
      Left            =   540
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   -1  'True
      FetchSorted     =   -1  'True
      FetchUnreadOnly =   0   'False
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "DO NOT TOUCH!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1260
      TabIndex        =   0
      Top             =   60
      Width           =   1035
   End
End
Attribute VB_Name = "frmBuzMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
