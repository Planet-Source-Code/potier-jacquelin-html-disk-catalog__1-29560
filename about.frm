VERSION 5.00
Begin VB.Form about 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4230
   Icon            =   "about.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4230
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
      Height          =   255
      Left            =   1388
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   120
      Picture         =   "about.frx":030A
      Top             =   120
      Width           =   1110
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Jacquelin POTIER jacquelin.potier@libertysurf.fr"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "HTML Disk Catalog is a freeware under GPL licence made by"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdok_Click()
    Unload Me
End Sub
