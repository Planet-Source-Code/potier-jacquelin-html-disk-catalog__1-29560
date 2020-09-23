VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form bgoptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Background Options"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   Icon            =   "bgoptions.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4950
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2400
      TabIndex        =   12
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   2775
      Left            =   -120
      TabIndex        =   13
      Top             =   -120
      Width           =   4935
      Begin VB.TextBox txtbgpicture 
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   600
         Width           =   2655
      End
      Begin VB.ComboBox combvpostype 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "bgoptions.frx":030A
         Left            =   3840
         List            =   "bgoptions.frx":0314
         TabIndex        =   10
         Text            =   "pixels"
         Top             =   2280
         Width           =   855
      End
      Begin VB.ComboBox combhpostype 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "bgoptions.frx":0323
         Left            =   3840
         List            =   "bgoptions.frx":032D
         TabIndex        =   8
         Text            =   "pixels"
         Top             =   1920
         Width           =   855
      End
      Begin VB.ComboBox combvpos 
         Height          =   315
         ItemData        =   "bgoptions.frx":033C
         Left            =   2400
         List            =   "bgoptions.frx":034C
         TabIndex        =   9
         Top             =   2280
         Width           =   1215
      End
      Begin VB.ComboBox combhpos 
         Height          =   315
         ItemData        =   "bgoptions.frx":036E
         Left            =   2400
         List            =   "bgoptions.frx":037E
         TabIndex        =   7
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CheckBox chkbgcolor 
         Caption         =   "Color"
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   735
      End
      Begin VB.PictureBox Picturebgcolor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   1
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox chkpicture 
         Caption         =   "Picture"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton cmdpicturebrowse 
         Caption         =   "browse"
         Height          =   255
         Left            =   3960
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.ComboBox combrepeat 
         Height          =   315
         ItemData        =   "bgoptions.frx":03A0
         Left            =   2400
         List            =   "bgoptions.frx":03B0
         TabIndex        =   5
         Top             =   1080
         Width           =   1215
      End
      Begin VB.ComboBox combattach 
         Height          =   315
         ItemData        =   "bgoptions.frx":03DB
         Left            =   2400
         List            =   "bgoptions.frx":03E5
         TabIndex        =   6
         Top             =   1440
         Width           =   1215
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3960
         Top             =   1080
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label3 
         Caption         =   "Repeat"
         Height          =   255
         Left            =   720
         TabIndex        =   17
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Attachement"
         Height          =   255
         Left            =   720
         TabIndex        =   16
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Horizontal Position"
         Height          =   255
         Left            =   720
         TabIndex        =   15
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Vertical Position"
         Height          =   255
         Left            =   720
         TabIndex        =   14
         Top             =   2280
         Width           =   1335
      End
   End
End
Attribute VB_Name = "bgoptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    On Error Resume Next
    If Me.chkbgcolor.Value = vbChecked Then
        style_background.use_color = True
    Else
        style_background.use_color = False
    End If
    
    If Me.chkpicture.Value = vbChecked Then
        style_background.use_picture = True
    Else
        style_background.use_picture = False
    End If
    
    style_background.color = Me.Picturebgcolor.BackColor
    style_background.picture_name = Me.txtbgpicture
    style_background.repeat = Me.combrepeat.Text
    style_background.attachement = Me.combattach.Text
    style_background.hpos = Me.combhpos.Text
    style_background.hpostype = Me.combhpostype.Text
    style_background.vpos = Me.combvpos.Text
    style_background.vpostype = Me.combvpostype.Text

    Unload Me
End Sub

Private Sub cmdpicturebrowse_Click()
    Me.CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    Me.CommonDialog1.Filter = "Pictures (*.gif;*.jpg;*.png)|*.gif;*.jpg;*.png| All (*.*) |*.*"
    Me.CommonDialog1.Flags = cdlOFNFileMustExist
    Me.CommonDialog1.ShowOpen
    Me.txtbgpicture.Text = Me.CommonDialog1.filename
    Exit Sub

ErrHandler:
   Exit Sub
End Sub

Private Sub combhpos_Click()
    If Me.combhpos.Text <> "left" And Me.combhpos.Text <> "center" And Me.combhpos.Text <> "right" Then
        Me.combhpostype.Enabled = True
    Else
        Me.combhpostype.Enabled = False
    End If
End Sub



Private Sub combvpos_Click()
    If Me.combvpos.Text <> "top" And Me.combvpos.Text <> "center" And Me.combvpos.Text <> "bottom" Then
        Me.combvpostype.Enabled = True
    Else
        Me.combvpostype.Enabled = False
    End If
End Sub


Private Sub Form_Load()
    On Error Resume Next
    If style_background.use_color Then Me.chkbgcolor.Value = vbChecked
    If style_background.use_picture Then Me.chkpicture.Value = vbChecked
    Me.Picturebgcolor.BackColor = style_background.color
    Me.txtbgpicture = style_background.picture_name
    Me.combrepeat.Text = style_background.repeat
    Me.combattach.Text = style_background.attachement
    Me.combhpos.Text = style_background.hpos
    Me.combhpostype.Text = style_background.hpostype
    Me.combvpos.Text = style_background.vpos
    Me.combvpostype.Text = style_background.vpostype
End Sub

Private Sub Picturebgcolor_Click()
   Dim tmp As Long
   Me.CommonDialog1.CancelError = True
   On Error GoTo ErrHandler
   Me.CommonDialog1.Flags = cdlCCRGBInit
   Me.CommonDialog1.ShowColor
   tmp = CommonDialog1.color
   Me.Picturebgcolor.BackColor = tmp
   Exit Sub

ErrHandler:
   Exit Sub
End Sub
