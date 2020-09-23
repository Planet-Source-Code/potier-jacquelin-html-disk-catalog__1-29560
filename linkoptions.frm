VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form linkoptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Links Options"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4320
   Icon            =   "linkoptions.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   4320
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Visited"
      Height          =   2295
      Left            =   120
      TabIndex        =   24
      Top             =   4680
      Width           =   4095
      Begin VB.CheckBox chklinethroughvisited 
         Caption         =   "Line-through"
         Height          =   255
         Left            =   1680
         TabIndex        =   31
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtsizevisited 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   30
         Top             =   360
         Width           =   615
      End
      Begin VB.PictureBox Picturevisited 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   29
         Top             =   360
         Width           =   255
      End
      Begin VB.CheckBox chkboldvisited 
         Caption         =   "Bold"
         Height          =   255
         Left            =   1680
         TabIndex        =   28
         Top             =   840
         Width           =   1095
      End
      Begin VB.CheckBox chkitalicvisited 
         Caption         =   "Italic"
         Height          =   255
         Left            =   1680
         TabIndex        =   27
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CheckBox chkunderlinevisited 
         Caption         =   "Underline"
         Height          =   255
         Left            =   1680
         TabIndex        =   26
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CheckBox chknonevisited 
         Caption         =   "none"
         Height          =   255
         Left            =   1680
         TabIndex        =   25
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "size"
         Height          =   255
         Left            =   720
         TabIndex        =   34
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Decoration"
         Height          =   255
         Left            =   600
         TabIndex        =   33
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Color"
         Height          =   255
         Left            =   2160
         TabIndex        =   32
         Top             =   360
         Width           =   495
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3360
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2280
      TabIndex        =   15
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Hover"
      Height          =   2175
      Left            =   120
      TabIndex        =   20
      Top             =   2400
      Width           =   4095
      Begin VB.CheckBox chknonehover 
         Caption         =   "None"
         Height          =   195
         Left            =   1680
         TabIndex        =   13
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CheckBox chkunderlinehover 
         Caption         =   "Underline"
         Height          =   255
         Left            =   1680
         TabIndex        =   11
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CheckBox chkitalichover 
         Caption         =   "Italic"
         Height          =   255
         Left            =   1680
         TabIndex        =   10
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CheckBox chkboldhover 
         Caption         =   "Bold"
         Height          =   255
         Left            =   1680
         TabIndex        =   9
         Top             =   840
         Width           =   1095
      End
      Begin VB.PictureBox Picturehover 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   8
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox txtsizehover 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Top             =   360
         Width           =   615
      End
      Begin VB.CheckBox chklinethroughhover 
         Caption         =   "Line-through"
         Height          =   255
         Left            =   1680
         TabIndex        =   12
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Color"
         Height          =   255
         Left            =   2160
         TabIndex        =   23
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Decoration"
         Height          =   255
         Left            =   600
         TabIndex        =   22
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "size"
         Height          =   255
         Left            =   720
         TabIndex        =   21
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Link"
      Height          =   2295
      Left            =   120
      TabIndex        =   16
      Top             =   0
      Width           =   4095
      Begin VB.CheckBox chknonelink 
         Caption         =   "none"
         Height          =   255
         Left            =   1680
         TabIndex        =   6
         Top             =   1800
         Width           =   735
      End
      Begin VB.CheckBox chkunderlinelink 
         Caption         =   "Underline"
         Height          =   255
         Left            =   1680
         TabIndex        =   4
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CheckBox chkitaliclink 
         Caption         =   "Italic"
         Height          =   255
         Left            =   1680
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CheckBox chkboldlink 
         Caption         =   "Bold"
         Height          =   255
         Left            =   1680
         TabIndex        =   2
         Top             =   840
         Width           =   1095
      End
      Begin VB.PictureBox Picturelink 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   1
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox txtlinksize 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   360
         Width           =   615
      End
      Begin VB.CheckBox chklinethroughlink 
         Caption         =   "Line-through"
         Height          =   255
         Left            =   1680
         TabIndex        =   5
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Color"
         Height          =   255
         Left            =   2160
         TabIndex        =   19
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Decoration"
         Height          =   255
         Left            =   600
         TabIndex        =   18
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "size"
         Height          =   255
         Left            =   720
         TabIndex        =   17
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
      Height          =   255
      Left            =   840
      TabIndex        =   14
      Top             =   7080
      Width           =   1335
   End
End
Attribute VB_Name = "linkoptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chknonehover_Click()
    If Me.chknonehover.Value = vbChecked Then
        Me.chkunderlinehover.Enabled = False
        Me.chklinethroughhover.Enabled = False
    Else
        Me.chkunderlinehover.Enabled = True
        Me.chklinethroughhover.Enabled = True
    End If
End Sub

Private Sub chknonelink_Click()
    If Me.chknonelink.Value = vbChecked Then
        Me.chkunderlinelink.Enabled = False
        Me.chklinethroughlink.Enabled = False
    Else
        Me.chkunderlinelink.Enabled = True
        Me.chklinethroughlink.Enabled = True
    End If
End Sub

Private Sub chknonevisited_Click()
    If Me.chknonevisited.Value = vbChecked Then
        Me.chkunderlinevisited.Enabled = False
        Me.chklinethroughvisited.Enabled = False
    Else
        Me.chkunderlinevisited.Enabled = True
        Me.chklinethroughvisited.Enabled = True
    End If
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    On Error Resume Next
    style_link.size = val(Me.txtlinksize)
    style_link.color = Me.Picturelink.BackColor
    If Me.chkboldlink.Value = vbChecked Then
        style_link.bold = True
    Else
        style_link.bold = False
    End If
    If Me.chklinethroughlink.Value = vbChecked Then
        style_link.line_through = True
    Else
        style_link.line_through = False
    End If
    If Me.chkitaliclink.Value = vbChecked Then
        style_link.italic = True
    Else
        style_link.italic = False
    End If
    If Me.chknonelink.Value = vbChecked Then
        style_link.none = True
    Else
        style_link.none = False
    End If
    If Me.chkunderlinelink.Value = vbChecked Then
        style_link.underline = True
    Else
        style_link.underline = False
    End If

    style_hover.size = val(Me.txtsizehover)
    style_hover.color = Me.Picturehover.BackColor
    If Me.chkboldhover.Value = vbChecked Then
        style_hover.bold = True
    Else
        style_hover.bold = False
    End If
    If Me.chklinethroughhover.Value = vbChecked Then
        style_hover.line_through = True
    Else
        style_hover.line_through = False
    End If
    If Me.chkitalichover.Value = vbChecked Then
        style_hover.italic = True
    Else
        style_hover.italic = False
    End If
    If Me.chknonehover.Value = vbChecked Then
        style_hover.none = True
    Else
        style_hover.none = False
    End If
    If Me.chkunderlinehover.Value = vbChecked Then
        style_hover.underline = True
    Else
        style_hover.underline = False
    End If



    style_visited.size = val(Me.txtsizevisited)
    style_visited.color = Me.Picturevisited.BackColor
    If Me.chkboldvisited.Value = vbChecked Then
        style_visited.bold = True
    Else
        style_visited.bold = False
    End If
    If Me.chklinethroughvisited.Value = vbChecked Then
        style_visited.line_through = True
    Else
        style_visited.line_through = False
    End If
    If Me.chkitalicvisited.Value = vbChecked Then
        style_visited.italic = True
    Else
        style_visited.italic = False
    End If
    If Me.chknonevisited.Value = vbChecked Then
        style_visited.none = True
    Else
        style_visited.none = False
    End If
    If Me.chkunderlinevisited.Value = vbChecked Then
        style_visited.underline = True
    Else
        style_visited.underline = False
    End If


    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    If style_link.size <> 0 Then Me.txtlinksize = style_link.size
    Me.Picturelink.BackColor = style_link.color
    If style_link.bold Then Me.chkboldlink.Value = vbChecked
    If style_link.italic Then Me.chkitaliclink.Value = vbChecked
    If style_link.underline Then Me.chkunderlinelink.Value = vbChecked
    If style_link.line_through Then Me.chklinethroughlink.Value = vbChecked
    If style_link.none Then Me.chknonelink.Value = vbChecked
    
    
    If style_hover.size <> 0 Then Me.txtsizehover = style_hover.size
    Me.Picturehover.BackColor = style_hover.color
    If style_hover.bold Then Me.chkboldhover.Value = vbChecked
    If style_hover.italic Then Me.chkitalichover.Value = vbChecked
    If style_hover.underline Then Me.chkunderlinehover.Value = vbChecked
    If style_hover.line_through Then Me.chklinethroughhover.Value = vbChecked
    If style_hover.none Then Me.chknonehover.Value = vbChecked
    
    If style_visited.size <> 0 Then Me.txtsizevisited = style_visited.size
    Me.Picturevisited.BackColor = style_visited.color
    If style_visited.bold Then Me.chkboldvisited.Value = vbChecked
    If style_visited.italic Then Me.chkitalicvisited.Value = vbChecked
    If style_visited.underline Then Me.chkunderlinevisited.Value = vbChecked
    If style_visited.line_through Then Me.chklinethroughvisited.Value = vbChecked
    If style_visited.none Then Me.chknonevisited.Value = vbChecked
End Sub

Private Sub Picturehover_Click()
   Dim tmp As Long
   Me.CommonDialog1.CancelError = True
   On Error GoTo ErrHandler
   Me.CommonDialog1.Flags = cdlCCRGBInit
   Me.CommonDialog1.ShowColor
   tmp = CommonDialog1.color
   Me.Picturehover.BackColor = tmp
   Exit Sub

ErrHandler:
   Exit Sub
End Sub

Private Sub Picturelink_Click()
   Dim tmp As Long
   Me.CommonDialog1.CancelError = True
   On Error GoTo ErrHandler
   Me.CommonDialog1.Flags = cdlCCRGBInit
   Me.CommonDialog1.ShowColor
   tmp = CommonDialog1.color
   Me.Picturelink.BackColor = tmp
   Exit Sub

ErrHandler:
   Exit Sub
End Sub

Private Sub Picturevisited_Click()
   Dim tmp As Long
   Me.CommonDialog1.CancelError = True
   On Error GoTo ErrHandler
   Me.CommonDialog1.Flags = cdlCCRGBInit
   Me.CommonDialog1.ShowColor
   tmp = CommonDialog1.color
   Me.Picturevisited.BackColor = tmp
   Exit Sub

ErrHandler:
   Exit Sub
End Sub
