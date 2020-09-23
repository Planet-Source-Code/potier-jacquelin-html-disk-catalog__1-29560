VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form textoptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Page Options"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7140
   Icon            =   "textoptions.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   7140
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   3803
      TabIndex        =   25
      Top             =   4440
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3360
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
      Height          =   255
      Left            =   2123
      TabIndex        =   24
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "Table properties"
      Height          =   4335
      Left            =   4200
      TabIndex        =   36
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton cmdbrowse 
         Caption         =   "Browse"
         Height          =   255
         Left            =   1800
         TabIndex        =   45
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox txtbgpicture 
         Height          =   285
         Left            =   240
         TabIndex        =   44
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox txtsizetable 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   16
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtcellspacetable 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   23
         Top             =   3960
         Width           =   615
      End
      Begin VB.TextBox txtcellpadtable 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   22
         Top             =   3600
         Width           =   615
      End
      Begin VB.ComboBox combaligntable 
         Height          =   315
         ItemData        =   "textoptions.frx":030A
         Left            =   960
         List            =   "textoptions.frx":031A
         TabIndex        =   21
         Text            =   "Center"
         Top             =   3120
         Width           =   1095
      End
      Begin VB.ComboBox combwtypetable 
         Height          =   315
         ItemData        =   "textoptions.frx":033C
         Left            =   1680
         List            =   "textoptions.frx":0346
         TabIndex        =   20
         Text            =   "pixels"
         Top             =   2760
         Width           =   855
      End
      Begin VB.ComboBox combhtypetable 
         Height          =   315
         ItemData        =   "textoptions.frx":0355
         Left            =   1680
         List            =   "textoptions.frx":035F
         TabIndex        =   18
         Text            =   "pixels"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtwidthtable 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         TabIndex        =   19
         Top             =   2760
         Width           =   615
      End
      Begin VB.TextBox txtheighttable 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         TabIndex        =   17
         Top             =   2280
         Width           =   615
      End
      Begin VB.PictureBox Picturebordercolortable 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1920
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   15
         Top             =   600
         Width           =   255
      End
      Begin VB.CheckBox chkbordercolortable 
         Caption         =   "Border color"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Width           =   1575
      End
      Begin VB.PictureBox Picturebgcolortable 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1920
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   13
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox chkbgtable 
         Caption         =   "Background color"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Background picture"
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "Border size"
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Cellspace"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   3960
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Cellpad"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "Align"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "Width"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "Height"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   2280
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Title"
      Height          =   2295
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   4095
      Begin VB.CheckBox chklinethroughtitle 
         Caption         =   "Line-through"
         Height          =   255
         Left            =   1920
         TabIndex        =   6
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txttitlesize 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   840
         Width           =   615
      End
      Begin VB.CheckBox chkboldtitle 
         Caption         =   "Bold"
         Height          =   255
         Left            =   1920
         TabIndex        =   3
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CheckBox chkitalictitle 
         Caption         =   "Italic"
         Height          =   255
         Left            =   1920
         TabIndex        =   4
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CheckBox chkunderlinetitle 
         Caption         =   "Underline"
         Height          =   255
         Left            =   1920
         TabIndex        =   5
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txttitle 
         Height          =   285
         Left            =   600
         TabIndex        =   0
         Top             =   360
         Width           =   2175
      End
      Begin VB.PictureBox Picture_title 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   2
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Title"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "size"
         Height          =   255
         Left            =   720
         TabIndex        =   34
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Decoration"
         Height          =   255
         Left            =   720
         TabIndex        =   33
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Color"
         Height          =   255
         Left            =   2160
         TabIndex        =   32
         Top             =   840
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Folder names"
      Height          =   1935
      Left            =   0
      TabIndex        =   26
      Top             =   2400
      Width           =   4095
      Begin VB.CheckBox chklinethroughfolder 
         Caption         =   "Line-through"
         Height          =   255
         Left            =   1680
         TabIndex        =   11
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtsizefolder 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Top             =   360
         Width           =   615
      End
      Begin VB.PictureBox Picture_folders 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   27
         Top             =   360
         Width           =   255
      End
      Begin VB.CheckBox chkboldfolder 
         Caption         =   "Bold"
         Height          =   255
         Left            =   1680
         TabIndex        =   8
         Top             =   840
         Width           =   1095
      End
      Begin VB.CheckBox chkitalicfolder 
         Caption         =   "Italic"
         Height          =   255
         Left            =   1680
         TabIndex        =   9
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CheckBox chkunderlinefolder 
         Caption         =   "Underline"
         Height          =   255
         Left            =   1680
         TabIndex        =   10
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "size"
         Height          =   255
         Left            =   720
         TabIndex        =   30
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "Decoration"
         Height          =   255
         Left            =   600
         TabIndex        =   29
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Color"
         Height          =   255
         Left            =   2160
         TabIndex        =   28
         Top             =   360
         Width           =   495
      End
   End
End
Attribute VB_Name = "textoptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdbrowse_Click()
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

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    On Error Resume Next
    style_title.title = Me.txttitle
    style_title.style.size = val(Me.txttitlesize)
    style_title.style.color = Me.Picture_title.BackColor
    If Me.chkboldtitle.Value = vbChecked Then
        style_title.style.bold = True
    Else
        style_title.style.bold = False
    End If
    If Me.chkitalictitle.Value = vbChecked Then
        style_title.style.italic = True
    Else
        style_title.style.italic = False
    End If
    If Me.chkunderlinetitle.Value = vbChecked Then
        style_title.style.underline = True
    Else
        style_title.style.underline = False
    End If
    If Me.chklinethroughtitle.Value = vbChecked Then
        style_title.style.line_through = True
    Else
        style_title.style.line_through = False
    End If
    
    style_folder.size = val(Me.txtsizefolder)
    style_folder.color = Me.Picture_folders.BackColor
    If Me.chkboldfolder.Value = vbChecked Then
        style_folder.bold = True
    Else
        style_folder.bold = False
    End If
    If Me.chkitalicfolder.Value = vbChecked Then
        style_folder.italic = True
    Else
        style_folder.italic = False
    End If
    If Me.chkunderlinefolder.Value = vbChecked Then
        style_folder.underline = True
    Else
        style_folder.underline = False
    End If
    If Me.chklinethroughfolder.Value = vbChecked Then
        style_folder.line_through = True
    Else
        style_folder.line_through = False
    End If
    If Me.chkbgtable.Value = vbChecked Then
        style_table.use_bgcolor = True
    Else
        style_table.use_bgcolor = False
    End If
    If Me.chkbordercolortable.Value = vbChecked Then
        style_table.use_bordercolor = True
    Else
        style_table.use_bordercolor = False
    End If
    style_table.bgcolor = Me.Picturebgcolortable.BackColor
    style_table.bordercolor = Me.Picturebordercolortable.BackColor
    style_table.bordersize = val(Me.txtsizetable.Text)
    style_table.height = val(Me.txtheighttable.Text)
    style_table.width = val(Me.txtwidthtable.Text)
    style_table.cellpad = val(Me.txtcellpadtable.Text)
    style_table.cellspace = val(Me.txtcellspacetable.Text)
    style_table.align = Me.combaligntable.Text
    style_table.heighttype = Me.combhtypetable.Text
    style_table.widthtype = Me.combwtypetable.Text
    style_table.bgpicture = Me.txtbgpicture.Text
    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Me.txttitle = style_title.title
    If style_title.style.size <> 0 Then Me.txttitlesize = style_title.style.size
    Me.Picture_title.BackColor = style_title.style.color
    If style_title.style.bold Then Me.chkboldtitle.Value = vbChecked
    If style_title.style.italic Then Me.chkitalictitle.Value = vbChecked
    If style_title.style.underline Then Me.chkunderlinetitle.Value = vbChecked
    If style_title.style.line_through Then Me.chklinethroughtitle.Value = vbChecked
    
    If style_folder.size <> 0 Then Me.txtsizefolder = style_folder.size
    Me.Picture_folders.BackColor = style_folder.color
    If style_folder.bold Then Me.chkboldfolder.Value = vbChecked
    If style_folder.italic Then Me.chkitalicfolder.Value = vbChecked
    If style_folder.underline Then Me.chkunderlinefolder.Value = vbChecked
    If style_folder.line_through Then Me.chklinethroughfolder.Value = vbChecked
    
    If style_table.use_bgcolor Then Me.chkbgtable.Value = vbChecked
    If style_table.use_bordercolor Then Me.chkbordercolortable.Value = vbChecked
    Me.Picturebgcolortable.BackColor = style_table.bgcolor
    Me.Picturebordercolortable.BackColor = style_table.bordercolor
    If style_table.bordersize <> 0 Then Me.txtsizetable.Text = style_table.bordersize
    If style_table.height <> 0 Then Me.txtheighttable.Text = style_table.height
    If style_table.width <> 0 Then Me.txtwidthtable.Text = style_table.width
    If style_table.cellpad <> 0 Then Me.txtcellpadtable.Text = style_table.cellpad
    If style_table.cellspace <> 0 Then Me.txtcellspacetable.Text = style_table.cellspace
    Me.combaligntable.Text = style_table.align
    Me.combhtypetable.Text = style_table.heighttype
    Me.combwtypetable.Text = style_table.widthtype
    Me.txtbgpicture.Text = style_table.bgpicture
End Sub

Private Sub Picture_folders_Click()
   Dim tmp As Long
   Me.CommonDialog1.CancelError = True
   On Error GoTo ErrHandler
   Me.CommonDialog1.Flags = cdlCCRGBInit
   Me.CommonDialog1.ShowColor
   tmp = CommonDialog1.color
   Me.Picture_folders.BackColor = tmp
   Exit Sub

ErrHandler:
   Exit Sub
End Sub

Private Sub Picture_title_Click()
   Dim tmp As Long
   Me.CommonDialog1.CancelError = True
   On Error GoTo ErrHandler
   Me.CommonDialog1.Flags = cdlCCRGBInit
   Me.CommonDialog1.ShowColor
   tmp = CommonDialog1.color
   Me.Picture_title.BackColor = tmp
   Exit Sub

ErrHandler:
   Exit Sub
End Sub

Private Sub Picturebgcolortable_Click()
   Dim tmp As Long
   Me.CommonDialog1.CancelError = True
   On Error GoTo ErrHandler
   Me.CommonDialog1.Flags = cdlCCRGBInit
   Me.CommonDialog1.ShowColor
   tmp = CommonDialog1.color
   Me.Picturebgcolortable.BackColor = tmp
   Exit Sub

ErrHandler:
   Exit Sub
End Sub

Private Sub Picturebordercolortable_Click()
   Dim tmp As Long
   Me.CommonDialog1.CancelError = True
   On Error GoTo ErrHandler
   Me.CommonDialog1.Flags = cdlCCRGBInit
   Me.CommonDialog1.ShowColor
   tmp = CommonDialog1.color
   Me.Picturebordercolortable.BackColor = tmp
   Exit Sub

ErrHandler:
   Exit Sub
End Sub
