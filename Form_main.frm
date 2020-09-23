VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form_main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HTML Disc Catalog"
   ClientHeight    =   4140
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6240
   Icon            =   "Form_main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5280
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      Caption         =   "Filters"
      Height          =   3375
      Left            =   0
      TabIndex        =   15
      Top             =   720
      Width           =   4695
      Begin VB.TextBox txtnotword 
         Height          =   285
         Left            =   600
         TabIndex        =   5
         Top             =   2400
         Width           =   3375
      End
      Begin VB.TextBox txtorword 
         Height          =   285
         Left            =   600
         TabIndex        =   4
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox txtandword 
         Height          =   285
         Left            =   600
         TabIndex        =   3
         Top             =   1200
         Width           =   3375
      End
      Begin VB.TextBox txtmaxsize 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3600
         TabIndex        =   7
         Text            =   "0"
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox txtminsize 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Text            =   "0"
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox txtextension 
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Text            =   "exe;zip;ace;rar"
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label12 
         Caption         =   "Max size (kb)"
         Height          =   255
         Left            =   2520
         TabIndex        =   21
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Min size (kb)"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "All of the following words musn't be in filename"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   2160
         Width           =   3375
      End
      Begin VB.Label Label9 
         Caption         =   "At least one of the following word must be in filenames"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1560
         Width           =   3855
      End
      Begin VB.Label Label8 
         Caption         =   "All the following words must be in filenames"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label Label7 
         Caption         =   "Files Extensions"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1695
      Left            =   4800
      TabIndex        =   14
      Top             =   1560
      Width           =   1335
      Begin VB.CommandButton cmdsaveoptions 
         Caption         =   "Save options"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdexit 
         Caption         =   "Exit"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdgenerate 
         Caption         =   "Generate"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmdpreview 
         Caption         =   "Preview"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   6255
      Begin VB.CommandButton cmdbrowse 
         Caption         =   "Browse"
         Height          =   255
         Left            =   5160
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtfolder 
         Height          =   285
         Left            =   1800
         TabIndex        =   0
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Select drive or folder"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnupreview 
         Caption         =   "&Preview"
      End
      Begin VB.Menu mnugenerate 
         Caption         =   "&Generate"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnusaveopt 
         Caption         =   "&Save options"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuoptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuoptbackground 
         Caption         =   "&Background options"
      End
      Begin VB.Menu mnuoptlink 
         Caption         =   "&Links options"
      End
      Begin VB.Menu mnuopttext 
         Caption         =   "&Page options"
      End
      Begin VB.Menu mnud1 
         Caption         =   "-"
      End
      Begin VB.Menu mnugenerateabsolute 
         Caption         =   "Generate with absolute path"
      End
      Begin VB.Menu mnugeneraterelative 
         Caption         =   "Generate with relative path"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuabout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Form_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub get_current_values()
    last_folder = Me.txtfolder.Text
    opt_extensions = Me.txtextension.Text
    opt_all_and_words = Me.txtandword.Text
    opt_all_or_words = Me.txtorword.Text
    opt_all_not_words = Me.txtnotword.Text
    min_file_size = Me.txtminsize.Text
    max_file_size = Me.txtmaxsize.Text
End Sub

Private Sub cmdbrowse_Click()
    Me.txtfolder.Text = GetFolderName()
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub cmdgenerate_Click()
    On Error Resume Next


    Call get_current_values
    If last_folder = "" Then
        MsgBox "Please fill the folder field", vbExclamation, "HTML Disk Catalog"
        Exit Sub
    End If
    
    fill_restrictions_array min_file_size, max_file_size, opt_all_and_words, _
                            opt_all_or_words, opt_all_not_words, opt_extensions
    find_files last_folder
    subQuickSort 0, UBound(files_found) - 1
    
    form_entries.Show vbModal

End Sub

Private Sub cmdpreview_Click()
    On Error Resume Next
    Dim ShellExecutereturn  As Long
    Call get_current_values
    If last_folder = "" Then
        MsgBox "Please fill the folder field", vbExclamation, "HTML Disk Catalog"
        Exit Sub
    End If
    fill_restrictions_array min_file_size, max_file_size, opt_all_and_words, _
                            opt_all_or_words, opt_all_not_words, opt_extensions
    find_files last_folder
    subQuickSort 0, UBound(files_found) - 1
    make_html App.Path & "\preview.htm", True, True
    ShellExecutereturn = ShellExecute(Me.hwnd, "Open", App.Path & "\preview.htm", "", App.Path, 1)
End Sub

Private Sub cmdsaveoptions_Click()
    Call get_current_values
    Call update_ini
End Sub

Private Sub Form_Load()
    On Error Resume Next
    ChDir (App.Path)
    ReDim known_files_restrictions_array(1)
    ini_file_name = App.Path & "\html_cat.ini"
    Call get_ini_value
    Me.txtfolder.Text = last_folder
    Me.txtextension.Text = opt_extensions
    Me.txtandword.Text = opt_all_and_words
    Me.txtorword.Text = opt_all_or_words
    Me.txtnotword.Text = opt_all_not_words
    Me.txtminsize.Text = min_file_size
    Me.txtmaxsize.Text = max_file_size
End Sub

Private Sub mnugenerateabsolute_Click()
    mnugenerateabsolute.Checked = True
    mnugeneraterelative.Checked = False
End Sub

Private Sub mnugeneraterelative_Click()
    mnugeneraterelative.Checked = True
    mnugenerateabsolute.Checked = False
End Sub

Private Sub mnuoptbackground_Click()
    bgoptions.Show vbModal
End Sub

Private Sub mnuoptlink_Click()
    linkoptions.Show vbModal
End Sub

Private Sub mnuopttext_Click()
    textoptions.Show vbModal
End Sub

Private Sub mnuabout_Click()
    about.Show vbModal
End Sub

Private Sub mnuaddremoveentry_Click()
    addremoveentry.Show vbModal
End Sub

Private Sub mnupreview_Click()
    Call cmdpreview_Click
End Sub

Private Sub mnusaveopt_Click()
    Call cmdsaveoptions_Click
End Sub

Private Sub mnuexit_Click()
    Call cmdexit_Click
End Sub

Private Sub mnugenerate_Click()
    Call cmdgenerate_Click
End Sub
