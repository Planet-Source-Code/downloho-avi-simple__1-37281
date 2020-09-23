VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Avi SImple 1.0"
   ClientHeight    =   4230
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   282
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   370
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   3735
      Index           =   1
      Left            =   120
      ScaleHeight     =   249
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   353
      TabIndex        =   1
      Top             =   360
      Width           =   5295
      Begin VB.CommandButton cmdStop 
         Caption         =   ""
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   9
         ToolTipText     =   "Stop"
         Top             =   3240
         Width           =   495
      End
      Begin VB.CommandButton cmdPlay 
         Caption         =   "|>"
         Height          =   375
         Left            =   2520
         TabIndex        =   8
         ToolTipText     =   "Play"
         Top             =   3240
         Width           =   495
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   ">> |"
         Height          =   375
         Left            =   1920
         TabIndex        =   7
         ToolTipText     =   "Last Frame"
         Top             =   3240
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">>"
         Height          =   375
         Left            =   1320
         TabIndex        =   6
         ToolTipText     =   "Next Frame"
         Top             =   3240
         Width           =   495
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "<<"
         Height          =   375
         Left            =   720
         TabIndex        =   5
         ToolTipText     =   "Previous Frame"
         Top             =   3240
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "| <<"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "First Frame"
         Top             =   3240
         Width           =   495
      End
      Begin VB.Frame fraPre 
         Caption         =   "Options"
         Height          =   3615
         Left            =   3720
         TabIndex        =   3
         Top             =   0
         Width           =   1575
         Begin VB.ComboBox cmbAlign 
            Height          =   315
            ItemData        =   "frmMain.frx":0000
            Left            =   120
            List            =   "frmMain.frx":001F
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   3120
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox txtCap 
            Height          =   645
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   19
            Top             =   2160
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox txtHeight 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   720
            MaxLength       =   3
            TabIndex        =   17
            Text            =   "240"
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox txtWidth 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   720
            MaxLength       =   3
            TabIndex        =   15
            Text            =   "320"
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txtFPS 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   720
            MaxLength       =   2
            TabIndex        =   12
            Text            =   "30"
            Top             =   600
            Width           =   375
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Caption Align:"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   20
            Top             =   2880
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Caption:"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   18
            Top             =   1920
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Height: "
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   16
            Top             =   1365
            Width           =   555
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Width: "
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   14
            Top             =   1005
            Width           =   510
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "p/s"
            Height          =   195
            Index           =   2
            Left            =   1200
            TabIndex        =   13
            Top             =   645
            Width           =   240
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Frames: "
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   11
            Top             =   645
            Width           =   600
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Frames: 00"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   780
         End
      End
      Begin VB.PictureBox picPre 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2775
         Left            =   0
         ScaleHeight     =   181
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   237
         TabIndex        =   2
         Top             =   120
         Width           =   3615
      End
      Begin VB.HScrollBar slPre 
         Height          =   255
         Left            =   0
         Max             =   1
         Min             =   1
         TabIndex        =   30
         Top             =   2940
         Value           =   1
         Width           =   3615
      End
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   3735
      Index           =   2
      Left            =   120
      ScaleHeight     =   249
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   353
      TabIndex        =   22
      Top             =   360
      Visible         =   0   'False
      Width           =   5295
      Begin VB.CommandButton cmdClear 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   14.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   33
         ToolTipText     =   "Clear List"
         Top             =   1680
         Width           =   375
      End
      Begin VB.CommandButton cmdAddM 
         Caption         =   "Add Multiple Files"
         Height          =   375
         Left            =   3240
         TabIndex        =   32
         Top             =   3120
         Width           =   1575
      End
      Begin VB.PictureBox picMini 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   0
         ScaleHeight     =   41
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   41
         TabIndex        =   28
         Top             =   3120
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add File"
         Height          =   375
         Left            =   2160
         TabIndex        =   27
         Top             =   3120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdRem 
         Caption         =   "û"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   14.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   26
         ToolTipText     =   "Remove"
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton cmdDown 
         Caption         =   "ê"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   25
         ToolTipText     =   "Move Down"
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton cmdUp 
         Caption         =   "é"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   24
         ToolTipText     =   "Move Up"
         Top             =   120
         Width           =   375
      End
      Begin MSComctlLib.ListView lvwFiles 
         Height          =   2895
         Left            =   0
         TabIndex        =   23
         Top             =   120
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   5106
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "FilePath"
            Object.Width           =   38100
         EndProperty
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   5040
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   7
   End
   Begin VB.PictureBox picSave 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1080
      Left            =   1800
      ScaleHeight     =   72
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   31
      Top             =   4440
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.PictureBox picTemp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   120
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   29
      Top             =   4440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSComctlLib.TabStrip tabMain 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   7435
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Preview"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Files"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileSaveAvi 
         Caption         =   "Save AVI"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileNPrj 
         Caption         =   "New Project"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileSavePrj 
         Caption         =   "Save Project"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileOPrj 
         Caption         =   "Open Project"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuLine33 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetRect Lib "user32.dll" (ByRef lprc As AVI_RECT, ByVal xLeft As Long, ByVal yTop As Long, ByVal xRight As Long, ByVal yBottom As Long) As Long     'BOOL
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetLongPathName Lib "kernel32" Alias "GetLongPathNameA" (ByVal lpszLongPath As String, ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long

Private Const SRCCOPY = &HCC0020   'Copies the source over the destination

Dim mStop As Boolean

Private Function FileExist(ByVal filename As String) As Boolean
On Error GoTo 1
Dim l&
 l& = Len(CStr(FileLen(filename$)))
 FileExist = True
Exit Function
1
FileExist = False
End Function

Private Sub Pause(duration As Single)
'this will Pause your program for givin amount of time
'Usage Pause 1
    
    Dim Current As Single
    Current! = Timer
    Do Until Timer - Current! >= duration!
     If mStop = True Then Exit Sub
       DoEvents
    Loop
    
End Sub

Private Sub Play()
Dim l As Integer
mStop = False
 For l = 1 To lvwFiles.ListItems.Count
  DoEvents
  If mStop = True Then Exit Sub
  slPre.Value = l
  Call Pause(IIf(CInt(txtFPS.Text) > 25, 0.01, (1 / CInt(txtFPS.Text))))
 Next l
End Sub

Private Sub printCap(ByRef xPic As PictureBox)
If txtCap.Text = "" Then Exit Sub

Select Case cmbAlign.ListIndex
 Case 0
  xPic.CurrentX = (xPic.Width / 2) - (xPic.TextWidth(txtCap.Text) / 2)
  xPic.CurrentY = (xPic.Height / 2) - (xPic.TextHeight(txtCap.Text) / 2)
 Case 1
  xPic.CurrentX = (xPic.Width / 2) - (xPic.TextWidth(txtCap.Text) / 2)
  xPic.CurrentY = 0
 Case 2
  xPic.CurrentX = (xPic.Width / 2) - (xPic.TextWidth(txtCap.Text) / 2)
  xPic.CurrentY = xPic.Height - xPic.TextHeight(txtCap.Text)
 Case 3
  xPic.CurrentX = 0
  xPic.CurrentY = (xPic.Height / 2) - (xPic.TextHeight(txtCap.Text) / 2)
 Case 4
  xPic.CurrentX = 0
  xPic.CurrentY = 0
 Case 5
  xPic.CurrentX = 0
  xPic.CurrentY = xPic.Height - xPic.TextHeight(txtCap.Text)
 Case 6
  xPic.CurrentX = xPic.Width - xPic.TextWidth(txtCap.Text) - 2
  xPic.CurrentY = (xPic.Height / 2) - (xPic.TextHeight(txtCap.Text) / 2)
 Case 7
  xPic.CurrentX = xPic.Width - xPic.TextWidth(txtCap.Text) - 2
  xPic.CurrentY = 0
 Case 8
  xPic.CurrentX = xPic.Width - xPic.TextWidth(txtCap.Text) - 2
  xPic.CurrentY = xPic.Height - xPic.TextHeight(txtCap.Text)
End Select

xPic.Print txtCap.Text
End Sub

Private Sub WriteAVI(ByVal filename As String, Optional ByVal FrameRate As Integer = 1)
    Dim s$
    Dim InitDir As String
    Dim szOutputAVIFile As String
    Dim res As Long
    Dim pfile As Long 'ptr PAVIFILE
    Dim bmp As cDIB
    Dim ps As Long 'ptr PAVISTREAM
    Dim psCompressed As Long 'ptr PAVISTREAM
    Dim strhdr As AVI_STREAM_INFO
    Dim BI As BITMAPINFOHEADER
    Dim opts As AVI_COMPRESS_OPTIONS
    Dim pOpts As Long
    Dim I As Long
    
    'get an avi filename from user
    szOutputAVIFile = filename$
'    Open the file for writing
    res = AVIFileOpen(pfile, szOutputAVIFile, OF_WRITE Or OF_CREATE, 0&)
    If (res <> AVIERR_OK) Then GoTo error

    'Get the first bmp in the list for setting format
    Set bmp = New cDIB
     s$ = App.Path & IIf(right$(App.Path, 1) <> "\", "\", "") & "temp.bmp"
     Set picTemp.Picture = LoadPicture(lvwFiles.ListItems(1).Text)
     Call StretchBlt(picSave.hdc, 0, 0, picSave.ScaleWidth, picSave.ScaleHeight, picTemp.hdc, 0, 0, picTemp.ScaleWidth, picTemp.ScaleHeight, SRCCOPY)
     picSave.Refresh
     Set picSave.Picture = picSave.Image
     Call SavePicture(picSave.Picture, s$)

    If bmp.CreateFromFile(s$) <> True Then
        MsgBox "Could not load first bitmap file in list!", vbExclamation, App.Title
        GoTo error
    End If

'   Fill in the header for the video stream
    With strhdr
        .fccType = mmioStringToFOURCC("vids", 0&)    '// stream type video
        .fccHandler = 0&                             '// default AVI handler
        .dwScale = 1
        .dwRate = Val(FrameRate%)                        '// fps
        .dwSuggestedBufferSize = bmp.SizeImage       '// size of one frame pixels
        Call SetRect(.rcFrame, 0, 0, bmp.Width, bmp.Height)       '// rectangle for stream
    End With
    
    'validate user input
    If strhdr.dwRate < 1 Then strhdr.dwRate = 1
    If strhdr.dwRate > 30 Then strhdr.dwRate = 30

'   And create the stream
    res = AVIFileCreateStream(pfile, ps, strhdr)
    If (res <> AVIERR_OK) Then GoTo error

    'get the compression options from the user
    'Careful! this API requires a pointer to a pointer to a UDT
    pOpts = VarPtr(opts)
    res = AVISaveOptions(lHwnd, ICMF_CHOOSE_KEYFRAME Or ICMF_CHOOSE_DATARATE, 1, ps, pOpts)
    'returns TRUE if User presses OK, FALSE if Cancel, or error code
    If res <> 1 Then 'In C TRUE = 1
      Call AVISaveOptionsFree(1, pOpts)
      GoTo error
    End If
    
    'make compressed stream
    res = AVIMakeCompressedStream(psCompressed, ps, opts, 0&)
    If res <> AVIERR_OK Then GoTo error
    
    'set format of stream according to the bitmap
    With BI
        .biBitCount = bmp.BitCount
        .biClrImportant = bmp.ClrImportant
        .biClrUsed = bmp.ClrUsed
        .biCompression = bmp.Compression
        .biHeight = bmp.Height
        .biWidth = bmp.Width
        .biPlanes = bmp.Planes
        .biSize = bmp.SizeInfoHeader
        .biSizeImage = bmp.SizeImage
        .biXPelsPerMeter = bmp.XPPM
        .biYPelsPerMeter = bmp.YPPM
    End With
    
    'set the format of the compressed stream
    res = AVIStreamSetFormat(psCompressed, 0, ByVal bmp.PointerToBitmapInfo, bmp.SizeBitmapInfo)
    If (res <> AVIERR_OK) Then GoTo error

    For I = 1 To lvwFiles.ListItems.Count
     s$ = App.Path & IIf(right$(App.Path, 1) <> "\", "\", "") & "temp.bmp"
     Set picTemp.Picture = LoadPicture(lvwFiles.ListItems(I).Text)
     Call StretchBlt(picSave.hdc, 0, 0, picSave.ScaleWidth, picSave.ScaleHeight, picTemp.hdc, 0, 0, picTemp.ScaleWidth, picTemp.ScaleHeight, SRCCOPY)
     picSave.Refresh
     Set picSave.Picture = picSave.Image
     Call SavePicture(picSave.Picture, s$)

     bmp.CreateFromFile (s$) 'load the bitmap (ignore errors)
     res = AVIStreamWrite(psCompressed, I - 1, 1, bmp.PointerToBits, bmp.SizeImage, AVIIF_KEYFRAME, ByVal 0&, ByVal 0&)
     If res <> AVIERR_OK Then GoTo error
     slPre.Value = I
    Next

error:
'   Now close the file
    Set bmp = Nothing
    
    If (ps <> 0) Then Call AVIStreamClose(ps)

    If (psCompressed <> 0) Then Call AVIStreamClose(psCompressed)

    If (pfile <> 0) Then Call AVIFileClose(pfile)

    Call AVIFileExit

    If (res <> AVIERR_OK) Then
      MsgBox "There was an error writing the file.", vbInformation, App.Title
    End If
End Sub


Private Sub cmdAdd_Click()
On Error GoTo 1
cd.filename = ""
cd.DialogTitle = "Add Image to List"
cd.Filter = "All Image Files|*.bmp;*.gif;*.jpg;*.jpeg|Bitmaps (*.bmp)|*.bmp|GIF (*.gif)|*.gif|JPEG (*.jpg, *.jpeg)|*.jpg;*.jpeg|"
cd.ShowOpen

If cd.filename <> "" Then
  If FileExist(cd.filename) = False Then GoTo 1
  Dim l As ListItem
  Set picTemp.Picture = LoadPicture(cd.filename)
  Set l = lvwFiles.ListItems.Add(, , cd.filename)
  l.Selected = True
 Call StretchBlt(picMini.hdc, 0, 0, picMini.ScaleWidth, picMini.ScaleHeight, picTemp.hdc, 0, 0, picTemp.ScaleWidth, picTemp.ScaleHeight, SRCCOPY)
 picMini.Refresh

 slPre.Max = lvwFiles.ListItems.Count
 'slPre.Enabled = True
End If
Exit Sub
1
End Sub

Private Sub cmdAddM_Click()
On Error GoTo 1
'Const OFN_ALLOWMULTISELECT = &H200&
'cd.filename = ""
'cd.Filter = "All Image Files|*.bmp;*.gif;*.jpg;*.jpeg|Bitmaps (*.bmp)|*.bmp|GIF (*.gif)|*.gif|JPEG (*.jpg, *.jpeg)|*.jpg;*.jpeg|"
'cd.Flags = OFN_ALLOWMULTISELECT + 7
'cd.DialogTitle = "Add Image to List " & iss%
'cd.ShowOpen
'cd.Flags = 7
frmFile.Show vbModal

If frmFile.mOk = False Then GoTo 1

If frmFile.mFile <> "" Then

  Dim s As String, arr$(), v As Variant
  s$ = String(260, 0)
  arr$() = Split(frmFile.mFile, " ")
  Call GetLongPathName(Trim$(arr$(0)), s$, 260)

  s$ = Replace(s$, Chr(0), "", , , vbTextCompare)
 ' s$ = arr$(0)
  s$ = IIf(right$(s$, 1) <> "\", s$ & "\", s$)

Dim l As ListItem

 For Each v In arr$()
  If FileExist(s$ & v) <> False Then
   Set picTemp.Picture = LoadPicture(s$ & v)
   Set l = lvwFiles.ListItems.Add(, , s$ & v)
   l.Selected = True
   Call StretchBlt(picMini.hdc, 0, 0, picMini.ScaleWidth, picMini.ScaleHeight, picTemp.hdc, 0, 0, picTemp.ScaleWidth, picTemp.ScaleHeight, SRCCOPY)
   picMini.Refresh
  End If
 Next v

 slPre.Max = lvwFiles.ListItems.Count
 'slPre.Enabled = True
End If
Exit Sub
1
If Err.Number = 20476 Then MsgBox "Please select less files.", vbCritical, "Error"
End Sub

Private Sub cmdBack_Click()
If lvwFiles.ListItems.Count = 0 Or slPre.Value - 1 < 1 Then Exit Sub
slPre.Value = slPre.Value - 1
End Sub

Private Sub cmdClear_Click()
Call lvwFiles.ListItems.Clear
End Sub

Private Sub cmdDown_Click()
On Error GoTo 1
Dim l As Integer, t As String
Dim f As ListItem

 l = lvwFiles.SelectedItem.Index
 t = lvwFiles.SelectedItem.Text
 If l = lvwFiles.ListItems.Count Then GoTo 1
 Call lvwFiles.ListItems.Remove(l)
 Set f = lvwFiles.ListItems.Add(l + 1, , t)
 f.Selected = True
1
End Sub

Private Sub cmdFirst_Click()
If lvwFiles.ListItems.Count = 0 Then Exit Sub
slPre.Value = 1
End Sub

Private Sub cmdLast_Click()
If lvwFiles.ListItems.Count = 0 Then Exit Sub
slPre.Value = lvwFiles.ListItems.Count
End Sub

Private Sub cmdNext_Click()
If lvwFiles.ListItems.Count = 0 Or slPre.Value + 1 > lvwFiles.ListItems.Count Then Exit Sub
slPre.Value = slPre.Value + 1
End Sub

Private Sub cmdPlay_Click()
Call Play
End Sub

Private Sub cmdRem_Click()
On Error GoTo 1
Dim l As Integer
 l = lvwFiles.SelectedItem.Index
 Call lvwFiles.ListItems.Remove(l)
 slPre.Max = lvwFiles.ListItems.Count
 If lvwFiles.ListItems.Count = 0 Then slPre.Enabled = False
 
1
End Sub

Private Sub cmdStop_Click()
mStop = True
End Sub

Private Sub cmdUp_Click()
On Error GoTo 1
Dim l As Integer, t As String
Dim f As ListItem

 l = lvwFiles.SelectedItem.Index
 t = lvwFiles.SelectedItem.Text
 If l = 1 Then GoTo 1
 Call lvwFiles.ListItems.Remove(l)
 Set f = lvwFiles.ListItems.Add(l - 1, , t)
 f.Selected = True
 
1
End Sub

Private Sub Form_Load()
 lvwFiles.ColumnHeaders(1).Width = lvwFiles.Width - 30
 cmbAlign.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub lvwFiles_Click()
On Error GoTo 1
Dim s$
s$ = lvwFiles.SelectedItem.Text
  If FileExist(s$) = False Then GoTo 1
  Set picTemp.Picture = LoadPicture(s$)

 Call StretchBlt(picMini.hdc, 0, 0, picMini.ScaleWidth, picMini.ScaleHeight, picTemp.hdc, 0, 0, picTemp.ScaleWidth, picTemp.ScaleHeight, SRCCOPY)
 picMini.Refresh
1
End Sub

Private Sub mnuFileExit_Click()
Call Unload(Me)
End Sub

Private Sub mnuFileNPrj_Click()
If MsgBox("Are you sure you want to start a new Project?", vbQuestion + vbYesNo, "New Project") = vbYes Then
 txtFPS.Text = "30"
 txtWidth.Text = "320"
 txtHeight.Text = "240"

Call lvwFiles.ListItems.Clear
Set picPre.Picture = LoadPicture()
Set picMini.Picture = LoadPicture()
End If
End Sub

Private Sub mnuFileOPrj_Click()
On Error GoTo 1
cd.filename = ""
cd.DialogTitle = "Open Project"
cd.Filter = "AVI Simple Projects (*.avs)|*.avs|"
cd.ShowOpen

If cd.filename <> "" Then
 Dim arr$(), l&, s$

Open cd.filename For Input As #1
 s$ = Input(LOF(1), #1)
Close #1

arr$() = Split(s$, vbCrLf)

txtFPS.Text = arr$(0)
txtWidth.Text = arr$(1)
txtHeight.Text = arr$(2)

For l& = 3 To UBound(arr$())
 If arr$(l&) <> "" Then lvwFiles.ListItems.Add , , arr$(l&)
Next l&
End If
1
End Sub

Private Sub mnuFileSaveAvi_Click()
On Error GoTo 1
cd.filename = ""
cd.DialogTitle = "Save Project As"
cd.Filter = "AVI Movies (*.avi)|*.avi|"
cd.ShowSave

If cd.filename <> "" Then
 Call AVIFileInit
 picSave.Width = CInt(txtWidth.Text)
 picSave.Height = CInt(txtHeight.Text)
 Call WriteAVI(cd.filename, CInt(txtFPS.Text))
 Call AVIFileExit
End If
Exit Sub
1
End Sub

Private Sub mnuFileSavePrj_Click()
On Error GoTo 1
cd.filename = ""
cd.DialogTitle = "Save As"
cd.Filter = "AVI Simple Projects (*.avs)|*.avs|"
cd.ShowSave

If cd.filename <> "" Then
 Dim s$, l&

 s$ = txtFPS.Text & vbCrLf & _
      txtWidth.Text & vbCrLf & _
      txtHeight.Text & vbCrLf

For l& = 1 To lvwFiles.ListItems.Count
 s$ = s$ & lvwFiles.ListItems(l&).Text & vbCrLf
Next l&

s$ = left$(s$, Len(s$) - 2)

Open cd.filename For Output As #1
 Print #1, s$
Close #1
End If
1
End Sub

Private Sub slPre_Change()
On Error GoTo 1
  s$ = lvwFiles.ListItems(slPre.Value).Text
  If FileExist(s$) = False Then GoTo 1
  Set picTemp.Picture = LoadPicture(s$)

 Call StretchBlt(picPre.hdc, 0, 0, picPre.ScaleWidth, picPre.ScaleHeight, picTemp.hdc, 0, 0, picTemp.ScaleWidth, picTemp.ScaleHeight, SRCCOPY)
 picPre.Refresh
 Set picPre.Picture = picPre.Image
 Call printCap(picPre)
 picPre.Refresh
 
lbl(0).Caption = "Frames: " & slPre.Max
lvwFiles.ListItems(slPre.Value).Selected = True
 
1
End Sub

Private Sub tabMain_Click()
Dim I%
For I% = 1 To pic.Count
 pic(I%).Visible = False
Next I%

If tabMain.SelectedItem.Index = 1 And lvwFiles.ListItems.Count <> 0 Then
 slPre.Enabled = True
 slPre.Max = lvwFiles.ListItems.Count
 slPre.Value = 1
 Call slPre_Change
End If

 pic(tabMain.SelectedItem.Index).Visible = True

1
End Sub

Private Sub txtFPS_Change()
If txtFPS.Text = "" Then Exit Sub
If IsNumeric(txtFPS.Text) = False Then txtFPS.Text = "30"
End Sub

Private Sub txtFPS_LostFocus()
If IsNumeric(txtFPS.Text) = False Then txtFPS.Text = "30"

If CLng(txtFPS.Text) > 30 Then MsgBox "Frames Per Second must not exceed 30.", vbInformation, "CamSpy": txtFPS.Text = "30"
If CLng(txtFPS.Text) < 1 Then MsgBox "Frames Per Second must be greater than 0.", vbInformation, "CamSpy": txtFPS.Text = "1"
End Sub

Private Sub txtHeight_Change()
If txtHeight.Text = "" Then Exit Sub
If IsNumeric(txtHeight.Text) = False Then txtHeight.Text = "240"
End Sub

Private Sub txtHeight_LostFocus()
If IsNumeric(txtHeight.Text) = False Then txtHeight.Text = "320"

If CLng(txtHeight.Text) > 640 Then MsgBox "Size must not exceed 640 Pixels.", vbInformation, "CamSpy": txtHeight.Text = "640"
If CLng(txtHeight.Text) < 96 Then MsgBox "Size must be greater than 96 Pixels.", vbInformation, "CamSpy": txtHeight.Text = "96"
End Sub

Private Sub txtWidth_Change()
If txtWidth.Text = "" Then Exit Sub
If IsNumeric(txtWidth.Text) = False Then txtWidth.Text = "320"
End Sub

Private Sub txtWidth_LostFocus()
If IsNumeric(txtWidth.Text) = False Then txtWidth.Text = "320"

If CLng(txtWidth.Text) > 640 Then MsgBox "Size must not exceed 640 Pixels.", vbInformation, "CamSpy": txtWidth.Text = "640"
If CLng(txtWidth.Text) < 96 Then MsgBox "Size must be greater than 96 Pixels.", vbInformation, "CamSpy": txtWidth.Text = "96"
End Sub
