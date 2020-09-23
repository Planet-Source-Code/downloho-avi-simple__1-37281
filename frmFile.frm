VERSION 5.00
Begin VB.Form frmFile 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Open Files"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbList 
      Height          =   315
      ItemData        =   "frmFile.frx":0000
      Left            =   120
      List            =   "frmFile.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   2400
      TabIndex        =   2
      Top             =   2640
      Width           =   2175
   End
   Begin VB.FileListBox File1 
      Height          =   2430
      Left            =   120
      MultiSelect     =   2  'Extended
      Pattern         =   "*.jpg; *.bmp; *.gif"
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   2400
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Image imgPre 
      Height          =   855
      Left            =   4680
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1095
   End
End
Attribute VB_Name = "frmFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Public mOk As Boolean, mFile As String

Private Sub cmbList_Click()
File1.Pattern = cmbList.Text
End Sub

Private Sub cmdCan_Click()
mOk = False
Me.Hide
End Sub

Private Sub cmdOk_Click()
Dim s$
s$ = Space(256)
Call GetShortPathName(Dir1.Path, s$, 256)
mFile = Trim$(s$) & " "

For I% = 0 To File1.ListCount - 1
 If File1.Selected(I%) = True Then mFile$ = mFile$ & File1.List(I%) & " "
Next I%
mFile$ = left$(mFile$, Len(mFile$) - 1)
mOk = True
Call Me.Hide
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
Set imgPre.Picture = LoadPicture(Dir1.Path & IIf(right$(Dir1.Path, 1) <> "\", "\", "") & File1.List(File1.ListIndex))
End Sub

Private Sub Form_Load()
cmbList.ListIndex = 0
End Sub
