VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Membuka Kotak Dialog ""Browse For Folder"""
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   7845
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   2880
      TabIndex        =   0
      Top             =   2640
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim lpIDList As Long
Dim sBuffer As String
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo
'Ganti 'This Is My Title' dengan judul yang ingin Anda 'letakkan pada kotak dialog "Browse For Folders" 'tersebut.
  szTitle = "This Is My Title"
  With tBrowseInfo
     .hWndOwner = Me.hWnd
     .lpszTitle = lstrcat(szTitle, "")
     .ulFlags = BIF_RETURNONLYFSDIRS + _
                BIF_DONTGOBELOWDOMAIN
  End With
  lpIDList = SHBrowseForFolder(tBrowseInfo)
  If (lpIDList) Then
     sBuffer = Space(MAX_PATH)
     SHGetPathFromIDList lpIDList, sBuffer
     'Nilai sBuffer adalah directori yang dipilih oleh
     'user pada kotak dialog.
     sBuffer = Left(sBuffer, InStr(sBuffer, _
               vbNullChar) - 1)
     MsgBox sBuffer
  End If
End Sub


Private Sub Form_Load()
    Command1.Caption = "Buka Browse For Folder"
End Sub
