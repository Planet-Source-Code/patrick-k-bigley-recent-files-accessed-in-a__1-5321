VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recent Files in Menu Box Example program"
   ClientHeight    =   3735
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2040
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin VB.TextBox Text1 
      Height          =   2895
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   720
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "NOTE: This sample code will display recent files in the ""Drop Menu"", stored in the ""MyProg.INI"" file in the Windows directory."
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   4215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "-"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "RecentFile1"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "RecentFile2"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "RecentFile3"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "RecentFile4"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "RecentFile5"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "RecentFile6"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "RecentFile7"
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "RecentFile8"
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
GetRecentFiles
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuOpen_Click()

Dim FileNum
Close #1
FileNum = FreeFile

CommonDialog1.CancelError = True
On Error GoTo ErrHandler2
CommonDialog1.Filter = "Text Files|*.txt"
CommonDialog1.Flags = &H400 Or &H1000 Or &H800
CommonDialog1.ShowOpen

Open CommonDialog1.FileName For Binary As #FileNum
  Text1.Text = Input$(LOF(FileNum), #FileNum)
Close #1

UpdateFileMenu CommonDialog1.FileName

Exit Sub
ErrHandler2:
Close #1
If Err = 32755 Then
'User pressed CANCEL
Else
  MsgBox Err.Description & vbCrLf & vbCrLf & "Could not load file", vbExclamation, "ERR# " & Err
End If

End Sub


Private Sub mnuRecentFile_Click(Index As Integer)
Dim FileNum

FileNum = FreeFile

    
On Error GoTo ErrHandler3

    Open mnuRecentFile(Index).Caption For Input As #1
    If Err Then
        MsgBox "Can't open file: " + FileName
        Close #1
        Exit Sub
    End If
    Close #1


Open mnuRecentFile(Index).Caption For Binary As #FileNum
  Text1.Text = Input$(LOF(FileNum), #FileNum)
Close #1

Exit Sub


ErrHandler3:
Close #1
MsgBox Err.Description & vbCrLf & vbCrLf & "Cannot open this file.", vbExclamation, "ERR# " & Err
End Sub

Private Sub mnuSave_Click()
CommonDialog1.CancelError = True
On Error GoTo ErrHandler1
CommonDialog1.Filter = "Text Files|*.txt"
CommonDialog1.Flags = &H400 Or &H1000 Or &H800 Or &H2
CommonDialog1.ShowSave



Open CommonDialog1.FileName For Output As #1
Print #1, Text1.Text
Close #1

UpdateFileMenu (CommonDialog1.FileName)
Exit Sub
ErrHandler1:
Close #1
If Err = 32755 Then
'User pressed CANCEL
Else
  MsgBox Err.Description & vbCrLf & vbCrLf & "Could not save file", vbExclamation, "ERR# " & Err
End If

End Sub
