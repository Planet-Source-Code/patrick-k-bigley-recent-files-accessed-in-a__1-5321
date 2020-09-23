Attribute VB_Name = "Module1"

Sub UpdateFileMenu(FileName)
        Dim retval
        ' Check if OpenFileName is already on MRU list.
        retval = OnRecentFilesList(FileName)
        If Not retval Then
          ' Write OpenFileName to MyProg.INI
          WriteRecentFiles (FileName)
        End If
        ' Update menus for most recent file list.
        GetRecentFiles
End Sub
Function OnRecentFilesList(FileName) As Integer
  Dim i

  For i = 1 To 8
    If Form1.mnuRecentFile(i).Caption = FileName Then
      OnRecentFilesList = True
      Exit Function
    End If
  Next i
    OnRecentFilesList = False
End Function

