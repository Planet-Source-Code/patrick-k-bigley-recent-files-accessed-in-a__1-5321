Attribute VB_Name = "Module2"

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lplFileName As String) As Integer
Sub GetRecentFiles()
  Dim retval, key, i, j
  Dim IniString As String

  ' This variable must be large enough to hold the return string
  ' from the GetPrivateProfileString API.
  IniString = String(255, 0)

  ' Get recent file strings from MyProg.INI
  For i = 1 To 8
    key = "RecentFile" & i
    retval = GetPrivateProfileString("Recent Files", key, "Not Used", IniString, Len(IniString), "MyProg.INI")
    If retval And Left(IniString, 8) <> "Not Used" Then
      ' Update the MDI form's menu.
      Form1.mnuRecentFile(0).Visible = True
      Form1.mnuRecentFile(i).Caption = IniString
      Form1.mnuRecentFile(i).Visible = True

    End If
  Next i

End Sub
Sub WriteRecentFiles(OpenFileName)
  Dim i, j, key, retval
  Dim IniString As String
  IniString = String(255, 0)

  ' Copy RecentFile1 to RecentFile2, etc.
  For i = 7 To 1 Step -1
    key = "RecentFile" & i
    retval = GetPrivateProfileString("Recent Files", key, "Not Used", IniString, Len(IniString), "MyProg.INI")
    If retval And Left(IniString, 8) <> "Not Used" Then
      key = "RecentFile" & (i + 1)
      retval = WritePrivateProfileString("Recent Files", key, IniString, "MyProg.INI")
    End If
  Next i
  
  ' Write openfile to first Recent File.
    retval = WritePrivateProfileString("Recent Files", "RecentFile1", OpenFileName, "MyProg.INI")

End Sub

