Attribute VB_Name = "modFileTools"
Option Explicit
' File(s) related procedures
' --------------------------
' Move a file API
' Copy a file
' Delete to recycling bin API
Private Type SHFILEOPSTRUCT
  hWnd                                   As Long
  wFunc                                  As Long
  pFrom                                  As String
  pTo                                    As String
  fFlags                                 As Integer
  fAborted                               As Boolean
  hNameMaps                              As Long
  sProgress                              As String
End Type
'
''Private Const FO_DELETE                     As Long = &H3
''Private Const FOF_ALLOWUNDO                 As Long = &H40
''Private Const FOF_NOCONFIRMATION            As Long = &H10
''Private Const FOF_SILENT                    As Long = &H4
'
' File properties Constants and API
Private Type SHELLEXECUTEINFO
  cbSize                                 As Long
  fMask                                  As Long
  hWnd                                   As Long
  lpVerb                                 As String
  lpFile                                 As String
  lpParameters                           As String
  lpDirectory                            As String
  nShow                                  As Long
  hInstApp                               As Long
  lpIDList                               As Long    ' Optional parameter
  lpClass                                As String ' Optional parameter
  hkeyClass                              As Long    ' Optional parameter
  dwHotKey                               As Long    ' Optional parameter
  hIcon                                  As Long    ' Optional parameter
  hProcess                               As Long    ' Optional parameter
End Type
'
''Private Const SEE_MASK_INVOKEIDLIST         As Long = &HC
''Private Const SEE_MASK_NOCLOSEPROCESS       As Long = &H40
''Private Const SEE_MASK_FLAG_NO_UI           As Long = &H400
'
' Browse for Folder Dialog
Public Type BrowseInfo
  hOwner                                 As Long
  pIDLRoot                               As Long
  pszDisplayName                         As String
  lpszTitle                              As String
  ulFlags                                As Long
  lpfn                                   As Long
  lParam                                 As Long
  iImage                                 As Long
End Type
'
' BROWSEINFO.ulFlags values:
Private Const BIF_RETURNONLYFSDIRS       As Long = &H1
'
Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, _
                                                                    ByVal lpNewFileName As String) As Long
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, _
                                                                    ByVal lpNewFileName As String, _
                                                                    ByVal bFailIfExists As Long) As Long
''Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
''Private Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" (SEI As SHELLEXECUTEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, _
                                                                                             ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BrowseInfo) As Long

' Concates file to a path (checks if a backslash is required)
Public Function AttachPath(sFileName As String, _
                           sPath As String) As String

  If Len(Trim$(ExtractPath(sFileName))) = 0 Then
    AttachPath = FixPath(sPath) & sFileName
   Else
    AttachPath = sFileName
  End If

End Function

Public Function ExtractFileName(ByVal sFileIn As String, _
                                Optional ByVal bIncludeExt As Boolean = True) As String

  Dim I As Long

  For I = Len(sFileIn) To 1 Step -1
    If InStr(":\", Mid$(sFileIn, I, 1)) Then
      Exit For
    End If
  Next I
  sFileIn = Mid$(sFileIn, I + 1, Len(sFileIn) - I)
  If Not bIncludeExt Then
    For I = Len(sFileIn) To 1 Step -1
      If InStr(".", Mid$(sFileIn, I, 1)) Then
        Exit For
      End If
    Next I
    If I > 0 Then
      sFileIn = Left$(sFileIn, I - 1)
    End If
  End If
  ExtractFileName = sFileIn

End Function

' Extracts the path section of a file-string
Public Function ExtractPath(sPathIn As String) As String

  Dim I As Long

  For I = Len(sPathIn) To 1 Step -1
    If InStr(":\", Mid$(sPathIn, I, 1)) Then
      Exit For
    End If
  Next I
  ExtractPath = Left$(sPathIn, I)

End Function

Public Sub File2BAK(sFile As String, _
                    Optional ByVal bKeepSource As Boolean = False)

  ' Creates a backup (.bak) file of given file
  
  Dim sBAKFile As String

  'Dim i        As Long
  'Dim dl As Long
  If FileExist(sFile) Then
    sBAKFile = Left$(sFile, InStr(sFile, ".") - 1) & "BAK" & Mid$(sFile, InStr(sFile, "."))
    If FileExist(sBAKFile) Then
      Kill sBAKFile
    End If
    If bKeepSource Then
      CopyFile sFile, sBAKFile, False
     Else
      MoveFile sFile, sBAKFile
    End If
  End If

End Sub

Public Function FileExist(ByVal sFile As String) As Boolean

  ' Checks whether file exist (handles wildcards too)

  On Error GoTo ExistErrorHandler
  If Len(Trim$(sFile)) Then
    If LenB(Dir(sFile)) Then
      ' Not there...
      FileExist = True
    End If
  End If

Exit Function

ExistErrorHandler:
  On Error GoTo 0

End Function

' Adds a backslash if required
Private Function FixPath(ByVal sPath As String) As String

  If Len(Trim$(sPath)) = 0 Then
    FixPath = ""
   ElseIf Right$(sPath, 1) <> "\" Then
    FixPath = sPath & "\"
   Else
    FixPath = sPath
  End If

End Function

' Displays the Browse For Folder dialog and
' returns the folder that was chosen.
Public Function FolderBrowser(Optional ByVal sTitle As String = "Please select a folder:", _
                              Optional ByVal OwnerhWnd As Long = 0) As String

  Dim bInf    As BrowseInfo
  Dim nPathID As Long
  Dim sPath   As String

  'Dim nOffset As Integer
  ' Set the properties of the folder dialog
  With bInf
    .hOwner = OwnerhWnd
    .lpszTitle = sTitle
    .ulFlags = BIF_RETURNONLYFSDIRS
    'Show the Browse For Folder dialog
  End With 'bInf
  nPathID = SHBrowseForFolder(bInf)
  sPath = Space$(512)
  If SHGetPathFromIDList(ByVal nPathID, ByVal sPath) Then
    ' Trim off the null chars ending the path of the returned folder
    FolderBrowser = Left$(sPath, InStr(sPath, vbNullChar) - 1)
   Else
    FolderBrowser = ""
  End If

End Function

'-----------------------------------------------------------
' FUNCTION: FolderExist
'
' Determines whether the specified directory name exists.
' This function is used (for example) to determine whether
' an installation floppy is in the drive by passing in
' something like 'A:\'.
'
' IN: [sDirName] - name of directory to check for
'
' Returns: True if the directory exists, False otherwise
'-----------------------------------------------------------
'
Public Function FolderExist(ByVal sDirName As String) As Boolean

  Const WILDCARD As String = "*.*"

  'Dim sDummy     As String
  On Error Resume Next
  FolderExist = LenB(Dir(FixPath(sDirName) & WILDCARD, vbDirectory))
  Err.Number = 0
  On Error GoTo 0

End Function

':)Code Fixer V2.9.2 (1/02/2005 4:13:50 PM) 66 + 160 = 226 Lines Thanks Ulli for inspiration and lots of code.

