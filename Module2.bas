Attribute VB_Name = "Module2"
Option Explicit
'
' Brought to you by Brad Martinez
'   http://members.aol.com/btmtz/vb
'
' Common Controls Replacement Project
'   http://www.mvps.org/ccrp

Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

' Maximum long filename path length
Public Const MAX_PATH = 260

' Special folder values for SHGetSpecialFolderLocation and
' SHGetSpecialFolderPath (Shell32.dll v4.71)

Public Enum SpecialShellFolderIDs
  CSIDL_FAVORITES = &H6
  CSIDL_COOKIES = &H21
  CSIDL_HISTORY = &H22
End Enum

' Retrieves the path of a special folder.
' The docs say it returns NOERROR if successful, or an OLE-defined
' error result otherwise, *but* with both Shell32.dll v4.71 and v4.72 I
' have only seen it return 1 if successful, or 0 otherwise.
Declare Function SHGetSpecialFolderPath Lib "shell32" Alias "SHGetSpecialFolderPathA" _
                              (ByVal hwndOwner As Long, _
                              ByVal pszPath As String, _
                              ByVal nFolder As SpecialShellFolderIDs, _
                              ByVal fCreate As Boolean) As Long

' Retrieves the location of a special (system) folder.
' Returns NOERROR if successful or an OLE-defined error result otherwise.
Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" _
                              (ByVal hwndOwner As Long, _
                              ByVal nFolder As SpecialShellFolderIDs, _
                              pidl As Long) As Long

' Indicates a successful HRESULT
Public Const NOERROR = 0

' Converts an item identifier list to a file system path.
' Returns TRUE if successful or FALSE if an error occurs, for example, if
' the location specified by the pidl parameter is not part of the file system.
Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" _
                              (ByVal pidl As Long, _
                              ByVal pszPath As String) As Long

' Frees memory allocated by the shell
Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
'

' ====== Begin pidl util calls ===============================

' Returns the size in bytes of the first item ID in a pidl.
' Returns 0 if the pidl is the desktop's pidl or is the last
' item ID in the pidl (the zero terminator), or is invalid.

Public Function GetItemIDSize(ByVal pidl As Long) As Integer
  ' If we try to access memory at address 0, then it's bye-bye...
  If pidl Then MoveMemory GetItemIDSize, ByVal pidl, 2
End Function

' Returns a pointer to the next item ID in a pidl.
' Returns 0 if the next item ID is the pidl's zero value terminating 2 bytes.

Public Function GetNextItemID(ByVal pidl As Long) As Long
  Dim cb As Integer   ' SHITEMID.cb, 2 bytes
  cb = GetItemIDSize(pidl)
  ' Make sure it's not the zero value terminator.
  If cb Then GetNextItemID = pidl + cb
End Function

' If successful, returns the size in bytes of the memory occcupied by a pidl,
' including it's 2 byte zero terminator. Returns 0 otherwise.

Public Function GetPIDLSize(ByVal pidl As Long) As Integer
  Dim cb As Integer
  ' Error handle in case we get a bad pidl and overflow cb.
  ' (most item IDs are roughly 20 bytes in size, and since an item ID represents
  ' a folder, a pidl will never have more than 260 folders, or 5200 bytes).
  On Error GoTo Out
  
  If pidl Then
    Do While pidl
      cb = cb + GetItemIDSize(pidl)
      pidl = GetNextItemID(pidl)
    Loop
    ' Add 2 bytes for the zero terminator
    GetPIDLSize = cb + 2
  End If
  
Out:
End Function

' Retrieves a special shell folder's path.
'   hWnd    - handle of window that will own any displayed msgboxes
'   nFolder  - one of the CSIDL_* folder ID values.

' If successful, returns the special shell folder's path,
' or an empty string on failure.

Public Function GetSpecialFolderPath(hWnd As Long, _
                                                             nFolder As SpecialShellFolderIDs) As String
  Dim pidl As Long
  Dim sPath As String * MAX_PATH
  
  ' If the version of Shell32.dll is <v4.71 then SHGetSpecialFolderPath
  ' won't be exported and we'll get VB error 453.
  On Error GoTo NotExported
  
  ' Since we're not sure what the call's return value is, we'll
  ' just check where the first Null char is in the path below.
  Call SHGetSpecialFolderPath(hWnd, sPath, nFolder, 0)
  ' Return the path (if any)
  If InStr(sPath, vbNullChar) > 1 Then
    GetSpecialFolderPath = Left$(sPath, InStr(sPath, vbNullChar) - 1)
    Exit Function
  End If
  
NotExported:

  ' Get the pointer to the folder's item ID list from
  ' it's specified folder ID, returns 0 on success
  If SHGetSpecialFolderLocation(hWnd, nFolder, pidl) = NOERROR Then
    If pidl Then
      ' Get the path from the pointer to the item id list,
      ' returns True on success.
      If SHGetPathFromIDList(pidl, sPath) Then
        ' Return the path
        GetSpecialFolderPath = Left$(sPath, InStr(sPath, vbNullChar) - 1)
      End If
      ' Free the memory the shell allocated for the pidl
      Call CoTaskMemFree(pidl)
    End If
  End If

End Function


