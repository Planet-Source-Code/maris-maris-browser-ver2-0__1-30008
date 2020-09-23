VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAutoComplete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents Text1 As TextBox
Attribute Text1.VB_VarHelpID = -1

Private m_fInSearch As Boolean
Private m_fCanChange As Boolean

Private Const vbBackslash = "\"
Private Const vbSlashes = "\/"
Private Const vbAllFileSpec = "*.*"
Private Const vbAtChar = "@"

' =================================================================
' kernel32.dll definitions

' LocalAlloc uFlags
Private Const LMEM_FIXED = &H0
Private Const LMEM_ZEROINIT = &H40
Private Const LPTR = (LMEM_FIXED Or LMEM_ZEROINIT)

Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

Private Declare Function lstrlenA Lib "kernel32" (lpString As Any) As Long
Private Declare Function lstrcpyA Lib "kernel32" (lpString1 As Any, lpString2 As Any) As Long

Private Const MAX_PATH = 260

Private Type FILETIME   ' ft
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA   ' wfd
  dwFileAttributes As Long
  ftCreationTime As FILETIME
  ftLastAccessTime As FILETIME
  ftLastWriteTime As FILETIME
  nFileSizeHigh As Long
  nFileSizeLow As Long
  dwReserved0 As Long
  dwReserved1 As Long
  cFileName As String * MAX_PATH
  cAlternateFileName As String * 14
End Type

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Boolean

' FindFirstFile error return value
Private Const INVALID_HANDLE_VALUE = -1

' Return values: -1 = first < second (first is older), 0 = first = second, 1 = first > second (first is newer)
Private Declare Function CompareFileTime Lib "kernel32" (lpFileTime1 As FILETIME, lpFileTime2 As FILETIME) As Long

' =======================================================================
' wininet (URL cache) definitions

' The following blurb is taken directly from the top of the Cachevw.cpp file included in:
' "Q172607 - SAMPLE: CacheVW.exe Enumerates and Displays Wininet Cache"
' http://support.microsoft.com/support/kb/articles/Q172/6/07.ASP
'
' <quote>
' In IE3.0x, cache entry info structure is limited to 2,048 bytes, and
' wininet.h for IE3.0x has the following define:
' #define MAX_CACHE_ENTRY_INFO_SIZE 4096
'
' FindNextUrlCacheEntry in IE3.0x wininet.dll has a bug where an item is
' skipped if the initial buffer is too small. This bug will be
' fixed in IE4.0 wininet.dll. Also IE4.0 does not have the 4K limit for cache
' entry info structure and the MAX_CACHE_ENTRY_INFO_SIZE flag is removed from
' the IE4.0 wininet.h file. If the app is designed to run only with the IE4.0
' version of Wininet.dll, the app could use a much smaller initial buffer, for
' example, 512 bytes, to enumerate the cache. This is usually sufficient for
' most of cache entries.

' To make our app works with both IE3.0x and IE4.0 wininet.dll, we
' set our intial buffer size to 4K.
Private Const MAX_CACHE_ENTRY_INFO_SIZE = 4096
' </quote>

' There is no cache entry size limit, so applications that need to enumerate the cache must be prepared to allocate
' variable-sized buffers. For more information, see Appendix B: Using Buffers (above).

Private Type INTERNET_CACHE_ENTRY_INFO  ' A
  dwStructSize As Long                 ' Size, in bytes, of this structure.
  lpszSourceUrlName As Long       ' pointer to the URL name.
  lpszLocalFileName As Long         ' pointer to the local file name.
  CacheEntryType As Long           ' cache type bit mask.
  dwUseCount As Long                 ' current user count of the cache entry.
  dwHitRate As Long                     ' number of times the cache entry was retrieved.
  dwSizeLow As Long                    ' low DWORD of the file size.
  dwSizeHigh As Long                   ' high DWORD of the file size.
  LastModifiedTime As FILETIME  ' last modified time of the file in GMT format.
  ExpireTime As FILETIME            ' expire time of the file in GMT format
  LastAccessTime As FILETIME   ' last accessed time in GMT format
  LastSyncTime As FILETIME      ' last time the URL was synchronized with the source
  lpHeaderInfo As Long                ' embedded pointer to the header info.  LPBYTE
  dwHeaderInfoSize As Long       ' size of the above header.
  lpszFileExtension As Long         ' File extension used to retrive the urldata as a file.
'  union {
'    DWORD dwReserved;
'    DWORD dwExemptDelta;
'  };
  dwExemptDelta As Long           ' Exemption time, in seconds, from the last accessed time.
End Type

Private Declare Function FindFirstUrlCacheEntry Lib "wininet" Alias "FindFirstUrlCacheEntryA" _
                            (ByVal lpszUrlSearchPattern As String, _
                            lpNextCacheEntryInfo As Any, _
                            lpdwNextCacheEntryInfoBufferSize As Long) As Long   ' rtns handle

Private Declare Function FindNextUrlCacheEntry Lib "wininet" Alias "FindNextUrlCacheEntryA" _
                            (ByVal hEnumHandle As Long, _
                            lpNextCacheEntryInfo As Any, _
                            lpdwNextCacheEntryInfoBufferSize As Long) As Boolean

Private Declare Function FindCloseUrlCache Lib "wininet" (ByVal hEnumHandle As Long) As Boolean
'

Public Sub Attach(tb As TextBox)
  Set Text1 = tb
End Sub

' =======================================================================
' sunk TextBox events

Private Sub Text1_KeyPress(KeyAscii As Integer)
  
  ' Set the flag if KeyAscii is in the range of printable ANSII chars
  ' (allowing no changes on backspace, tab, cut/paste, etc)
  m_fCanChange = (KeyAscii >= 32) And (KeyAscii <= 122)

End Sub

Private Sub Text1_Change()
  Dim sText As String
  Dim sPath As String
  Dim nChars As Integer
  
  ' If currently searching (or changing Text1.Text) below, or the flag was
  ' not set in Text1_KeyPress above, or Text1 has no chars, then exit.
  If m_fInSearch Or (m_fCanChange = False) Or (Len(Text1) = 0) Then
    m_fInSearch = False
    Exit Sub
  End If
  
  ' Get the entry's text length and clear the flag set in Text1_KeyPress
  sText = Text1
  m_fCanChange = False
  
  ' If the textbox entry does not end with a slash, and is not
  ' a net UNC path (which can take too long to resolve...)
  If (InStr(vbSlashes, Right(sText, 1)) = 0) And (Left$(sText, 1) <> vbBackslash) Then
    
    ' Set the flag and start searching. If the search flag was cleared
    ' above because the user entered new text, call this proc again
    ' and search for that new text.
    m_fInSearch = True
    
    sPath = CompleteURL(sText)
    If (m_fInSearch = False) Then Call Text1_Change
    
    If (Len(sPath) = 0) Then sPath = CompletePath(sText)
    If (m_fInSearch = False) Then Call Text1_Change
    
    nChars = Len(sText)
    
    ' If we got a path and it's longer than the textbox entry...
    If Len(sPath) And (Len(sPath) > nChars) Then
      ' Append the next path element with any sslash to the textbox
      ' entry (or complete the current path element), and select the
      ' appended text.
      Text1 = Text1 & Mid$(sPath, nChars + 1, GetNextSlashPos(sPath, nChars + 1) - nChars)
      Text1.SelStart = nChars
      Text1.SelLength = Len(Text1) - nChars
    End If
    
    m_fInSearch = False
    
  End If   ' InStr
  
End Sub

' =======================================================================
' private helper functions

Private Function CompleteURL(sURL As String) As String
  Dim nChars As Integer
  Dim dwSize As Long
  Dim lpice As Long   ' pointer to ice
  Dim ice As INTERNET_CACHE_ENTRY_INFO
  Dim hFind As Long
  Dim sTmp As String
  Dim nPos As Integer
  Dim sSave As String
  Dim ftSave As FILETIME
  
'Dim n As Long

  nChars = Len(sURL)
  dwSize = MAX_CACHE_ENTRY_INFO_SIZE
  
  lpice = LocalAlloc(LMEM_FIXED, dwSize)
  If lpice Then
  
    ' set the struct's size member in its allocated pointer
    ' (dwStructSize is the first member)
    MoveMemory ByVal lpice, dwSize, 4
  '  ice.dwStructSize = dwSize
  '  MoveMemory ByVal lpice, ice, Len(ice)
  
    ' Get only historry ("visited:") entries
    hFind = FindFirstUrlCacheEntry("visited:", ByVal lpice, dwSize)
    If hFind Then
      Do
        
        ' If new text was typed in by the user, cancel the search
        ' (and close the find handle), and exit .
        DoEvents
        If (m_fInSearch = False) Then Exit Do
        
        ' fill the struct from its pointer
        MoveMemory ice, ByVal lpice, Len(ice)
        
        ' Get the entry and strip any leading "Visited: <logon-name>@" string
        ' (this string is not added when using IUrlHistoryStg, which despite
        ' being more efficient, is only available on at least IE4).
        sTmp = GetStrFromPtrA(ice.lpszSourceUrlName)
        nPos = InStr(sTmp, vbAtChar)
        If nPos And (nPos < Len(sTmp)) Then
          sTmp = Mid$(sTmp, nPos + 1, Len(sTmp) - nPos)
        End If
        
        ' If the current URL begins with the specified URL...
        If (StrComp(sURL, Left$(sTmp, nChars), vbTextCompare) = 0) Then
          ' If the current URL is newer than the last saved URL, save it and its date
          If (CompareFileTime(ice.LastAccessTime, ftSave) = 1) Then
            ftSave = ice.LastAccessTime
            sSave = sTmp
          End If
        End If
        
'n = n + 1
        dwSize = MAX_CACHE_ENTRY_INFO_SIZE
      Loop While FindNextUrlCacheEntry(hFind, ByVal lpice, dwSize)
      
      Call FindCloseUrlCache(hFind)
    End If   ' hFind
  
    Call LocalFree(lpice)
  End If   ' lpice
  
'Debug.Print n

  ' Returns the last saved URL
  If m_fInSearch Then CompleteURL = sSave
  
End Function

Private Function CompletePath(ByVal sPath As String) As String
  Dim sFile As String
  Dim hFind As Long
  Dim wfd As WIN32_FIND_DATA
  Dim nChars As Integer
  
  ' If the specified path is valid, then there's nothing to complete, exit.
  If IsValidFilePath(sPath) Or IsValidDirPath(sPath) Then
    CompletePath = NormalizePath(sPath)
    Exit Function
  End If
  
  ' Get the last incomplete dir/file in the specified path
  sFile = GetLastItemInPath(sPath)
  If Len(sFile) Then
    ' Remove the last incomplete dir/file in the specified
    ' path, leaving the trailing backslash
    sPath = Left$(sPath, Len(sPath) - Len(sFile))
  Else
    ' Add a trailing backslash to the specified path
    sPath = NormalizePath(sPath)
  End If
  
  If IsValidDirPath(sPath) Then
    nChars = Len(sFile)
    
    hFind = FindFirstFile(sPath & vbAllFileSpec, wfd)
    If (hFind <> INVALID_HANDLE_VALUE) Then
      Do
        
        ' If new text was typed in by the user, cancel the search
        ' (and close the find handle), and exit .
        DoEvents
        If (m_fInSearch = False) Then Exit Do
        
        ' If the current dir/file begins with the specified incomplete
        ' dir/file, append it to the specified path and return it.
        If (StrComp(Left$(wfd.cFileName, nChars), sFile, vbTextCompare) = 0) Then
          If (wfd.dwFileAttributes And vbDirectory) Then
            CompletePath = NormalizePath(sPath & GetStrFromBufferA(wfd.cFileName))
          Else
            CompletePath = sPath & GetStrFromBufferA(wfd.cFileName)
          End If
          Exit Function
        End If
      Loop While FindNextFile(hFind, wfd)
    
      Call FindClose(hFind)
    End If   ' (hFind <> INVALID_HANDLE_VALUE)
  
  End If   ' IsValidFilePath
  
End Function

' Returns the string before first null char encountered (if any) from an ANSII string.

Private Function GetStrFromBufferA(sz As String) As String
  If InStr(sz, vbNullChar) Then
    GetStrFromBufferA = Left$(sz, InStr(sz, vbNullChar) - 1)
  Else
    ' If sz had no null char, the Left$ function
    ' above would return a zero length string ("").
    GetStrFromBufferA = sz
  End If
End Function

' Returns an ANSI string from a pointer to an ANSII string.

Private Function GetStrFromPtrA(lpszA As Long) As String
  GetStrFromPtrA = Space$(lstrlenA(ByVal lpszA))
  Call lstrcpyA(ByVal GetStrFromPtrA, ByVal lpszA)
End Function

Private Function NormalizePath(sPath As String) As String
  If ((Len(sPath) > 2) And (Right$(sPath, 1) <> vbBackslash)) Or _
        ((Len(sPath) = 2) And (Right$(sPath, 1) = ":")) Or (sPath = vbBackslash) Then
    NormalizePath = sPath & vbBackslash
  Else
    NormalizePath = sPath
  End If
End Function

Private Function IsValidFilePath(sPath As String) As Boolean
  Dim wfd As WIN32_FIND_DATA
  IsValidFilePath = FindClose(FindFirstFile(sPath, wfd))
End Function

Private Function IsValidDirPath(sPath As String) As Boolean
  Dim wfd As WIN32_FIND_DATA
  If Len(sPath) Then   ' vbAllFileSpec is a valid relative search
    IsValidDirPath = FindClose(FindFirstFile(NormalizePath(sPath) & vbAllFileSpec, wfd))
  End If
End Function

' Returns the file or folder name after last "\", if any.

Private Function GetLastItemInPath(sPath) As String
  Dim nPos1 As Integer
  Dim nPos2 As Integer
  
  nPos1 = InStr(sPath, vbBackslash)
  
  If (nPos1 = 0) Then
    GetLastItemInPath = sPath
  Else
    
    Do
      nPos2 = InStr(nPos1, sPath, vbBackslash)
      If nPos2 Then nPos1 = nPos2 + 1
    Loop While nPos2
    
    GetLastItemInPath = Mid$(sPath, nPos1, (Len(sPath) - nPos1) + 1)
  
  End If

End Function

Private Function GetNextSlashPos(s As String, iStart As Integer) As Integer
  Dim nChars As Integer
  Dim i As Integer
  Dim fFound As Boolean
  
  nChars = Len(s)
  If (nChars < iStart) Then Exit Function
  
  For i = iStart To nChars
    If (fFound = False) Then
      fFound = InStr(vbSlashes, Mid$(s, i, 1))
    Else
      If (InStr(vbSlashes, Mid$(s, i, 1)) = 0) Then
        GetNextSlashPos = i - 1
        Exit Function
      End If
    End If
  Next
  
  GetNextSlashPos = nChars
  
End Function
