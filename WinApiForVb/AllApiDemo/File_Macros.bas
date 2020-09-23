Attribute VB_Name = "File_Macros"
Option Explicit

'Bunch of file related functions, read/write,... All Api
'Put together by MH
'mehr13@hotmail.com
'Last updated: Sept 2004

'Used in File_FileSizeToString
Private Const C10 As Currency = 1024 '2 ^ 10
Private Const C20 As Currency = 1048576 '2 ^ 20
Private Const C30 As Currency = 1073741824 '2 ^ 30
Private Const C40 As Currency = 1099511627776# '2^ 40
Private WFD As A_WIN32_FIND_DATA
Private hFile As Long

Public Sub File_Terminate()
    hFile = NUM_ZERO
End Sub

'=================File Read Write Related

'Writes sData items to sFilename
'If successful return true else false
Public Function WriteToFileStringArray(sFileName As String, sData() As String, Optional bAppend As Boolean = False) As Boolean
    If LenB(sFileName) = NUM_ZERO Then Exit Function
    If File_Macros.WriteToFileTextApi(sFileName, Join(sData, vbCrLf), bAppend) > NUM_ZERO Then WriteToFileStringArray = True
End Function

'Returns bytes written to file or 0
Public Function WriteToFileTextApi(sFileName As String, sData As String, Optional bAppend As Boolean = False, Optional bAddCRLF As Boolean = True) As Long
    On Error GoTo WriteToFileTextApi_Error
    Dim lRet As Long
    'Dim tmpBuffer() As Byte
    Dim nSize As Long
    
    If LenB(sData) = NUM_ZERO Then Exit Function
        If bAppend Then
            'Opens the file, if it exists. If the file does not exist, the function creates the file
            hFile = A_CreateFile(sFileName, GENERIC_WRITE, FILE_SHARE_READ, ByVal CLng(0), OPEN_ALWAYS, 0&, 0&)
        Else
            'Create a new file or other object. Overwrite the file or other object (i.e., delete the old one first) if it already exists
            hFile = A_CreateFile(sFileName, GENERIC_WRITE, FILE_SHARE_READ, ByVal CLng(0), CREATE_ALWAYS, 0&, 0&)
        End If
        If hFile <> INVALID_HANDLE_VALUE Then
            'nSize = GetFileSize(hFile, 0)
            'If appending then go to the end of file
            If bAppend Then
                If SetFilePointer(hFile, 0&, 0&, FILE_END) = INVALID_SET_FILE_POINTER Then
                    CloseHandle hFile
                    Exit Function
                End If
            End If
            If bAddCRLF Then sData = sData & vbCrLf  'StrConv(sData & vbCrLf, vbFromUnicode)
            
            WriteFileStr hFile, sData, Len(sData), lRet, ByVal CLng(0)
            CloseHandle hFile
            WriteToFileTextApi = lRet
        End If
    Exit Function
WriteToFileTextApi_Error:
    If hFile > NUM_ZERO Then CloseHandle hFile
End Function

'Returns bytes written to file or 0
'sData must be in ASNI format
Public Function WriteToFileANSIByteApi(sFileName As String, sData() As Byte, Optional bAppend As Boolean = False, Optional lBytesToWrite As Long = NUM_ZERO) As Long
    On Error GoTo WriteToFileANSIByteApi_Error
    Dim lRet As Long
    'Dim tmpBuffer() As Byte
    Dim nSize As Long
    Dim lWrite As Long
    
        If bAppend Then
            'Opens the file, if it exists. If the file does not exist, the function creates the file
            hFile = A_CreateFile(sFileName, GENERIC_WRITE, FILE_SHARE_READ, ByVal CLng(0), OPEN_ALWAYS, 0&, 0&)
        Else
            'Create a new file or other object. Overwrite the file or other object (i.e., delete the old one first) if it already exists
            hFile = A_CreateFile(sFileName, GENERIC_WRITE, FILE_SHARE_READ, ByVal CLng(0), CREATE_ALWAYS, 0&, 0&)
        End If
        If hFile <> INVALID_HANDLE_VALUE Then
            'nSize = GetFileSize(hFile, 0)
            'If appending then go to the end of file
            If bAppend Then
                If SetFilePointer(hFile, 0&, 0&, FILE_END) = INVALID_SET_FILE_POINTER Then
                    CloseHandle hFile
                    Exit Function
                End If
            End If
            lWrite = lBytesToWrite
            If lWrite = NUM_ZERO Then lWrite = UBound(sData) + 1
            WriteFile hFile, ByVal VarPtr(sData(NUM_ZERO)), lWrite, lRet, ByVal CLng(0)
            CloseHandle hFile
            WriteToFileANSIByteApi = lRet
        End If
    Exit Function
WriteToFileANSIByteApi_Error:
    If hFile > NUM_ZERO Then CloseHandle hFile
End Function

'reads the entire file in one shot to a byte array and returns number of bytes read
'or reads a part of it
Public Function BytesFromFileApi(fInStream As String, bBytes() As Byte, Optional lStart As Long = -1, Optional lHowManyBytes As Long = -1) As Long
    Dim nSize As Long
    Dim lRet As Long
    
    On Error GoTo BytesFromFileApi_Error

    If LenB(fInStream) = NUM_ZERO Then Exit Function    'Or Not FilePresent(fInStream) Then Exit Function
    
    hFile = A_CreateFile(fInStream, GENERIC_READ, FILE_SHARE_READ, ByVal CLng(0), OPEN_EXISTING, 0&, 0&)

    If hFile <> INVALID_HANDLE_VALUE Then
        nSize = GetFileSize(hFile, 0)
        If nSize > NUM_ZERO Then
            If lStart > NUM_MINUS_ONE And nSize > lStart Then
                If SetFilePointer(hFile, lStart, 0&, FILE_END) = INVALID_SET_FILE_POINTER Then
                    CloseHandle hFile
                    Exit Function
                End If
                ReDim bBytes(lHowManyBytes - NUM_ONE) As Byte
                nSize = ReadFile(hFile, ByVal VarPtr(bBytes(NUM_ZERO)), lHowManyBytes, lRet, ByVal CLng(0))
                CloseHandle hFile
                BytesFromFileApi = lRet
            Else
                ReDim bBytes(nSize - NUM_ONE) As Byte
                nSize = ReadFile(hFile, ByVal VarPtr(bBytes(NUM_ZERO)), nSize, lRet, ByVal CLng(0))
                CloseHandle hFile
                BytesFromFileApi = lRet
            End If
        End If
    End If
    
    Exit Function
BytesFromFileApi_Error:
    If hFile > NUM_ZERO Then CloseHandle hFile
End Function

'reads the entire file in one shot to a string
Public Function TextFromFileApi(fInStream As String) As String
    'Returns text in file fInStream
    'Example usage: stemp =  TextFromFileApi (App.path & "\popups.dat")
    
    On Error GoTo TrapTextFromFile
    Dim nSize As Long
    Dim lRet As Long, lRet1 As Long
    Dim bBytes() As Byte
    Dim i As Long, strText As String
    
    If LenB(fInStream) = NUM_ZERO Then Exit Function 'Or Not FilePresent(fInStream) Then Exit Function
    
    i = FreeFile
    strText = CHAR_ZERO_LENGTH_STRING
    
    hFile = A_CreateFile(fInStream, GENERIC_READ, FILE_SHARE_READ, ByVal CLng(0), OPEN_EXISTING, 0&, 0&)
    
    If hFile <> INVALID_HANDLE_VALUE Then
        nSize = GetFileSize(hFile, 0)
        If nSize > NUM_ZERO Then
            strText = Space(nSize)
            lRet1 = ReadFileStr(hFile, strText, nSize, lRet, ByVal CLng(0))
            CloseHandle hFile
            If lRet <= NUM_ZERO Then Exit Function
        End If
    End If
    
    'If file is empty then exit
    If LenB(strText) = NUM_ZERO Then Exit Function
    
    'Get rid of all vbcrlf at the end of file, if exists
    'len(vbCrLf) = 2
    Do While Right$(strText, NUM_TWO) = vbCrLf
        strText = Left$(strText, Len(strText) - NUM_TWO)
    Loop

    TextFromFileApi = strText
    Exit Function
    
TrapTextFromFile:
    If hFile > NUM_ZERO Then CloseHandle hFile
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Utility functions, must call CloseFileApi when finished
Public Function FileApi_OpenFile(sFileName As String, hFileHandle As Long, Optional bWriteOperation As Boolean = True, Optional bAppend As Boolean = False) As Boolean
    On Error GoTo OpenFileToWriteApi_Error
    
    If LenB(sFileName) = NUM_ZERO Then Exit Function
    
    If bWriteOperation Then
        If bAppend Then
            'Opens the file, if it exists. If the file does not exist, the function creates the file
            hFileHandle = A_CreateFile(sFileName, GENERIC_WRITE, FILE_SHARE_READ, ByVal CLng(0), OPEN_ALWAYS, 0&, 0&)
        Else
            'Create a new file or other object. Overwrite the file or other object (i.e., delete the old one first) if it already exists
            hFileHandle = A_CreateFile(sFileName, GENERIC_WRITE, FILE_SHARE_READ, ByVal CLng(0), CREATE_ALWAYS, 0&, 0&)
        End If
    Else
        hFileHandle = A_CreateFile(sFileName, GENERIC_READ, FILE_SHARE_READ, ByVal CLng(0), OPEN_EXISTING, 0&, 0&)
    End If
    FileApi_OpenFile = iif(hFileHandle <> INVALID_HANDLE_VALUE, True, False)
    Exit Function
OpenFileToWriteApi_Error:
End Function

Public Function FileApi_CloseFile(hFileHandle As Long) As Long
    On Error GoTo CloseFileApi_Error

    If hFileHandle <= NUM_ZERO Then Exit Function
    FileApi_CloseFile = CloseHandle(hFileHandle)
    Exit Function
CloseFileApi_Error:
End Function

Public Function FileApi_GetSize(hFileHandle As Long) As Long
    FileApi_GetSize = GetFileSize(hFile, 0)
End Function

'Four bytes
Public Function FileApi_WriteLong(hFileHandle As Long, lValue As Long, lBytWritten As Long) As Long
    FileApi_WriteLong = WriteFile(hFileHandle, lValue, NUM_FOUR, lBytWritten, ByVal CLng(0))
End Function

Public Function FileApi_ReadLong(hFileHandle As Long, lValue As Long, lBytRead As Long) As Long
    FileApi_ReadLong = WriteFile(hFileHandle, lValue, NUM_FOUR, lBytRead, ByVal CLng(0))
End Function

Public Function FileApi_WriteString(hFileHandle As Long, sValue As String, lBytWritten As Long) As Long
    FileApi_WriteString = WriteFileStr(hFileHandle, sValue, Len(sValue), lBytWritten, ByVal CLng(0))
End Function

'ReadFile(hFile, ByVal VarPtr(bBytes(NUM_ZERO)), nSize, lRet, ByVal CLng(0))
Public Function FileApi_ReadByte(hFileHandle As Long, bBytes() As Byte, lHowManyBytsToRead As Long, lBytRead As Long) As Long
    FileApi_ReadByte = ReadFile(hFile, ByVal VarPtr(bBytes(NUM_ZERO)), lHowManyBytsToRead, lBytRead, ByVal CLng(0))
End Function

'=================PATH Related

'Adds slash to the end of directories if necessary
Public Function AddSlash(sPath As String) As String
    AddSlash = sPath
    'One for character and one for the embedded null
    If Asc(RightB$(sPath, NUM_TWO)) <> ASCII_BACK_SLASH Then AddSlash = sPath & CHAR_BACK_SLASH     '92
End Function

Public Function RemoveSlash(sTarget As String) As String
    RemoveSlash = sTarget
    If Asc(RightB$(sTarget, NUM_TWO)) = ASCII_BACK_SLASH Then RemoveSlash = Left$(sTarget, Len(sTarget) - NUM_ONE)
End Function

'C:\Mike\file.txt => C:\Mike\file
Public Function RemoveExtension(mPath As String) As String
    Dim i As Integer
    
    RemoveExtension = mPath
    i = InStrRev(mPath, CHAR_DOT)
    If i > NUM_ONE Then RemoveExtension = Left(mPath, i - NUM_ONE)
End Function

'Changes given file extension to another
Public Function ChangeExtension(mPath As String, mExt As String) As String
    
    ChangeExtension = A_PathRenameExtension(mPath, mExt)
    
'    Dim sTmpPath As String
'    sTmpPath = mExt
'    If Asc(LeftB$(sTmpPath, NUM_ONE)) <> ASCII_DOT Then sTmpPath = CHAR_DOT & sTmpPath
'    ChangeExtension = RemoveExtension(mPath) & sTmpPath
End Function

''mFileSpec = .txt .bmp .frm .bat
''mPath = C:\Program Files\folder\file.txt => true
''mPath = C:\Program Files\folder\file.dat => true
''mPath = file.frm => true

'mFileSpec = *.txt;*.bmp;*.f?m;*.?at
'mPath = C:\Program Files\folder\file.txt => true
'mPath = C:\Program Files\folder\file.dat => true
'mPath = file.frm => true
'mFileSpec can be in format "*.txt;*.bmp;*.f?m;*.?at"
Public Function DoesFileExtMatchFileSpec(mPath As String, mFileSpec As String) As Boolean
    DoesFileExtMatchFileSpec = CBool(A_PathMatchSpec(mPath, mFileSpec))
'    Dim iPos As Long
'    Dim strExt As String
'
'    'Do we have an extension
'    iPos = InStrRev(mPath, CHAR_DOT)
'    If iPos >= NUM_ONE Then
'        strExt = LCase$(Right$(mPath, Len(mPath) - iPos + NUM_ONE)) '.txt
'        'Look for extension
'        If InStr(NUM_ONE, mFileSpec, strExt) > NUM_ZERO Then DoesFileExtMatchFileSpec = True
'    End If
End Function

'Attempts to extract the portion of the file name or cur dir from a given file or url path
'http://209.195.98.125/ads/fishing/fishing468x60-0.gif ==> fishing468x60-0.gif
'C:\WINNT\Profiles\Administrator.000\Favorites\Links.url ==> Links.url
'C:\folder1\folder2\folder3 ==> folder3
Public Function ExtractFilePart(ByVal sPathFile As String, Optional sSep As String = "\") As String
    Dim iIndex As Integer
    
    If Right$(sPathFile, NUM_ONE) = sSep Then sPathFile = Left$(sPathFile, Len(sPathFile) - NUM_ONE)
    ExtractFilePart = CHAR_ZERO_LENGTH_STRING
    iIndex = InStrRev(sPathFile, sSep)
    If iIndex > NUM_ONE Then ExtractFilePart = Right$(sPathFile, Len(sPathFile) - iIndex)
End Function

'C:\Program Files\fred\Allan\Betty\Greg.txt => C:\Program Files\fred\Allan\Betty\
'C:/Program Files/fred/Allan/Betty/Greg.txt => C:/Program Files/fred/Allan/Betty/
Public Function ExtractPathPart(sPathFile As String, Optional sSep As String = "\") As String
    Dim iIndex As Integer
    
    ExtractPathPart = CHAR_ZERO_LENGTH_STRING
    iIndex = InStrRev(sPathFile, sSep)
    If iIndex > 1 Then ExtractPathPart = Left$(sPathFile, iIndex)
End Function

'MSN.com => .com
'Returns empty string if no ext found
Public Function ExtractFileExt(sFileName As String) As String
    Dim iPos As Long
    
    ExtractFileExt = CHAR_ZERO_LENGTH_STRING
    
    iPos = InStrRev(sFileName, CHAR_DOT)
    If iPos > NUM_ONE Then ExtractFileExt = Right$(sFileName, Len(sFileName) - iPos + NUM_ONE)
End Function

'sFileName name.ext
'Returns name
Public Function ExtractFileNameWithoutExtension(sFileName As String) As String
    Dim i As Integer
    Dim sTmpPath As String
    Dim iPos As Long
    
    ExtractFileNameWithoutExtension = CHAR_ZERO_LENGTH_STRING
    iPos = InStr(NUM_ONE, sFileName, CHAR_DOT)
    If iPos > NUM_ZERO Then ExtractFileNameWithoutExtension = Left$(sTmpPath, iPos - NUM_ONE)
End Function

'C:\Folder\temp.txt => if present true else false
Public Function FilePresent(sFileName As String) As Boolean
    'FilePresent = CBool(Dir(sFileName) <> "")
    hFile = A_FindFirstFile(sFileName, WFD)
    If hFile <> INVALID_HANDLE_VALUE Then FilePresent = True
    Call FindClose(hFile)
End Function

'if exists C:\Folder\ or C:\Folder then returns true
Public Function DirPresent(ByVal sSource As String) As Boolean
    'DirPresent = CBool(Dir(AddSlash(sSource)) <> "")
    Dim FileName As String

    sSource = AddSlash(sSource) & CHAR_FIND_ALL_FILES_FILTER
    hFile = A_FindFirstFile(sSource, WFD)
    If hFile <> INVALID_HANDLE_VALUE Then DirPresent = True
    Call FindClose(hFile)
End Function

'if ready A or A: or A:\ then returns true
'on locaal drives acts like isDirectoryEmpty
Public Function IsDriveReady(ByVal sDrive As String) As Boolean
    Dim lLen As Long
    
    lLen = Len(sDrive)
    'Get the root A:\
    If lLen > NUM_THREE Then
        A_PathStripToRoot sDrive
        sDrive = RTrimNull(sDrive) & CHAR_FIND_ALL_FILES_FILTER
    ElseIf lLen = NUM_TWO Then
        sDrive = sDrive & CHAR_BACK_SLASH & CHAR_FIND_ALL_FILES_FILTER
    ElseIf lLen = NUM_ONE Then
        sDrive = sDrive & CHAR_COLON & CHAR_BACK_SLASH & CHAR_FIND_ALL_FILES_FILTER
    End If
    hFile = A_FindFirstFile(sDrive, WFD)
    If hFile <> INVALID_HANDLE_VALUE Then IsDriveReady = True
    Call FindClose(hFile)
End Function

'Ret = SearchTreeForFile("c:\", "myfile.ext", tempStr)
'Searches a dir tree below sRootPath for a sFileName file
'if found returns the full path to the sFileName file else returns an empty string
Public Function FilePresentInTree(ByVal sRootPath As String, ByVal sFileName As String) As String
    Dim tempStr As String
    Dim ret As Long
    
    'create a buffer string
    tempStr = String(MAX_PATH, NUM_ZERO)
    sRootPath = AddSlash(sRootPath)
    'returns 1 when successfull, 0 when failed
    ret = SearchTreeForFile(sRootPath, sFileName, tempStr)
    If ret <> NUM_ZERO Then
        FilePresentInTree = RTrimNull(tempStr)
    Else
        FilePresentInTree = CHAR_ZERO_LENGTH_STRING
    End If
End Function

Public Function FitPath(mPath As String, numChars As Long) As String
    FitPath = SysAllocStringLen(0&, numChars + NUM_ONE)
    If A_PathAddEllipseEx(FitPath, mPath, numChars, 0&) <> NUM_ZERO Then
        FitPath = RTrimNull(FitPath)
    Else
        FitPath = Left$(mPath, numChars) & CHAR_TRIPLE_DOT
    End If
'    Dim sTmpPath As String
'    Dim SmPath As String * MAX_PATH
'    sTmpPath = mPath
'    A_PathAddEllipseEx SmPath, sTmpPath, numChars, 0&
'    FitPath = RTrimNull(SmPath) 'StripNulls(SmPath)
End Function

Public Function GetDosPath(ByVal LongPath As String) As String
    Dim S As String
    Dim i As Long
    Dim PathLength As Long
    
    i = Len(LongPath) + NUM_ONE
    S = String(i, vbNullChar)
    PathLength = A_GetShortPathName(LongPath, S, i)
    GetDosPath = Left$(S, PathLength)
End Function

Public Function GetLongFilename(ByVal sShortFilename As String) As String
    Dim lRet As Long
    Dim sLongFilename As String
    
    sLongFilename = String$(1024, vbNullChar)
    lRet = A_GetLongPathName(sShortFilename, sLongFilename, Len(sLongFilename))
    'Do we need more room
    If lRet > Len(sLongFilename) Then
        sLongFilename = String$(lRet + NUM_ONE, vbNullChar)
        lRet = A_GetLongPathName(sShortFilename, sLongFilename, Len(sLongFilename))
    End If
    If lRet > NUM_ZERO Then
        GetLongFilename = Left$(sLongFilename, lRet)
    End If
End Function

'get UNC path
Public Function GetUNCPath(ByVal strDriveLetter As String) As String
    On Local Error GoTo fGetUNCPath_Err

    Dim MSG As String, lngReturn As Long
    Dim lpszLocalName As String
    Dim lpszRemoteName As String
    Dim cbRemoteName As Long
    
    lpszLocalName = strDriveLetter
    lpszRemoteName = String$(MAX_PATH, Chr$(32))
    cbRemoteName = Len(lpszRemoteName)
    lngReturn = A_WNetGetConnection(lpszLocalName, lpszRemoteName, _
                                       cbRemoteName)
    Select Case lngReturn
        Case ERROR_BAD_DEVICE
            MSG = "Error: Bad Device"
        Case ERROR_CONNECTION_UNAVAIL
            MSG = "Error: Connection Un-Available"
        Case ERROR_EXTENDED_ERROR
            MSG = "Error: Extended Error"
        Case ERROR_MORE_DATA
               MSG = "Error: More Data"
        Case ERROR_NOT_SUPPORTED
               MSG = "Error: Feature not Supported"
        Case ERROR_NO_NET_OR_BAD_PATH
               MSG = "Error: No Network Available or Bad Path"
        Case ERROR_NO_NETWORK
               MSG = "Error: No Network Available"
        Case ERROR_NOT_CONNECTED
               MSG = "Error: Not Connected"
        Case NO_ERROR
               ' all is successful...
    End Select
    If LenB(MSG) > NUM_ZERO Then
        MsgBox MSG
    Else
        GetUNCPath = Left$(lpszRemoteName, cbRemoteName)
    End If
fGetUNCPath_End:
    Exit Function
fGetUNCPath_Err:
    Debug.Print "GetUNCPath " & Err.Description
    Resume fGetUNCPath_End
End Function

Public Function IsFileOpen(sFile As String, _
                            Optional bForWriting As Boolean = False, _
                            Optional bForAppending As Boolean = True) As Boolean
    Dim tmpFile As Long
    
    'if bForWriting = true then the file will be opened/created and wiped cleaned
    'Unless bForAppending is true
    IsFileOpen = FileApi_OpenFile(sFile, tmpFile, bForWriting, bForAppending)
    If IsFileOpen Then
        FileApi_CloseFile tmpFile
    End If
'    On Error Resume Next
'    Dim iFreeFile As Integer
'
'    If LenB(sFile) = NUM_ZERO Then Exit Function
'    iFreeFile = FreeFile
'    If bReadAccess Then
'        Open sFile For Input Lock Read As #iFreeFile
'    Else
'        'This action will also vipe the file clean!!!!!!!!!
'        Open sFile For Output Lock Write As #iFreeFile
'    End If
'    If Err <> NUM_ZERO Then IsFileOpen = True: Exit Function
'    Close #iFreeFile
End Function

'Make a directory sAppPath & "Fav"
Public Function MkDirectory(ByVal sPath As String) As Boolean
    Dim Security As SECURITY_ATTRIBUTES
    Dim sTmpPath As String
    
    If Len(sPath) = NUM_ZERO Then Exit Function
    sTmpPath = AddSlash(sPath)
    If A_CreateDirectory(sTmpPath, Security) <> NUM_ZERO Then MkDirectory = True
End Function

Public Function StripNulls(OriginalStr As String) As String
    Dim iPos As Long

    iPos = InStr(OriginalStr, vbNullChar)
    
    If iPos = NUM_ONE Then
        StripNulls = CHAR_ZERO_LENGTH_STRING
    ElseIf iPos > NUM_ONE Then
        StripNulls = Left$(OriginalStr, iPos - NUM_ONE)
        Exit Function
    End If
    
    StripNulls = OriginalStr
End Function

Public Sub ClearBadCharsFromDirFileName(sName As String)
    If Len(sName) = NUM_ZERO Then Exit Sub
    
    If InStr(NUM_ONE, sName, CHAR_FORWARD_SLASH) > NUM_ZERO Then sName = Replace$(sName, CHAR_FORWARD_SLASH, CHAR_ZERO_LENGTH_STRING)
    If InStr(NUM_ONE, sName, CHAR_BACK_SLASH) > NUM_ZERO Then sName = Replace$(sName, CHAR_BACK_SLASH, CHAR_ZERO_LENGTH_STRING)
    If InStr(NUM_ONE, sName, CHAR_DOUBLE_QUATAION) > NUM_ZERO Then sName = Replace$(sName, CHAR_DOUBLE_QUATAION, CHAR_ZERO_LENGTH_STRING)
    If InStr(NUM_ONE, sName, CHAR_STAR) > NUM_ZERO Then sName = Replace$(sName, CHAR_STAR, CHAR_ZERO_LENGTH_STRING)
    If InStr(NUM_ONE, sName, CHAR_COLON) > NUM_ZERO Then sName = Replace$(sName, CHAR_COLON, CHAR_ZERO_LENGTH_STRING)
    If InStr(NUM_ONE, sName, CHAR_QUESTION) > NUM_ZERO Then sName = Replace$(sName, CHAR_QUESTION, CHAR_ZERO_LENGTH_STRING)
    If InStr(NUM_ONE, sName, CHAR_LEFT_ARROW) > NUM_ZERO Then sName = Replace$(sName, CHAR_LEFT_ARROW, CHAR_ZERO_LENGTH_STRING)
    If InStr(NUM_ONE, sName, CHAR_RIGHT_ARROW) > NUM_ZERO Then sName = Replace$(sName, CHAR_RIGHT_ARROW, CHAR_ZERO_LENGTH_STRING)
    If InStr(NUM_ONE, sName, CHAR_PIPE) > NUM_ZERO Then sName = Replace$(sName, CHAR_PIPE, CHAR_ZERO_LENGTH_STRING)
End Sub

Public Function PathHasBadChars(sPath As String) As Boolean
    'Dim lCount As Long
    
    If Len(sPath) = NUM_ZERO Then Exit Function
    
    If InStr(NUM_ONE, sPath, CHAR_FORWARD_SLASH) > NUM_ZERO Then PathHasBadChars = True: Exit Function
    If InStr(NUM_ONE, sPath, CHAR_BACK_SLASH) > NUM_ZERO Then PathHasBadChars = True: Exit Function
    If InStr(NUM_ONE, sPath, CHAR_DOUBLE_QUATAION) > NUM_ZERO Then PathHasBadChars = True: Exit Function
    If InStr(NUM_ONE, sPath, CHAR_STAR) > NUM_ZERO Then PathHasBadChars = True: Exit Function
    If InStr(NUM_ONE, sPath, CHAR_COLON) > NUM_ZERO Then PathHasBadChars = True: Exit Function
    If InStr(NUM_ONE, sPath, CHAR_QUESTION) > NUM_ZERO Then PathHasBadChars = True: Exit Function
    If InStr(NUM_ONE, sPath, CHAR_LEFT_ARROW) > NUM_ZERO Then PathHasBadChars = True: Exit Function
    If InStr(NUM_ONE, sPath, CHAR_RIGHT_ARROW) > NUM_ZERO Then PathHasBadChars = True: Exit Function
    If InStr(NUM_ONE, sPath, CHAR_PIPE) > NUM_ZERO Then PathHasBadChars = True: Exit Function

End Function

'Make an empty file
Public Function CreateEmptyFile(sFileName As String) As Boolean
    Dim tmpFile As Long
    
    CreateEmptyFile = FileApi_OpenFile(sFileName, tmpFile, True)
    If CreateEmptyFile Then
        FileApi_CloseFile tmpFile
    End If
End Function

'Sample call for file/dir
'ShowProps "c:\config.sys", Me.hwnd
'ShowProps "C:\Dir\",Me.hwnd
Public Sub File_ShowProperties(sPath As String)
    On Error Resume Next
    If LenB(sPath) > NUM_ZERO Then 'ShowProps curDirPath, UserControl.hwnd
        Dim SEI As SHELLEXECUTEINFO
        Dim r As Long

        With SEI
            'Set the structure's size
            .cbSize = Len(SEI)
            'Seet the mask
            .fMask = SEE_MASK_NOCLOSEPROCESS Or _
             SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
            'Set the owner window
            .hWnd = GetDesktopWindow
            'Show the properties
            .lpzVerb = StrPtr(StrConv("properties", vbFromUnicode)) 'or use stringtobyte passing varptr of the first byte element
            'Set the filename
            .lpzFile = StrPtr(StrConv(sPath, vbFromUnicode))
            .lpzParameters = 0& 'vbNullChar
            .lpzDirectory = 0& 'vbNullChar
            .nShow = 0
            .hInstApp = 0
            .lpIDList = 0
        End With
        'Returns TRUE if successful, or FALSE otherwise. Call GetLastError for error information.
'If a_ used it does not work, if w_ then it works unless we use strptr(strconv( combo
        r = A_ShellExecuteEx(SEI)
    End If
End Sub

Public Function File_GetFileSizeByName(strFilename As String) As Currency
    Dim WFD As A_WIN32_FIND_DATA
    Dim hFindFirst As Long
    Dim Hi As Currency, Lo As Currency
    Dim sTmpPath As String

    If LenB(strFilename) = 0 Then Exit Function
    sTmpPath = strFilename
    hFindFirst = A_FindFirstFile(sTmpPath, WFD)

    If hFindFirst <> INVALID_HANDLE_VALUE Then
        FindClose hFindFirst
        Hi = WFD.nFileSizeHigh
        Lo = WFD.nFileSizeLow
        If Hi < 0 Then Hi = Hi + (2 ^ 32)
        If Lo < 0 Then Lo = Lo + (2 ^ 32)
        File_GetFileSizeByName = (Hi * (2 ^ 32)) + Lo
    End If
End Function

Public Function File_FileSizeToString(NumberOfBytes As Currency) As String
    Dim strFileSize As String
    Select Case NumberOfBytes
        Case Is < C10 '2 ^ 10 'if less than 1024 return size in bytes.
            strFileSize = NumberOfBytes & " Bytes"
        Case Is < C20 '2 ^ 20 'if less than 1048576 return size in kilobytes.
            strFileSize = Str(Format(NumberOfBytes / C10, "#.##")) & " KB"
        Case Is < C30 '2 ^ 30 'if less than 1073741824 return size in megabytes.
            strFileSize = Str(Format(NumberOfBytes / C20, "#.##")) & " MB"
        Case Is < C40 '2 ^ 40 'if less than 1099511627776 return size in gigabytes.
            strFileSize = Str(Format(NumberOfBytes / C30, "#.##")) & " GB"
        Case Else 'otherwise, return size in terabytes, the largest unit.
            strFileSize = Str(Format(NumberOfBytes / C40, "#.##")) & " TB"
    End Select
    File_FileSizeToString = Trim(strFileSize)
End Function

'Returns file size over 2GB correctly in all wins
Public Function File_GetFileSizeByWFD(WFD As A_WIN32_FIND_DATA) As Currency
    Dim Hi As Currency, Lo As Currency

    On Error Resume Next
    Hi = WFD.nFileSizeHigh
    Lo = WFD.nFileSizeLow
    If Hi < 0 Then Hi = Hi + (2 ^ 32)
    If Lo < 0 Then Lo = Lo + (2 ^ 32)
    '2 ^ 32 = 4294967296
    'Max long
    '2 147 483 647
    File_GetFileSizeByWFD = (Hi * (2 ^ 32)) + Lo
End Function



