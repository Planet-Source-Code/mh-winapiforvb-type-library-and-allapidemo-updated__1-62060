Attribute VB_Name = "Shell_Macros"
Option Explicit

Private foBuff() As Byte
Private mSource As String
Private mDestination As String

Public Sub Shell_Terminate()
    Erase foBuff
    mSource = CHAR_ZERO_LENGTH_STRING
    mDestination = CHAR_ZERO_LENGTH_STRING
End Sub

Public Function Shell_Move(mS As String, mD As String, Optional mFlags As Long = &H100) As Long
    On Error GoTo Shell_Move_Error
    Dim SHFileOp As SHFILEOPSTRUCT
    Dim lRet As Long
    Dim lenFileop As Long
    
    lRet = NUM_MINUS_ONE
    Shell_Move = lRet
    
    'Root Drive
    If Len(mS) < NUM_FOUR Then
        Err.Raise 1001, "Class File Operations", "Unable to move root drive " & mS
        Exit Function
    End If
    
    lenFileop = LenB(SHFileOp) ' double word alignment increase
    ReDim foBuff(1 To lenFileop) ' the size of the structure.
    
    mSource = mS
    mDestination = mD
    
    If Right(mSource, NUM_ONE) = CHAR_BACK_SLASH Then
        mSource = Left(mSource, Len(mSource) - NUM_ONE) & vbNullChar & vbNullChar
    Else
        mSource = mSource & vbNullChar & vbNullChar
    End If
    If Right(mDestination, NUM_ONE) = CHAR_BACK_SLASH Then
        mDestination = Left(mDestination, Len(mDestination) - NUM_ONE) & vbNullChar & vbNullChar
    Else
        mDestination = mDestination & vbNullChar & vbNullChar
    End If

    With SHFileOp
        .hWnd = GetDesktopWindow 'frmProfiles.hWnd
        .wFunc = FO_MOVE
        .lpszFrom = StrPtr(mSource) 'or strptr(strconv(msource,vbfromunicode)) for ANSI
        .lpszTo = StrPtr(mDestination)
        .fFlags = mFlags
    End With

    ' Now we need to copy the structure into a byte array
    Call CopyMemory(foBuff(1), SHFileOp, lenFileop)

    ' Next we move the last 12 bytes by 2 to byte align the data
    Call CopyMemory(foBuff(19), foBuff(21), 12)

    lRet = W_SHFileOperation(SHFileOp)
    Shell_Move = lRet

    mSource = CHAR_ZERO_LENGTH_STRING
    mDestination = CHAR_ZERO_LENGTH_STRING
    Erase foBuff

    Exit Function
Shell_Move_Error:
    Shell_Move = lRet
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Shell_Move of Class Module clsMehrFileOp"
End Function

Public Function Shell_Copy(mS As String, mD As String, Optional mFlags As Long = &H100) As Long
    On Error GoTo Shell_Copy_Error
    Dim SHFileOp As SHFILEOPSTRUCT
    Dim lRet As Long
    Dim lenFileop As Long
    lRet = -1

    lenFileop = LenB(SHFileOp) ' double word alignment increase
    ReDim foBuff(1 To lenFileop) ' the size of the structure.

    mSource = mS
    mDestination = mD

    If Right(mSource, NUM_ONE) = CHAR_BACK_SLASH Then
        mSource = Left(mSource, Len(mSource) - NUM_ONE) & vbNullChar & vbNullChar
    Else
        mSource = mSource & vbNullChar & vbNullChar
    End If
    If Right(mDestination, NUM_ONE) = CHAR_BACK_SLASH Then
        mDestination = Left(mDestination, Len(mDestination) - NUM_ONE) & vbNullChar & vbNullChar
    Else
        mDestination = mDestination & vbNullChar & vbNullChar
    End If

    With SHFileOp
        .hWnd = GetDesktopWindow 'frmProfiles.hWnd
        .wFunc = FO_COPY
        .lpszFrom = StrPtr(mSource) 'or strptr(strconv(msource,vbfromunicode)) for ANSI
        .lpszTo = StrPtr(mDestination)
        .fFlags = mFlags
    End With

    ' Now we need to copy the structure into a byte array
    Call CopyMemory(foBuff(1), SHFileOp, lenFileop)

    ' Next we move the last 12 bytes by 2 to byte align the data
    Call CopyMemory(foBuff(19), foBuff(21), 12)

    lRet = W_SHFileOperation(SHFileOp)
    Shell_Copy = lRet

    mSource = CHAR_ZERO_LENGTH_STRING
    mDestination = CHAR_ZERO_LENGTH_STRING
    Erase foBuff

    Exit Function
Shell_Copy_Error:
    Shell_Copy = lRet
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Shell_Copy of Class Module clsMehrFileOp"
End Function

'&H100 simple progress
Public Function Shell_Rename(mS As String, mD As String, Optional mFlags As Long = &H100) As Long
    On Error GoTo Shell_Rename_Error
    Dim SHFileOp As SHFILEOPSTRUCT
    Dim lRet As Long
    Dim lenFileop As Long
    lRet = -1

    lenFileop = LenB(SHFileOp) ' double word alignment increase
    ReDim foBuff(1 To lenFileop) ' the size of the structure.

    mSource = mS
    mDestination = mD

    If Right(mSource, NUM_ONE) = CHAR_BACK_SLASH Then
        mSource = Left(mSource, Len(mSource) - NUM_ONE) & vbNullChar & vbNullChar  'Chr$(0) & Chr$(0)
    Else
        mSource = mSource & vbNullChar & vbNullChar 'Chr$(0) & Chr$(0)
    End If
    If Right(mDestination, NUM_ONE) = CHAR_BACK_SLASH Then
        mDestination = Left(mDestination, Len(mDestination) - NUM_ONE) & vbNullChar & vbNullChar ' Chr$(0) & Chr$(0)
    Else
        mDestination = mDestination & vbNullChar & vbNullChar ' Chr$(0) & Chr$(0)
    End If
'    mSource = RemoveSlash(mSource) '& Chr$(0) & Chr$(0)
'    mDestination = RemoveSlash(mDestination) '& Chr$(0) & Chr$(0)
'
'Debug.Print "mSource " & mSource & "="
'Debug.Print "mDestination " & mDestination & "="
'
'Name mSource As mDestination
'lRet = NUM_ZERO

    With SHFileOp
        .hWnd = GetDesktopWindow 'frmProfiles.hWnd
        .wFunc = FO_RENAME
        .lpszFrom = StrPtr(mSource) 'or strptr(strconv(msource,vbfromunicode)) for ANSI
        .lpszTo = StrPtr(mDestination)
        .fFlags = mFlags
        'Causes a GPF
        '.lpszProgressTitle = "Proceed to rename:" & vbNullChar & vbNullChar
    End With

    ' Now we need to copy the structure into a byte array
    Call CopyMemory(foBuff(1), SHFileOp, lenFileop)

    ' Next we move the last 12 bytes by 2 to byte align the data
    Call CopyMemory(foBuff(19), foBuff(21), 12)

    lRet = W_SHFileOperation(SHFileOp)
    Shell_Rename = lRet
'
''            If lRet <> NUM_ZERO Then  ' Operation failed
''               MsgBox Err.LastDllError 'Show the error returned from
''                                       'the API.
''               Else
''               If SHFileOp.fAnyOperationsAborted <> 0 Then
''                  MsgBox "Operation Failed"
''               End If
''            End If

    mSource = CHAR_ZERO_LENGTH_STRING
    mDestination = CHAR_ZERO_LENGTH_STRING
    Erase foBuff

    Exit Function
Shell_Rename_Error:
    Shell_Rename = lRet
    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Shell_Rename of Class Module clsMehrFileOp" & vbCrLf & "Source =" & mSource & "=" & vbCrLf & "Destination =" & mDestination & "="
End Function

'&H40 allow undo
'Can delete entire dir with all it's sub dirs. Does not care about readonly or system attributes
'Can delete entire subdir structure => \*.* if flag fof_filesonly then only files are deleted
'Can delete partial files => \*.txt
Public Function Shell_Delete(mS As String, Optional mFlags As Long = &H40) As Long
    On Error GoTo Shell_Delete_Error
    Dim SHFileOp As SHFILEOPSTRUCT
    Dim lRet As Long
    Dim lenFileop As Long

    lRet = NUM_MINUS_ONE
    Shell_Delete = lRet

    'Root Drive
    If Len(mS) < NUM_FOUR Then
        Err.Raise 1001, "Class File Operations", "Unable to delete root drive: " & mS
        Exit Function
    End If

'    'Check for special folders
'    If modSysDirs.IsFolderSystemFolder(mSource) Then
'        Err.Raise 1002, "Class File Operations", "Unable to delete system directory" & vbCrLf & mS
'        Exit Function
'    End If

    lenFileop = LenB(SHFileOp) ' double word alignment increase
    ReDim foBuff(1 To lenFileop) ' the size of the structure.

    mSource = mS

    If Right(mSource, NUM_ONE) = CHAR_BACK_SLASH Then
        mSource = Left(mSource, Len(mSource) - NUM_ONE) & vbNullChar & vbNullChar  'Chr$(0) & Chr$(0)
    Else
        mSource = mSource & vbNullChar & vbNullChar 'Chr$(0) & Chr$(0)
    End If

'    mSource = RemoveSlash(mSource) & Chr$(0) & Chr$(0)

    With SHFileOp
        .hWnd = GetDesktopWindow 'frmProfiles.hWnd
        .wFunc = FO_DELETE
        .lpszFrom = StrPtr(mSource) 'or strptr(strconv(msource,vbfromunicode)) for ANSI
        '.lpszTo = vbNullChar 'not used
        .fFlags = mFlags
    End With

    ' Now we need to copy the structure into a byte array
    Call CopyMemory(foBuff(1), SHFileOp, lenFileop)

    ' Next we move the last 12 bytes by 2 to byte align the data
    Call CopyMemory(foBuff(19), foBuff(21), 12)

    lRet = W_SHFileOperation(SHFileOp)
    Shell_Delete = lRet

    mSource = CHAR_ZERO_LENGTH_STRING
    mDestination = CHAR_ZERO_LENGTH_STRING
    Erase foBuff

    Exit Function
Shell_Delete_Error:
    Shell_Delete = lRet
    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Shell_Delete of Class Module clsMehrFileOp" & vbCrLf & "Source " & mSource
End Function




'ShellExecLaunch "C:\mike\temp.txt"
Public Function Shell_ShellExecLaunch(strPathFile As String, _
                                Optional strParameters As String = vbNullString, _
                                Optional strOpenInPath As String = vbNullString, _
                                Optional strOperations As String = "open", _
                                Optional lShowCmd As Long = SW_NORMAL, _
                                Optional bDisplayErrors As Boolean = True) As Long

    Dim Scr_hDC As Long
    Dim lRet As Long
    Dim strErrMsg As String
    
    'Get the Desktop handle
    Scr_hDC = GetDesktopWindow()

    'Launch File
    lRet = A_ShellExecute(Scr_hDC, strOperations, strPathFile, strParameters, strOpenInPath, lShowCmd)
    If lRet < 33 Then
        'There was an error
        Select Case lRet
            Case SE_ERR_FNF
                strErrMsg = "File not found"
            Case SE_ERR_PNF
                strErrMsg = "Path not found"
            Case SE_ERR_ACCESSDENIED
                strErrMsg = "Access denied"
            Case SE_ERR_OOM
                strErrMsg = "Out of memory"
            Case SE_ERR_DLLNOTFOUND
                strErrMsg = "DLL not found"
            Case SE_ERR_SHARE
                strErrMsg = "A sharing violation occurred"
            Case SE_ERR_ASSOCINCOMPLETE
                strErrMsg = "Incomplete or invalid file association"
            Case SE_ERR_DDETIMEOUT
                strErrMsg = "DDE Time out"
            Case SE_ERR_DDEFAIL
                strErrMsg = "DDE transaction failed"
            Case SE_ERR_DDEBUSY
                strErrMsg = "DDE busy"
            Case SE_ERR_NOASSOC
                strErrMsg = "No association for file extension"
            Case ERROR_BAD_FORMAT
                strErrMsg = "Invalid EXE file or error in EXE image"
            Case ERROR_FILE_NOT_FOUND
                strErrMsg = "The specified file was not found."
            Case ERROR_PATH_NOT_FOUND
                strErrMsg = "The specified path was not found."
            Case ERROR_BAD_EXE_FORMAT
                strErrMsg = "The .exe file is invalid (non-Win32® .exe or error in .exe image)."
            Case Else
                strErrMsg = "Unknown error"
        End Select
        If bDisplayErrors Then A_MessageBox Scr_hDC, strErrMsg, "Error", MB_OK Or MB_ICONERROR
    End If
    Shell_ShellExecLaunch = lRet
End Function


