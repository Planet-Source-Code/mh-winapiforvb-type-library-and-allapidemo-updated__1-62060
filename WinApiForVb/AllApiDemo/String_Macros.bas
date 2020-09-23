Attribute VB_Name = "String_Macros"
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''
''''''''''A bunch of usefull string macros
'''''''''''''''''''''''''''''''''''''''''''

Private Const sTimeFormat As String = "HH:MM:SS"

Public Function String_ByteArrayToString(Bytes() As Byte) As String
    Dim iUnicode As Long, i As Long, j As Long
    'Dim sUni As String
    
    On Error Resume Next
    i = UBound(Bytes)
    
    If (i < NUM_ONE) Then
        'ANSI, just convert to unicode and return
        String_ByteArrayToString = StrConv(Bytes, vbUnicode)
        Exit Function
    End If
    i = i + NUM_ONE
    
    'Examine the first two bytes
    CopyMemory iUnicode, ByVal VarPtr(Bytes(0)), 2
    
    If iUnicode = Bytes(NUM_ZERO) Then 'Unicode
        Debug.Print "Unicode " & iUnicode
        'Account for terminating null
        If (i Mod NUM_TWO) Then i = i - NUM_ONE
        'Set up a buffer to recieve the string
        String_ByteArrayToString = String$(i / NUM_TWO, 0)
        'Copy to string
        CopyMemoryToStr String_ByteArrayToString, ByVal VarPtr(Bytes(0)), i
    Else 'ANSI
        String_ByteArrayToString = StrConv(Bytes, vbUnicode)
    End If
                    
End Function

Public Function String_ToByteArray(strInput As String, _
                                Optional bReturnAsUnicode As Boolean = False, _
                                Optional bAddToArray As Boolean = False, _
                                Optional lHowManyCharactrs As Long = 0) As Byte()
    
    Dim lRet As Long, lUsed As Long
    Dim bytBuffer() As Byte
    Dim lLenB As Long
    
    If bReturnAsUnicode Then
        'Number of bytes
        lLenB = LenB(strInput)
        'Resize buffer?
        If bAddToArray Then
            ReDim bytBuffer(lLenB + lHowManyCharactrs)
        Else
            ReDim bytBuffer(lLenB - NUM_ONE)
        End If
        'Copy characters from string to byte array
        CopyMemoryFromStr ByVal VarPtr(bytBuffer(0)), strInput, lLenB
    Else
'        'METHOD ONE
'        lLenB = Len(strInput)
'        'Get rid of embedded nulls
'        If bAddNullTerminator Then
'            ReDim bytBuffer(lLenB + 1)
'        Else
'            ReDim bytBuffer(lLenB)
'        End If
'        CopyMemory bytBuffer(0), ByVal StrPtr(StrConv(strInput, vbFromUnicode)), lLenB

        'METHOD TWO
        'Num of characters
        lLenB = Len(strInput)
        If bAddToArray Then
            ReDim bytBuffer(lLenB + lHowManyCharactrs)
        Else
            ReDim bytBuffer(lLenB - NUM_ONE)
        End If
        lRet = WideCharToMultiByte(CP_ACP, 0&, strInput, -1, ByVal VarPtr(bytBuffer(0)), lLenB, 0&, lUsed)
    End If
    
    String_ToByteArray = bytBuffer
    
End Function

'Does not include terminating null
'Returns the address of first element of a ANSI string (ANSI character array)
'This function has to be called like so
'CopyMemory ByVal VarPtr(bTmp(0)), ByVal StringToAnsiPtr(sTmp), lLen
'if used like var = StringToAnsiPtr(str) then var used to pass to an API call, it fails
Public Function String_ToAnsiPtr(strInput As String) As Long
    String_ToAnsiPtr = StrPtr(StrConv(strInput, vbFromUnicode))
End Function

Public Sub String_SwapStrings(pbString1 As String, pbString2 As String)
    Dim l_Hold As Long
    'Copy BSTR value (address of pbString1 character array) into l_Hold
    CopyMemory l_Hold, ByVal VarPtr(pbString1), 4
    'Copy BSTR value (address of pbString1 character array) into BSTR value (address of pbString2 character array)
    'So now pbString2 address is pointing to pbString1 character array
    'Copy value of VarPtr(pbString2) to VarPtr(pbString1) value, chnaging what BSTR's are pointing to
    CopyMemory ByVal VarPtr(pbString1), ByVal VarPtr(pbString2), 4
    'Now change the address of BSTR pbString1 to point to pbString2 former character array'
    CopyMemory ByVal VarPtr(pbString2), l_Hold, 4
End Sub

Public Function String_GetDay(ByVal iDate As Long) As String
    'Weekday starts from sunday
    Select Case iDate
        Case 1
            String_GetDay = "Sunday"
        Case 2
            String_GetDay = "Monday"
        Case 3
            String_GetDay = "Tuesday"
        Case 4
            String_GetDay = "Wednesday"
        Case 5
            String_GetDay = "Thursday"
        Case 6
            String_GetDay = "Friday"
        Case 7
            String_GetDay = "Saturday"
    End Select
End Function

Public Function String_GetMonthToString(ByVal iMonth As Long) As String
    Select Case iMonth
        Case 1
            String_GetMonthToString = "January"
        Case 2
            String_GetMonthToString = "Feburary"
        Case 3
            String_GetMonthToString = "March"
        Case 4
            String_GetMonthToString = "April"
        Case 5
            String_GetMonthToString = "May"
        Case 6
            String_GetMonthToString = "June"
        Case 7
            String_GetMonthToString = "July"
        Case 8
            String_GetMonthToString = "August"
        Case 9
            String_GetMonthToString = "September"
        Case 10
            String_GetMonthToString = "October"
        Case 11
            String_GetMonthToString = "November"
        Case 12
            String_GetMonthToString = "December"
    End Select

End Function

'Determines if a string starts with the same characters as
'CheckFor string
Public Function String_StartsWith(ByVal strValue As String, _
  CheckFor As String, Optional CompareType As VbCompareMethod _
   = vbBinaryCompare) As Boolean

'True if starts with CheckFor, false otherwise
'Case sensitive by default.  If you want non-case sensitive, set
'last parameter to vbTextCompare
    
    'Examples:
    'msgbox StringStartsWith("Test", "TE") 'false
    'msgbox StringStartsWith("Test", "TE", vbTextCompare) 'True
    On Error GoTo TrapStringStartsWith
  Dim sCompare As String
  Dim lLen As Long
   
  lLen = Len(CheckFor)
  If lLen > Len(strValue) Then Exit Function
  sCompare = Left(strValue, lLen)
  String_StartsWith = (StrComp(sCompare, CheckFor, CompareType) = NUM_ZERO)
  Exit Function
TrapStringStartsWith:

End Function

 'Determines if a string ends with the same characters as
 'CheckFor string
Public Function String_EndsWith(ByVal strValue As String, _
   CheckFor As String, Optional CompareType As VbCompareMethod _
   = vbBinaryCompare) As Boolean
 
 'True if end with CheckFor, false otherwise
 'Case sensitive by default.  If you want non-case sensitive, set
 'last parameter to vbTextCompare
 
  'Examples
  'msgbox StringEndsWith("Test", "ST") 'False
  'msgbox StringEndsWith("Test", "ST", vbTextCompare) 'True

  On Error GoTo TrapStringEndsWith

  Dim sCompare As String
  Dim lLen As Long

  lLen = Len(CheckFor)
  If lLen > Len(strValue) Then Exit Function
  sCompare = Right(strValue, lLen)
  String_EndsWith = (StrComp(sCompare, CheckFor, CompareType) = NUM_ZERO)
  Exit Function
    
TrapStringEndsWith:

End Function

'If the lenght of a string is more than 50 it replaces the rest with ....
'=shortenString(somestr)
Public Function String_Shorten(strString As String, lNewLen As Long) As String
    On Error Resume Next
    String_Shorten = strString
    If lNewLen <= NUM_ZERO Or lNewLen < Len(strString) Then Exit Function
    String_Shorten = Left(strString, lNewLen) & CHAR_TRIPLE_DOT
End Function

Public Function String_ConvertTime(TheTime As Single) As String
    Dim NewTime As String
    Dim Sec As Single
    Dim Min As Single
    Dim h As Single
    
    If TheTime <= NUM_ZERO Then
        String_ConvertTime = "00:00:00"
        Exit Function
    End If

    If TheTime > NUM_SIXTY Then
        Sec = Round(TheTime)
        Min = Sec / NUM_SIXTY
        Min = Int(Min)
        Sec = Sec - Min * NUM_SIXTY
        h = Int(Min / NUM_SIXTY)
        Min = Min - h * NUM_SIXTY
        NewTime = h & CHAR_COLON & Min & CHAR_COLON & Sec
        If h < NUM_ZERO Then h = NUM_ZERO
        If Min < NUM_ZERO Then Min = NUM_ZERO
        If Sec < NUM_ZERO Then Sec = NUM_ZERO
        NewTime = Format(NewTime, sTimeFormat)
        String_ConvertTime = NewTime
    End If

    If TheTime < NUM_SIXTY Then
        NewTime = "00:00:" & Round(TheTime)
        NewTime = Format(NewTime, sTimeFormat)
        String_ConvertTime = NewTime
    End If
End Function

' Convert 'pointer to (wide-)null-terminated Unicode string' to a 'VB stringValue '
' lBytesCopied: the number of bytes copied
Public Function String_LPWSTRtoBSTR(lStrptr As Long, lBytesCopied As Long) As String
    lBytesCopied = W_lstrlenPtr(lStrptr)
    String_LPWSTRtoBSTR = String$(lBytesCopied, 0)
    lBytesCopied = lBytesCopied * 2
    CopyMemoryToStr String_LPWSTRtoBSTR, ByVal lStrptr, lBytesCopied
End Function

' Convert 'pointer to Ansi string' to a 'VB string value'
Public Function String_LPSTRtoBSTR(lStrptr As Long, lBytesCopied As Long) As String
    Dim cChars As Long
    lBytesCopied = A_lstrlenPtr(lStrptr)
    If lBytesCopied = NUM_ZERO Then Exit Function
    String_LPSTRtoBSTR = String$(lBytesCopied, vbNullChar)
    'CopyMemoryToStr String_LPSTRtoBSTR, ByVal lStrptr, lBytesCopied
    A_lstrcpynPtrStr String_LPSTRtoBSTR, ByVal lStrptr, lBytesCopied
    String_LPSTRtoBSTR = StrConv(String_LPSTRtoBSTR, vbUnicode)
End Function


Public Function String_IsCharBadPath(KeyAscii As Integer) As Boolean
'34 "
'42 *
'47 /
'58 :
'60 <
'62 <
'63 ?
'92 \
'124 |
    If KeyAscii = CHAR_DOUBLE_QUATAION Or KeyAscii = ASCII_STAR Or KeyAscii = ASCII_FORWARD_SLASH Or _
        KeyAscii = ASCII_COLON Or KeyAscii = ASCII_LEFT_ARROW Or KeyAscii = ASCII_RIGHT_ARROW Or _
        KeyAscii = ASCII_QUESTION Or KeyAscii = ASCII_BACK_SLASH Or KeyAscii = ASCII_PIPE Then
            String_IsCharBadPath = True
    End If
End Function

Public Function String_IsCharLower(KeyAscii As Integer) As Boolean
'a = 97 to z = 122
    'If KeyAscii > 96 And KeyAscii < 123 Then IsKeyAsciiCharLower = True
    String_IsCharLower = CBool(A_IsCharLower(KeyAscii))
End Function

Public Function String_IsCharUpper(KeyAscii As Integer) As Boolean
'A = 65 to Z = 90
    'If KeyAscii > 64 And KeyAscii < 91 Then IsKeyAsciiCharUpper = True
    String_IsCharUpper = CBool(A_IsCharUpper(KeyAscii))
End Function

'determines whether a character is either an alphabetical or a numeric character
Public Function String_IsCharNumber(KeyAscii As Integer) As Boolean
'0  = 48 to 9 = 57
    If KeyAscii > 47 And KeyAscii < 58 Then String_IsCharNumber = True
End Function

Public Function String_IsCharAlpha(KeyAscii As Integer) As Boolean
    String_IsCharAlpha = CBool(A_IsCharAlpha(KeyAscii))
End Function

Public Function TranslateHTTPResponseCode(lHttpCode As Long) As String
  If lHttpCode = HTTP_STATUS_CONTINUE Then
    TranslateHTTPResponseCode = " - Continue"
      ElseIf lHttpCode = HTTP_STATUS_SWITCH_PROTOCOLS Then
        TranslateHTTPResponseCode = " - Switching Protocols"
      ElseIf lHttpCode = HTTP_STATUS_OK Then
        TranslateHTTPResponseCode = " - OK"
      ElseIf lHttpCode = HTTP_STATUS_CREATED Then
        TranslateHTTPResponseCode = " - Created"
      ElseIf lHttpCode = HTTP_STATUS_ACCEPTED Then
        TranslateHTTPResponseCode = " - Accepted"
      ElseIf lHttpCode = 203 Then
        TranslateHTTPResponseCode = " - Non-Authoritative Information"
      ElseIf lHttpCode = 204 Then
        TranslateHTTPResponseCode = " - No Content"
      ElseIf lHttpCode = 205 Then
        TranslateHTTPResponseCode = " - Reset Content"
      ElseIf lHttpCode = 206 Then
        TranslateHTTPResponseCode = " - Partial Content"
      ElseIf lHttpCode = 300 Then
        TranslateHTTPResponseCode = " - Multiple Choices"
      ElseIf lHttpCode = 301 Then
        TranslateHTTPResponseCode = " - Moved Permanently"
      ElseIf lHttpCode = 302 Then
        TranslateHTTPResponseCode = " - Moved Temporarily"
      ElseIf lHttpCode = 303 Then
        TranslateHTTPResponseCode = " - See Other"
      ElseIf lHttpCode = 304 Then
        TranslateHTTPResponseCode = " - Not Modified"
      ElseIf lHttpCode = 305 Then
        TranslateHTTPResponseCode = " - Use Proxy"
      ElseIf lHttpCode = 400 Then
        TranslateHTTPResponseCode = " - Bad Request/Invalid syntax"
      ElseIf lHttpCode = 401 Then
        TranslateHTTPResponseCode = " - Unauthorized"
      ElseIf lHttpCode = 402 Then
        TranslateHTTPResponseCode = " - Payment Required"
      ElseIf lHttpCode = 403 Then
        TranslateHTTPResponseCode = " - Forbidden"
      ElseIf lHttpCode = 404 Then
        TranslateHTTPResponseCode = " - Not Found"
      ElseIf lHttpCode = 405 Then
        TranslateHTTPResponseCode = " - Method Not Allowed"
      ElseIf lHttpCode = 406 Then
        TranslateHTTPResponseCode = " - Not Acceptable"
      ElseIf lHttpCode = 407 Then
        TranslateHTTPResponseCode = " - Proxy Authentication Required"
      ElseIf lHttpCode = 408 Then
        TranslateHTTPResponseCode = " - Request Time-Out"
      ElseIf lHttpCode = 409 Then
        TranslateHTTPResponseCode = " - Conflict"
      ElseIf lHttpCode = 410 Then
        TranslateHTTPResponseCode = " - Gone"
      ElseIf lHttpCode = 411 Then
        TranslateHTTPResponseCode = " - Length Required"
      ElseIf lHttpCode = 412 Then
        TranslateHTTPResponseCode = " - Precondition Failed"
      ElseIf lHttpCode = 413 Then
        TranslateHTTPResponseCode = " - Request Entity Too Large"
      ElseIf lHttpCode = 414 Then
        TranslateHTTPResponseCode = " - Request-URL Too Large"
      ElseIf lHttpCode = 415 Then
        TranslateHTTPResponseCode = " - Unsupported Media Type"
      ElseIf lHttpCode = 500 Then
        TranslateHTTPResponseCode = " - Server error"
      ElseIf lHttpCode = 501 Then
        TranslateHTTPResponseCode = " - Not Implemented"
      ElseIf lHttpCode = 502 Then
        TranslateHTTPResponseCode = " - Bad Gateway"
      ElseIf lHttpCode = 503 Then
        TranslateHTTPResponseCode = " - Out of Resources"
      ElseIf lHttpCode = 504 Then
        TranslateHTTPResponseCode = " - Gateway Time-Out"
      ElseIf lHttpCode = 505 Then
        TranslateHTTPResponseCode = " - HTTP Not Version"
    Else
      TranslateHTTPResponseCode = " - N/A"
  End If
End Function




