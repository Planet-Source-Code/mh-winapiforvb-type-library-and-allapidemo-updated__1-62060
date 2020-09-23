VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "URLHistoryItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'Taken from: Tweak & Clean
'http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=56482&lngWId=1

Dim m_STAT As STATURL

'*********************************************************************************************
'
' Returns a string from a unicode string pointer.
'
'*********************************************************************************************
Private Function Ptr2StrW(Ptr As Long) As String
    Dim sRtn() As Byte

    ' Check if the pointer is valid
    If Ptr <> 0 Then
    	'Modified
        ReDim sRtn(W_lstrlenPtr(Ptr) * 2)
        
        ' Copy the string to the byte array
        CopyMemory sRtn(0), ByVal Ptr, UBound(sRtn)
        Ptr2StrW = sRtn()
        
    End If
    
End Function

'*********************************************************************************************
'
' Initializes the object data
'
'*********************************************************************************************
Friend Sub Init(STAT As STATURL)

    m_STAT = STAT
    
End Sub

'*********************************************************************************************
'
' Returns the last visited date
'
'*********************************************************************************************
Public Property Get LastVisited() As Date
Dim ST As SYSTEMTIME, SerialDate As Double

    FileTimeToLocalFileTime m_STAT.ftLastVisited, m_STAT.ftLastVisited
    FileTimeToSystemTime m_STAT.ftLastVisited, ST
    SystemTimeToVariantTime ST, SerialDate
    
    LastVisited = SerialDate
    
End Property

'*********************************************************************************************
'
' Expires
'
'*********************************************************************************************
Public Property Get Expires() As Date
Dim ST As SYSTEMTIME, SerialDate As Double

    FileTimeToLocalFileTime m_STAT.ftExpires, m_STAT.ftExpires
    FileTimeToSystemTime m_STAT.ftExpires, ST
    SystemTimeToVariantTime ST, SerialDate
    
    Expires = SerialDate
    
End Property

'*********************************************************************************************
'
' Returns the last update date of the URL
'
'*********************************************************************************************
Public Property Get LastUpdated() As Date
Dim ST As SYSTEMTIME, SerialDate As Double

    FileTimeToLocalFileTime m_STAT.ftLastUpdated, m_STAT.ftLastUpdated
    FileTimeToSystemTime m_STAT.ftLastUpdated, ST
    SystemTimeToVariantTime ST, SerialDate
    
    LastUpdated = SerialDate
    
End Property

'*********************************************************************************************
'
' Returns the URL
'
'*********************************************************************************************
Public Property Get URL() As String

    URL = Ptr2StrW(m_STAT.pwcsUrl)
    
End Property

'*********************************************************************************************
'
' Returns the URL flags
'
'*********************************************************************************************
Public Property Get flags() As STATURLFLAGS

    flags = m_STAT.dwFlags
    
End Property

'*********************************************************************************************
'
' Returns the URL title
'
'*********************************************************************************************
Public Property Get Title() As String

    Title = Ptr2StrW(m_STAT.pwcsTitle)
    
    If Title = "" Then Title = "EMPTY"
        
End Property

Private Sub Class_Terminate()

    ' Release the strings from
    ' m_STAT variable
    CoTaskMemFree m_STAT.pwcsTitle
    CoTaskMemFree m_STAT.pwcsUrl
    
End Sub

