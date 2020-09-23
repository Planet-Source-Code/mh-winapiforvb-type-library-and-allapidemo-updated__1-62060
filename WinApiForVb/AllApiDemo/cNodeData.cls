VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cNodeData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'A class to manage extra info for each node in tree
'I thought of using cNode and cNodes type of class hirearchy
'Then I decided to go with one class and a bunch of arrays to manage this task
'This is faster and less resource hungry (I know in the age of P4's who cares, right?)
'Also, you can add/remove items from arrays dynamically

Private hNodes() As Long 'Node handles
Private lLevels() As Long 'Root = 0, ...
Private lBackColors() As Long
Private lForeColors() As Long
Private hFonts() As Long
Private lFontStyles() As Long 'CFE_BOLD, CFE_ITALIC
Private sTags() As String

Private mElements As Long
'Temp vars used in loops and functions
Private mCounter As Long
Private mUbound As Long
Private mIndex As Long

'================GET/SET NODE ATTRIBUTES
'In Get procedures, either the lNode or lIndex must be valid

'Get individual node attributes
Public Function GetNodeLevel(Optional lNode As Long = -1, Optional lIndex As Long = -1) As Long
    If lIndex > -1 Then
        GetNodeLevel = lLevels(lIndex)
    Else
        GetNodeLevel = -1
        mIndex = FindNodeIndex(lNode)
        If mIndex > -1 Then GetNodeLevel = lLevels(mIndex)
    End If
End Function

Public Sub SetNodeLevel(lNewLevel As Long, Optional lNode As Long = -1, Optional lIndex As Long = -1)
    If lIndex > -1 Then
        lLevels(lIndex) = lNewLevel
    Else
        mIndex = FindNodeIndex(lNode)
        If mIndex > -1 Then lLevels(mIndex) = lNewLevel
    End If
End Sub

Public Function GetNodeBackColor(Optional lNode As Long = -1, Optional lIndex As Long = -1) As Long
    If lIndex > -1 Then
        GetNodeBackColor = lBackColors(lIndex)
    Else
        GetNodeBackColor = -1
        mIndex = FindNodeIndex(lNode)
        If mIndex > -1 Then GetNodeBackColor = lBackColors(mIndex)
    End If
End Function

Public Sub SetNodeBackColor(lNewColor As Long, Optional lNode As Long = -1, Optional lIndex As Long = -1)
    If lIndex > -1 Then
        lBackColors(lIndex) = lNewColor
    Else
        mIndex = FindNodeIndex(lNode)
        If mIndex > -1 Then lBackColors(mIndex) = lNewColor
    End If
End Sub

Public Function GetNodeForeColor(Optional lNode As Long = -1, Optional lIndex As Long = -1) As Long
    If lIndex > -1 Then
        GetNodeForeColor = lForeColors(lIndex)
    Else
        GetNodeForeColor = -1
        mIndex = FindNodeIndex(lNode)
        If mIndex > -1 Then GetNodeForeColor = lForeColors(mIndex)
    End If
End Function

Public Sub SetNodeForeColor(lNewColor As Long, Optional lNode As Long = -1, Optional lIndex As Long = -1)
    If lIndex > -1 Then
        lForeColors(lIndex) = lNewColor
    Else
        mIndex = FindNodeIndex(lNode)
        If mIndex > -1 Then lForeColors(mIndex) = lNewColor
    End If
End Sub

'Use either lNode or lIndex
Public Function GetNodeFont(Optional lNode As Long = -1, Optional lIndex As Long = -1) As Long
    If lIndex > -1 Then
        GetNodeFont = hFonts(lIndex)
    Else
        GetNodeFont = -1
        mIndex = FindNodeIndex(lNode)
        If mIndex > -1 Then GetNodeFont = hFonts(mIndex)
    End If
End Function

Public Sub SetNodeFont(hNewFont As Long, Optional lNode As Long = -1, Optional lIndex As Long = -1)
    If lIndex > -1 Then
        hFonts(lIndex) = hNewFont
    Else
        mIndex = FindNodeIndex(lNode)
        If mIndex > -1 Then hFonts(mIndex) = hNewFont
    End If
End Sub

Public Function GetNodeFontStyle(Optional lNode As Long = -1, Optional lIndex As Long = -1) As Long
    If lIndex > -1 Then
        GetNodeFontStyle = lFontStyles(lIndex)
    Else
        GetNodeFontStyle = -1
        mIndex = FindNodeIndex(lNode)
        If mIndex > -1 Then GetNodeFontStyle = lFontStyles(mIndex)
    End If
End Function

Public Sub SetNodeFontStyle(hNewStyle As Long, Optional lNode As Long = -1, Optional lIndex As Long = -1)
    If lIndex > -1 Then
        lFontStyles(lIndex) = hNewStyle
    Else
        mIndex = FindNodeIndex(lNode)
        If mIndex > -1 Then lFontStyles(mIndex) = hNewStyle
    End If
End Sub

Public Function GetNodeTag(Optional lNode As Long = -1, Optional lIndex As Long = -1) As String
    If lIndex > -1 Then
        GetNodeTag = sTags(lIndex)
    Else
        mIndex = FindNodeIndex(lNode)
        If mIndex > -1 Then GetNodeTag = sTags(mIndex)
    End If
End Function

Public Sub SetNodeTag(sNewTag As String, Optional lNode As Long = -1, Optional lIndex As Long = -1)
    If lIndex > -1 Then
        sTags(lIndex) = sNewTag
    Else
        mIndex = FindNodeIndex(lNode)
        If mIndex > -1 Then sTags(mIndex) = sNewTag
    End If
End Sub

'Get all node attributes at one shot
Public Function GetNodeData(sTag As String, _
                                lLevel As Long, _
                                lBackColor As Long, _
                                lForeColor As Long, _
                                hFont As Long, _
                                lFontStyle As Long, _
                                Optional lNode As Long = -1, _
                                Optional lIndex As Long = -1) As Long
    If lIndex = -1 Then lIndex = FindNodeIndex(lNode)
    GetNodeData = lIndex
    If GetNodeData > -1 Then
        lLevel = lLevels(GetNodeData)
        lBackColor = lBackColors(GetNodeData)
        lForeColor = lForeColors(GetNodeData)
        hFont = hFonts(GetNodeData)
        lFontStyle = lFontStyles(GetNodeData)
        sTag = sTags(GetNodeData)
    End If
End Function

Public Sub SetNodeData(Optional lNode As Long = -1, _
                        Optional lIndex As Long = -1, _
                        Optional lLevel As Long = 0, _
                        Optional lBackColor As Long = 0, _
                        Optional lForeColor As Long = 0, _
                        Optional hFont As Long = 0, _
                        Optional CFE_FontStyle As Long = 0, _
                        Optional sTag As String = "")
    If lIndex = -1 Then mIndex = FindNodeIndex(lNode)
    If mIndex = -1 Then Exit Sub
    'Set values
    If lLevel > 0 Then lLevels(mIndex) = lLevel
    If lBackColor > 0 Then lBackColors(mIndex) = lBackColor
    If lForeColor > 0 Then lForeColors(mIndex) = lForeColor
    If hFont <> 0 Then hFonts(mIndex) = hFont
    If CFE_FontStyle > 0 Then lFontStyles(mIndex) = CFE_FontStyle
    If LenB(sTag) > 0 Then sTags(mIndex) = sTag
End Sub

'================END GET/SET NODE ATTRIBUTES

'================ADDING RESIZING FINDING,...

'Returns -1 if not found
Public Function FindNodeIndex(lNode As Long) As Long
    FindNodeIndex = -1
    If mElements = 0 Then Exit Function
    
    mUbound = mElements - 1
    'Debug.Print "lNode: " & lNode & " melements: " & mElements
    For mCounter = 0 To mUbound
        'Debug.Print "hNodes(" & CStr(mCounter) & "): " & hNodes(mCounter)
        If hNodes(mCounter) = lNode Then
            FindNodeIndex = mCounter
            Exit Function
        End If
    Next
End Function

Public Sub AddNodeData(hNode As Long, _
                        Optional lLevel As Long = 0, _
                        Optional lBackColor As Long = 0, _
                        Optional lForeColor As Long = 0, _
                        Optional hFont As Long = 0, _
                        Optional CFE_FontStyle As Long = 0, _
                        Optional sTag As String = "")
    'Do we need to increase storage
    If mElements >= UBound(hNodes) Then SetArraySize
    
    'Set values
    hNodes(mElements) = hNode
    lLevels(mElements) = lLevel
    lBackColors(mElements) = lBackColor
    lForeColors(mElements) = lForeColor
    hFonts(mElements) = hFont
    lFontStyles(mElements) = CFE_FontStyle
    sTags(mElements) = sTag
    'Increase count
    mElements = mElements + 1
    'Debug.Print "hNode: " & hNode & "mElements: " & mElements & " hFont: " & hFont
End Sub

' Return the current array size
Public Property Get ArraySize() As Long
    ArraySize = UBound(hNodes)
End Property

' Return current amount of elements in the array
Public Property Get NodesCount() As Long
    NodesCount = mElements
End Property

'You can allocate a storage prior to adding nodes to speed things up
Public Sub SetArraySize(Optional lHowMany As Long = 1)
    If mElements > 0 Then
        'Ubound + 1 = Array Count
        lHowMany = UBound(hNodes) + lHowMany + 1
        ReDim Preserve hNodes(lHowMany)
        ReDim Preserve lLevels(lHowMany)
        ReDim Preserve lBackColors(lHowMany)
        ReDim Preserve lForeColors(lHowMany)
        ReDim Preserve hFonts(lHowMany)
        ReDim Preserve lFontStyles(lHowMany)
        ReDim Preserve sTags(lHowMany)
    Else 'First time
        lHowMany = lHowMany - 1 '0 one element,...
        ReDim hNodes(lHowMany)
        ReDim lLevels(lHowMany)
        ReDim lBackColors(lHowMany)
        ReDim lForeColors(lHowMany)
        ReDim hFonts(lHowMany)
        ReDim lFontStyles(lHowMany)
        ReDim sTags(lHowMany)
    End If
End Sub

'Removes the corresponding array elements
Public Sub DeleteNodeData(lNode As Long)
    Dim iLBound As Long, iUbound As Long, iTemp  As Long
    
    If mElements = 0 Then Exit Sub
    mIndex = FindNodeIndex(lNode)
    If mIndex > -1 Then Exit Sub
    
    iLBound = LBound(hNodes)
    iUbound = UBound(hNodes)
    
    'Get string pointer
    iTemp = StrPtr(sTags(mIndex))
    'Shift elements one up
    CopyMemory ByVal VarPtr(sTags(mIndex)), ByVal VarPtr(sTags(mIndex + 1)), (iUbound - mIndex) * 4
    ' we do this to have VB unalloc the string to evade memory leaks
    CopyMemory ByVal VarPtr(sTags(iUbound)), iTemp, 4
    'Resize
    ReDim Preserve sTags(iUbound - 1)
    
    'Just for a bit of speed
    iTemp = iUbound - iLBound - mIndex
    iLBound = mIndex + 1
    iUbound = iUbound - 1
    
    'With long arrays, we simply shift elements and resize
    'Len(hNodes(mIndex)) returns the number of bytes required to store this variable
    CopyMemory hNodes(mIndex), hNodes(iLBound), iTemp * Len(hNodes(mIndex))
    ReDim Preserve hNodes(iUbound)
    
    CopyMemory lLevels(mIndex), lLevels(iLBound), iTemp * Len(lLevels(mIndex))
    ReDim Preserve lLevels(iUbound)
    
    CopyMemory lBackColors(mIndex), lBackColors(iLBound), iTemp * Len(lBackColors(mIndex))
    ReDim Preserve lBackColors(iUbound)
    
    CopyMemory lForeColors(mIndex), lForeColors(iLBound), iTemp * Len(lForeColors(mIndex))
    ReDim Preserve lForeColors(iUbound)
    
    CopyMemory hFonts(mIndex), hFonts(iLBound), iTemp * Len(hFonts(mIndex))
    ReDim Preserve hFonts(iUbound)
    
    CopyMemory lFontStyles(mIndex), lFontStyles(iLBound), iTemp * Len(lFontStyles(mIndex))
    ReDim Preserve lFontStyles(iUbound)
    
    'Adjust count
    mElements = mElements - 1
End Sub

Public Sub EraseNodeData()
    On Error Resume Next
    
    'Make sure we delete all created fonts
    If mElements > 0 Then
        mUbound = mElements - 1
        For mCounter = 0 To mUbound
            If hFonts(mCounter) <> 0 Then
                DeleteObject hFonts(mCounter)
            End If
        Next
    End If
    Erase hNodes
    Erase lLevels
    Erase lBackColors
    Erase lForeColors
    Erase hFonts
    Erase lFontStyles
    Erase sTags
    mElements = 0
    mCounter = 0
    mUbound = 0
    mIndex = 0
End Sub

Private Sub Class_Terminate()
    EraseNodeData
End Sub

