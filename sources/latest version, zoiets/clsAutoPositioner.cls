VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAutoPositioner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum tPOSITION_TYPE
    RELATIVE_POS_RIGHT = 1             'Keep horizontal pos relative to the right side of another ctrl/form
    RELATIVE_POS_BOTTOM = 2            'Keep vertical pos relative to the bottom of another ctrl/form
    DELTA_WIDTH_RIGHT = 4              'Vary width relative to the right side of another ctrl/form
    DELTA_HEIGHT_BOTTOM = 8            'Vary height relative to the bottom of another ctrl/form
    RELATIVE_SAME_POS_HORIZONTAL = 16  'Keep same horizontal pos relative to another ctrl/form
    RELATIVE_SAME_POS_VERTICAL = 32    'Keep same vertical pos relative to another ctrl/form
    PROP_WIDTH = 64
    PROP_HEIGHT = 128
End Enum

Dim m_oAssignments As New Collection

Public Function AddAssignment(ctl As Object, ctlRelative As Object, tPosType As tPOSITION_TYPE)
    Dim objPosItem As New clsAutoPositionerItem
    
    With objPosItem
        .setParams ctl, ctlRelative, tPosType
    
        If .isRelativePosRight Then .ValueX = getWidth(.oREL) - getLeft(.oCTL)
        If .isRelativePosBottom Then .ValueY = getHeight(.oREL) - getTop(.oCTL)
                    
        If .isDeltaWidthRight Then .ValueX = getWidth(.oREL) - (getLeft(.oCTL) + getWidth(.oCTL))
        If .isDeltaHeightBottom Then .ValueY = getHeight(.oREL) - (getTop(.oCTL) + getHeight(.oCTL))
                    
        If .isRelativeSamePosHorizontal Then .ValueX = getLeft(.oCTL) - getLeft(.oREL)
        If .isRelativeSamePosVertical Then .ValueY = getTop(.oCTL) - getTop(.oREL)
        
        If .isPropWidth Then
            .ValueX = Fix(100 * (getWidth(.oCTL) / getWidth(.oREL)))
        End If
        If .isPropHeight Then
            .ValueY = Fix(100 * (getHeight(.oCTL) / getHeight(.oREL)))
        End If
    End With
                
    m_oAssignments.Add objPosItem
End Function

Public Function RefreshPositions()
    Dim i As Long
    
    For i = 1 To m_oAssignments.Count
        With m_oAssignments.Item(i)
            If .isRelativePosRight Then setLeft .oCTL, getWidth(.oREL) - .ValueX
            If .isRelativePosBottom Then setTop .oCTL, getHeight(.oREL) - .ValueY
                    
            If .isDeltaWidthRight Then setWidth .oCTL, getWidth(.oREL) - getLeft(.oCTL) - .ValueX
            If .isDeltaHeightBottom Then setHeight .oCTL, getHeight(.oREL) - getTop(.oCTL) - .ValueY
                    
            If .isRelativeSamePosHorizontal Then setTop .oCTL, getTop(.oREL) + .ValueX
            If .isRelativeSamePosVertical Then setLeft .oCTL, getLeft(.oREL) + .ValueY
        
            If .isPropWidth Then
                setWidth .oCTL, Fix(getWidth(.oREL) * (.ValueX / 100))
            End If
            If .isPropHeight Then
                setHeight .oCTL, Fix(getHeight(.oREL) * (.ValueY / 100))
            End If
        End With
    Next
End Function

Private Function isLine(ctl As Object) As Boolean
    isLine = (StrComp(TypeName(ctl), "Line", vbTextCompare) = 0)
End Function

Private Function getWidth(ctl As Object) As Long
    If isLine(ctl) Then
        getWidth = ctl.X2 - ctl.X1
    Else
        getWidth = ctl.Width
    End If
End Function

Private Function getLeft(ctl As Object) As Long
    If isLine(ctl) Then
        getLeft = ctl.X1
    Else
        getLeft = ctl.Left
    End If
End Function

Private Function getHeight(ctl As Object) As Long
    If isLine(ctl) Then
        getHeight = ctl.Y2 - ctl.Y1
    Else
        getHeight = ctl.Height
    End If
End Function

Private Function getTop(ctl As Object) As Long
    If isLine(ctl) Then
        getTop = ctl.Y1
    Else
        getTop = ctl.Top
    End If
End Function

Private Sub setLeft(ctl As Object, ByVal newLeft As Long)
    If isLine(ctl) And (newLeft >= 0) Then
        ctl.X2 = ctl.X2 + (newLeft - ctl.X1)
        ctl.X1 = newLeft
    ElseIf newLeft > 0 Then
        ctl.Left = newLeft
    End If
End Sub

Private Sub setTop(ctl As Object, ByVal newTop As Long)
    If isLine(ctl) And (newTop >= 0) Then
        ctl.Y2 = ctl.Y2 + (newTop - ctl.Y1)
        ctl.Y1 = newTop
    ElseIf newTop > 0 Then
        ctl.Top = newTop
    End If
End Sub

Private Sub setWidth(ctl As Object, ByVal newWidth As Long)
    If isLine(ctl) And (newWidth >= 0) Then
        ctl.X2 = ctl.X1 + newWidth
    ElseIf newWidth > 0 Then
        ctl.Width = newWidth
    End If
End Sub

Private Sub setHeight(ctl As Object, ByVal newHeight As Long)
    If isLine(ctl) And (newHeight >= 0) Then
        ctl.Y2 = ctl.Y1 + newHeight
    ElseIf newHeight > 0 Then
        ctl.Height = newHeight
    End If
End Sub

