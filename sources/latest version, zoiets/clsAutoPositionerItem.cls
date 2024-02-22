VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAutoPositionerItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public oCTL As Object
Public oREL As Object
Private tPosType As tPOSITION_TYPE
Private lValueX As Long
Private lValueY As Long

Public Function setParams(objThisCTL As Object, objRefCTL As Object, tPositionType As tPOSITION_TYPE)
    Set oCTL = objThisCTL
    Set oREL = objRefCTL
    tPosType = tPositionType
End Function

Property Let ValueX(lValue As Long)
    lValueX = lValue
End Property
Property Get ValueX() As Long
    ValueX = lValueX
End Property

Property Let ValueY(lValue As Long)
    lValueY = lValue
End Property
Property Get ValueY() As Long
    ValueY = lValueY
End Property

Property Get isRelativePosRight() As Boolean
    isRelativePosRight = ((tPosType And RELATIVE_POS_RIGHT) <> 0)
End Property

Property Get isRelativePosBottom() As Boolean
    isRelativePosBottom = ((tPosType And RELATIVE_POS_BOTTOM) <> 0)
End Property
    
Property Get isDeltaWidthRight() As Boolean
    isDeltaWidthRight = ((tPosType And DELTA_WIDTH_RIGHT) <> 0)
End Property

Property Get isDeltaHeightBottom() As Boolean
    isDeltaHeightBottom = ((tPosType And DELTA_HEIGHT_BOTTOM) <> 0)
End Property

Property Get isRelativeSamePosHorizontal() As Boolean
    isRelativeSamePosHorizontal = ((tPosType And RELATIVE_SAME_POS_HORIZONTAL) <> 0)
End Property

Property Get isRelativeSamePosVertical() As Boolean
    isRelativeSamePosVertical = ((tPosType And RELATIVE_SAME_POS_VERTICAL) <> 0)
End Property

Property Get isPropWidth() As Boolean
    isPropWidth = ((tPosType And PROP_WIDTH) <> 0)
End Property

Property Get isPropHeight() As Boolean
    isPropHeight = ((tPosType And PROP_HEIGHT) <> 0)
End Property

