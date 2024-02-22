Attribute VB_Name = "ModPixels"
Option Explicit

' ****************************************************
' For information on this module and other graphics
' programming, see the book "Visual Basic Graphics
' Programming, Second Edition." For more information,
' go to:
'
'       http://www.vb-helper.com/vbgp.htm
' ****************************************************

' ------------------------
' Bitmap Array Information
' ------------------------
Public Type RGBTriplet
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
End Type

' ------------------
' Bitmap Information
' ------------------
Public Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Public Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Public Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Public Enum bmphErrors
    bmphInvalidBitmapBits = vbObjectError + 1001
    bmphPaletteError
End Enum

' -------------------
' Palette Information
' -------------------
Private Type PALETTEENTRY
    peRed As Byte
    peGreen As Byte
    peBlue As Byte
    peFlags As Byte
End Type
Private Declare Function GetNearestPaletteIndex Lib "gdi32" (ByVal hPalette As Long, ByVal crColor As Long) As Long
Private Declare Function GetPaletteEntries Lib "gdi32" (ByVal hPalette As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hdc As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function ResizePalette Lib "gdi32" (ByVal hPalette As Long, ByVal nNumEntries As Long) As Long
Private Declare Function SetPaletteEntries Lib "gdi32" (ByVal hPalette As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Const MAX_PALETTE_SIZE = 256
Private Const PC_NOCOLLAPSE = &H4    ' Do not match color existing entries.

' -------------------------------
' System Capabilities Information
' -------------------------------
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Const NUMRESERVED = 106  ' Number of reserved entries in system palette.
Private Const SIZEPALETTE = 104  ' Size of system palette.

' Copy memory quickly. Used for 24-bit images.
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

' Load the control's palette so it matches the
' system palette.
Private Sub MatchColorPalette(ByVal pic As PictureBox)
Dim log_hpal As Long
Dim sys_pal(0 To MAX_PALETTE_SIZE - 1) As PALETTEENTRY
Dim orig_pal(0 To MAX_PALETTE_SIZE - 1) As PALETTEENTRY
Dim i As Integer
Dim sys_pal_size As Long
Dim num_static_colors As Long
Dim static_color_1 As Long
Dim static_color_2 As Long

    ' Make sure pic has the foreground palette.
    pic.ZOrder
    RealizePalette pic.hdc
    DoEvents

    ' Get system palette size and # static colors.
    sys_pal_size = GetDeviceCaps(pic.hdc, SIZEPALETTE)
    num_static_colors = GetDeviceCaps(pic.hdc, NUMRESERVED)
    static_color_1 = num_static_colors \ 2 - 1
    static_color_2 = sys_pal_size - num_static_colors \ 2

    ' Get the system palette entries.
    GetSystemPaletteEntries pic.hdc, 0, _
        sys_pal_size, sys_pal(0)

    ' Make the logical palette as big as possible.
    log_hpal = pic.Picture.hpal
    If ResizePalette(log_hpal, sys_pal_size) = 0 Then
        Err.Raise bmphPaletteError, _
            "DDBHelper.MatchColorPalette", _
            "Error matching bitmap palette"
    End If

    ' Blank the non-static colors.
    For i = 0 To static_color_1
        orig_pal(i) = sys_pal(i)
    Next i
    For i = static_color_1 + 1 To static_color_2 - 1
        With orig_pal(i)
            .peRed = 0
            .peGreen = 0
            .peBlue = 0
            .peFlags = PC_NOCOLLAPSE
        End With
    Next i
    For i = static_color_2 To 255
        orig_pal(i) = sys_pal(i)
    Next i
    SetPaletteEntries log_hpal, 0, sys_pal_size, orig_pal(0)

    ' Insert the non-static colors.
    For i = static_color_1 + 1 To static_color_2 - 1
        orig_pal(i) = sys_pal(i)
        orig_pal(i).peFlags = PC_NOCOLLAPSE
    Next i
    SetPaletteEntries log_hpal, static_color_1 + 1, static_color_2 - static_color_1 - 1, orig_pal(static_color_1 + 1)

    ' Realize the new palette.
    RealizePalette pic.hdc
End Sub
' Return a binary representation of the byte.
' This helper function is useful for understanding
' byte values.
Public Function BinaryByte(ByVal value As Byte) As String
Dim i As Integer
Dim txt As String

    For i = 1 To 8
        If value And 1 Then
            txt = "1" & txt
        Else
            txt = "0" & txt
        End If
        value = value \ 2
    Next i

    BinaryByte = txt
End Function

' Load the bits from this PictureBox into a
' two-dimensional array of RGB values. Set
' bits_per_pixel to be the number of bits per pixel.
Public Sub GetBitmapPixels(ByVal pic As PictureBox, ByRef pixels() As RGBTriplet, ByRef bits_per_pixel As Integer)
' Uncomment the following to make the routine
' display information about the bitmap.
' #Const DEBUG_PRINT_BITMAP = True

Dim hbm As Long
Dim bm As BITMAP
Dim l As Single
Dim t As Single
Dim old_color As Long
Dim bytes() As Byte
Dim num_pal_entries As Long
Dim pal_entries(0 To MAX_PALETTE_SIZE - 1) As PALETTEENTRY
Dim pal_index As Integer
Dim wid As Integer
Dim hgt As Integer
Dim x As Integer
Dim Y As Integer
Dim two_bytes As Long

    ' Get the bitmap information.
    hbm = pic.Image
    GetObject hbm, Len(bm), bm
    bits_per_pixel = bm.bmBitsPixel

    ' If bits_per_pixel is 16, see if it's really
    ' 15 or 16 bits per pixel.
    If bits_per_pixel = 16 Then
        ' Make the upper left pixel white.
        l = pic.ScaleLeft
        t = pic.ScaleTop
        old_color = pic.Point(l, t)
        pic.PSet (l, t), vbWhite

        ' See what color was set.
        ReDim bytes(0 To 0, 0 To 0)
        GetBitmapBits hbm, 2, bytes(0, 0)
        If (bytes(0, 0) And &H80) = 0 Then
            ' It's really a 15-bit image.
            bits_per_pixel = 15
        End If

        ' Restore the pixel's original color.
        pic.PSet (l, t), old_color
    End If

    #If DEBUG_PRINT_BITMAP Then
        Debug.Print "*** BITMAP Data ***"
        Debug.Print "bmType       "; bm.bmType
        Debug.Print "bmWidth      "; bm.bmWidth
        Debug.Print "bmHeight     "; bm.bmHeight
        Debug.Print "bmWidthBytes "; bm.bmWidthBytes
        Debug.Print "bmPlanes     "; bm.bmPlanes
        Debug.Print "bmBitsPixel  "; bm.bmBitsPixel
        Debug.Print "BitsPerPixel "; bits_per_pixel
    #End If

    ' Get the bits.
    If (bits_per_pixel = 8) Or _
       (bits_per_pixel = 15) Or _
       (bits_per_pixel = 16) Or _
       (bits_per_pixel = 24) Or _
       (bits_per_pixel = 32) _
    Then
        ' Get the bits.
        ReDim bytes(0 To bm.bmWidthBytes - 1, 0 To bm.bmHeight - 1)
        GetBitmapBits hbm, bm.bmWidthBytes * bm.bmHeight, bytes(0, 0)
    Else
        ' We don't know how to read this format.
        Err.Raise bmphInvalidBitmapBits, _
            "DDBHelper.GetBitmapPixels", _
            "Invalid number of bits per pixel: " _
            & Format$(bits_per_pixel)
    End If

    ' Create the pixels array.
    wid = bm.bmWidth
    hgt = bm.bmHeight
    ReDim pixels(0 To wid - 1, 0 To hgt - 1)
    Select Case bits_per_pixel
        Case 8
            ' Match pic's palette to the system palette.
            MatchColorPalette pic

            ' Get the image's palette entries.
            num_pal_entries = GetPaletteEntries( _
                pic.Picture.hpal, 0, _
                MAX_PALETTE_SIZE, pal_entries(0))

            ' Get the RGB color components.
            For Y = 0 To hgt - 1
                For x = 0 To wid - 1
                    With pixels(x, Y)
                        pal_index = bytes(x, Y)
                        .rgbRed = pal_entries(pal_index).peRed
                        .rgbGreen = pal_entries(pal_index).peGreen
                        .rgbBlue = pal_entries(pal_index).peBlue
                    End With
                Next x
            Next Y

        Case 15
            For Y = 0 To hgt - 1
                For x = 0 To wid - 1
                    With pixels(x, Y)
                        ' Get the combined 2 bytes for this pixel.
                        two_bytes = bytes(x * 2, Y) + bytes(x * 2 + 1, Y) * 256&

                        ' Separate the pixel's components.
                        .rgbBlue = two_bytes Mod 32
                        two_bytes = two_bytes \ 32
                        .rgbGreen = two_bytes Mod 32
                        two_bytes = two_bytes \ 32
                        .rgbRed = two_bytes
                    End With
                Next x
            Next Y

        Case 16
            For Y = 0 To hgt - 1
                For x = 0 To wid - 1
                    With pixels(x, Y)
                        ' Get the combined 2 bytes for this pixel.
                        two_bytes = bytes(x * 2, Y) + bytes(x * 2 + 1, Y) * 256&

                        ' Separate the pixel's components.
                        .rgbBlue = two_bytes Mod 32
                        two_bytes = two_bytes \ 32
                        .rgbGreen = two_bytes Mod 64
                        two_bytes = two_bytes \ 64
                        .rgbRed = two_bytes
                    End With
                Next x
            Next Y

        Case 24
            ' Blast the data from the pixels array
            ' to the bytes array using CopyMemory.
            For Y = 0 To hgt - 1
                CopyMemory pixels(0, Y), bytes(0, Y), wid * 3
            Next Y

        Case 32
            For Y = 0 To hgt - 1
                For x = 0 To wid - 1
                    With pixels(x, Y)
                        .rgbBlue = bytes(x * 4, Y)
                        .rgbGreen = bytes(x * 4 + 1, Y)
                        .rgbRed = bytes(x * 4 + 2, Y)
                    End With
                Next x
            Next Y

    End Select
End Sub
' Set the bits in this PictureBox using a 0-based
' two-dimensional array of RGBTriplets. The pixels must
' have the right dimensions to match the picture.
Public Sub SetBitmapPixels(ByVal pic As PictureBox, ByVal bits_per_pixel As Integer, pixels() As RGBTriplet)
Dim wid_bytes As Long
Dim wid As Integer
Dim hgt As Integer
Dim x As Integer
Dim Y As Integer
Dim bytes() As Byte
Dim hpal As Long
Dim two_bytes As Long

    ' See how big the image must be.
    wid = UBound(pixels, 1) + 1
    hgt = UBound(pixels, 2) + 1

    ' See how many bytes per row we need.
    Select Case bits_per_pixel
        Case 8
            wid_bytes = wid
        Case 15, 16
            wid_bytes = wid * 2
        Case 24
            wid_bytes = wid * 3
        Case 32
            wid_bytes = wid * 4
        Case Else
            ' We don't understand this format.
            Err.Raise bmphInvalidBitmapBits, _
                "DDBHelper.GetBitmapPixels", _
                "Invalid number of bits per pixel: " _
                & Format$(bits_per_pixel)
    End Select

    ' Make sure it's even.
    If wid_bytes Mod 2 = 1 Then wid_bytes = wid_bytes + 1

    ' Create the bitmap bytes array.
    ReDim bytes(0 To wid_bytes - 1, 0 To hgt - 1)

    ' Set the bitmap byte values.
    Select Case bits_per_pixel
        Case 8
            ' Use the nearest palette entries.
            hpal = pic.Picture.hpal

            ' Get the RGB color components.
            For Y = 0 To hgt - 1
                For x = 0 To wid - 1
                    With pixels(x, Y)
                        bytes(x, Y) = (&HFF And _
                            GetNearestPaletteIndex(hpal, _
                                RGB(.rgbRed, .rgbGreen, .rgbBlue) _
                            + &H2000000))
                    End With
                Next x
            Next Y

        Case 15
            For Y = 0 To hgt - 1
                For x = 0 To wid - 1
                    With pixels(x, Y)
                        ' Keep the values in bounds.
                        If .rgbRed > &H1F Then .rgbRed = &H1F
                        If .rgbGreen > &H1F Then .rgbGreen = &H1F
                        If .rgbBlue > &H1F Then .rgbBlue = &H1F

                        ' Combine the values in 2 bytes.
                        two_bytes = .rgbBlue + 32 * (.rgbGreen + CLng(.rgbRed) * 32)

                        ' Set the byte values.
                        bytes(x * 2, Y) = (two_bytes Mod 256) And &HFF
                        bytes(x * 2 + 1, Y) = (two_bytes \ 256) And &HFF
                    End With
                Next x
            Next Y

        Case 16
            For Y = 0 To hgt - 1
                For x = 0 To wid - 1
                    With pixels(x, Y)
                        ' Keep the values in bounds.
                        If .rgbRed > &H1F Then .rgbRed = &H1F
                        If .rgbGreen > &H3F Then .rgbGreen = &H3F
                        If .rgbBlue > &H1F Then .rgbBlue = &H1F

                        ' Combine the values in 2 bytes.
                        two_bytes = .rgbBlue + 32 * (.rgbGreen + CLng(.rgbRed) * 64)

                        ' Set the byte values.
                        bytes(x * 2, Y) = (two_bytes Mod 256) And &HFF
                        bytes(x * 2 + 1, Y) = (two_bytes \ 256) And &HFF

                    End With
                Next x
            Next Y

        Case 24
            ' Blast the data from the bytes array
            ' to the pixels array using CopyMemory.
            For Y = 0 To hgt - 1
                CopyMemory bytes(0, Y), pixels(0, Y), wid * 3
            Next Y

        Case 32
            For Y = 0 To hgt - 1
                For x = 0 To wid - 1
                    With pixels(x, Y)
                        bytes(x * 4, Y) = .rgbBlue
                        bytes(x * 4 + 1, Y) = .rgbGreen
                        bytes(x * 4 + 2, Y) = .rgbRed
                    End With
                Next x
            Next Y

    End Select

    ' Set the picture's bitmap bits.
    SetBitmapBits pic.Image, wid_bytes * hgt, _
        bytes(0, 0)
    pic.Refresh
End Sub


Public Sub RotatePicture(fr_pic As PictureBox, to_pic As PictureBox, ByVal angle As Double)
Dim fr_pixels() As RGBTriplet
Dim c0 As RGBTriplet
Dim c1 As RGBTriplet
Dim c2 As RGBTriplet
Dim c3 As RGBTriplet

Dim to_pixels() As RGBTriplet
Dim bits_per_pixel As Integer
Dim fr_wid As Long
Dim fr_hgt As Long
Dim to_wid As Long
Dim to_hgt As Long
Dim x As Integer
Dim Y As Integer
Dim a As Double, r As Double
Dim p1x As Long, p1y As Long
Dim x1 As Integer, y1 As Integer

Const Pi = 3.1415926536
    ' Get the picture's image.
    GetBitmapPixels fr_pic, fr_pixels, bits_per_pixel
    
    ' Get the picture's size.
    fr_wid = UBound(fr_pixels, 1) + 1
    fr_hgt = UBound(fr_pixels, 2) + 1
    If angle = 0 Or angle = 180 Then
        to_wid = fr_wid
        to_hgt = fr_hgt
    Else
        to_wid = fr_hgt
        to_hgt = fr_wid
    End If
    
    ' Size the output picture to fit.
    to_pic.Width = to_pic.Parent.ScaleX(to_wid, vbPixels, to_pic.Parent.ScaleMode) + _
        to_pic.Width - to_pic.ScaleWidth
    to_pic.Height = to_pic.Parent.ScaleY(to_hgt, vbPixels, to_pic.Parent.ScaleMode) + _
        to_pic.Height - to_pic.ScaleHeight
    
    to_pic.Cls
    Dim cx As Long
    Dim cy As Long
    Dim px As Long
    Dim py As Long
    cx = to_wid / 2
    cy = to_hgt / 2
    
    ' Copy the rotated pixels.
    ReDim to_pixels(0 To to_wid - 1, 0 To to_hgt - 1)
    For x = 0 To fr_wid - 1
        For Y = 0 To fr_hgt - 1
            to_pixels(x, Y) = fr_pixels(1, 1)
        Next
    Next
    
    Dim c1x As Integer  ' Center of pic1.
    Dim c1y As Integer  '   "
    Dim c2x As Integer  ' Center of pic2.
    Dim c2y As Integer  '   "
    Dim p2x As Integer  ' Position on pic2.
    Dim p2y As Integer  '   "
    Dim n As Integer    ' Max width or height of pic2.
    
   ' Compute the centers.
    c1x = fr_pic.ScaleWidth / 2
    c1y = fr_pic.ScaleHeight / 2
    c2x = to_pic.ScaleWidth / 2
    c2y = to_pic.ScaleHeight / 2

   ' Compute the image size.
    n = to_pic.ScaleWidth
    If n < to_pic.ScaleHeight Then n = to_pic.ScaleHeight
    n = n / 2 - 1

    For p2x = 0 To fr_wid - 1

        For p2y = 0 To fr_hgt - 1
         ' Compute polar coordinate of p2.
         If p2x = 0 Then
           a = Pi / 2
         Else
           a = Atn(p2y / p2x)
         End If
         r = Sqr(1& * p2x * p2x + 1& * p2y * p2y)

         ' Compute rotated position of p1.
         p1x = r * Cos(a + angle)
         p1y = r * Sin(a + angle)

         ' Copy pixels, 4 quadrants at once.
         On Error Resume Next
         c0 = fr_pixels(c1x + p1x, c1y + p1y)
         c1 = fr_pixels(c1x - p1x, c1y - p1y)
         c2 = fr_pixels(c1x + p1y, c1y - p1x)
         c3 = fr_pixels(c1x - p1y, c1y + p1x)
         to_pixels(c2x + p2x, c2y + p2y) = c0
         to_pixels(c2x - p2x, c2y - p2y) = c1
         to_pixels(c2x + p2y, c2y - p2x) = c2
         to_pixels(c2x - p2y, c2y + p2x) = c3
      Next
    Next


    ' Display the result.
    SetBitmapPixels to_pic, bits_per_pixel, to_pixels

    ' Make the image permanent.
    to_pic.Refresh
    to_pic.Picture = to_pic.Image
End Sub

