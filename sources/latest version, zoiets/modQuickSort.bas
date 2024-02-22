Attribute VB_Name = "modQuickSort"
Public Sub QuickSort(iArray As Variant, l, r)

Dim i, j
Dim x
Dim y

i = l
j = r
x = iArray((l + r) / 2)

While (i <= j)

   While (iArray(i) < x And i < r)
      i = i + 1
   Wend

   While (x < iArray(j) And j > l)
      j = j - 1
   Wend

   If (i <= j) Then
      y = iArray(i)
      iArray(i) = iArray(j)
      iArray(j) = y
      i = i + 1
      j = j - 1
   End If

Wend

If (l < j) Then QuickSort iArray, l, j

If (i < r) Then QuickSort iArray, i, r

End Sub




