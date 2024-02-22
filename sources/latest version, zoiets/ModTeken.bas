Attribute VB_Name = "ModTeken"
Public Sub TekenCirkel(pic As PictureBox, vM As TVECTOR, r1 As Double, r2 As Double, kleur As Long)
pic.FillStyle = 0
pic.FillColor = kleur
pic.Circle (pic.ScaleWidth / 2 - vM.x * r1, pic.ScaleHeight / 2 + vM.Y * r2), 30, kleur
End Sub
Public Sub TekenCirkelKlein(pic As PictureBox, vM As TVECTOR, r1 As Double, r2 As Double, kleur As Long, size As Double)
pic.FillStyle = 0
pic.FillColor = kleur
pic.Circle (pic.ScaleWidth / 2 - vM.x * r1, pic.ScaleHeight / 2 + vM.Y * r2), size, kleur
End Sub
Public Sub TekenCirkelWit(vM As TVECTOR, r1 As Double, r2 As Double, kleur As Long)
frmJupiterMoons.picHidden.FillStyle = 0
frmJupiterMoons.picHidden.FillColor = kleur
If Sqr(vM.x * vM.x + vM.Y * vM.Y) * r1 + 50 >= r1 Then 'de image van jupiter wordt getekend waardoor de witte cirkels in
                                                       'dat gebied niet meer nodig zijn
    frmJupiterMoons.picHidden.Circle (frmJupiterMoons.picHidden.ScaleWidth / 2 - vM.x * r1, frmJupiterMoons.picHidden.ScaleHeight / 2 + vM.Y * r2), 30, kleur
End If
End Sub
Public Sub TekenCirkel2(vM As TVECTOR)
End Sub


