

Dim hsf As HybridShapeFactory
Dim part As part



Sub CATMain()

Set part = CATIA.ActiveDocument.part
Set hsf = part.HybridShapeFactory
Set sel = CATIA.ActiveDocument.Selection

For i = 1 To sel.Count
    Dim crv As HybridShape
    Set crv = sel.Item(i).Value
    Call ShapeExtrapol(crv)
Next
End Sub



Sub ShapeExtrapol(ByVal crv As HybridShape)

Dim p0 As HybridShape
Set p0 = hsf.AddNewPointOnCurveFromPercent(crv, 1, False)
p0.Compute

Dim p1 As HybridShape
Set p1 = hsf.AddNewPointOnCurveFromPercent(crv, 0, False)
p1.Compute

Dim crv1 As HybridShapeExtrapol
Set crv1 = hsf.AddNewExtrapolLength(p0, crv, 1000)
crv1.SetAssemble True


Dim crv2 As HybridShapeExtrapol
Set crv2 = hsf.AddNewExtrapolLength(p1, crv1, 1000)
crv2.Compute
crv2.SetAssemble True


part.MainBody.InsertHybridShape crv2



End Sub
