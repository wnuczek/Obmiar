Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput

Public Class LineTypeCreate

    <CommandMethod("CCL")>
    Public Shared Sub CreateComplexLinetype(name, desc, text)
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurdb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor

        Dim tr As Transaction = acCurdb.TransactionManager.StartTransaction()

        'We'll use the textstyle table to access the "Standard" textstyle for our textsegment

        Dim tt As TextStyleTable = tr.GetObject(acCurdb.TextStyleTableId, OpenMode.ForRead)

        'Get the linetype table from the drawing

        Dim lt As LinetypeTable = tr.GetObject(acCurdb.LinetypeTableId, OpenMode.ForWrite)

        'Create our new linetype table record...

        Dim ltr As LinetypeTableRecord = New LinetypeTableRecord()

        '... and set its properties

        ltr.Name = name
        ltr.AsciiDescription = desc
        ltr.PatternLength = 0.9
        ltr.NumDashes = 3

        'Dash #1

        ltr.SetDashLengthAt(0, 0.5)

        'Dash #2

        ltr.SetDashLengthAt(1, -0.2)
        ltr.SetShapeStyleAt(1, tt("Standard"))
        ltr.SetShapeNumberAt(1, 0)
        ltr.SetShapeOffsetAt(1, New Vector2d(-0.1, -0.05))
        ltr.SetShapeScaleAt(1, 0.1)
        ltr.SetShapeIsUcsOrientedAt(1, False)
        ltr.SetShapeRotationAt(1, 0)
        ltr.SetTextAt(1, text)

        'Dash #3

        ltr.SetDashLengthAt(2, -0.2)

        'Add the new linetype to the linetype table

        Dim ltId As ObjectId = lt.Add(ltr)
        tr.AddNewlyCreatedDBObject(ltr, True)

        'Create a test line with this linetype

        Dim bt As BlockTable = tr.GetObject(acCurdb.BlockTableId, OpenMode.ForRead)
        Dim btr As BlockTableRecord = tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

        Dim ln As Line = New Line(New Point3d(0, 0, 0), New Point3d(10, 10, 0))

        ln.SetDatabaseDefaults(acCurdb)
        ln.LinetypeId = ltId

        btr.AppendEntity(ln)
        tr.AddNewlyCreatedDBObject(ln, True)

        tr.Commit()
    End Sub
End Class
