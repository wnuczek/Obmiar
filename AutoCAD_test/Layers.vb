Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Colors

Public Class Layers

    <CommandMethod("CreateAndAssignALayer")>
    Public Shared Sub CreateAndAssignALayer(LayerName)

        '' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        '' Lock the new document
        Using acLckDoc As DocumentLock = acDoc.LockDocument()

            '' Start a transaction
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                '' Open the Layer table for read
                Dim acLyrTbl As LayerTable
                acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId,
                                             OpenMode.ForRead)

                Dim sLayerName As String = LayerName

                If acLyrTbl.Has(sLayerName) = False Then
                    Dim acLyrTblRec As LayerTableRecord = New LayerTableRecord()

                    '' Assign the layer the ACI color 1 and a name
                    acLyrTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, Paleta.ColorCode)
                    Dim lineweightcode As String = "LineWeight." & Paleta.WeightCode
                    acLyrTblRec.LineWeight = LineWeight.LineWeight040
                    acLyrTblRec.Name = sLayerName

                    '' Upgrade the Layer table for write
                    acLyrTbl.UpgradeOpen()

                    '' Append the new layer to the Layer table and the transaction
                    acLyrTbl.Add(acLyrTblRec)
                    acTrans.AddNewlyCreatedDBObject(acLyrTblRec, True)
                End If
                ''Set current layer
                acCurDb.Clayer = acLyrTbl(sLayerName)

                '' Open the Block table for read
                'Dim acBlkTbl As BlockTable
                'acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                'OpenMode.ForRead)

                '' Open the Block table record Model space for write
                'Dim acBlkTblRec As BlockTableRecord
                'acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace),
                'OpenMode.ForWrite)

                '' Create a circle object
                'Dim acCirc As Circle = New Circle()
                'acCirc.Center = New Point3d(2, 2, 0)
                'acCirc.Radius = 1
                'acCirc.Layer = sLayerName

                'acBlkTblRec.AppendEntity(acCirc)
                'acTrans.AddNewlyCreatedDBObject(acCirc, True)

                '' Save the changes and dispose of the transaction
                acTrans.Commit()
            End Using
        End Using
    End Sub
End Class

