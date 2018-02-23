Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Colors

Public Class Colors
    Public Shared Sub ColorGrade(LayerName)
        '' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        '' Lock the new document
        Using acLckDoc As DocumentLock = acDoc.LockDocument()

            '' Start a transaction
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                '' Open the Layer table for read
                Dim acLyrTbl As LayerTable
                acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)

                Dim sLayerName As String = LayerName

                If acLyrTbl.Has(sLayerName) = False Then
                    Dim acLyrTblRec As LayerTableRecord = New LayerTableRecord()

                    '' Assign the layer the ACI color 1 and a name

                    'acLyrTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 1)



                    '' Upgrade the Layer table for write
                    acLyrTbl.UpgradeOpen()

                    '' Append the new layer to the Layer table and the transaction
                    acLyrTbl.Add(acLyrTblRec)
                    acTrans.AddNewlyCreatedDBObject(acLyrTblRec, True)
                End If
            End Using
        End Using
    End Sub

End Class
