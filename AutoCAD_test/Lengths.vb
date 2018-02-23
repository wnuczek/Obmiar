Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Colors
Imports Autodesk.AutoCAD.EditorInput
Imports System.Linq
Imports System.Text

Public Class Lengths

    Public Shared Sub GetLayerLenght()
        '' Get the current document and database, and start a transaction
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            '' This example returns the layer table for the current database
            Dim acLyrTbl As LayerTable
            acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)

            '' Step through the Layer table and print each layer name

            For Each acObjId As ObjectId In acLyrTbl
                Dim acLyrTblRec As LayerTableRecord
                acLyrTblRec = acTrans.GetObject(acObjId, OpenMode.ForRead)
                If acLyrTblRec.Name.Contains(".OBMR") = True Then
                    acDoc.Editor.WriteMessage(vbLf & acLyrTblRec.Name)

                    'SelectEntitiesWithProperties()
                    testTotalLengthVB(acLyrTblRec.Name)
                End If

            Next

            '' Dispose of the transaction
        End Using
    End Sub

    <CommandMethod("LV")>
    Public Shared Sub ListVertices()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        Dim db As Database = doc.Database

        Dim per As PromptEntityResult = ed.GetEntity("Select a polyline")

        If per.Status = PromptStatus.OK Then
            Dim tr As Transaction = db.TransactionManager.StartTransaction()
            Using tr
                Dim obj As DBObject = tr.GetObject(per.ObjectId, OpenMode.ForRead)

                ' If a "lightweight" (or optimized) polyline
                Dim lwp As Polyline = TryCast(obj, Polyline)
                If lwp IsNot Nothing Then
                    ' Use a for loop to get each vertex, one by one
                    Dim vn As Integer = lwp.NumberOfVertices
                    For i As Integer = 0 To vn - 1
                        ' Could also get the 3D point here
                        Dim pt As Point2d = lwp.GetPoint2dAt(i)
                        ed.WriteMessage(vbLf + pt.ToString())
                    Next
                Else
                    ' If an old-style, 2D polyline
                    Dim p2d As Polyline2d = TryCast(obj, Polyline2d)
                    If p2d IsNot Nothing Then
                        ' Use foreach to get each contained vertex
                        For Each vId As ObjectId In p2d
                            Dim v2d As Vertex2d = DirectCast(tr.GetObject(vId, OpenMode.ForRead), Vertex2d)
                            ed.WriteMessage(vbLf + v2d.Position.ToString())
                        Next
                    Else
                        ' If an old-style, 3D polyline
                        Dim p3d As Polyline3d = TryCast(obj, Polyline3d)
                        If p3d IsNot Nothing Then
                            ' Use foreach to get each contained vertex
                            For Each vId As ObjectId In p3d
                                Dim v3d As PolylineVertex3d = DirectCast(tr.GetObject(vId, OpenMode.ForRead), PolylineVertex3d)
                                ed.WriteMessage(vbLf + v3d.Position.ToString())
                            Next
                        End If
                    End If
                End If
                ' Committing is cheaper than aborting
                tr.Commit()
            End Using
        End If
    End Sub

    <CommandMethod("SEWP")>
    Public Shared Function SelectEntitiesWithProperties(layername)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor

        Dim tvs As TypedValue() = New TypedValue() {New TypedValue(CInt(DxfCode.Start), "*POLYLINE"),
        New TypedValue(CInt(DxfCode.LayerName), layername)}
        Dim sf As New SelectionFilter(tvs)
        Dim psr As PromptSelectionResult = ed.SelectAll(sf)



        ed.WriteMessage(vbLf & "Found {0} entit{1}.", psr.Value.Count, (If(psr.Value.Count = 1, "y", "ies")))

        Return ed.SelectAll(sf)
    End Function

    <CommandMethod("TotalLengthVB")>
    Public Sub testTotalLengthVB()
        Dim db As Database = HostApplicationServices.WorkingDatabase
        Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
        Dim sb As New StringBuilder
        sb.AppendLine(vbTab + "Total Length:" + vbTab + "------------------------------------")
        Using tr As Transaction = db.TransactionManager.StartTransaction
            Dim tv() As TypedValue = {New TypedValue(0, "LINE,ARC,SPLINE,CIRCLE,ELLIPSE,*POLYLINE")}
            Dim sf As New SelectionFilter(tv)
            Dim psr As PromptSelectionResult = ed.GetSelection(sf)
            If psr.Status <> PromptStatus.OK Then Return

            ''  See here for more: http://msdn.microsoft.com/en-us/vstudio/bb688088.aspx        101 LINQ Samples VB.NET
            ''  Grouping objects by name with sum of lengths
            Dim subtotals = From id In psr.Value.GetObjectIds
                            Let curve = CType(tr.GetObject(id, OpenMode.ForRead), Curve)
                            Group curve By curve.GetRXClass().DxfName Into grp = Group
                            Select ObjectName = grp.First, Subtotal = grp.Sum(Function(x) Math.Abs(x.GetDistanceAtParameter(x.EndParam - x.StartParam)))

            For Each n In subtotals
                sb.AppendLine(String.Format("{0}" + vbTab + "{1:f6}", n.ObjectName, n.Subtotal))
            Next

            System.Windows.Forms.MessageBox.Show(sb.ToString())
        End Using
    End Sub

End Class

