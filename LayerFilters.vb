
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.LayerManager

Namespace LayerFilters
    Public Class Commands
        <CommandMethod("LLFS")>
        Public Shared Sub ListLayerFilters()
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor

            ' List the nested layer filters

            Dim lfc As LayerFilterCollection = db.LayerFilters.Root.NestedFilters

            For i As Integer = 0 To lfc.Count - 1
                Dim lf As LayerFilter = lfc(i)
                ed.WriteMessage(vbLf & "{0} - {1} (can{2} be deleted)", i + 1, lf.Name, (If(lf.AllowDelete, "", "not")))
            Next
        End Sub

        <CommandMethod("CLFS")>
        Public Shared Sub CreateLayerFilters()
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor

            Try
                ' Get the existing layer filters
                ' (we will add to them and set them back)

                Dim lft As LayerFilterTree = db.LayerFilters
                Dim lfc As LayerFilterCollection = lft.Root.NestedFilters

                ' Create three new layer filters

                Dim lf1 As New LayerFilter()
                lf1.Name = "Unlocked Layers"
                lf1.FilterExpression = "LOCKED==""False"""

                Dim lf2 As New LayerFilter()
                lf2.Name = "White Layers"
                lf2.FilterExpression = "COLOR==""7"""

                Dim lf3 As New LayerFilter()
                lf3.Name = "Visible Layers"
                lf3.FilterExpression = "OFF==""False"" AND FROZEN==""False"""

                ' Add them to the collection

                lfc.Add(lf1)
                lfc.Add(lf2)
                lfc.Add(lf3)

                ' Set them back on the Database

                db.LayerFilters = lft

                ' List the layer filters, to see the new ones

                ListLayerFilters()
            Catch ex As Exception
                ed.WriteMessage(vbLf & "Exception: {0}", ex.Message)
            End Try
        End Sub

        <CommandMethod("DLF")>
        Public Shared Sub DeleteLayerFilter()
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor

            ListLayerFilters()

            Try
                ' Get the existing layer filters
                ' (we will add to them and set them back)

                Dim lft As LayerFilterTree = db.LayerFilters
                Dim lfc As LayerFilterCollection = lft.Root.NestedFilters

                ' Prompt for the index of the filter to delete

                Dim pio As New PromptIntegerOptions(vbLf & vbLf & "Enter index of filter to delete")
                pio.LowerLimit = 1
                pio.UpperLimit = lfc.Count

                Dim pir As PromptIntegerResult = ed.GetInteger(pio)

                ' Get the selected filter

                Dim lf As LayerFilter = lfc(pir.Value - 1)

                ' If it's possible to delete it, do so

                If Not lf.AllowDelete Then
                    ed.WriteMessage(vbLf & "Layer filter cannot be deleted.")
                Else
                    lfc.Remove(lf)
                    db.LayerFilters = lft

                    ListLayerFilters()
                End If
            Catch ex As Exception
                ed.WriteMessage(vbLf & "Exception: {0}", ex.Message)
            End Try
        End Sub
    End Class
End Namespace

'=======================================================
'Service provided by Telerik (www.telerik.com)
'Conversion powered by NRefactory.
'Twitter: @telerik
'Facebook: facebook.com/telerik
'=======================================================
