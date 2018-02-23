
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.LayerManager
Imports System.Collections.Generic

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

        <CommandMethod("CLG")>
        Public Shared Sub CreateLayerGroup()
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor

            ' A list of the layers' names & IDs contained
            ' in the current database, sorted by layer name

            Dim ld As New SortedList(Of String, ObjectId)()

            ' A list of the selected layers' IDs

            Dim lids As New ObjectIdCollection()

            ' Start by populating the list of names/IDs
            ' from the LayerTable

            Dim tr As Transaction = db.TransactionManager.StartTransaction()
            Using tr
                Dim lt As LayerTable = DirectCast(tr.GetObject(db.LayerTableId, OpenMode.ForRead), LayerTable)
                For Each lid As ObjectId In lt
                    Dim ltr As LayerTableRecord = DirectCast(tr.GetObject(lid, OpenMode.ForRead), LayerTableRecord)
                    ld.Add(ltr.Name, lid)
                Next
            End Using

            ' Display a numbered list of the available layers

            ed.WriteMessage(vbLf & "Layers available for group:")

            Dim i As Integer = 1
            For Each kv As KeyValuePair(Of String, ObjectId) In ld
                ed.WriteMessage(vbLf & "{0} - {1}", System.Math.Max(System.Threading.Interlocked.Increment(i), i - 1), kv.Key)
            Next

            ' We will ask the user to select from the list

            Dim pio As New PromptIntegerOptions(vbLf & "Enter number of layer to add: ")
            pio.LowerLimit = 1
            pio.UpperLimit = ld.Count
            pio.AllowNone = True

            ' And will do so in a loop, waiting for
            ' Escape or Enter to terminate

            Dim pir As PromptIntegerResult
            Do
                ' Select one from the list

                pir = ed.GetInteger(pio)

                If pir.Status = PromptStatus.OK Then
                    ' Get the layer's name

                    Dim ln As String = ld.Keys(pir.Value - 1)

                    ' And then its ID

                    Dim lid As ObjectId
                    ld.TryGetValue(ln, lid)

                    ' Add the layer'd ID to the list, is it's not
                    ' already on it

                    If lids.Contains(lid) Then
                        ed.WriteMessage(vbLf & "Layer ""{0}"" has already been selected.", ln)
                    Else
                        lids.Add(lid)
                        ed.WriteMessage(vbLf & "Added ""{0}"" to selected layers.", ln)
                    End If
                End If
            Loop While pir.Status = PromptStatus.OK

            ' Now we've selected our layers, let's create the group

            Try
                If lids.Count > 0 Then
                    ' Get the existing layer filters
                    ' (we will add to them and set them back)

                    Dim lft As LayerFilterTree = db.LayerFilters
                    Dim lfc As LayerFilterCollection = lft.Root.NestedFilters

                    ' Create a new layer group

                    Dim lg As New LayerGroup()
                    lg.Name = "My Layer Group"

                    ' Add our layers' IDs to the list

                    For Each id As ObjectId In lids
                        lg.LayerIds.Add(id)
                    Next

                    ' Add the group to the collection

                    lfc.Add(lg)

                    ' Set them back on the Database

                    db.LayerFilters = lft

                    ed.WriteMessage(vbLf & """{0}"" group created containing {1} layers." & vbLf, lg.Name, lids.Count)

                    ' List the layer filters, to see the new group

                    ListLayerFilters()
                End If
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
