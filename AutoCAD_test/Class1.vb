Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports MgdAcApplication = Autodesk.AutoCAD.ApplicationServices.Application
Imports Autodesk.AutoCAD.DatabaseServices
Imports ObmiarPW.Layers
Imports ObmiarPW.Paleta
Imports ObmiarPW.LineTypeCreate
Imports Autodesk.AutoCAD.Windows
Imports System.Drawing

Public Class Class1
    'auto-enable our toolpalette for AutoCAD
    Implements Autodesk.AutoCAD.Runtime.IExtensionApplication

    Public Shared ReadOnly Property ThisDrawing() As AcadDocument
        Get
            Return DocumentExtension.GetAcadDocument(MgdAcApplication.DocumentManager.MdiActiveDocument)
        End Get
    End Property

    Public Shared LayerName As String

    Public Shared Sub AddLayer(LayerName, desc, text)
        CreateAndAssignALayer(LayerName)
        CreateComplexLinetype(LayerName, desc, text)
        'colorgrade 

    End Sub

    <CommandMethod("AddDocColEvent")>
    Public Sub AddDocColEvent()
        AddHandler Application.DocumentManager.DocumentActivated,
      AddressOf docColDocAct
    End Sub

    <CommandMethod("RemoveDocColEvent")>
    Public Sub RemoveDocColEvent()
        RemoveHandler Application.DocumentManager.DocumentActivated,
      AddressOf docColDocAct
    End Sub

    Public Sub docColDocAct(ByVal senderObj As Object,
                        ByVal docColDocActEvtArgs As DocumentCollectionEventArgs)
        AddHandler Application.DocumentManager.DocumentCreated,
      AddressOf docColDocAct
        Application.ShowAlertDialog("Sprawdz skale rysunku" & docColDocActEvtArgs.Document.Name)
    End Sub




    <CommandMethod("SetLayerCurrent")>
    Public Shared Sub SetLayerCurrent(LayerName)
        '' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            '' Open the Layer table for read
            Dim acLyrTbl As LayerTable
            acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)

            Dim sLayerName As String = LayerName

            If acLyrTbl.Has(sLayerName) = True Then
                '' Set the layer Center current
                acCurDb.Clayer = acLyrTbl(sLayerName)

                '' Save the changes
                acTrans.Commit()

            End If

            '' Dispose of the transaction
        End Using
    End Sub

    Private Shared Function GetEmbeddedIcon(ByVal sName As String) As Icon
        'pulls embedded resource
        Return New Icon(System.Reflection.Assembly.GetExecutingAssembly.GetManifestResourceStream(sName))
    End Function


    'ensure single instance of this app... 
    Friend Shared m_ps As Autodesk.AutoCAD.Windows.PaletteSet = Nothing

    ' Define command 
    <CommandMethod("Obmiar")>
    Public Sub DoIt()

        'check to see if paletteset is already created 
        If m_ps Is Nothing Then
            'no so create it 
            m_ps = New Autodesk.AutoCAD.Windows.PaletteSet("Narzędzia obmiarowe (© 2016-2019 Paweł Wnuk)", New Guid("{CCBFEC73-9FE4-4aa2-8E4B-3068E94A2BFC}"))
            'create new instance of user control 
            Dim myPalette As Paleta = New Paleta()
            'add it to the paletteset 
            m_ps.Add("Narzędzia obmiarowe (© 2016-2019 Paweł Wnuk)", myPalette)
        End If
        'turn it on 

        m_ps.Style = PaletteSetStyles.ShowPropertiesMenu Or PaletteSetStyles.ShowAutoHideButton Or PaletteSetStyles.ShowCloseButton

        'm_ps.Icon = GetEmbeddedIcon("favicon.ico")

        m_ps.Size = New Size(210, 390)

        m_ps.Dock = DockSides.Right
        m_ps.Visible = True
    End Sub

    Private Shared Sub Paleta_Load(ByVal sender As Object, ByVal e As Autodesk.AutoCAD.Windows.PalettePersistEventArgs)
        'demo loading user data
        Dim a As Double =
    CType(e.ConfigurationSection.ReadProperty("Narzędzia obmiarowe (© 2016-2019 Paweł Wnuk)", 22.3), Double)
    End Sub

    Private Shared Sub ps_Save(ByVal sender As Object, ByVal e As Autodesk.AutoCAD.Windows.PalettePersistEventArgs)
        'demo saving user data 
        e.ConfigurationSection.WriteProperty("Narzędzia obmiarowe (© 2016-2019 Paweł Wnuk)", 32.3)
    End Sub

    Public Sub Initialize() Implements IExtensionApplication.Initialize
        'add anything that needs to be instantiated on startup
    End Sub
    Public Sub Terminate() Implements IExtensionApplication.Terminate
        'handle closing down a link to a database/etc.
    End Sub

End Class
