Imports Autodesk.AutoCAD.Windows
Imports ObmiarPW.Layers
Imports ObmiarPW.Class1
Imports ObmiarPW.LineTypeCreate
Imports ObmiarPW.Lengths


Public Class Paleta
    'DEKLARACJA ZMIENNYCH/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Public Shared inst As String
    Public Shared mat As String
    Public Shared atr(3) As String
    Public Shared atrs As String
    Public Shared sred As String
    Public Shared LayerName As String
    Public Shared LayerNameVent As String

    Public Shared sysname As String
    Public Shared atrVent(2) As String
    Public Shared wymA As String
    Public Shared wymB As String
    Public Shared sredVent As String

    Public Shared actualListboxWymA As Integer
    Public Shared actualListboxWymB As Integer


    'TYPOWE WYMIARY KANAŁÓW////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Public Shared TABwymA() As String = {"200", "250", "300", "315", "350", "400", "500", "600", "700", "800", "1000", "1200", "1500", "2000", "2500"}
    Public Shared TABwymB() As String = {"200", "250", "300", "315", "350", "400", "500", "600", "700", "800", "1000", "1200", "1500", "2000", "2500"}
    Public Shared TABsr() As String = {"80", "100", "125", "160", "200", "250", "315", "350", "400"}

    'TYPOWE WYMIARY RUR////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Public Shared TABsrDN() As String = {"10", "15", "20", "25", "32", "40", "50", "65", "80", "100", "125", "150", "200", "250", "300", "350", "400", "450", "500"}
    Public Shared TABsrFi() As String = {"12", "16", "20", "25", "32", "40", "50", "63", "75", "90", "110", "125", "160", "200", "250"}

    'DODANIE ELEMENTU DO TABLICY////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Public Function ArrayAdd(arr() As String, item As String)
        Dim duplicate As Integer = 0
        For Each arrayitem As String In arr
            If arrayitem = item Then
                duplicate = 1
            End If
        Next

        If duplicate = 1 Then
            'MsgBox("item already exists: " & item)
            Return arr
        Else
            'MsgBox("before adding: " & Join(arr, vbCrLf))
            Array.Resize(arr, arr.Length + 1)
            'MsgBox("adding: " & item)
            arr(arr.Length - 1) = item
            'MsgBox("after adding: " & Join(arr, vbCrLf))
            Dim stringArray = arr
            Dim intArray = Array.ConvertAll(stringArray, Function(str) Int32.Parse(str))
            Array.Sort(intArray)
            stringArray = Array.ConvertAll(Of Integer, String)(intArray, Function(x) x.ToString())
            'MsgBox(intArray(2))
            'MsgBox("after sorting: " & Join(stringArray, vbCrLf))
            arr = stringArray
            Return arr
        End If
    End Function

    'LISTA INSTALACJI////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Dim separator1() As Char = {vbCrLf}
    Dim separator2() As Char = {","}
    'Dim plikWoda As String = My.Computer.FileSystem.ReadAllText("C:\Users\pawwnu\Dropbox\ObmiarAutoCAD\InstWoda.txt")
    'Dim plikKan As String = My.Computer.FileSystem.ReadAllText("C:\Users\pawwnu\Dropbox\ObmiarAutoCAD\InstKan.txt")
    'Dim plikHC As String = My.Computer.FileSystem.ReadAllText("C:\Users\pawwnu\Dropbox\ObmiarAutoCAD\InstHC.txt")
    Dim instWoda() As String = {"Woda zimna", "Woda ciepła", "Cyrkulacja", "Woda szara", "Woda uzdatniona", "Podlewanie zieleni", "Hydranty", "Tryskacze"}
    Dim instKan() As String = {"Kan. sanitarna", "Kan. technologiczna", "Kan. deszczowa", "Kan. deszczowa podciśn.", "Kan. tłuszczowa", "Kan. ciśnieniowa", "Kan. parkingu/garażu"}
    Dim instHC() As String = {"CO", "CT", "CT AHU", "WL", "WL AHU"}

    Private Sub Paleta_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MsgBox("Narzędzie obmiarowe - © 2016-2018 Paweł Wnuk" & vbNewLine & "(pawelwnuk.pl, mail@pawelwnuk.pl)")
    End Sub

    'WYPEŁNIANIE LISTY Z PLIKU TXT///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Public Sub ListBoxInstFill(plik)
        ListBoxSred.Items.Clear()
        Dim TabInst() As String = plik
        'Dim TabInst() As String = plik.Split(separator1, StringSplitOptions.RemoveEmptyEntries)
        Dim TabInstCount As Integer = TabInst.Count
        For i = 0 To TabInstCount - 1
            TabInst(i) = TabInst(i).Trim
            ListBoxSred.Items.Add(TabInst(i))
        Next
    End Sub

    'WYPEŁNIANIE LISTY Z TABLICY///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Public Sub ListBoxFill(listbox, table)
        listbox.Items.Clear()
        Dim Tab() As String = table
        Dim TabCount As Integer = Tab.Count
        For i = 0 To TabCount - 1
            listbox.Items.Add(Tab(i))
        Next
    End Sub

    'TWORZENIE NAZW WARSTW///////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Public Sub CreateLayerName(inst, mat, atrs, sred)
        LayerName = ".OBMR_" & inst & "_" & sred & "_" & mat & "_" & atrs
    End Sub

    'INSTALACJA//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Private Sub ListBoxSred_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBoxSred.SelectedIndexChanged
        sred = ListBoxSred.SelectedItem
    End Sub


    Private Sub txtInst_TextChanged(sender As Object, e As EventArgs) Handles rInst.TextChanged
        inst = rInst.Text
    End Sub

    Private Sub rInst_Click(sender As Object, e As EventArgs) Handles rInst.Click
        If rInst.Text = "Nazwa instalacji" Then
            rInst.Clear()
        End If
    End Sub

    'MATERIAŁY/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Private Sub ChangeDiameter()
        If RBtnStal.Checked = True Or RBtnStalNierdz.Checked = True Or RBtnZeliwo.Checked = True Then
            rbtnDN.Checked = True
        Else
            rbtnFi.Checked = True
        End If
    End Sub

    Private Sub RBtnStal_CheckedChanged(sender As Object, e As EventArgs) Handles RBtnStal.CheckedChanged
        mat = "stal"
        ChangeDiameter()
    End Sub

    Private Sub RBtnPP_CheckedChanged(sender As Object, e As EventArgs) Handles RBtnPP.CheckedChanged
        mat = "PP"
        ChangeDiameter()
    End Sub

    Private Sub RBtnStalNierdz_CheckedChanged(sender As Object, e As EventArgs) Handles RBtnStalNierdz.CheckedChanged
        mat = "stal.nierdz"
        ChangeDiameter()
    End Sub

    Private Sub RBtnZeliwo_CheckedChanged(sender As Object, e As EventArgs) Handles RBtnZeliwo.CheckedChanged
        mat = "zeliwo"
        ChangeDiameter()
    End Sub

    Private Sub RBtnPVC_CheckedChanged(sender As Object, e As EventArgs) Handles RBtnPVC.CheckedChanged
        mat = "PVC"
        ChangeDiameter()
    End Sub

    Private Sub RBtnHDPE_CheckedChanged(sender As Object, e As EventArgs) Handles RBtnHDPE.CheckedChanged
        mat = "HDPE"
        ChangeDiameter()
    End Sub

    'ATRYBUTY////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Private Sub ChkBoxKabel_CheckedChanged(sender As Object, e As EventArgs) Handles ChkBoxKabel.CheckedChanged
        If ChkBoxKabel.Checked = True Then
            atr(0) = "Kabel.grzejny"
        Else
            atr(0) = ""
        End If
    End Sub

    Private Sub ChkBoxPlaszcz_CheckedChanged(sender As Object, e As EventArgs) Handles ChkBoxPlaszcz.CheckedChanged
        If ChkBoxPlaszcz.Checked = True Then
            atr(1) = "Płaszcz"
        Else
            atr(1) = ""
        End If
    End Sub

    Private Sub ChkBoxZewn_CheckedChanged(sender As Object, e As EventArgs) Handles ChkBoxZewn.CheckedChanged
        If ChkBoxZewn.Checked = True Then
            atr(2) = "Zewn"
        Else
            atr(2) = ""
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles ChkBoxInne.CheckedChanged
        If ChkBoxInne.Checked = True Then
            TxtBoxInne.Enabled = True
        Else
            TxtBoxInne.Enabled = False
            TxtBoxInne.Text = "Inne"
        End If
    End Sub



    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TxtBoxInne.TextChanged
        atr(3) = txtWymA.Text
    End Sub

    Private Sub txtboxInne_Click(sender As Object, e As EventArgs) Handles TxtBoxInne.Click
        If TxtBoxInne.Text = "Inne" Then
            TxtBoxInne.Clear()
        End If
    End Sub

    'ŚREDNICA////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


    'POLECENIE RYSUJ//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDraw.Click

        'KARTA RURY////////////////////////////////////////
        If TabControl1.SelectedTab Is TabPage1 Then

            If rInst.Text = "Nazwa instalacji" Or rInst.Text = "" Then
                MsgBox("Nie podano nazwy instalacji")
            End If

            If txtInnaSred.Text = "Inna średnica" Then

            Else

                If rbtnDN.Checked = True Then

                    ListBoxFill(ListBoxSred, ArrayAdd(TABsrDN, txtInnaSred.Text))
                    ' ListBoxWymA.SelectedIndex = actualListboxWymA
                    TABsrDN = ArrayAdd(TABsrDN, txtInnaSred.Text)

                ElseIf rbtnFi.Checked = True Then

                    ListBoxFill(ListBoxSred, ArrayAdd(TABsrFi, txtInnaSred.Text))
                    ' ListBoxWymA.SelectedIndex = actualListboxWymA
                    TABsrFi = ArrayAdd(TABsrFi, txtInnaSred.Text)

                End If

            End If


            Dim atrs As String = Join(atr, "_")
            CreateLayerName(inst, mat, atrs, sred)
            'MsgBox(LayerName)
            CreateAndAssignALayer(LayerName)

            ' CreateComplexLinetype(LayerName, "opis", "TXT")
            ' CreateLayerName(inst, mat, atrs, sred)
            'AddLayer(LayerName, "opis", "SKR_DN")
            'CreateAndAssignALayer(inst & "_" & mat & "_" & atrx & "_" & sred)
            'ThisDrawing.SendCommand("pl" & vbCr)

            'KARTA KANAŁY//////////////////////////////////////
        ElseIf TabControl1.SelectedTab Is TabPage2 Then
            'actualListboxWymA = ListBoxWymA.SelectedIndex
            'actualListboxWymB = ListBoxWymB.SelectedIndex

            If txtWymA.Text = "Wymiar A [mm]" Then
                If rbtnSpiro.Checked = True Or rbtnFlex.Checked = True Or rbtnAkustik.Checked = True Then
                    sredVent = ListBoxWymA.Text
                Else
                    wymA = ListBoxWymA.Text
                    sredVent = ""
                End If

            Else

                If rbtnSpiro.Checked = True Or rbtnFlex.Checked = True Or rbtnAkustik.Checked = True Then
                    sredVent = txtWymA.Text
                Else
                    wymA = txtWymA.Text
                    sredVent = ""
                End If


                If rbtnProst.Checked = True Or rbtnPD.Checked = True Then

                    ListBoxFill(ListBoxWymA, ArrayAdd(TABwymA, txtWymA.Text))
                    ' ListBoxWymA.SelectedIndex = actualListboxWymA
                    TABwymA = ArrayAdd(TABwymA, txtWymA.Text)
                Else

                    ListBoxFill(ListBoxWymA, ArrayAdd(TABsr, txtWymA.Text))
                    ' ListBoxWymA.SelectedIndex = actualListboxWymA
                    TABsr = ArrayAdd(TABsr, txtWymA.Text)
                End If

            End If

            If txtWymB.Text = "Wymiar B [mm]" Then
                wymB = ListBoxWymB.Text
            Else
                wymB = txtWymB.Text
                ListBoxFill(ListBoxWymB, ArrayAdd(TABwymB, txtWymB.Text))
                'ListBoxWymB.SelectedIndex = actualListboxWymB
                TABwymB = ArrayAdd(TABwymB, txtWymB.Text)
            End If




            Dim atrsVent As String = Join(atrVent, "_")
            If sysname = "" Then
                MsgBox("Nie podano nazwy systemu.")
            End If

            If ListBoxWymB.Enabled = True Then
                If wymA = "" Then
                    MsgBox("Nie podano wymiaru kanału.")
                ElseIf wymB = "" Then
                    MsgBox("Nie podano wymiaru kanału.")
                Else
                    LayerNameVent = ".OBMR_" & sysname & "_" & wymA & "_" & wymB & "_" & sredVent & "_" & atrsVent
                    CreateAndAssignALayer(LayerNameVent)
                End If

            ElseIf ListBoxWymB.Enabled = False Then

                If sredVent = "" Then
                    MsgBox("Nie podano średnicy kanału.")
                Else
                    LayerNameVent = ".OBMR_" & sysname & "_" & wymA & "_" & wymB & "_" & sredVent & "_" & atrsVent
                    CreateAndAssignALayer(LayerNameVent)
                End If
            End If
        End If


    End Sub

    'POLECENIE ANULUJ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles BtnCancel.Click
        'ThisDrawing.SendCommand(Chr(27))
    End Sub


    'KANAŁY//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles txtSysname.TextChanged
        sysname = txtSysname.Text
    End Sub

    Private Sub RadioButton9_CheckedChanged(sender As Object, e As EventArgs) Handles rbtnProst.CheckedChanged
        If rbtnProst.Checked = True Then
            atrVent(0) = "prostokątne"
            rbtnSpiro.Checked = False
            rbtnFlex.Checked = False
            rbtnAkustik.Checked = False
            ListBoxWymB.Enabled = True
            txtWymB.Enabled = True
            ListBoxFill(ListBoxWymA, TABwymA)
            ListBoxFill(ListBoxWymB, TABwymB)
            txtWymA.Text = "Wymiar A [mm]"
            txtWymB.Text = "Wymiar B [mm]"
        ElseIf rbtnSpiro.checked = True Or rbtnFlex.Checked = True Or rbtnAkustik.Checked = True Then
            ListBoxFill(ListBoxWymA, TABsr)
            ListBoxWymB.Items.Clear()
        Else
            ListBoxWymA.Items.Clear()
            ListBoxWymB.Items.Clear()
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles rbtnPD.CheckedChanged
        If rbtnPD.Checked = True Then
            atrVent(0) = "PD"
            rbtnSpiro.Checked = False
            rbtnFlex.Checked = False
            rbtnAkustik.Checked = False
            ListBoxWymB.Enabled = True
            txtWymB.Enabled = True
            ListBoxFill(ListBoxWymA, TABwymA)
            ListBoxFill(ListBoxWymB, TABwymB)
            txtWymA.Text = "Wymiar A [mm]"
            txtWymB.Text = "Wymiar B [mm]"
        ElseIf rbtnSpiro.checked = True Or rbtnFlex.Checked = True Or rbtnAkustik.Checked = True Then
            ListBoxFill(ListBoxWymA, TABsr)
            ListBoxWymB.Items.Clear()
        Else
            atrVent(0) = ""
            ListBoxWymA.Items.Clear()
            ListBoxWymB.Items.Clear()
        End If
    End Sub

    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles rbtnSpiro.CheckedChanged
        If rbtnSpiro.Checked = True Then
            atrVent(0) = "spiro"
            rbtnProst.Checked = False
            rbtnPD.Checked = False
            ListBoxWymB.Enabled = False
            txtWymB.Enabled = False
            wymB = ""
            'ListBoxFill(ListBoxWymA, TABsr)
            txtWymA.Text = "Wymiar A [mm]"
            txtWymB.Text = "Wymiar B [mm]"
        ElseIf rbtnFlex.Checked = True Then
            atrVent(0) = "flex"
        ElseIf rbtnAkustik.Checked = True Then
            atrVent(0) = "akustik"
        Else
            ListBoxWymA.Items.Clear()
        End If
    End Sub

    Private Sub RadioButton10_CheckedChanged(sender As Object, e As EventArgs) Handles rbtnFlex.CheckedChanged
        If rbtnFlex.Checked = True Then
            atrVent(0) = "flex"
            rbtnProst.Checked = False
            rbtnPD.Checked = False
            ListBoxWymB.Enabled = False
            txtWymB.Enabled = False
            wymB = ""
            'ListBoxFill(ListBoxWymA, TABsr)
            txtWymA.Text = "Wymiar A [mm]"
            txtWymB.Text = "Wymiar B [mm]"
        ElseIf rbtnSpiro.Checked = True Then
            atrVent(0) = "spiro"
        ElseIf rbtnAkustik.Checked = True Then
            atrVent(0) = "akustik"

        Else
            ListBoxWymA.Items.Clear()
        End If

    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles rbtnAkustik.CheckedChanged
        If rbtnAkustik.Checked = True Then
            atrVent(0) = "akustik"
            rbtnProst.Checked = False
            rbtnPD.Checked = False
            ListBoxWymB.Enabled = False
            txtWymB.Enabled = False
            wymB = ""
            'ListBoxFill(ListBoxWymA, TABsr)
            txtWymA.Text = "Wymiar A [mm]"
            txtWymB.Text = "Wymiar B [mm]"
        ElseIf rbtnFlex.Checked = True Then
            atrVent(0) = "flex"
        ElseIf rbtnSpiro.Checked = True Then
            atrVent(0) = "spiro"
        Else
            ListBoxWymA.Items.Clear()
        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles rbtnPromat.CheckedChanged
        If rbtnPromat.Checked = True Then
            atrVent(1) = "Promat"
        Else
            atrVent(1) = ""
        End If
    End Sub

    Private Sub rbtnConlit_CheckedChanged(sender As Object, e As EventArgs) Handles rbtnConlit.CheckedChanged
        If rbtnConlit.Checked = True Then
            atrVent(1) = "Conlit"
        Else
            atrVent(1) = ""
        End If
    End Sub

    Private Sub rbtnEIS120_CheckedChanged(sender As Object, e As EventArgs) Handles rbtnEIS120.CheckedChanged
        If rbtnEIS120.Checked = True Then
            atrVent(1) = "EIS120"
        Else
            atrVent(1) = ""
        End If
    End Sub

    Private Sub txtWymA_Click(sender As Object, e As EventArgs) Handles txtWymA.Click
        If txtWymA.Text = "Wymiar A [mm]" Then
            txtWymA.Clear()
        End If
    End Sub

    Private Sub txtWymB_Click(sender As Object, e As EventArgs) Handles txtWymB.Click
        If txtWymB.Text = "Wymiar B [mm]" And txtWymB.Enabled = True Then
            txtWymB.Clear()
        End If
    End Sub

    Private Sub ListBoxWymA_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBoxWymA.SelectedIndexChanged

        If rbtnSpiro.Checked = True Or rbtnFlex.Checked = True Or rbtnAkustik.Checked = True Then
            sredVent = ListBoxWymA.Text
            wymA = ""
            txtWymA.Text = "Wymiar A [mm]"
        Else
            wymA = ListBoxWymA.Text
            txtWymA.Text = "Wymiar A [mm]"
        End If

    End Sub

    Private Sub ListBoxWymB_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBoxWymB.SelectedIndexChanged
        wymB = ListBoxWymB.Text
        txtWymB.Text = "Wymiar B [mm]"
    End Sub

    Private Sub chkboxInneVent_CheckedChanged(sender As Object, e As EventArgs) Handles chkboxInneVent.CheckedChanged
        If chkboxInneVent.Checked = True Then
            txtInneVent.Enabled = True
            txtInneVent.Text = "Inne"
        Else
            txtInneVent.Enabled = False
            txtInneVent.Text = "Inne"
        End If
    End Sub

    Private Sub txtInneVent_TextChanged(sender As Object, e As EventArgs) Handles txtInneVent.TextChanged
        If chkboxInneVent.Checked = True Then
            atrVent(2) = txtInneVent.Text
        Else
            atrVent(2) = ""
        End If

    End Sub

    Private Sub txtInneVent_Click(sender As Object, e As EventArgs) Handles txtInneVent.Click
        If txtInneVent.Text = "Inne" Then
            txtInneVent.Clear()
        End If

    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs)
        'MsgBox("Funkcja w fazie rozwojowej.")

        GetLayerLenght()

    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs)
        MsgBox("Funkcja w fazie rozwojowej.")
    End Sub



    Private Sub chkBoxPPoz_CheckedChanged(sender As Object, e As EventArgs) Handles chkBoxPPoz.CheckedChanged
        If chkBoxPPoz.Checked = True Then
            rbtnConlit.Enabled = True
            rbtnPromat.Enabled = True
            rbtnEIS120.Enabled = True
            RadioButton4.Enabled = True
            RadioButton3.Enabled = True
            CheckBox3.Enabled = True
            TextBox1.Enabled = True
        Else
            atrVent(1) = ""
            rbtnConlit.Enabled = False
            rbtnPromat.Enabled = False
            rbtnEIS120.Enabled = False
            RadioButton4.Enabled = False
            RadioButton3.Enabled = False
            CheckBox3.Enabled = False
            TextBox1.Enabled = False
            rbtnConlit.Checked = False
            rbtnPromat.Checked = False
            rbtnEIS120.Checked = False
            RadioButton4.Checked = False
            RadioButton3.Checked = False
            CheckBox3.Checked = False
        End If
    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub txtWymA_TextChanged(sender As Object, e As EventArgs) Handles txtWymA.TextChanged

    End Sub

    Private Sub RadioButton2_CheckedChanged_1(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            atrVent(0) = "PD"
            rbtnSpiro.Checked = False
            rbtnFlex.Checked = False
            rbtnAkustik.Checked = False
            ListBoxWymB.Enabled = True
            txtWymB.Enabled = True
            ListBoxFill(ListBoxWymA, TABwymA)
            ListBoxFill(ListBoxWymB, TABwymB)
            txtWymA.Text = "Wymiar A [mm]"
            txtWymB.Text = "Wymiar B [mm]"
        ElseIf rbtnSpiro.Checked = True Or rbtnFlex.Checked = True Or rbtnAkustik.Checked = True Then
            ListBoxFill(ListBoxWymA, TABsr)
            ListBoxWymB.Items.Clear()
        Else
            atrVent(0) = ""
            ListBoxWymA.Items.Clear()
            ListBoxWymB.Items.Clear()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles rInneChkBox.CheckedChanged
        If rInneChkBox.Checked = True Then
            rInne.Enabled = True
            rInne.Text = "Inne"
        Else
            rInne.Enabled = False
            rInne.Text = "Inne"
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged

    End Sub

    Private Sub txtWymB_TextChanged(sender As Object, e As EventArgs) Handles txtWymB.TextChanged

    End Sub

    Private Sub txtInnaSred_TextChanged(sender As Object, e As EventArgs) Handles txtInnaSred.TextChanged

    End Sub

    Private Sub txtInnaSred_Click(sender As Object, e As EventArgs) Handles txtInnaSred.Click
        If txtInnaSred.Text = "Inna średnica" Then
            txtInnaSred.Clear()
        End If
    End Sub

    Private Sub rbtnDN_CheckedChanged(sender As Object, e As EventArgs) Handles rbtnDN.CheckedChanged
        If rbtnDN.Checked = True Then
            ListBoxFill(ListBoxSred, TABsrDN)
            txtInnaSred.Text = "Inna średnica"
        End If
    End Sub

    Private Sub rbtnFi_CheckedChanged(sender As Object, e As EventArgs) Handles rbtnFi.CheckedChanged
        If rbtnFi.Checked = True Then
            ListBoxFill(ListBoxSred, TABsrFi)
            txtInnaSred.Text = "Inna średnica"
        End If
    End Sub

    Private Sub rInne_TextChanged(sender As Object, e As EventArgs) Handles rInne.TextChanged

    End Sub

    Private Sub rInne_Click(sender As Object, e As EventArgs) Handles rInne.Click
        If rInne.Text = "Inne" Then
            rInne.Clear()
        End If
    End Sub

    Private Sub chkBoxMaterial_CheckedChanged(sender As Object, e As EventArgs) Handles chkBoxMaterial.CheckedChanged
        If chkBoxMaterial.Checked = True Then
            RBtnStal.Enabled = True
            RBtnPP.Enabled = True
            RBtnPVC.Enabled = True
            RBtnStalNierdz.Enabled = True
            RBtnHDPE.Enabled = True
            RBtnZeliwo.Enabled = True

        Else
            mat = ""
            RBtnStal.Enabled = False
            RBtnPP.Enabled = False
            RBtnPVC.Enabled = False
            RBtnStalNierdz.Enabled = False
            RBtnHDPE.Enabled = False
            RBtnZeliwo.Enabled = False
        End If
    End Sub

    Private Sub GrpInst_Enter(sender As Object, e As EventArgs) Handles GrpInst.Enter

    End Sub

    Private Sub Button2_Click_2(sender As Object, e As EventArgs) Handles Button2.Click
        ThisDrawing.SendCommand("LOS" & vbCr)
        ThisDrawing.SendCommand("TLEN" & vbCr)
    End Sub

    Private Sub Button1_Click_2(sender As Object, e As EventArgs) Handles Button1.Click
        ThisDrawing.SendCommand("LOS" & vbCr)
        ThisDrawing.SendCommand("ADDLEN" & vbCr)
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub
End Class

Friend Class LineType
    Friend Shared Function m_ps() As Object
        Throw New NotImplementedException()
    End Function
End Class
