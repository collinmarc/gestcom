Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports vini_DB
Imports System.Collections.Generic
Imports System.IO
Imports Microsoft.Office.Interop
Imports System.Data
<TestClass()> Public Class TestImportEXcel
    Inherits test_Base
    Private _TestContext As TestContext
    Public Property NewProperty() As TestContext
        Get
            Return _TestContext
        End Get
        Set(ByVal value As TestContext)
            _TestContext = value
        End Set
    End Property

    <TestInitialize()> Public Overrides Sub TestInitialize()
        MyBase.TestInitialize()

        If System.IO.File.Exists("./testImportEXCEL.xlsx") Then
            System.IO.File.Delete("./testImportEXCEL.xlsx")
        End If
        Dim objApp As New Microsoft.Office.Interop.Excel.Application()
        Try
            objApp.Workbooks().Add()
            Dim oSheet As Excel.Worksheet
            Dim oWB As Excel.Workbook
            oWB = objApp.Workbooks(1)
            oSheet = oWB.Sheets(1)
            oSheet.Cells(1, 2) = "Colonne2"
            oSheet.Cells(1, 19) = "ColonneS"
            Dim n As Integer
            For n = 10 To 15
                oSheet.Cells(n, 2).Value = "9999" & n
                oSheet.Cells(n, 19).Value = "123.45"
            Next
            oSheet.Cells(n + 1, 2).Value = "999910"
            oSheet.Cells(n + 1, 2).Value = "TESTVALEURINCONNUE"

            oSheet.Cells(n + 2, 2).Value = "999910"
            objApp.Workbooks(1).SaveAs(Environment.CurrentDirectory & "/testImportExcel.xlsx")
            Console.WriteLine("fichier créé en " & Environment.CurrentDirectory & "/testImportExcel.xlsx")
        Catch ex As Exception
            Console.WriteLine("Erreur en ecriture du fichier EXCEL" & ex.Message)

        End Try
        objApp.Quit

    End Sub
    <TestCleanup()> Public Overrides Sub TestCleanup()

        MyBase.TestCleanup()

    End Sub
    ''' <summary>
    ''' Test de l'export d'une souscommande vers QuadraFacturation
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()>
    Public Sub T101_ImportTarif()
        Dim oFRN As New Fournisseur("FRNTEST", "TEST")
        oFRN.Save()
        Dim oProduit As New Produit("999910", oFRN, "2017")
        oProduit.TarifA = 741.85
        oProduit.TarifA = 852.96
        oProduit.TarifA = 963.74
        oProduit.save()
        Dim idProduit As Integer
        idProduit = oProduit.id

        Dim oProduit2 As New Produit("999915", oFRN, "2017")
        oProduit2.TarifA = 741.85
        oProduit2.TarifA = 852.96
        oProduit2.TarifA = 963.74
        oProduit2.save()
        Dim idProduit2 As Integer
        idProduit2 = oProduit.id

        Dim oImport As New ImportTarifGESTCOM(Environment.CurrentDirectory & "/testImportExcel.xlsx", "Feuil1", 2, 19)

        oImport.ImportTarif()
        oProduit = Produit.createandload(idProduit)

        Assert.AreEqual(123.45D, oProduit.TarifA)
        Assert.AreEqual(123.45D, oProduit.TarifB)
        Assert.AreEqual(123.45D, oProduit.TarifC)

        oProduit2 = Produit.createandload(idProduit2)

        Assert.AreEqual(123.45D, oProduit2.TarifA)
        Assert.AreEqual(123.45D, oProduit2.TarifB)
        Assert.AreEqual(123.45D, oProduit2.TarifC)

        oProduit.delete()
        oProduit2.delete()
        oFRN.delete()

    End Sub

    <TestMethod()>
    Public Sub T101_GetNbreLignes()

        Dim oImport As New ImportTarifGESTCOM(Environment.CurrentDirectory & "/testImportExcel.xlsx", "Feuil1", 2, 19)

        Assert.AreEqual(18, oImport.getNbreLignes())

    End Sub


End Class