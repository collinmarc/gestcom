Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports vini_DB
Imports System.Data.OleDb

<TestClass()> Public Class TestFactCol

    <TestMethod(), Ignore()> Public Sub TestMethod1()

        Dim ocsBld As New System.Data.OleDb.OleDbConnectionStringBuilder(My.Settings.ConnectionString)
        ocsBld("InitialCatalog") = "vnc520190402"
        Dim oCS As String = ocsBld.ConnectionString

        Persist.ConnectionString = oCS

        Dim oPrd As Produit
        Dim ds As New dsVinicom()
        oPrd = Produit.createandload(2945)
        oPrd.loadcolmvtStockDepuisLeDernierMouvementInventaire()
        oPrd.GenereDataSetRecapColisage(CDate("2019-2-1"), CDate("2019-2-28"), 7, ds)
        Assert.AreEqual(1, ds.RECAPCOLISAGEJOURN.Count)
        Dim oRow As dsVinicom.RECAPCOLISAGEJOURNRow = ds.RECAPCOLISAGEJOURN(0)

        Dim qte As Decimal
        qte = oRow.RC_S01 + oRow.RC_S02 + oRow.RC_S03 + oRow.RC_S04 + oRow.RC_S05 + oRow.RC_S06 + oRow.RC_S07 + oRow.RC_S08 + oRow.RC_S09 + oRow.RC_S10 + oRow.RC_S11 + oRow.RC_S12 + oRow.RC_S13 + oRow.RC_S14 + oRow.RC_S15 + oRow.RC_S16 + oRow.RC_S17 + oRow.RC_S18 + oRow.RC_S19 + oRow.RC_S20 + oRow.RC_S21 + oRow.RC_S22 + oRow.RC_S23 + oRow.RC_S24 + oRow.RC_S25 + oRow.RC_S26 + oRow.RC_S27 + oRow.RC_S28 + oRow.RC_S29 + oRow.RC_S30 + oRow.RC_S31
        Assert.AreEqual(6408D, qte)



    End Sub


    <TestMethod()> Public Sub TestPeriode()
        Dim Periode As String
        Dim dDeb As Date
        Periode = DateAndTime.Now.ToString("MMMM yyyy")
        dDeb = CDate(Periode)
        Assert.AreEqual(1, dDeb.Day)


    End Sub

End Class