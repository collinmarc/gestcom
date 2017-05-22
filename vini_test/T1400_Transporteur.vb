Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports NUnit.Extensions.Forms
Imports vini_DB
Imports vini_App


<TestClass()> Public Class T1400_Transporteur
    Inherits test_Base

    Private m_objTransporteur As Transporteur
    <TestInitialize()> Public Overrides Sub TestInitialize()
        MyBase.TestInitialize()
        m_objTransporteur.shared_connect()
        m_objTransporteur = New Transporteur()
    End Sub
    <TestCleanup()> Public Overrides Sub TestCleanup()
        Dim oTA As dsVinicomTableAdapters.TRANSPORTEURTableAdapter
        Dim oDT As dsVinicom.TRANSPORTEURDataTable
        Dim oRow As dsVinicom.TRANSPORTEURRow

        oTA = New dsVinicomTableAdapters.TRANSPORTEURTableAdapter()
        oDT = oTA.GetData()
        For Each oRow In oDT.Rows
            If oRow.TRP_CODE = "TRPTEST" Then
                oRow.Delete()
            End If
        Next
        oTA.Update(oDT)
        m_objTransporteur.shared_disconnect()
        MyBase.TestCleanup()
    End Sub
    <TestMethod()> Public Sub T01_CREATE()
        Dim n As Integer

        m_objTransporteur = New Transporteur()
        m_objTransporteur.code = "TRPTEST"
        m_objTransporteur.nom = "Transporteur 1"
        m_objTransporteur.AdresseLivraison.nom = m_objTransporteur.nom
        m_objTransporteur.AdresseLivraison.rue1 = " "
        m_objTransporteur.AdresseLivraison.rue2 = " "
        m_objTransporteur.AdresseLivraison.cp = " "
        m_objTransporteur.AdresseLivraison.ville = " "


        n = Transporteur.colTransporteur().Count
        Assert.IsTrue(m_objTransporteur.Save(), m_objTransporteur.getErreur())
        ''Sans rechragement
        Assert.AreEqual(n, Transporteur.colTransporteur.Count())
        ''Avec rechargement
        Assert.AreEqual(n + 1, Transporteur.colTransporteur(True).Count())

    End Sub

    <TestMethod()> Public Sub T02_CREATE_DataSet()
        Dim oTA As dsVinicomTableAdapters.TRANSPORTEURTableAdapter
        Dim oDt As New dsVinicom.TRANSPORTEURDataTable
        Dim oRow As dsVinicom.TRANSPORTEURRow

        oTA = New dsVinicomTableAdapters.TRANSPORTEURTableAdapter()
        oDt = oTA.GetData()
        oRow = oDt.NewTRANSPORTEURRow()
        oRow.TRP_CODE = "TRPTEST"
        oRow.TRP_NOM = "TRP2"
        oRow.TRP_LIV_RUE1 = "TRP2"
        oRow.TRP_LIV_RUE2 = "TRP2"
        oRow.TRP_LIV_CP = "TRP2"
        oRow.TRP_LIV_VILLE = "TRP2"

        oTA.Update(oDt)


    End Sub

End Class


