'Test de la classe dbConnection
Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports vini_DB



<TestClass()> Public Class T1020_Etat
    Inherits test_Base


    Private m_oCmd As CommandeClient
    <TestInitialize()> Public Overrides Sub TestInitialize()
        MyBase.TestInitialize()
        m_oCmd = New CommandeClient(New Client("", ""))
    End Sub
    <TestCleanup()> Public Overrides Sub TestCleanup()
        m_oCmd = Nothing
        MyBase.TestCleanup()
    End Sub
    <TestMethod()> Public Sub T10_Object()

        Assert.IsTrue(m_oCmd.etat.codeEtat = vncEnums.vncEtatCommande.vncEnCoursSaisie)
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionValider)
        Assert.IsTrue(m_oCmd.etat.codeEtat = vncEnums.vncEtatCommande.vncValidee)
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionAnnValider)
        Assert.IsTrue(m_oCmd.etat.codeEtat = vncEnums.vncEtatCommande.vncEnCoursSaisie)
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(m_oCmd.etat.codeEtat = vncEnums.vncEtatCommande.vncLivree)
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionAnnLivrer)
        Assert.IsTrue(m_oCmd.etat.codeEtat = vncEnums.vncEtatCommande.vncValidee)
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionAnnValider)
        Assert.IsTrue(m_oCmd.etat.codeEtat = vncEnums.vncEtatCommande.vncEnCoursSaisie)


        'VALIDEE
        'Action Permise
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionValider)
        Assert.IsTrue(m_oCmd.etat.codeEtat = vncEnums.vncEtatCommande.vncValidee)
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(m_oCmd.etat.codeEtat = vncEnums.vncEtatCommande.vncLivree)
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionAnnLivrer)
        Assert.IsTrue(m_oCmd.etat.codeEtat = vncEnums.vncEtatCommande.vncValidee)
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionAnnValider)
        Assert.IsTrue(m_oCmd.etat.codeEtat = vncEnums.vncEtatCommande.vncEnCoursSaisie)
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionValider)
        Assert.IsTrue(m_oCmd.etat.codeEtat = vncEnums.vncEtatCommande.vncValidee)
        'Action non permise
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionValider)
        Assert.IsTrue(m_oCmd.etat.codeEtat = vncEnums.vncEtatCommande.vncValidee)
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionEclater)
        Assert.IsTrue(m_oCmd.etat.codeEtat = vncEnums.vncEtatCommande.vncValidee)
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionAnnEclater)
        Assert.IsTrue(m_oCmd.etat.codeEtat = vncEnums.vncEtatCommande.vncValidee)

        'LIVREE
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(m_oCmd.etat.codeEtat = vncEnums.vncEtatCommande.vncLivree)
        Assert.IsTrue(m_oCmd.getActionMvtStock = vncEnums.vncGenererSupprimer.vncGenerer)
        'Action Permise
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionEclater)
        Assert.IsTrue(m_oCmd.etat.codeEtat = vncEnums.vncEtatCommande.vncEclatee)
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionAnnEclater)
        Assert.IsTrue(m_oCmd.etat.codeEtat = vncEnums.vncEtatCommande.vncLivree)
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionAnnLivrer)
        Assert.IsTrue(m_oCmd.etat.codeEtat = vncEnums.vncEtatCommande.vncValidee)
        Assert.IsTrue(m_oCmd.getActionMvtStock = vncEnums.vncGenererSupprimer.vncSupprimer)
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(m_oCmd.etat.codeEtat = vncEnums.vncEtatCommande.vncLivree)
        'Action non permise
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(m_oCmd.etat.codeEtat = vncEnums.vncEtatCommande.vncLivree)
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionValider)
        Assert.IsTrue(m_oCmd.etat.codeEtat = vncEnums.vncEtatCommande.vncLivree)
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionAnnEclater)
        Assert.IsTrue(m_oCmd.etat.codeEtat = vncEnums.vncEtatCommande.vncLivree)
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionAnnValider)
        Assert.IsTrue(m_oCmd.etat.codeEtat = vncEnums.vncEtatCommande.vncLivree)

    End Sub
End Class



