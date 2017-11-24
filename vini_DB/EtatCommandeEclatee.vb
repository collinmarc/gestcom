Public Class EtatCommandeEclatee
    Inherits EtatCommande

    Public Sub New(Optional ByVal pactionMvtStock As vncGenererSupprimer = vncEnums.vncGenererSupprimer.vncRien, Optional ByVal pactionSousCommande As vncGenererSupprimer = vncEnums.vncGenererSupprimer.vncRien)
        MyBase.New(pactionMvtStock, pactionSousCommande)
        codeEtat = vncEnums.vncEtatCommande.vncEclatee
        m_libelle = ETAT_ECLATEE
    End Sub

    Protected Overrides Function action(ByVal vncAction As vncActionEtatCommande) As vncEtatCommande
        Dim nReturn As vncActionEtatCommande
        Select Case vncAction
            Case vncEnums.vncActionEtatCommande.vncActionAnnEclater
                m_actionSousCommande = vncEnums.vncGenererSupprimer.vncSupprimer
                nReturn = vncEnums.vncEtatCommande.vncLivree
            Case vncEnums.vncActionEtatCommande.vncActionTransmettre
                nReturn = vncEnums.vncEtatCommande.vncTransmiseQuadra
            Case Else
                nReturn = vncEnums.vncEtatCommande.vncEclatee
        End Select
        Return nReturn
    End Function

End Class
