Public Class EtatCommandeTransmiseQuadra
    Inherits EtatCommande

    Public Sub New(Optional ByVal pactionMvtStock As vncGenererSupprimer = vncEnums.vncGenererSupprimer.vncRien, Optional ByVal pactionSousCommande As vncGenererSupprimer = vncEnums.vncGenererSupprimer.vncRien)
        MyBase.New(pactionMvtStock, pactionSousCommande)
        codeEtat = vncEnums.vncEtatCommande.vncTransmiseQuadra
        m_libelle = ETAT_TRANSMISE
    End Sub
    ''' <summary>
    ''' Si annluEclatement => Livrée , sinon on reste dans l'état présent
    ''' </summary>
    ''' <param name="vncAction"></param>
    ''' <returns></returns>
    Protected Overrides Function action(ByVal vncAction As vncActionEtatCommande) As vncEtatCommande
        Dim nReturn As vncActionEtatCommande
        Select Case vncAction
            Case vncEnums.vncActionEtatCommande.vncActionAnnEclater
                m_actionSousCommande = vncEnums.vncGenererSupprimer.vncSupprimer
                nReturn = vncEnums.vncEtatCommande.vncLivree
            Case Else
                nReturn = vncEnums.vncEtatCommande.vncTransmiseQuadra
        End Select
        Return nReturn
    End Function

End Class
