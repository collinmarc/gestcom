Public Class EtatCommandeLivree
    Inherits EtatCommande

    Public Sub New(Optional ByVal pactionMvtStock As vncGenererSupprimer = vncEnums.vncGenererSupprimer.vncRien, Optional ByVal pactionSousCommande As vncGenererSupprimer = vncEnums.vncGenererSupprimer.vncRien)
        MyBase.New(pactionMvtStock, pactionSousCommande)
        codeEtat = vncEnums.vncEtatCommande.vncLivree
        m_libelle = ETAT_LIVREE
    End Sub

    Protected Overrides Function action(ByVal vncAction As vncActionEtatCommande) As vncEtatCommande
        Dim nReturn As vncEtatCommande
        Select Case vncAction

            Case vncEnums.vncActionEtatCommande.vncActionAnnLivrer
                m_actionMvtStock = vncEnums.vncGenererSupprimer.vncSupprimer
                nReturn = vncEnums.vncEtatCommande.vncValidee
            Case vncEnums.vncActionEtatCommande.vncActionEclater
                m_actionSousCommande = vncEnums.vncGenererSupprimer.vncGenerer
                nReturn = vncEnums.vncEtatCommande.vncEclatee
            Case Else
                nReturn = vncEnums.vncEtatCommande.vncLivree
        End Select

        Return nReturn
    End Function

End Class
