Public Class EtatCommandeRien
    Inherits EtatCommande

    Friend Sub New(Optional ByVal pactionMvtStock As vncGenererSupprimer = vncEnums.vncGenererSupprimer.vncRien, Optional ByVal pactionSousCommande As vncGenererSupprimer = vncEnums.vncGenererSupprimer.vncRien)
        MyBase.New(pactionMvtStock, pactionSousCommande)
        m_codeEtat = vncEnums.vncEtatCommande.vncRien
        m_libelle = "TOUS"
    End Sub
    Protected Overrides Function action(ByVal vncAction As vncActionEtatCommande) As vncEtatCommande
        Return vncEtatCommande.vncRien
    End Function

End Class
