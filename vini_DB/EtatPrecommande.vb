Public Class EtatPrecommande
    Inherits EtatCommande

    Public Sub New()
        MyBase.New(vncEnums.vncGenererSupprimer.vncRien, vncEnums.vncGenererSupprimer.vncRien)
        codeEtat = vncEnums.vncEtatCommande.vncRien
        m_libelle = "Tous"
    End Sub

    Protected Overrides Function action(ByVal vncAction As vncActionEtatCommande) As vncEtatCommande
        Dim nReturn As vncEtatCommande
        nReturn = codeEtat
        Return nReturn
    End Function

End Class
