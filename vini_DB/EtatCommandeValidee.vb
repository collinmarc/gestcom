Public Class EtatCommandeValidee
    Inherits EtatCommande

    Public Sub New(Optional ByVal pactionMvtStock As vncGenererSupprimer = vncEnums.vncGenererSupprimer.vncRien, Optional ByVal pactionSousCommande As vncGenererSupprimer = vncEnums.vncGenererSupprimer.vncRien)
        MyBase.New(pactionMvtStock, pactionSousCommande)
        codeEtat = vncEnums.vncEtatCommande.vncValidee
        m_libelle = ETAT_VALIDEE
    End Sub

    Protected Overrides Function action(ByVal vncAction As vncActionEtatCommande) As vncEtatCommande
        Dim nReturn As vncEtatCommande
        Select Case vncAction
            Case vncEnums.vncActionEtatCommande.vncActionAnnValider
                nReturn = vncEnums.vncEtatCommande.vncEnCoursSaisie
            Case vncEnums.vncActionEtatCommande.vncActionLivrer
                nReturn = vncEnums.vncEtatCommande.vncLivree
                m_actionMvtStock = vncEnums.vncGenererSupprimer.vncGenerer
            Case Else
                nReturn = vncEnums.vncEtatCommande.vncValidee
        End Select
        Return nReturn
    End Function

End Class
