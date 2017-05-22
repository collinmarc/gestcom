Public Class EtatCommandeEnCoursSaisie
    Inherits EtatCommande

    Public Sub New(Optional ByVal pactionMvtStock As vncGenererSupprimer = vncEnums.vncGenererSupprimer.vncRien, Optional ByVal pactionSousCommande As vncGenererSupprimer = vncEnums.vncGenererSupprimer.vncRien)
        MyBase.New(pactionMvtStock, pactionSousCommande)
        m_codeEtat = vncEnums.vncEtatCommande.vncEnCoursSaisie
        m_libelle = ETAT_ENCOURS
    End Sub
    Protected Overrides Function action(ByVal vncAction As vncActionEtatCommande) As vncEtatCommande
        Dim nReturn As vncActionEtatCommande
        Select Case vncAction
            Case vncEnums.vncActionEtatCommande.vncActionValider
                nReturn = vncEnums.vncEtatCommande.vncValidee
            Case vncEnums.vncActionEtatCommande.vncActionLivrer
                m_actionMvtStock = vncEnums.vncGenererSupprimer.vncGenerer
                nReturn = vncEnums.vncEtatCommande.vncLivree
            Case Else
                nReturn = vncEnums.vncEtatCommande.vncEnCoursSaisie
        End Select
        Return nReturn
    End Function

End Class
