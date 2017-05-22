Public Class EtatBAEnCoursSaisie
    Inherits EtatCommande

    Public Sub New(Optional ByVal pactionMvtStock As vncGenererSupprimer = vncEnums.vncGenererSupprimer.vncRien, Optional ByVal pactionSousCommande As vncGenererSupprimer = vncEnums.vncGenererSupprimer.vncRien)
        MyBase.New(pactionMvtStock, pactionSousCommande)
        m_codeEtat = vncEnums.vncEtatCommande.vncBAEnCours
        m_libelle = ETAT_ENCOURS
    End Sub
    Protected Overrides Function action(ByVal vncAction As vncActionEtatCommande) As vncEtatCommande
        Dim nReturn As vncActionEtatCommande
        Select Case vncAction
            Case vncEnums.vncActionEtatCommande.vncActionBALivrer
                nReturn = vncEnums.vncEtatCommande.vncBALivre
                m_actionMvtStock = vncEnums.vncGenererSupprimer.vncGenerer
                m_actionSousCommande = vncEnums.vncGenererSupprimer.vncRien
            Case vncEnums.vncActionEtatCommande.vncActionBAAnnLivrer
                nReturn = vncEnums.vncEtatCommande.vncBAEnCours
                m_actionMvtStock = vncEnums.vncGenererSupprimer.vncRien
                m_actionSousCommande = vncEnums.vncGenererSupprimer.vncRien
            Case Else
                nReturn = m_codeEtat
        End Select
        Return nReturn
    End Function

End Class
