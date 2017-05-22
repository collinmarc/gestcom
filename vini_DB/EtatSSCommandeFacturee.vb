Public Class EtatSSCommandeFacturee
    Inherits EtatCommande

    Public Sub New(Optional ByVal pactionMvtStock As vncGenererSupprimer = vncEnums.vncGenererSupprimer.vncRien, Optional ByVal pactionSousCommande As vncGenererSupprimer = vncEnums.vncGenererSupprimer.vncRien)
        MyBase.New(pactionMvtStock, pactionSousCommande)
        codeEtat = vncEnums.vncEtatCommande.vncSCMDFacturee
        m_libelle = ETAT_SS_FACTUREE
    End Sub

    Protected Overrides Function action(ByVal vncAction As vncActionEtatCommande) As vncEtatCommande
        Dim nReturn As vncEtatCommande
        Select Case vncAction
            Case vncEnums.vncActionEtatCommande.vncActionSCMDAnnFacturer
                nReturn = vncEnums.vncEtatCommande.vncSCMDRapprochee
            Case Else
                nReturn = codeEtat
        End Select
        Return nReturn
    End Function

End Class
