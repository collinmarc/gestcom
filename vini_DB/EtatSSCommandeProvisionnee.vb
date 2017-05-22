Public Class EtatSSCommandeProvisionnee2
    'Inherits EtatCommande

    'Public Sub New(Optional ByVal pactionMvtStock As vncGenererSupprimer = vncEnums.vncGenererSupprimer.vncRien, Optional ByVal pactionSousCommande As vncGenererSupprimer = vncEnums.vncGenererSupprimer.vncRien)
    '    MyBase.New(pactionMvtStock, pactionSousCommande)
    '    codeEtat = vncEnums.vncEtatCommande.vncSCMDProvisionnee
    '    m_libelle = ETAT_SS_PROVIONNEE
    'End Sub

    'Protected Overrides Function action(ByVal vncAction As vncActionEtatCommande) As vncEtatCommande
    '    Dim nReturn As vncEtatCommande
    '    Select Case vncAction
    '        Case vncEnums.vncActionEtatCommande.vncActionSCMDAnnProvisionner
    '            nReturn = vncEnums.vncEtatCommande.vncSCMDRapprochee
    '        Case vncEnums.vncActionEtatCommande.vncActionSCMDFacturer
    '            nReturn = vncEnums.vncEtatCommande.vncSCMDFacturee
    '        Case Else
    '            nReturn = codeEtat
    '    End Select
    '    Return nReturn
    'End Function

End Class
