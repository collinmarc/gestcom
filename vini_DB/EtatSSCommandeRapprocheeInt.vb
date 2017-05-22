Public Class EtatSSCommandeRapprocheeInt
    Inherits EtatCommande

    Public Sub New(Optional ByVal pactionMvtStock As vncGenererSupprimer = vncEnums.vncGenererSupprimer.vncRien, Optional ByVal pactionSousCommande As vncGenererSupprimer = vncEnums.vncGenererSupprimer.vncRien)
        MyBase.New(pactionMvtStock, pactionSousCommande)
        codeEtat = vncEnums.vncEtatCommande.vncSCMDRapprocheeInt
        m_libelle = ETAT_SS_RAPPROCHEEInt
    End Sub

    Protected Overrides Function action(ByVal vncAction As vncActionEtatCommande) As vncEtatCommande
        Dim nReturn As vncEtatCommande
        Select Case vncAction
            Case vncEnums.vncActionEtatCommande.vncActionSCMDAnnImportInternet
                nReturn = vncEnums.vncEtatCommande.vncSCMDExporteeInt
            Case vncEnums.vncActionEtatCommande.vncActionSCMDFacturer
                nReturn = vncEnums.vncEtatCommande.vncSCMDFacturee
            Case Else
                nReturn = codeEtat
        End Select
        Return nReturn
    End Function

End Class
