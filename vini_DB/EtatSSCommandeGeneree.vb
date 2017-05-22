Public Class EtatSSCommandeGeneree
    Inherits EtatCommande

    Public Sub New(Optional ByVal pactionMvtStock As vncGenererSupprimer = vncEnums.vncGenererSupprimer.vncRien, Optional ByVal pactionSousCommande As vncGenererSupprimer = vncEnums.vncGenererSupprimer.vncRien)
        MyBase.New(pactionMvtStock, pactionSousCommande)
        codeEtat = vncEnums.vncEtatCommande.vncSCMDGeneree
        m_libelle = ETAT_SS_GENEREE
    End Sub

    Protected Overrides Function action(ByVal vncAction As vncActionEtatCommande) As vncEtatCommande
        Dim nReturn As vncEtatCommande
        Select Case vncAction
            Case vncEnums.vncActionEtatCommande.vncActionSCMDFaxer
                nReturn = vncEnums.vncEtatCommande.vncSCMDtransmiseFax
            Case vncEnums.vncActionEtatCommande.vncActionSCMDExportInternet
                nReturn = vncEnums.vncEtatCommande.vncSCMDExporteeInt
            Case vncEnums.vncActionEtatCommande.vncActionSCMDRapprocher
                nReturn = vncEnums.vncEtatCommande.vncSCMDRapprochee
            Case Else
                nReturn = codeEtat
        End Select
        Return nReturn
    End Function

End Class
