Public Class EtatFactTRPExportee
    Inherits EtatCommande

    Public Sub New()
        MyBase.New(vncEnums.vncGenererSupprimer.vncRien, vncEnums.vncGenererSupprimer.vncRien)
        codeEtat = vncEnums.vncEtatCommande.vncFactTRPExportee
        m_libelle = ETAT_EXPORTEE
    End Sub

    Protected Overrides Function action(ByVal vncAction As vncActionEtatCommande) As vncEtatCommande
        Dim nReturn As vncEtatCommande
        Select Case vncAction
            Case vncEnums.vncActionEtatCommande.vncActionFactTRPAnnExporter
                nReturn = vncEnums.vncEtatCommande.vncFactTRPGeneree
            Case Else
                nReturn = codeEtat
        End Select
        Return nReturn
    End Function

End Class
