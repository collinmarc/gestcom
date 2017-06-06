Public Class EtatFactHBVExportee
    Inherits EtatCommande

    Public Sub New()
        MyBase.New(vncEnums.vncGenererSupprimer.vncRien, vncEnums.vncGenererSupprimer.vncRien)
        codeEtat = vncEnums.vncEtatCommande.vncFactHBVExportee
        m_libelle = ETAT_EXPORTEE
    End Sub

    Protected Overrides Function action(ByVal vncAction As vncActionEtatCommande) As vncEtatCommande
        Dim nReturn As vncEtatCommande
        Select Case vncAction
            Case vncEnums.vncActionEtatCommande.vncActionFactHBVAnnExporter
                nReturn = vncEnums.vncEtatCommande.vncFactHBVGeneree
            Case Else
                nReturn = codeEtat
        End Select
        Return nReturn
    End Function

End Class
