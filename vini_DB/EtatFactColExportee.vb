Public Class EtatFactcolExportee
    Inherits EtatCommande

    Public Sub New()
        MyBase.New(vncEnums.vncGenererSupprimer.vncRien, vncEnums.vncGenererSupprimer.vncRien)
        codeEtat = vncEnums.vncEtatCommande.vncFactCOLExportee
        m_libelle = ETAT_EXPORTEE
    End Sub

    Protected Overrides Function action(ByVal vncAction As vncActionEtatCommande) As vncEtatCommande
        Dim nReturn As vncEtatCommande
        Select Case vncAction
            Case vncEnums.vncActionEtatCommande.vncActionFactCOLAnnExporter
                nReturn = vncEnums.vncEtatCommande.vncFactCOLGeneree
            Case Else
                nReturn = codeEtat
        End Select
        Return nReturn
    End Function

End Class
