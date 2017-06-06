Public Class EtatFactHBVGeneree
    Inherits EtatCommande

    Public Sub New()
        MyBase.New(vncEnums.vncGenererSupprimer.vncRien, vncEnums.vncGenererSupprimer.vncRien)
        codeEtat = vncEnums.vncEtatCommande.vncFactHBVGeneree
        m_libelle = ETAT_GENEREE
    End Sub

    Protected Overrides Function action(ByVal vncAction As vncActionEtatCommande) As vncEtatCommande
        Dim nReturn As vncEtatCommande
        Select Case vncAction
            Case vncEnums.vncActionEtatCommande.vncActionFactHBVExporter
                nReturn = vncEnums.vncEtatCommande.vncFactHBVExportee
            Case Else
                nReturn = codeEtat
        End Select
        Return nReturn
    End Function 'Action

End Class
