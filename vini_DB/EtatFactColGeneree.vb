Public Class EtatFactColGeneree
    Inherits EtatCommande

    Public Sub New()
        MyBase.New(vncEnums.vncGenererSupprimer.vncGenerer, vncEnums.vncGenererSupprimer.vncRien)
        codeEtat = vncEnums.vncEtatCommande.vncFactCOLGeneree
        m_libelle = ETAT_GENEREE
    End Sub

    Protected Overrides Function action(ByVal vncAction As vncActionEtatCommande) As vncEtatCommande
        Dim nReturn As vncEtatCommande
        Select Case vncAction
            Case vncEnums.vncActionEtatCommande.vncActionFactCOLExporter
                nReturn = vncEnums.vncEtatCommande.vncFactCOLExportee
            Case Else
                nReturn = codeEtat
        End Select
        Return nReturn
    End Function 'Action

End Class
