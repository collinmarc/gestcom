Public Class EtatFactTRPGeneree
    Inherits EtatCommande

    Public Sub New()
        MyBase.New(vncEnums.vncGenererSupprimer.vncRien, vncEnums.vncGenererSupprimer.vncRien)
        codeEtat = vncEnums.vncEtatCommande.vncFactTRPGeneree
        m_libelle = ETAT_GENEREE
    End Sub

    Protected Overrides Function action(ByVal vncAction As vncActionEtatCommande) As vncEtatCommande
        Dim nReturn As vncEtatCommande
        Select Case vncAction
            Case vncEnums.vncActionEtatCommande.vncActionFactTRPExporter
                nReturn = vncEnums.vncEtatCommande.vncFactTRPExportee
            Case Else
                nReturn = codeEtat
        End Select
        Return nReturn
    End Function 'Action

End Class
