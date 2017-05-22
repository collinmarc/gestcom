Imports vini_DB
Public Class frmSaisieReglement
    Inherits vini_app.FrmVinicom

    Private m_oReglement As Reglement
    Private m_oFacture As Facture
    Protected Overrides Function frmLoad() As Boolean
        Return MyBase.frmLoad()
        dtDateReglement.Value = DateTime.Now
        m_oReglement = New Reglement()
    End Function
    Public Overrides Function getResume() As String
        Return "Saisie des règlements"
    End Function

    Protected Overrides Sub EnableControls(ByVal bEnabled As Boolean)
        cbAppliquer.Enabled = (m_oFacture IsNot Nothing)
    End Sub


    Private Sub cbRechercheFacture_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbRechercheFacture.Click
        rechercheFacture()
    End Sub

    Private Sub rechercheFacture()
        Dim frm As frmRechercheDB
        Dim obj As Facture

        Try

            frm = New frmRechercheDB()
            If (rbFactCommission.Checked) Then
                frm.setTypeDonnees(vini_DB.vncTypeDonnee.FACTCOMM_NONREGLEE)
            End If
            If (rbFactColisage.Checked) Then
                frm.setTypeDonnees(vini_DB.vncTypeDonnee.FACTCOL_NONREGLEE)
            End If
            If (rbFactureTransport.Checked) Then
                frm.setTypeDonnees(vini_DB.vncTypeDonnee.FACTTRP_NONREGLEE)
            End If
            frm.ShowDialog(Me)
            If frm.DialogResult = Windows.Forms.DialogResult.OK Then
                obj = frm.getElementSelectionne()
                tbNumFactCom.Text = obj.code
                afficheInfosFacture()
            End If
        Catch ex As Exception

        End Try


    End Sub


    Private Sub tbNumFactCom_Validated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbNumFactCom.Validated
        afficheInfosFacture()
    End Sub

    Private Sub afficheInfosFacture()
        Try

            If Not String.IsNullOrEmpty(tbNumFactCom.Text) Then
                If (rbFactCommission.Checked) Then
                    m_oFacture = FactCom.createandload(tbNumFactCom.Text)
                End If
                If (rbFactColisage.Checked) Then
                    m_oFacture = FactColisage.createandload(tbNumFactCom.Text)
                End If
                If (rbFactureTransport.Checked) Then
                    m_oFacture = FactTRP.createandload(tbNumFactCom.Text)
                End If

                If (m_oFacture.id <> 0) Then
                    m_oFacture.oTiers.load()
                    laTiers.Text = m_oFacture.oTiers.nom
                    laMontantFacture.Text = m_oFacture.totalTTC.ToString("c")
                    If (m_oFacture.getSolde <= 0) Then
                        MessageBox.Show("Facture déjà soldée")
                    End If

                    tbSolde.Text = m_oFacture.getSolde().ToString("c")
                    EnableControls(True)
                    tbMontant.Focus()
                End If
            End If
        Catch

        End Try

    End Sub
    Private Sub tbMontant_Validated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbMontant.Validated
        Dim oMontant As Decimal
        Try
            oMontant = CType(tbMontant.Text, Decimal)
            tbSolde.Text = (m_oFacture.getSolde() - oMontant).ToString("c")
        Catch ex As Exception

        End Try
    End Sub

    Private Sub cbAppliquer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAppliquer.Click

        m_oReglement = New Reglement
        m_oReglement.IdFact = m_oFacture.id
        m_oReglement.TypeFact = m_oFacture.typeDonnee
        m_oReglement.DateReglement = dtDateReglement.Value.ToShortDateString
        m_oReglement.Montant = tbMontant.Text
        m_oReglement.Reference = tbReferenceReglement.Text
        m_oReglement.Commentaire = tbCommentaire.Text

        m_oReglement.save()

        m_oReglement = Nothing
        m_oFacture = Nothing
        laTiers.Text = String.Empty
        laMontantFacture.Text = String.Empty
        tbNumFactCom.Text = String.Empty
        tbReferenceReglement.Text = String.Empty
        tbCommentaire.Text = String.Empty
        tbMontant.Text = 0
        tbSolde.Text = 0
        tbNumFactCom.Focus()
    End Sub

    Private Sub tbNumFactCom_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbNumFactCom.TextChanged

    End Sub

    Private Sub cbCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbCancel.Click
        Me.Close()
    End Sub
End Class