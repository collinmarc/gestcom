Imports vini_DB
Public Class frmRapproFactFourn
    Inherits vini_app.frmGestionSCMD

#Region " Code généré par le Concepteur Windows Form "

    Public Sub New()
        MyBase.New()

        'Cet appel est requis par le Concepteur Windows Form.
        InitializeComponent()

        'Ajoutez une initialisation quelconque après l'appel InitializeComponent()

    End Sub

    'La méthode substituée Dispose du formulaire pour nettoyer la liste des composants.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Requis par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer

    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Elle peut être modifiée en utilisant le Concepteur Windows Form.  
    'Ne la modifiez pas en utilisant l'éditeur de code.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.SuspendLayout()
        '
        'cbAfficher
        '
        '
        'frmRapproFactFourn
        '
        Me.Name = "frmRapproFactFourn"
        Me.Text = "Rapprochement des factures fournisseurs"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

#Region "Méthodes Redefinies"
    Public Overrides Function getResume() As String
        Return "Rapprochement des sous-commandes avec les factures fournisseurs"
    End Function
    Protected Overrides Sub EnableControls(ByVal bEnabled As Boolean)
        MyBase.EnableControls(bEnabled)
        ckTransmiseFax.Enabled = False
        cbTransmettreFax.Enabled = False
        ckExporteeInternet.Enabled = False
        cbExportInternet.Enabled = False
        ckFacturee.Enabled = False
        tbCMDTotalHT.Enabled = False
        tbCMDtotalTTC.Enabled = False
        tbComfournisseur.Enabled = False
    End Sub

    Protected Overrides Function setListeSousCommandes() As Boolean
        Dim ddeb As Date
        Dim dfin As Date
        Dim codeFourn As String
        Dim col As New Collection
        Dim bReturn As Boolean
        Dim nId As Integer
        Dim oScmd As SousCommande
        Dim strCode As String
        Try
            If Not getElementCourant() Is Nothing Then
                If getElementCourant().bUpdated Then
                    If MsgBox("La sous commande courante n'a pas été sauvegardée, voulez-vous conservez vos modifications", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                        appliqueModifications()
                    End If
                End If
            End If

            'Chargement de la liste de sous commandes en fonction des critères
            If rbDateSCMD.Checked Then
                ddeb = dtDatedeb.Value.ToShortDateString
                dfin = dtdateFin.Value.ToShortDateString
                codeFourn = tbCodeFournisseur.Text
                col = SousCommande.getListeTransmises(ddeb, dfin, codeFourn)
            End If
            If rbIDScmd.Checked Then
                nId = Integer.Parse(Me.tbIDScmd.Text)
                oSCMD = SousCommande.createandload(nId)
                If oSCMD.id = nId Then
                    col.Add(oSCMD)
                End If
            End If
            If rbCodeScmd.Checked Then
                strCode = Me.tbCodeScmd.Text
                oSCMD = SousCommande.createandload(strCode)
                If oSCMD.code.Equals(strCode) Then
                    col.Add(oSCMD)
                End If
            End If
            m_colSousCommandes = col
            bReturn = True
        Catch ex As Exception
            bReturn = False
            Debug.Assert(bReturn, ex.ToString)
        End Try

        Return bReturn
    End Function
    Protected Overrides Function selectionneSousCommande() As Boolean
        Dim BReturn As Boolean
        BReturn = MyBase.selectionneSousCommande()
        ckRapprochee.Focus()
        Return BReturn
    End Function


#End Region


End Class
