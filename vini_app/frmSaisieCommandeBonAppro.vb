Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
'Imports FAXCOMLib
Imports System.Windows.Forms.Cursors
Imports vini_DB
Public Class frmSaisieCommandeBonAppro
    Inherits vini_app.frmSaisieCommande
    Protected Shadows Function getCommandeCourante() As BonAppro
        Return CType(getElementCourant(), BonAppro)
    End Function
#Region " Code généré par le Concepteur Windows Form "


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
        components = New System.ComponentModel.Container
    End Sub

#End Region
    Public Sub New()
        MyBase.New()

        'Cet appel est requis par le Concepteur Windows Form.
        InitializeComponent()

        'Ajoutez une initialisation quelconque après l'appel InitializeComponent()

    End Sub

#Region "Méthodes Vinicom"
    Protected Overrides Sub setToolbarButtons()
    End Sub
    Protected Overrides Function creerElement() As Boolean
        Debug.Assert(Not isfrmUpdated, "La fenetre n'est pas libre")
        Dim bReturn As Boolean
        bReturn = setElementCourant2(New BonAppro(New Fournisseur))

        Return bReturn
    End Function

    Public Overrides Function setElementCourant2(ByVal pElement As Persist) As Boolean
        Debug.Assert(pElement.GetType().Name.Equals("BonAppro"))
        Return MyBase.setElementCourant2(pElement)
    End Function 'SetElement
    Public Overrides Function getResume() As String 'Rend le caption de la fenêtre
        If getCommandeCourante() Is Nothing Then
            Return "Gestion des Bons d'approvisionnement"
        Else
            Return "Bon Appro : " & getCommandeCourante.shortResume
        End If
    End Function 'getResume
    Public Overrides Function SauveElement() As Boolean
        Debug.Assert(Not getCommandeCourante() Is Nothing)
        Dim bReturn As Boolean
        bReturn = getCommandeCourante.save
        If bReturn Then
            tbCode.Text = getCommandeCourante.code
            tbCode.Enabled = True
            getResume()
        End If
        Debug.Assert(bReturn, "Erreur en sauvegarde")
        Return bReturn
    End Function
#End Region
#Region "Méthodes SaisieDeCommandes"
    Protected Overrides Sub ActionLivrer()
        getCommandeCourante().changeEtat(vncActionEtatCommande.vncActionBALivrer)
    End Sub
    Protected Overrides Sub ActionAnnLivrer()
        getCommandeCourante().changeEtat(vncActionEtatCommande.vncActionBAAnnLivrer)
    End Sub
    Protected Overrides Sub modifieruneLigneBL()
        If getCommandeCourante().etat.codeEtat = vncEtatCommande.vncBAEnCours Or _
            (getCommandeCourante().etat.codeEtat = vncEtatCommande.vncBALivre And getCommandeCourante().etat.actionMvtStock = vncGenererSupprimer.vncGenerer) Then
            MyBase.modifieruneLigneBL()
        Else
            MessageBox.Show(My.Resources.STR_LIGNE_NON_MODIFIABLE)
        End If
    End Sub
    Protected Overrides Sub valideCommande()
    End Sub

    Protected Overrides Function getbGestionMvtStock() As Boolean
        Return True
    End Function
    Protected Overrides Function getbGestionSousCommande() As Boolean
        Return True
    End Function

    Protected Overrides Sub blToutOK()
        '======================================================================================
        'Fonction : blToOK
        'Description : Valide la Livraison la Livraison => MAJ de la Qte livrée, de la date de livraison, changement d'état , Créationdes  des mvts de stocks à la prochaine Sauvegarde
        'Méthode refefinie car l'action a exercer sur l'état est différente selon le type de commande
        '======================================================================================
        'Méthode refefinie car l'action a exercer sur l'état est différente selon le type de commande
        Debug.Assert(Not getCommandeCourante() Is Nothing)
        Dim oLgCom As LgCommande
        Dim bReturn As Boolean

        For Each oLgCom In getCommandeCourante.colLignes
            If oLgCom.bResume Then
                bReturn = oLgCom.load()
                Debug.Assert(bReturn, LgCommande.getErreur())
            End If
            oLgCom.qteLiv = oLgCom.qteCommande
        Next oLgCom
        If getCommandeCourante.estEntierementLivree() Then
            ActionLivrer()
        End If
        AfficheEtat()
        affichecolLignes()
        setfrmUpdated()
    End Sub 'BLToutOK
    Protected Overrides Sub annulerLivraison()
        '======================================================================================
        'Fonction : AnnulerLivraison
        'Description : Annule la Livraison => changement d'état , suppression des mvt de stocks à la prochaine Sauvegarde
        'Méthode refefinie car l'action a exercer sur l'état est différente selon le type de commande
        '======================================================================================
        Dim oLgCom As LgCommande
        Debug.Assert(Not getCommandeCourante() Is Nothing)
        If getCommandeCourante.etat.codeEtat = vncEnums.vncEtatCommande.vncBALivre Then
            For Each oLgCom In getCommandeCourante.colLignes
                oLgCom.qteLiv = 0
            Next oLgCom
            ActionAnnLivrer()
            AfficheEtat()
            affichecolLignes()
        End If
        frmSave()
    End Sub
    Protected Overrides Function affichecrDetailCommande() As Boolean
        'Affichage de l'état 'Détail de Bon Appro'
        Dim objReport As ReportDocument
        Dim r1 As New crBonAppro()
        Dim strReportName As String = r1.ResourceName

        objReport = New ReportDocument
        objReport.Load(PATHTOREPORTS & strReportName)
        setcrDetailCommandeClientParameters(objReport)
        Persist.setReportConnection(objReport)
        crwDetailCommandeClient.ReportSource = objReport
        crwDetailCommandeClient.Zoom(1)
        Return True
    End Function 'Affichage de l'état de détail du bon Appro 
    Protected Overrides Function setcrDetailCommandeClientParameters(ByVal objReport As ReportDocument) As Boolean
        'Passe les paramètres à l'état Crystal Report
        objReport.SetParameterValue("IDAppro", getCommandeCourante.id)

    End Function 'Passe les paramètres à l'état Crystal Report
    Protected Overrides Function faxerValidationCommande() As Boolean
        'Dim diskOpts As New CrystalDecisions.Shared.DiskFileDestinationOptions
        'Dim objFax As clsFax
        'Dim objReport As ReportDocument

        'If Not crwDetailCommandeClient.ReportSource Is Nothing Then

        '    objReport = crwDetailCommandeClient.ReportSource
        '    objReport.ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.WordForWindows
        '    objReport.ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
        '    diskOpts.DiskFileName = FAX_DETAILCOMMANDE_PATH & "DETCMD_" & getCommandeCourante.code & ".doc"
        '    objReport.ExportOptions.DestinationOptions = diskOpts
        '    Try
        '        objReport.Export()
        '        objFax = New clsFax
        '        objFax.sendFax(FAX_NOM_INTERLOCUTEUR, FAX_TEL_INTERLOCUTEUR, FAX_BA_SUBJECT, FAX_BA_NOTES, FAX_DETAILCOMMANDE_BSENDCOVERPAGE, diskOpts.DiskFileName, tbNumFaxValidation.Text, getCommandeCourante.oTiers)
        '    Catch ex As Exception
        '        DisplayError("FaxerValidationBonAppro", "Erreur en Envoi de Fax" & ex.ToString)
        '    End Try
        'Else
        '    MsgBox("Vous devez d'abord visualiser la commande")
        'End If
        'Return True
    End Function 'Faxe la validation de commande
    Protected Overrides Function affichecrBL() As Boolean
        '===========================================================================
        'Function : AffichecrBL
        'Description : Affichage de l'état BL du Bon D'appro
        '===========================================================================
        Dim objReport As ReportDocument
        Dim r1 As New crBLBonAppro()
        Dim strReportName As String = r1.ResourceName
        objReport = New ReportDocument
        strReportName = PATHTOREPORTS & strReportName
        objReport.Load(strReportName)
        setcrBLParameters(objReport)
        Persist.setReportConnection(objReport)
        crwBL.ReportSource = objReport
        crwBL.Zoom(1)
        Return True
    End Function 'Affichage du Bon de Livraison
    Protected Sub setcrBLParameters(ByVal objReport As ReportDocument)
        '===========================================================================
        'Function : setcrBLParameters
        'Description : Passe les paramétres à L'état BLBonAppo
        '===========================================================================
        'Passe les paramètres à l'état Crystal Report
        objReport.SetParameterValue("IDAppro", getCommandeCourante.id)
    End Sub 'Passe les paramètres à l'état Crystal Report
    Protected Overrides Function faxerBL(ByVal pNumFax As String, Optional ByVal poTiers As Tiers = Nothing) As Boolean
        'Dim diskOpts As New CrystalDecisions.Shared.DiskFileDestinationOptions
        'Dim objFax As clsFax
        'Dim objReport As ReportDocument

        'If Not crwBL.ReportSource Is Nothing Then

        '    objReport = crwBL.ReportSource
        '    objReport.ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.WordForWindows
        '    objReport.ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
        '    diskOpts.DiskFileName = FAX_BLBA_PATH & "BLBA_" & getCommandeCourante.code & ".doc"
        '    objReport.ExportOptions.DestinationOptions = diskOpts
        '    Try
        '        objReport.Export()
        '        objFax = New clsFax
        '        objFax.sendFax(FAX_NOM_INTERLOCUTEUR, FAX_TEL_INTERLOCUTEUR, FAX_BLBA_SUBJECT, FAX_BLBA_NOTES, FAX_BLBA_BSENDCOVERPAGE, diskOpts.DiskFileName, pNumFax, poTiers)
        '    Catch ex As Exception
        '        DisplayError("FaxerBL", "Erreur en Envoi de Fax" & ex.ToString)
        '    End Try
        'Else
        '    MsgBox("Vous devez d'abord visualiser le Bon de Livraison")
        'End If
        'Return True
    End Function 'Faxe la Bon de Livraison au transporteur
#End Region
#Region "Méthodes privées"

#End Region

    Private Sub frmBonAppro_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        debAffiche()


        m_TypeDonnees = vncEnums.vncTypeDonnee.BA
        Me.Text = "Bon Approvisionnement"
        tpClient.Text = "Fournisseur"
        LaCodeTiers.Text = "Fournisseur"
        grpEntete.Text = "Bon Appro"
        SSTabCommandeClient.TabPages.Clear()
        SSTabCommandeClient.TabPages.Add(tpClient)
        SSTabCommandeClient.TabPages.Add(tpLignes)
        SSTabCommandeClient.TabPages.Add(tpCommentaires)
        SSTabCommandeClient.TabPages.Add(tpValidCmd)
        SSTabCommandeClient.TabPages.Add(tpBL)
        SSTabCommandeClient.TabPages.Add(tpRetourBL)
        grpTypeCommande.Enabled = False
        grpTypeCommande.Visible = False
        grpTypeTransport.Enabled = False
        grpTypeTransport.Visible = False
        grpFactTRP.Enabled = False
        grpFactTRP.Visible = False
        grpPrestashop.Enabled = False
        grpPrestashop.Visible = False
        'grpInfFactureTRP.Enabled = False
        'grpInfFactureTRP.Visible = False

        laMtComm.Visible = False
        tbMtComm.Visible = False

        finAffiche()
    End Sub

End Class
