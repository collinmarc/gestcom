Imports vini_DB
Imports System.Windows.Forms.Cursors
Imports CrystalDecisions.Windows.Forms
Imports CrystalDecisions.CrystalReports.Engine

Public Class FrmVinicom
    Inherits System.Windows.Forms.Form

    Public bTesting As Boolean = False
    Private oldCursor As Cursor
    Protected m_action As vncfrmAction
    'Private m_frmMain As frmMain
    'Public m_title As String
    Private m_AffichageEnCours As Integer = 0
    Private m_frmUpdated As Boolean
    Public m_ToolBarNewEnabled As Boolean = True
    Public m_ToolBarLoadEnabled As Boolean = True
    Public m_ToolBarSaveEnabled As Boolean = True
    Public m_ToolBarDelEnabled As Boolean = True
    Public m_ToolBarRefreshEnabled As Boolean = True
    Private m_ErrorMessage As String
    Private m_StatusMessage As String
    Public Event evt_ToolBarUpdated()
    Public Event evt_DisplayError()
    Public Event evt_DisplayStatus()


    Protected Sub DisposeCr(pcrw As CrystalReportViewer)
        Try

            If pcrw.ReportSource IsNot Nothing Then
                Dim oReport As ReportDocument
                oReport = pcrw.ReportSource
                If oReport IsNot Nothing Then
                    oReport.Dispose()
                End If
                pcrw.ReportSource = Nothing

            End If
            pcrw.Dispose()
        Catch ex As Exception

        End Try
    End Sub

    Protected Sub debAffiche()
        m_AffichageEnCours = m_AffichageEnCours + 1
    End Sub
    Protected Sub finAffiche()
        m_AffichageEnCours = m_AffichageEnCours - 1
    End Sub
    Protected Function bAffichageEnCours() As Boolean
        Return m_AffichageEnCours > 0
    End Function


    Protected Overridable Sub setToolbarButtons()
        m_ToolBarNewEnabled = True
        m_ToolBarLoadEnabled = True
        If isfrmUpdated() Then
            m_ToolBarSaveEnabled = True
        Else
            m_ToolBarSaveEnabled = False
        End If
        m_ToolBarDelEnabled = True
        m_ToolBarRefreshEnabled = True

    End Sub
    Private Sub setToolbarButtons_vinicom()
        setToolbarButtons()
        RaiseEvent evt_ToolBarUpdated()
    End Sub
    Public Function frmvncNew() As Boolean
        Trace.WriteLine(Me.Text & ":frmNew")

        m_action = vncEnums.vncfrmAction.FRMNEW
        frmNew()
        setToolbarButtons_vinicom()
    End Function
    Public Function frmvncLoad() As Boolean
        Trace.WriteLine(Me.Text & ":frmLoad")

        m_action = vncEnums.vncfrmAction.FRMLOAD
        frmLoad()
        setToolbarButtons_vinicom()
    End Function
    Public Function frmvncSave() As Boolean
        Trace.WriteLine(Me.Text & ":frmsave")

        m_action = vncEnums.vncfrmAction.FRMSAVE
        frmSave()
        setToolbarButtons_vinicom()
    End Function
    Public Function frmvncDel() As Boolean
        Trace.WriteLine(Me.Text & ":frmdel")

        m_action = vncEnums.vncfrmAction.FRMDEL
        frmDel()
        setToolbarButtons_vinicom()
    End Function
    Public Function frmvncRefresh() As Boolean
        Trace.WriteLine(Me.Text & ":frmRefresh")

        m_action = vncEnums.vncfrmAction.FRMREFRESH
        frmRefresh()
        setToolbarButtons_vinicom()
    End Function
    Protected Overridable Function frmNew() As Boolean
    End Function
    Protected Overridable Function frmLoad() As Boolean
    End Function
    Protected Overridable Function frmSave() As Boolean
    End Function
    Protected Overridable Function frmDel() As Boolean
    End Function
    Protected Overridable Function frmRefresh() As Boolean
    End Function
    '============================================================================
    'Function : DisplayError
    'Description : Affiche l'erreur dans la barre d'état + log + MsgBox si ERR_DEBUG = 1
    Public Sub DisplayError(ByVal Source As String, ByVal ErrorMessage As String)
        If Source = "" And ErrorMessage = "" Then
            m_ErrorMessage = ""
        Else
            m_ErrorMessage = "Error: " & Source & ":" & ErrorMessage
        End If
        RaiseEvent evt_DisplayError()
    End Sub 'DisplayError
    Public Sub DisplayStatus(ByVal strMessage As String)
        m_StatusMessage = strMessage
        RaiseEvent evt_DisplayStatus()
    End Sub 'DisplayStatus
    Public Sub ClearError()
        DisplayError("", "")
    End Sub 'ClearError

    Private Sub FrmVinicom_Actived(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        Dim ofrmMain As frmMain
        Trace.WriteLine(Me.Text & ":Activated")

        If Not bTesting Then

            'Debug.Assert(MdiParent.GetType().Name().Equals("frmMain"))
            ofrmMain = MdiParent
            setToolbarButtons_vinicom()
            If ofrmMain IsNot Nothing Then
                ofrmMain.setFrmActive(Me)
            End If

            If GLOBALCONNECTION.Equals("Medium") Then
                Persist.shared_connect()
            End If
        End If

    End Sub

    Private Sub FrmVinicom_Desactivated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Deactivate
        Dim ofrmMain As frmMain
        Trace.WriteLine(Me.Text & ":Desactivated")

        If Not bTesting Then
            Debug.Assert(MdiParent.GetType().Name().Equals("frmMain"))
            ofrmMain = MdiParent
            ofrmMain.setFrmActive(Nothing)
            'Connection à la Base de Données
            'Si la valeur GlobalConnection = "Medium"
            'Connection en début de fenêtre à la base de données
            'Sinon Chaque opération nécéessite une Connection à la base de données
            If GLOBALCONNECTION.Equals("Medium") Then
                Persist.shared_disconnect()
            End If
        End If
    End Sub

    Public Overridable Function getResume() As String
        Return "VNC"
    End Function

    Protected Function rechercheDonnee(ByVal ptypeDonnes As vncEnums.vncTypeDonnee, ByVal ptbcode As TextBox, Optional ByVal ptbnom As TextBox = Nothing, Optional ByVal ptbmotcle As TextBox = Nothing) As racine
        Dim col As Collection = Nothing
        Dim objSel As racine = Nothing
        Dim frm As frmRechercheDB
        Dim objClient As Client
        Dim objFournisseur As Fournisseur
        Dim objProduit As Produit

        If ptbcode.Text <> "" Then
            Select Case ptypeDonnes
                Case vncEnums.vncTypeDonnee.CLIENT
                    col = Client.getListe(ptbcode.Text)
                Case vncEnums.vncTypeDonnee.FOURNISSEUR
                    col = Fournisseur.getListe(ptbcode.Text)
                Case vncEnums.vncTypeDonnee.PRODUIT
                    col = Produit.getListe(vncEnums.vncTypeProduit.vncTous, ptbcode.Text)
            End Select
        Else
            col = New Collection
        End If
        If col.Count <> 1 Then
            'Création de la fenêtre de recherche
            frm = New frmRechercheDB
            frm.setTypeDonnees(ptypeDonnes)
            frm.tbCode.Text = ptbcode.Text
            frm.setListe(col)
            frm.setCode(ptbcode.Text)

            If Not ptbnom Is Nothing Then
                frm.setNom(ptbnom.Text)
            End If

            If Not ptbmotcle Is Nothing Then
                frm.setMotCle(ptbmotcle.Text)
            End If

            frm.displayListe()
            'Affichage de la fenêtre
            If frm.ShowDialog() = Windows.Forms.DialogResult.OK Then
                'Si on sort par OK
                objSel = frm.getElementSelectionne()
                Select Case ptypeDonnes
                    Case vncEnums.vncTypeDonnee.CLIENT
                        objClient = frm.getElementSelectionne()
                        ptbcode.Text = objClient.code
                        If Not ptbnom Is Nothing Then
                            ptbnom.Text = objClient.nom
                        End If
                    Case vncEnums.vncTypeDonnee.FOURNISSEUR
                        objFournisseur = frm.getElementSelectionne()
                        ptbcode.Text = objFournisseur.code
                        If Not ptbnom Is Nothing Then
                            ptbnom.Text = objFournisseur.nom
                        End If
                    Case vncEnums.vncTypeDonnee.PRODUIT
                        objProduit = frm.getElementSelectionne()
                        ptbcode.Text = objProduit.code
                        If Not ptbnom Is Nothing Then
                            ptbnom.Text = objProduit.nom
                        End If
                End Select
            End If
            frm.Close()
            frm = Nothing
        Else
            objSel = col(1)
        End If
        Return objSel
    End Function
    Protected Sub setcursorWait()
        oldCursor = Me.Cursor
        Me.Cursor = WaitCursor
    End Sub
    Protected Sub restoreCursor()
        Me.Cursor = oldCursor
    End Sub
    'Affiche la fenêtre client si le tag du lien est renseigné
    Protected Sub afficheFenetreClient(ByVal tag As String)
        If Not IsNumeric(tag) Then
            Exit Sub
        End If

        Dim objfrmClient As frmClient
        Dim objClient As Client
        Dim nid As Long
        Dim bReturn As Boolean

        objfrmClient = New frmClient
        nid = tag
        objClient = New Client("", "")
        Try
            objClient.load(nid)
            If objClient.id <> 0 Then
                objfrmClient.MdiParent = Me.MdiParent
                objfrmClient.Show()
                bReturn = objfrmClient.setElementCourant2(objClient)
                If (bReturn) Then
                    objfrmClient.AfficheElementCourant()
                End If
            End If

        Catch ex As Exception
            DisplayError("frmSaisieCommande.linkClient", Client.getErreur())
        End Try
    End Sub
    'Affiche la fenêtre Fournisseur si le tag du lien est renseigné
    Protected Sub afficheFenetreFournisseur(ByVal tag As String)
        If Not IsNumeric(tag) Then
            Exit Sub
        End If

        Dim objfrmFournisseur As frmFournisseurTab
        Dim objFournisseur As Fournisseur
        Dim nid As Long
        Dim bReturn As Boolean

        objfrmFournisseur = New frmFournisseurTab
        nid = tag
        objFournisseur = New Fournisseur("", "")
        Try
            objFournisseur.load(nid)
            If objFournisseur.id <> 0 Then
                objfrmFournisseur.MdiParent = Me.MdiParent
                objfrmFournisseur.Show()
                bReturn = objfrmFournisseur.setElementCourant2(objFournisseur)
                If bReturn Then
                    objfrmFournisseur.AfficheElementCourant()
                End If
            End If

        Catch ex As Exception
            DisplayError("frmSaisieCommande.afficheFenetreFournisseur", Fournisseur.getErreur())
        End Try

    End Sub 'afficheFenetreFournisseur
    'Affiche la fenêtre Produit si le tag du lien est renseigné
    Protected Sub afficheFenetreProduit(ByVal tag As String)
        If Not IsNumeric(tag) Then
            Exit Sub
        End If
        Dim objfrmProduit As frmProduit
        Dim objProduit As Produit
        Dim nid As Long
        Dim bReturn As Boolean

        objfrmProduit = New frmProduit
        nid = tag
        Try
            objProduit = Produit.createandload(nid)
            If objProduit.idCouleur <> 0 Then
                objfrmProduit.MdiParent = Me.MdiParent
                objfrmProduit.Show()
                bReturn = objfrmProduit.setElementCourant2(objProduit)
                If bReturn Then
                    objfrmProduit.AfficheElementCourant()
                End If
            End If
        Catch ex As Exception
            DisplayError("frmSaisieCommande.linkProduit", Produit.getErreur())
        End Try

    End Sub
    'Affiche la fenêtre Commande Client si le tag du lien est renseigné
    Protected Sub afficheFenetreCommandeClient(ByVal tag As String)
        If Not IsNumeric(tag) Then
            Exit Sub
        End If

        Dim objfrmCommandeclient As frmCommandeClient
        Dim objcommandeClient As CommandeClient
        Dim nid As Long
        Dim bReturn As Boolean

        objfrmCommandeclient = New frmCommandeClient
        nid = tag
        objcommandeClient = New CommandeClient(New Client("", ""))
        Try
            objcommandeClient.load(nid)
            If objcommandeClient.id <> 0 Then
                objfrmCommandeclient.MdiParent = Me.MdiParent
                objfrmCommandeclient.Show()
                bReturn = objfrmCommandeclient.setElementCourant2(objcommandeClient)
                If bReturn Then
                    objfrmCommandeclient.AfficheElementCourant()
                End If
            End If
        Catch
            DisplayError("frmSaisieCommande.afficheFenetreFournisseur", CommandeClient.getErreur())
        End Try
    End Sub 'afficheFenetreCommandeClient

    Protected Sub afficheFenetreFactCom(ByVal tag As String)
        'Affiche la fenêtre Facture de commission si le tag du lien est renseigné
        If Not IsNumeric(tag) Then
            Exit Sub
        End If

        Dim objfrmFactcom As frmGestFactCom
        Dim objFactcom As FactCom
        Dim nid As Long
        Dim bReturn As Boolean

        objfrmFactcom = New frmGestFactCom
        nid = tag
        Try
            objFactcom = FactCom.createandload(nid)
            If objFactcom.id <> 0 Then
                objfrmFactcom.MdiParent = Me.MdiParent
                objfrmFactcom.Show()
                bReturn = objfrmFactcom.setElementCourant2(objFactcom)
                If (bReturn) Then
                    objfrmFactcom.AfficheElementCourant()
                End If
            End If
        Catch
            DisplayError("frmVinicom.afficheFenetreFactCom", FactCom.getErreur())
        End Try
    End Sub 'afficheFenetreFactCom
    Protected Sub afficheFenetreFactTrp(ByVal tag As String)
        'Affiche la fenêtre Facture de commission si le tag du lien est renseigné
        If Not IsNumeric(tag) Then
            Exit Sub
        End If

        Dim objfrmFactcom As frmGestFactTRP
        Dim objFacture As FactTRP
        Dim nid As Long
        Dim bReturn As Boolean

        objfrmFactcom = New frmGestFactTRP
        nid = tag
        Try
            objFacture = FactTRP.createandload(nid)
            If objFacture.id <> 0 Then
                objfrmFactcom.MdiParent = Me.MdiParent
                objfrmFactcom.Show()
                bReturn = objfrmFactcom.setElementCourant2(objFacture)
                If (bReturn) Then
                    objfrmFactcom.AfficheElementCourant()
                End If
            End If
        Catch
            DisplayError("frmVinicom.afficheFenetreFactTRP", FactTRP.getErreur())
        End Try
    End Sub 'afficheFenetreFactTRP
    Protected Sub afficheFenetreFactColisage(ByVal tag As String)
        'Affiche la fenêtre Facture de Colisage si le tag du lien est renseigné
        If Not IsNumeric(tag) Then
            Exit Sub
        End If

        Dim objfrmFact As frmGestFactColisage
        Dim objFacture As FactColisageJ
        Dim nid As Long
        Dim bReturn As Boolean

        objfrmFact = New frmGestFactColisage
        nid = tag
        Try
            objFacture = FactColisageJ.createandload(nid)
            If objFacture.id <> 0 Then
                objfrmFact.MdiParent = Me.MdiParent
                objfrmFact.Show()
                bReturn = objfrmFact.setElementCourant2(objFacture)
                If (bReturn) Then
                    objfrmFact.AfficheElementCourant()
                End If
            End If
        Catch
            DisplayError("frmVinicom.afficheFenetreFactColisage", FactTRP.getErreur())
        End Try
    End Sub 'afficheFenetreFactCol
    Protected Sub ControlUpdated(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not bAffichageEnCours() Then
            setfrmUpdated()
        End If
    End Sub

    Private Sub frm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Not DesignMode Then
            If Not CONSTANTES_LOADED Then
                initConstantes()
                Param.LoadcolParams()
            End If
            m_action = vncfrmAction.FRMNOTHING
            m_ErrorMessage = ""
            m_StatusMessage = ""
            If Not bTesting Then
                AddHandlerValidated(Controls)
            End If
            If Not bTesting Then
                EnableControls(False)
            End If
        End If
        Me.WindowState = FormWindowState.Maximized
    End Sub
    Protected Overridable Sub setfrmUpdated()
        m_frmUpdated = True
        Text = getResume() & "*"
        DisplayStatus("Modifié")
        setToolbarButtons_vinicom()
    End Sub
    Protected Sub setfrmNotUpdated()
        'Debug.Assert(Not m_ElementCourant Is Nothing)
        m_frmUpdated = False
        Text = getResume()
        DisplayStatus("")
        setToolbarButtons_vinicom()
    End Sub

    Protected Overridable Function isfrmUpdated() As Boolean
        Return m_frmUpdated
    End Function
    Protected Overridable Sub AddHandlerValidated(ByVal ocol As System.Windows.Forms.Control.ControlCollection)
        Dim objControl As Control
        Dim objRadio As RadioButton
        Dim objCheck As CheckBox
        Dim objListBox As ListBox
        Dim objComboBox As ComboBox


        For Each objControl In ocol
            Select Case objControl.GetType().Name
                Case "RadioButton"
                    objRadio = CType(objControl, RadioButton)
                    AddHandler objRadio.CheckedChanged, AddressOf ControlUpdated
                Case "CheckBox"
                    objCheck = CType(objControl, CheckBox)
                    AddHandler objCheck.CheckedChanged, AddressOf ControlUpdated
                Case "ListBox"
                    objListBox = CType(objControl, ListBox)
                    AddHandler objListBox.SelectedValueChanged, AddressOf ControlUpdated
                Case "ComboBox"
                    objComboBox = CType(objControl, ComboBox)
                    AddHandler objComboBox.SelectedValueChanged, AddressOf ControlUpdated
                Case "TabPage"
                Case "TabControl"
                Case "GroupBox"
                Case "Label"
                Case "LinkLabel"
                Case "AxMSFlexGrid"
                Case "PageView"
                Case "ReportGroupTree"
                Case "TotallerTreeView"
                Case "CrystalReportViewer"
                Case "Splitter"
                Case "ToolStrip"
                    'Pas de Handler par defaut pour Les Crystal Report viewer


                Case Else
                    AddHandler objControl.Validated, AddressOf ControlUpdated
            End Select
            AddHandlerValidated(objControl.Controls)
        Next

    End Sub
    Protected Overridable Sub EnableControls(ByVal bEnabled As Boolean)
        'Procédure à écrire dans la fenêtre fille
        Dim objControl As Control
        For Each objControl In Me.Controls
            objControl.Enabled = bEnabled
        Next
    End Sub

    Public Function getErrrorMessage() As String
        Return m_ErrorMessage
    End Function
    Public Function getStatusMessage() As String
        Return m_StatusMessage
    End Function

    Private Sub FrmVinicom_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        disposeReport(Me)
    End Sub

    Private Sub disposeReport(ByVal pControl As Control)
        Dim objControl As Control
        For Each objControl In pControl.Controls
            Dim objcr As CrystalDecisions.Windows.Forms.CrystalReportViewer

            If TypeOf objControl Is CrystalDecisions.Windows.Forms.CrystalReportViewer Then
                objcr = CType(objControl, CrystalDecisions.Windows.Forms.CrystalReportViewer)
                If Not objcr.ReportSource Is Nothing Then
                    objcr.ReportSource.Dispose()
                End If
            Else
                disposeReport(objControl)
            End If
        Next
    End Sub

End Class
