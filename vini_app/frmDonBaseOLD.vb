Imports vini_DB
Public Class frmDonBase
    Inherits FrmVinicom

    Protected m_ElementCourant As Persist
    Protected m_TypeDonnees As vncTypeDonnee
    Protected m_BloquageElementCourant As Boolean = True

#Region " Code généré par le Concepteur Windows Form "

    Public Sub New()
        MyBase.New()

        'Cet appel est requis par le Concepteur Windows Form.
        InitializeComponent()

        'Ajoutez une initialisation quelconque après l'appel InitializeComponent()
        If (Not DesignMode) Then
            setfrmNotUpdated()
        End If
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
        '
        'frmDonBase
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1016, 734)
        Me.Name = "frmDonBase"
        Me.Text = "frmDonBase"

    End Sub

#End Region

#Region "Méthodes ViniCom"
    Protected Overrides Sub setToolbarButtons()
        m_ToolBarNewEnabled = True
        m_ToolBarLoadEnabled = True
        If Not m_ElementCourant Is Nothing Then
            m_ToolBarSaveEnabled = True
            m_ToolBarDelEnabled = True
        Else
            m_ToolBarSaveEnabled = False
            m_ToolBarDelEnabled = False
        End If
        m_ToolBarRefreshEnabled = True
    End Sub
    Protected Overrides Function frmNew() As Boolean
        Trace.WriteLine(Now() & Me.Text & "frmNew")
        Dim bReturn As Boolean

        If isFrmUpdated Then
            If MsgBox("L'élement courant a été modifié, Voulez-vous le sauvegarder ", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                frmSave()
            Else
                setfrmNotUpdated()
            End If
        End If

        bReturn = creerElement()
        If bReturn Then
            AfficheElementCourant()
            setfrmUpdated()
        End If
        Trace.WriteLine(Now() & Me.Text & "frmNew End" & bReturn)
        Return bReturn
    End Function
    'Les fenêtres filles doivent implémenter cette méthode qui fait un
    '        setElementCourant(New Client("", ""))

    Protected Overridable Function creerElement() As Boolean
        Debug.Assert(False, "Fonction non Implémentée")
        Return False
    End Function

    Protected Overrides Function frmLoad() As Boolean
        Trace.WriteLine(Now() & Me.Text & "frmLoad")
        '=======================================================================
        'Fonction : frmLoad()
        'Description : Chargement de l'élément à Partir de la fenêtre de recherche
        'Détails    :  
        'Retour : 
        '=======================================================================
        Dim frm As frmRechercheDB
        Dim bReturn As Boolean
        Dim oElement As Persist
        'Création de la fenêtre de recherche
        frm = New frmRechercheDB
        frm.setTypeDonnees(m_TypeDonnees)
        'Affichage de la fenêtre
        If frm.ShowDialog() = Windows.Forms.DialogResult.OK Then
            'Si on sort par OK
            oElement = frm.getElementSelectionne()
            DisplayStatus("Chargement de l'élément en cours")
            bReturn = oElement.load()
            DisplayStatus("Chargement Terminée")
            Debug.Assert(bReturn, "frmDonBase.frmLoad : Load Failed :" & Persist.getErreur())

            If isfrmUpdated() Then
                If MsgBox("L'élement courant a été modifié, Voulez-vous le sauvegarder ", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    frmSave()
                Else
                    setfrmNotUpdated()
                End If
            End If

            If Not oElement Is Nothing Then
                DisplayStatus("Positionnement de l'élement courant")
                If setElementCourant2(oElement) Then
                    DisplayStatus("Affichage de l'élement courant")
                    AfficheElementCourant()
                End If
            End If
            setfrmNotUpdated()
        Else
            bReturn = False
        End If
        DisplayStatus("")
        Trace.WriteLine(Now() & Me.Text & "frmLoad End" & bReturn)
        Return bReturn
    End Function
    Protected Overrides Function frmSave() As Boolean
        '=======================================================================
        'Fonction : frmSave()
        'Description : Sauvegarde de l'élément courant
        'Détails    :  
        'Retour : 
        '=======================================================================
        Trace.WriteLine(Now() & Me.Text & "frmSave")
        Dim bReturn As Boolean
        Dim bNouvelElement As Boolean
        bReturn = False

        Validate()
        setcursorWait()
        If Not m_ElementCourant Is Nothing Then
            If ControleAvantSauvegarde() Then
                If MAJElementCourant() Then
                    bNouvelElement = m_ElementCourant.bNew
                    If SauveElement() Then
                        If (bNouvelElement) Then
                            ' Bloquage du nouvel element crée (ID renseigné)
                            Bloque(m_ElementCourant)
                        End If
                        Text = getResume()
                        setfrmNotUpdated()
                        bReturn = True
                    Else
                        bReturn = False
                    End If
                End If
            End If
        End If
        restoreCursor()
        Trace.WriteLine(Now() & Me.Text & "frmSave End " & bReturn)
        Return bReturn
    End Function
    Protected Overrides Function frmRefresh() As Boolean
        Trace.WriteLine(Now() & Me.Text & "frmRefresh ")
        Dim objRacine As Persist
        objRacine = getElementCourant()
        If Not objRacine Is Nothing Then
            objRacine.load()
            AfficheElementCourant()
            setfrmNotUpdated()
        End If
        Trace.WriteLine(Now() & Me.Text & "frmRefresh End")
    End Function ' frmRefresh
    Protected Overrides Function frmDel() As Boolean
        Trace.WriteLine(Now() & Me.Text & "frmDel")
        If Not getElementCourant() Is Nothing Then
            Persist.shared_connect()
            If getElementCourant.checkForDelete() Then
                If MsgBox("Etes-vous sur de vouloir supprimer cet élement", MsgBoxStyle.YesNo) = vbYes Then
                    deleteElementCourant()
                    m_ElementCourant = Nothing
                    EnableControls(False)
                    Persist.shared_disconnect()
                End If

            Else
                MsgBox("Cet element n'est pas supprimable, car il a des éléments rattachés")
            End If
        End If
        Trace.WriteLine(Now() & Me.Text & "frmRefresh End")
    End Function ' frmDel
#End Region
    Public Overridable Sub deleteElementCourant()
        Debug.Assert(Not m_ElementCourant Is Nothing)
        Dim obj As Persist
        obj = getElementCourant()
        obj.bDeleted = True
        If SauveElement() Then
            EnableControls(False)
            setfrmNotUpdated()
        End If
    End Sub 'deleteElementCourant
    ''' <summary>
    ''' Renseigne l'élément courant en véfifiant s'il est libre d'utilisation
    ''' </summary>
    ''' <param name="pElement">L'element à traiter</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Function setElementCourant2(ByVal pElement As Persist) As Boolean
        Dim bReturn As Boolean
        Debug.Assert(Not isfrmUpdated(), "La fenêtre a été modifiée")
        Try

            If EstLibre(pElement) Then
                ' Libère l'élement précédent
                Libere(m_ElementCourant)
                'Bloque le nouvel element
                bReturn = Bloque(pElement)
                If bReturn Then
                    m_ElementCourant = pElement
                End If
            Else
                bReturn = False
            End If
        Catch ex As Exception
            DisplayError(System.Environment.StackTrace, ex.Message)
            bReturn = False
        End Try

        Return bReturn
    End Function 'setElementCourant2

    Private Function estIdentique_a_elementcourant(ByVal pElement As Persist) As Boolean
        Dim bReturn As Boolean
        bReturn = False
        If pElement Is Nothing Or m_ElementCourant Is Nothing Then
            ' l'un des deux est nothing
            If pElement Is Nothing And m_ElementCourant Is Nothing Then
                'Les deux sont nothing
                bReturn = True
            Else
                bReturn = False
            End If
        Else
            'Ni l'un ni l'autre ne sont nothing
            bReturn = (pElement.id = m_ElementCourant.id) And (pElement.typeDonnee = m_ElementCourant.typeDonnee)
        End If
        Return bReturn
    End Function
    Private Function EstLibre(ByVal pElement As Persist) As Boolean
        Dim bReturn As Boolean
        Try
            bReturn = True
            If (Not pElement Is Nothing) And m_BloquageElementCourant Then
                If (Not estIdentique_a_elementcourant(pElement)) Then
                    Dim ota As vini_DB.dsVinicomTableAdapters.LOCKTableAdapter
                    Dim odt As vini_DB.dsVinicom.LOCKDataTable
                    ota = New dsVinicomTableAdapters.LOCKTableAdapter
                    ota.Connection = Persist.oleDBConnection
                    odt = ota.GetDataByPersistID_dataType(pElement.id, pElement.typeDonnee)
                    If (odt.Count > 0) Then
                        MessageBox.Show("L'element est en cours d'utilisation par " & odt(0).LCK_NAME & " depuis " & odt(0).LCK_DATE)
                        bReturn = False
                    End If
                End If

            End If
        Catch ex As Exception
            DisplayError("frmDonBase.EstLibre", ex.Message)
            bReturn = False
        End Try
        Return bReturn
    End Function
    Private Function Bloque(ByVal pElement As Persist) As Boolean
        Dim bReturn As Boolean
        Try
            Dim ota As vini_DB.dsVinicomTableAdapters.LOCKTableAdapter
            If (Not pElement Is Nothing) Then
                If pElement.id <> 0 And m_BloquageElementCourant Then
                    ota = New dsVinicomTableAdapters.LOCKTableAdapter
                    ota.Connection = Persist.oleDBConnection
                    ota.Insert(pElement.id, pElement.typeDonnee, System.Environment.MachineName, System.DateTime.Now())
                End If
            End If
            bReturn = True
        Catch ex As Exception
            DisplayError("frmDonBase.Bloque", ex.Message)
            bReturn = False
        End Try
        Return bReturn
    End Function
    Private Function Libere(ByVal pElement As Persist) As Boolean
        Dim bReturn As Boolean
        Try
            Dim ota As vini_DB.dsVinicomTableAdapters.LOCKTableAdapter
            If (Not pElement Is Nothing) And m_BloquageElementCourant Then
                ota = New dsVinicomTableAdapters.LOCKTableAdapter
                ota.Connection = Persist.oleDBConnection
                ota.Delete(pElement.id, pElement.typeDonnee)
            End If
            bReturn = True
        Catch ex As Exception
            DisplayError("frmDonBase.Libere", ex.Message)
            bReturn = False
        End Try
        Return bReturn

    End Function

    Public Overridable Function getElementCourant() As Persist
        Return m_ElementCourant
    End Function
    Public Function AfficheElementCourant() As Boolean
        Dim bReturn As Boolean
        Try
            debAffiche()
            Debug.Assert(Not m_ElementCourant Is Nothing)

            bReturn = AfficheElement()
            If (bReturn) Then
                EnableControls(True)
                Text = getResume()
            End If
        Catch ex As Exception
            DisplayError("frmDonBase", ex.Message)
            bReturn = False
        Finally
            finAffiche()
        End Try
        Return bReturn

    End Function

    Public Overridable Function AfficheElement() As Boolean
        Return True
    End Function

    Public Function MAJElementCourant() As Boolean
        'On ne met pas a jour l'élément courant si on est en cours d'affichage
        Dim bReturn As Boolean
        Debug.Assert(Not m_ElementCourant Is Nothing)
        If Not bAffichageEnCours Then
            bReturn = MAJElement()
            Debug.Assert(bReturn, "MAJElement()")
            Text = getResume()
        End If
        Return bReturn
    End Function
    Public Overridable Function MAJElement() As Boolean
        Return True
    End Function
    Public Overridable Function SauveElement() As Boolean
        Dim bReturn As Boolean
        bReturn = m_ElementCourant.Save()
        Return bReturn
    End Function
    Public Overridable Function ControleAvantSauvegarde() As Boolean
        Return True
    End Function

    Public Overridable Function trtClosing() As Boolean
        Dim bReturn As Boolean
        Try
            bReturn = True
            If isfrmUpdated() Then
                If MsgBox("L'élément courant a été modifié, souhaitez-vous le sauvegarder?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    bReturn = frmSave()
                End If
            End If
            If (Not m_ElementCourant Is Nothing) Then
                Libere(m_ElementCourant)
            End If
        Catch ex As Exception
            DisplayError(System.Environment.StackTrace, ex.Message)
            bReturn = False
        End Try
        Return bReturn
    End Function

    Private Sub frmDonBase_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        trtClosing()
    End Sub

    '=========================================================================
    'Fonction : sauvegardeElementCourant
    'Description : Si la feneêtre est modifié , Pose la question de sauvegarder l'élement courant 
    '               Si oui Sauvegarde de l'élement
    '               Sinon Recherchement de l'élement courant
    ' Retourn    : True si la sauvegarde s'est bien déroulée
    '========================================================================
    Protected Overridable Function sauvegardeElementCourant() As Boolean
        Dim bReturn As Boolean
        bReturn = False
        If isfrmUpdated Then
            If MsgBox("Voulez-vous sauvegarder l'élement courant", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                bReturn = frmSave()
            Else
                setfrmNotUpdated()
                If Not getElementCourant() Is Nothing Then
                    If Not getElementCourant.bNew Then
                        getElementCourant.load()
                        AfficheElementCourant()
                    End If

                End If
            End If
        End If
        Return bReturn
    End Function 'SauvegardeElementcourant



End Class
