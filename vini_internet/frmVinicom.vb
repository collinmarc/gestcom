Imports vini_DB
Imports System.Windows.Forms.Cursors

Public Class FrmVinicom
    Inherits System.Windows.Forms.Form

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

    Protected Sub debAffiche()
        m_AffichageEnCours = m_AffichageEnCours + 1
    End Sub
    Protected Sub finAffiche()
        m_AffichageEnCours = m_AffichageEnCours - 1
    End Sub
    Protected Function bAffichageEnCours() As Boolean
        Return m_AffichageEnCours > 0
    End Function


#Region " Code généré par le Concepteur Windows Form "

    Public Sub New()
        MyBase.New()

        'Cet appel est requis par le Concepteur Windows Form.
        InitializeComponent()

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
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmVinicom))
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.SuspendLayout()
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Magenta
        Me.ImageList1.Images.SetKeyName(0, "")
        Me.ImageList1.Images.SetKeyName(1, "")
        '
        'FrmVinicom
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(712, 342)
        Me.Name = "FrmVinicom"
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub

#End Region
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
        MsgBox("New")
    End Function
    Protected Overridable Function frmLoad() As Boolean
        MsgBox("Load")
    End Function
    Protected Overridable Function frmSave() As Boolean
        MsgBox("Save")
    End Function
    Protected Overridable Function frmDel() As Boolean
        MsgBox("Del")
    End Function
    Protected Overridable Function frmRefresh() As Boolean
        MsgBox("Resfresh")
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


    Public Overridable Function getResume() As String
        Return "VNC"
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
    End Sub
    'Affiche la fenêtre Fournisseur si le tag du lien est renseigné
    Protected Sub afficheFenetreFournisseur(ByVal tag As String)
    End Sub 'afficheFenetreFournisseur
    'Affiche la fenêtre Produit si le tag du lien est renseigné
    Protected Sub afficheFenetreProduit(ByVal tag As String)

    End Sub
    'Affiche la fenêtre Commande Client si le tag du lien est renseigné
    Protected Sub ControlUpdated(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not bAffichageEnCours() Then
            setfrmUpdated()
        End If
    End Sub

    Private Sub frm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Ajoutez une initialisation quelconque après l'appel InitializeComponent()
        m_action = vncfrmAction.FRMNOTHING
        m_ErrorMessage = ""
        m_StatusMessage = ""
        AddHandlerValidated(Controls)
        EnableControls(False)
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

    Protected Function isfrmUpdated() As Boolean
        Return m_frmUpdated
    End Function
    Protected Overridable Sub AddHandlerValidated(ByVal ocol As System.Windows.Forms.Control.ControlCollection)
        Dim objControl As Control
        Dim objRadio As RadioButton
        Dim objTextBox As TextBox
        Dim objCheck As CheckBox
        Dim objListBox As ListBox
        Dim objComboBox As ComboBox
        Dim objTabPage As TabPage
        Dim objSSTAB As TabControl
        Dim objgrp As GroupBox
        Dim objLabel As Label


        For Each objControl In ocol
            Debug.Write(objControl.GetType().Name & "|")
            Select Case objControl.GetType().Name
                Case "TextBox"
                    objTextBox = CType(objControl, TextBox)
                    AddHandler objTextBox.Validated, AddressOf ControlUpdated
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
                    objTabPage = CType(objControl, TabPage)
                    AddHandler objTabPage.Click, AddressOf ControlUpdated
                Case "TabControl"
                    objSSTAB = CType(objControl, TabControl)
                Case "GroupBox"
                    objgrp = CType(objControl, GroupBox)
                Case "Label"
                    objLabel = CType(objControl, Label)
                Case "AxMSFlexGrid"
                    Debug.Write("Pas de Handler pour les FlexGrid")

                Case Else
                    AddHandler objControl.TextChanged, AddressOf ControlUpdated
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

    Private Sub ToolBar1_ButtonClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs)
        Select Case UCase(e.Button.Tag.ToString())
            Case "EXPORT"

        End Select
    End Sub
End Class
