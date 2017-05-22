' si ajout perte de ce type de code
' il faut ajouter dans InitializeComponent
'AddHandler MiAlgerian.Click, AddressOf Me.Font_Click
Imports System
Imports System.Collections
Imports System.ComponentModel
Imports System.Drawing
Imports System.data
Imports System.Drawing.Printing
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.Threading
Imports System.Text

<ToolboxBitmap(GetType(RichTextBoxJmb), "RichTextBoxJmb.bmp")> _
Public Class RichTextBoxJmb
    Inherits System.Windows.Forms.UserControl

    Private components As System.ComponentModel.IContainer

#Region "StampActions"
    Public Enum StampActions
        Editepar = 1
        DateHeure = 2
        Auchoix = 4
    End Enum 'StampActions
#End Region

#Region " Code généré par le Concepteur Windows Form "
    Public Sub New()
        ' This call is required by the Windows.Forms Form Designer.
        InitializeComponent()
        ' si je le mets dans InitializeComponent et que je modifie le menu
        ' je perds tout
        ' ceci permet de n'avoir qu'un évènement à traiter dans font_click, dans color_click etc..
        AddHandler miTimesNewRoman.Click, AddressOf Me.Font_Click
        AddHandler MiAlgerian.Click, AddressOf Me.Font_Click
        AddHandler miCourierNew.Click, AddressOf Me.Font_Click
        AddHandler miArial.Click, AddressOf Me.Font_Click
        AddHandler miTimesNewRoman.Click, AddressOf Me.Font_Click
        AddHandler miMicrosoftSansSerif.Click, AddressOf Me.Font_Click
        AddHandler miGaramond.Click, AddressOf Me.Font_Click
        AddHandler miTahoma.Click, AddressOf Me.Font_Click
        AddHandler miVerdana.Click, AddressOf Me.Font_Click
        ' fontSize

        AddHandler mi8.Click, AddressOf Me.FontSize_Click
        AddHandler mi9.Click, AddressOf Me.FontSize_Click
        AddHandler mi10.Click, AddressOf Me.FontSize_Click
        AddHandler mi11.Click, AddressOf Me.FontSize_Click
        AddHandler mi12.Click, AddressOf Me.FontSize_Click
        AddHandler mi14.Click, AddressOf Me.FontSize_Click
        AddHandler mi16.Click, AddressOf Me.FontSize_Click
        AddHandler mi18.Click, AddressOf Me.FontSize_Click
        AddHandler mi20.Click, AddressOf Me.FontSize_Click
        AddHandler mi22.Click, AddressOf Me.FontSize_Click
        AddHandler mi24.Click, AddressOf Me.FontSize_Click
        AddHandler mi26.Click, AddressOf Me.FontSize_Click
        AddHandler mi28.Click, AddressOf Me.FontSize_Click
        AddHandler mi36.Click, AddressOf Me.FontSize_Click
        AddHandler mi48.Click, AddressOf Me.FontSize_Click
        AddHandler mi72.Click, AddressOf Me.FontSize_Click
        ' couleurs
        AddHandler miBlack.Click, AddressOf Me.Color_Click
        AddHandler miBlue.Click, AddressOf Me.Color_Click
        AddHandler miRed.Click, AddressOf Me.Color_Click
        AddHandler miGreen.Click, AddressOf Me.Color_Click
        AddHandler Miblanc.Click, AddressOf Me.Color_Click
        AddHandler Mijaune.Click, AddressOf Me.Color_Click
        AddHandler Miorange.Click, AddressOf Me.Color_Click
        AddHandler MiMagenta.Click, AddressOf Me.Color_Click
        AddHandler migris.Click, AddressOf Me.Color_Click
        AddHandler mirose.Click, AddressOf Me.Color_Click
        AddHandler mibleupale.Click, AddressOf Me.Color_Click
        'Cet appel est requis par le Concepteur Windows Form.
        SCF_SELECTION = 1
        SCF_WORD = 2
        SCF_ALL = 4
        'Ajoutez une initialisation quelconque après l'appel InitializeComponent()

        'Update the graphics on the toolbar
        UpdateToolbar()

    End Sub 'New

    'La méthode substituée Dispose du formulaire pour nettoyer la liste des composants.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)

    End Sub
    Friend WithEvents Rtb1 As System.Windows.Forms.RichTextBox
    Friend WithEvents imgList1 As System.Windows.Forms.ImageList
    Friend WithEvents tb1 As System.Windows.Forms.ToolBar
    Friend WithEvents tbbSave As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbbOpen As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbbSeparator3 As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbbFont As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbbFontSize As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbbSeparator5 As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbbBold As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbbItalic As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbbUnderline As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbbStrikeout As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbbSeparator1 As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbbLeft As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbbCenter As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbbRight As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbbSeparator2 As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbbUndo As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbbRedo As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbbSeparator4 As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbbStamp As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbbColor As System.Windows.Forms.ToolBarButton
    Friend WithEvents cmColors As System.Windows.Forms.ContextMenu
    Friend WithEvents miBlack As System.Windows.Forms.MenuItem
    Friend WithEvents miBlue As System.Windows.Forms.MenuItem
    Friend WithEvents miRed As System.Windows.Forms.MenuItem
    Friend WithEvents miGreen As System.Windows.Forms.MenuItem
    Friend WithEvents ofd1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents sfd1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents cmFontSizes As System.Windows.Forms.ContextMenu
    Friend WithEvents mi8 As System.Windows.Forms.MenuItem
    Friend WithEvents mi9 As System.Windows.Forms.MenuItem
    Friend WithEvents mi10 As System.Windows.Forms.MenuItem
    Friend WithEvents mi11 As System.Windows.Forms.MenuItem
    Friend WithEvents mi12 As System.Windows.Forms.MenuItem
    Friend WithEvents mi14 As System.Windows.Forms.MenuItem
    Friend WithEvents mi16 As System.Windows.Forms.MenuItem
    Friend WithEvents mi18 As System.Windows.Forms.MenuItem
    Friend WithEvents mi20 As System.Windows.Forms.MenuItem
    Friend WithEvents mi22 As System.Windows.Forms.MenuItem
    Friend WithEvents mi24 As System.Windows.Forms.MenuItem
    Friend WithEvents mi26 As System.Windows.Forms.MenuItem
    Friend WithEvents mi28 As System.Windows.Forms.MenuItem
    Friend WithEvents mi36 As System.Windows.Forms.MenuItem
    Friend WithEvents mi48 As System.Windows.Forms.MenuItem
    Friend WithEvents mi72 As System.Windows.Forms.MenuItem
    Friend WithEvents cmFonts As System.Windows.Forms.ContextMenu
    Friend WithEvents miArial As System.Windows.Forms.MenuItem
    Friend WithEvents miCourierNew As System.Windows.Forms.MenuItem
    Friend WithEvents miGaramond As System.Windows.Forms.MenuItem
    Friend WithEvents miMicrosoftSansSerif As System.Windows.Forms.MenuItem
    Friend WithEvents miTahoma As System.Windows.Forms.MenuItem
    Friend WithEvents miTimesNewRoman As System.Windows.Forms.MenuItem
    Friend WithEvents miVerdana As System.Windows.Forms.MenuItem
    Friend WithEvents ToolBarButton1 As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbbprev As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbbprint As System.Windows.Forms.ToolBarButton
    Friend WithEvents PrintDocument1 As System.Drawing.Printing.PrintDocument
    Friend WithEvents tbbnormal As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButton2 As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbbQuit As System.Windows.Forms.ToolBarButton
    Friend WithEvents MiAlgerian As System.Windows.Forms.MenuItem
    Friend WithEvents Miblanc As System.Windows.Forms.MenuItem
    Friend WithEvents Mijaune As System.Windows.Forms.MenuItem
    Friend WithEvents Miorange As System.Windows.Forms.MenuItem
    Friend WithEvents MiMagenta As System.Windows.Forms.MenuItem
    Friend WithEvents migris As System.Windows.Forms.MenuItem
    Friend WithEvents mirose As System.Windows.Forms.MenuItem
    Friend WithEvents mibleupale As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(RichTextBoxJmb))
        Me.Rtb1 = New System.Windows.Forms.RichTextBox
        Me.imgList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.tb1 = New System.Windows.Forms.ToolBar
        Me.tbbSave = New System.Windows.Forms.ToolBarButton
        Me.tbbOpen = New System.Windows.Forms.ToolBarButton
        Me.tbbSeparator3 = New System.Windows.Forms.ToolBarButton
        Me.tbbFont = New System.Windows.Forms.ToolBarButton
        Me.cmFonts = New System.Windows.Forms.ContextMenu
        Me.MiAlgerian = New System.Windows.Forms.MenuItem
        Me.miArial = New System.Windows.Forms.MenuItem
        Me.miCourierNew = New System.Windows.Forms.MenuItem
        Me.miGaramond = New System.Windows.Forms.MenuItem
        Me.miMicrosoftSansSerif = New System.Windows.Forms.MenuItem
        Me.miTahoma = New System.Windows.Forms.MenuItem
        Me.miTimesNewRoman = New System.Windows.Forms.MenuItem
        Me.miVerdana = New System.Windows.Forms.MenuItem
        Me.tbbFontSize = New System.Windows.Forms.ToolBarButton
        Me.cmFontSizes = New System.Windows.Forms.ContextMenu
        Me.mi8 = New System.Windows.Forms.MenuItem
        Me.mi9 = New System.Windows.Forms.MenuItem
        Me.mi10 = New System.Windows.Forms.MenuItem
        Me.mi11 = New System.Windows.Forms.MenuItem
        Me.mi12 = New System.Windows.Forms.MenuItem
        Me.mi14 = New System.Windows.Forms.MenuItem
        Me.mi16 = New System.Windows.Forms.MenuItem
        Me.mi18 = New System.Windows.Forms.MenuItem
        Me.mi20 = New System.Windows.Forms.MenuItem
        Me.mi22 = New System.Windows.Forms.MenuItem
        Me.mi24 = New System.Windows.Forms.MenuItem
        Me.mi26 = New System.Windows.Forms.MenuItem
        Me.mi28 = New System.Windows.Forms.MenuItem
        Me.mi36 = New System.Windows.Forms.MenuItem
        Me.mi48 = New System.Windows.Forms.MenuItem
        Me.mi72 = New System.Windows.Forms.MenuItem
        Me.tbbSeparator5 = New System.Windows.Forms.ToolBarButton
        Me.tbbnormal = New System.Windows.Forms.ToolBarButton
        Me.tbbBold = New System.Windows.Forms.ToolBarButton
        Me.tbbItalic = New System.Windows.Forms.ToolBarButton
        Me.tbbUnderline = New System.Windows.Forms.ToolBarButton
        Me.tbbStrikeout = New System.Windows.Forms.ToolBarButton
        Me.tbbSeparator1 = New System.Windows.Forms.ToolBarButton
        Me.tbbLeft = New System.Windows.Forms.ToolBarButton
        Me.tbbCenter = New System.Windows.Forms.ToolBarButton
        Me.tbbRight = New System.Windows.Forms.ToolBarButton
        Me.tbbSeparator2 = New System.Windows.Forms.ToolBarButton
        Me.tbbUndo = New System.Windows.Forms.ToolBarButton
        Me.tbbRedo = New System.Windows.Forms.ToolBarButton
        Me.tbbSeparator4 = New System.Windows.Forms.ToolBarButton
        Me.tbbStamp = New System.Windows.Forms.ToolBarButton
        Me.tbbColor = New System.Windows.Forms.ToolBarButton
        Me.cmColors = New System.Windows.Forms.ContextMenu
        Me.miBlack = New System.Windows.Forms.MenuItem
        Me.miBlue = New System.Windows.Forms.MenuItem
        Me.miRed = New System.Windows.Forms.MenuItem
        Me.miGreen = New System.Windows.Forms.MenuItem
        Me.Miblanc = New System.Windows.Forms.MenuItem
        Me.Mijaune = New System.Windows.Forms.MenuItem
        Me.Miorange = New System.Windows.Forms.MenuItem
        Me.MiMagenta = New System.Windows.Forms.MenuItem
        Me.migris = New System.Windows.Forms.MenuItem
        Me.mirose = New System.Windows.Forms.MenuItem
        Me.mibleupale = New System.Windows.Forms.MenuItem
        Me.ToolBarButton1 = New System.Windows.Forms.ToolBarButton
        Me.tbbprev = New System.Windows.Forms.ToolBarButton
        Me.tbbprint = New System.Windows.Forms.ToolBarButton
        Me.ToolBarButton2 = New System.Windows.Forms.ToolBarButton
        Me.tbbQuit = New System.Windows.Forms.ToolBarButton
        Me.ofd1 = New System.Windows.Forms.OpenFileDialog
        Me.sfd1 = New System.Windows.Forms.SaveFileDialog
        Me.PrintDocument1 = New System.Drawing.Printing.PrintDocument
        Me.SuspendLayout()
        '
        'Rtb1
        '
        Me.Rtb1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Rtb1.Location = New System.Drawing.Point(2, 32)
        Me.Rtb1.Name = "Rtb1"
        Me.Rtb1.Size = New System.Drawing.Size(538, 248)
        Me.Rtb1.TabIndex = 0
        Me.Rtb1.Text = "RichTextBox1"
        '
        'imgList1
        '
        Me.imgList1.ColorDepth = System.Windows.Forms.ColorDepth.Depth32Bit
        Me.imgList1.ImageSize = New System.Drawing.Size(16, 16)
        Me.imgList1.ImageStream = CType(resources.GetObject("imgList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.imgList1.TransparentColor = System.Drawing.Color.Transparent
        '
        'tb1
        '
        Me.tb1.AutoSize = False
        Me.tb1.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.tbbSave, Me.tbbOpen, Me.tbbSeparator3, Me.tbbFont, Me.tbbFontSize, Me.tbbSeparator5, Me.tbbnormal, Me.tbbBold, Me.tbbItalic, Me.tbbUnderline, Me.tbbStrikeout, Me.tbbSeparator1, Me.tbbLeft, Me.tbbCenter, Me.tbbRight, Me.tbbSeparator2, Me.tbbUndo, Me.tbbRedo, Me.tbbSeparator4, Me.tbbStamp, Me.tbbColor, Me.ToolBarButton1, Me.tbbprev, Me.tbbprint, Me.ToolBarButton2, Me.tbbQuit})
        Me.tb1.ButtonSize = New System.Drawing.Size(16, 16)
        Me.tb1.Divider = False
        Me.tb1.DropDownArrows = True
        Me.tb1.ImageList = Me.imgList1
        Me.tb1.Location = New System.Drawing.Point(0, 0)
        Me.tb1.Name = "tb1"
        Me.tb1.ShowToolTips = True
        Me.tb1.Size = New System.Drawing.Size(542, 26)
        Me.tb1.TabIndex = 1
        '
        'tbbSave
        '
        Me.tbbSave.ImageIndex = 11
        Me.tbbSave.ToolTipText = "Sauver"
        '
        'tbbOpen
        '
        Me.tbbOpen.ImageIndex = 10
        Me.tbbOpen.ToolTipText = "Ouvrir"
        '
        'tbbSeparator3
        '
        Me.tbbSeparator3.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'tbbFont
        '
        Me.tbbFont.DropDownMenu = Me.cmFonts
        Me.tbbFont.ImageIndex = 16
        Me.tbbFont.Style = System.Windows.Forms.ToolBarButtonStyle.DropDownButton
        Me.tbbFont.ToolTipText = "Font"
        '
        'cmFonts
        '
        Me.cmFonts.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MiAlgerian, Me.miArial, Me.miCourierNew, Me.miGaramond, Me.miMicrosoftSansSerif, Me.miTahoma, Me.miTimesNewRoman, Me.miVerdana})
        '
        'MiAlgerian
        '
        Me.MiAlgerian.Index = 0
        Me.MiAlgerian.Text = "Algerian"
        '
        'miArial
        '
        Me.miArial.Index = 1
        Me.miArial.Text = "Arial"
        '
        'miCourierNew
        '
        Me.miCourierNew.Index = 2
        Me.miCourierNew.Text = "Courier New"
        '
        'miGaramond
        '
        Me.miGaramond.Index = 3
        Me.miGaramond.Text = "Garamond"
        '
        'miMicrosoftSansSerif
        '
        Me.miMicrosoftSansSerif.Index = 4
        Me.miMicrosoftSansSerif.Text = "Microsoft Sans Serif"
        '
        'miTahoma
        '
        Me.miTahoma.Index = 5
        Me.miTahoma.Text = "Tahoma"
        '
        'miTimesNewRoman
        '
        Me.miTimesNewRoman.Index = 6
        Me.miTimesNewRoman.Text = "Times New Roman"
        '
        'miVerdana
        '
        Me.miVerdana.Index = 7
        Me.miVerdana.Text = "Verdana"
        '
        'tbbFontSize
        '
        Me.tbbFontSize.DropDownMenu = Me.cmFontSizes
        Me.tbbFontSize.ImageIndex = 15
        Me.tbbFontSize.Style = System.Windows.Forms.ToolBarButtonStyle.DropDownButton
        Me.tbbFontSize.ToolTipText = "Font Size"
        '
        'cmFontSizes
        '
        Me.cmFontSizes.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mi8, Me.mi9, Me.mi10, Me.mi11, Me.mi12, Me.mi14, Me.mi16, Me.mi18, Me.mi20, Me.mi22, Me.mi24, Me.mi26, Me.mi28, Me.mi36, Me.mi48, Me.mi72})
        '
        'mi8
        '
        Me.mi8.Index = 0
        Me.mi8.Text = "8"
        '
        'mi9
        '
        Me.mi9.Index = 1
        Me.mi9.Text = "9"
        '
        'mi10
        '
        Me.mi10.Index = 2
        Me.mi10.Text = "10"
        '
        'mi11
        '
        Me.mi11.Index = 3
        Me.mi11.Text = "11"
        '
        'mi12
        '
        Me.mi12.Index = 4
        Me.mi12.Text = "12"
        '
        'mi14
        '
        Me.mi14.Index = 5
        Me.mi14.Text = "14"
        '
        'mi16
        '
        Me.mi16.Index = 6
        Me.mi16.Text = "16"
        '
        'mi18
        '
        Me.mi18.Index = 7
        Me.mi18.Text = "18"
        '
        'mi20
        '
        Me.mi20.Index = 8
        Me.mi20.Text = "20"
        '
        'mi22
        '
        Me.mi22.Index = 9
        Me.mi22.Text = "22"
        '
        'mi24
        '
        Me.mi24.Index = 10
        Me.mi24.Text = "24"
        '
        'mi26
        '
        Me.mi26.Index = 11
        Me.mi26.Text = "26"
        '
        'mi28
        '
        Me.mi28.Index = 12
        Me.mi28.Text = "28"
        '
        'mi36
        '
        Me.mi36.Index = 13
        Me.mi36.Text = "36"
        '
        'mi48
        '
        Me.mi48.Index = 14
        Me.mi48.Text = "48"
        '
        'mi72
        '
        Me.mi72.Index = 15
        Me.mi72.Text = "72"
        '
        'tbbSeparator5
        '
        Me.tbbSeparator5.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'tbbnormal
        '
        Me.tbbnormal.ImageIndex = 0
        Me.tbbnormal.ToolTipText = "Normal"
        '
        'tbbBold
        '
        Me.tbbBold.ImageIndex = 1
        Me.tbbBold.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
        Me.tbbBold.ToolTipText = "Gras"
        '
        'tbbItalic
        '
        Me.tbbItalic.ImageIndex = 2
        Me.tbbItalic.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
        Me.tbbItalic.ToolTipText = "Italique"
        '
        'tbbUnderline
        '
        Me.tbbUnderline.ImageIndex = 3
        Me.tbbUnderline.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
        Me.tbbUnderline.ToolTipText = "Souligné"
        '
        'tbbStrikeout
        '
        Me.tbbStrikeout.ImageIndex = 4
        Me.tbbStrikeout.ToolTipText = "Barré"
        '
        'tbbSeparator1
        '
        Me.tbbSeparator1.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'tbbLeft
        '
        Me.tbbLeft.ImageIndex = 5
        Me.tbbLeft.ToolTipText = "Gauche"
        '
        'tbbCenter
        '
        Me.tbbCenter.ImageIndex = 6
        Me.tbbCenter.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
        Me.tbbCenter.ToolTipText = "Centre"
        '
        'tbbRight
        '
        Me.tbbRight.ImageIndex = 7
        Me.tbbRight.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
        Me.tbbRight.ToolTipText = "Droit"
        '
        'tbbSeparator2
        '
        Me.tbbSeparator2.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'tbbUndo
        '
        Me.tbbUndo.ImageIndex = 13
        Me.tbbUndo.ToolTipText = "Annuler"
        '
        'tbbRedo
        '
        Me.tbbRedo.ImageIndex = 14
        Me.tbbRedo.ToolTipText = "Rétablir"
        '
        'tbbSeparator4
        '
        Me.tbbSeparator4.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'tbbStamp
        '
        Me.tbbStamp.ImageIndex = 9
        Me.tbbStamp.ToolTipText = "Editer Signature"
        '
        'tbbColor
        '
        Me.tbbColor.DropDownMenu = Me.cmColors
        Me.tbbColor.ImageIndex = 8
        Me.tbbColor.Style = System.Windows.Forms.ToolBarButtonStyle.DropDownButton
        Me.tbbColor.ToolTipText = "Couleur"
        '
        'cmColors
        '
        Me.cmColors.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.miBlack, Me.miBlue, Me.miRed, Me.miGreen, Me.Miblanc, Me.Mijaune, Me.Miorange, Me.MiMagenta, Me.migris, Me.mirose, Me.mibleupale})
        '
        'miBlack
        '
        Me.miBlack.Index = 0
        Me.miBlack.Text = "Noir"
        '
        'miBlue
        '
        Me.miBlue.Index = 1
        Me.miBlue.Text = "Bleu"
        '
        'miRed
        '
        Me.miRed.Index = 2
        Me.miRed.Text = "Rouge"
        '
        'miGreen
        '
        Me.miGreen.Index = 3
        Me.miGreen.Text = "Vert"
        '
        'Miblanc
        '
        Me.Miblanc.Index = 4
        Me.Miblanc.Text = "Blanc"
        '
        'Mijaune
        '
        Me.Mijaune.Index = 5
        Me.Mijaune.Text = "Jaune"
        '
        'Miorange
        '
        Me.Miorange.Index = 6
        Me.Miorange.Text = "Orange"
        '
        'MiMagenta
        '
        Me.MiMagenta.Index = 7
        Me.MiMagenta.Text = "Violet"
        '
        'migris
        '
        Me.migris.Index = 8
        Me.migris.Text = "Gris"
        '
        'mirose
        '
        Me.mirose.Index = 9
        Me.mirose.Text = "Rose"
        '
        'mibleupale
        '
        Me.mibleupale.Index = 10
        Me.mibleupale.Text = "Bleu pale"
        '
        'ToolBarButton1
        '
        Me.ToolBarButton1.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'tbbprev
        '
        Me.tbbprev.ImageIndex = 18
        Me.tbbprev.ToolTipText = "Aperçu avant impression"
        '
        'tbbprint
        '
        Me.tbbprint.ImageIndex = 19
        Me.tbbprint.ToolTipText = "Imprime"
        '
        'ToolBarButton2
        '
        Me.ToolBarButton2.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'tbbQuit
        '
        Me.tbbQuit.ImageIndex = 20
        Me.tbbQuit.ToolTipText = "Ferme"
        '
        'ofd1
        '
        Me.ofd1.DefaultExt = "rtf"
        Me.ofd1.Filter = "Rich Text files (*.RTF)|*.rtf|Plain Text File|*.txt"
        Me.ofd1.Title = "Open File"
        '
        'sfd1
        '
        Me.sfd1.DefaultExt = "rtf"
        Me.sfd1.Filter = "Rich Text files (*.RTF)|*.rtf|Plain Text File|*.txt"
        Me.sfd1.Title = "Save As"
        '
        'PrintDocument1
        '
        '
        'RichTextBoxJmb
        '
        Me.Controls.Add(Me.tb1)
        Me.Controls.Add(Me.Rtb1)
        Me.Name = "RichTextBoxJmb"
        Me.Size = New System.Drawing.Size(542, 282)
        Me.ResumeLayout(False)

    End Sub

#End Region
    <Description("Ce qui se produit quand le bouton Signature est cliqué "), Category("Behavior")> _
 Public Event Stamp As System.EventHandler
    Protected Overridable Sub OnStamp(ByVal e As EventArgs)

        RaiseEvent Stamp(Me, e)
        Select Case StampAction
            Case StampActions.Editepar
                If (True) Then
                    Dim stamp As New StringBuilder("")  'Votre texte
                    If Rtb1.Text.Length > 0 Then
                        stamp.Append(ControlChars.Cr + ControlChars.Lf + ControlChars.Cr + ControlChars.Lf)
                        'ajoute 2 lignes vierges
                    End If
                    stamp.Append("Edité Par ")
                    'utilise le username du login
                    If Thread.CurrentPrincipal Is Nothing Or Thread.CurrentPrincipal.Identity Is Nothing Or Thread.CurrentPrincipal.Identity.Name.Length <= 0 Then
                        stamp.Append(Environment.UserName)
                    Else
                        stamp.Append(Thread.CurrentPrincipal.Identity.Name)
                    End If
                    stamp.Append((" le " + DateTime.Now.ToLongDateString() + ControlChars.Cr + ControlChars.Lf))

                    Rtb1.SelectionLength = 0 'déselecte texte
                    Rtb1.SelectionStart = Rtb1.Text.Length 'Nouvelle sélection à la fin du texte
                    Rtb1.SelectionColor = Me.StampColor 'la sélection est en bleu
                    Rtb1.SelectionFont = New Font(Rtb1.SelectionFont, FontStyle.Bold) 'set the selection font and style
                    Rtb1.AppendText(stamp.ToString()) 'ajoute la signature
                    Rtb1.Focus() 'met focus sur la richtextbox
                End If
            Case StampActions.DateHeure
                If (True) Then
                    Dim stamp As New StringBuilder("") 'Votre texte
                    If Rtb1.Text.Length > 0 Then
                        stamp.Append(ControlChars.Cr + ControlChars.Lf + ControlChars.Cr + ControlChars.Lf) 'add two lines for space
                    End If
                    stamp.Append((DateTime.Now.ToLongDateString() + ControlChars.Cr + ControlChars.Lf))
                    Rtb1.SelectionLength = 0 'déselecte texte
                    Rtb1.SelectionStart = Rtb1.Text.Length 'Nouvelle sélection à la fin du texte
                    Rtb1.SelectionColor = Me.StampColor 'la sélection est en bleu
                    Rtb1.SelectionFont = New Font(Rtb1.SelectionFont, FontStyle.Bold) 'met en place la police
                    Rtb1.AppendText(stamp.ToString()) 'ajoute la signature
                    Rtb1.Focus() 'met focus sur la richtextbox
                End If
            Case Else
                Dim stamp As New StringBuilder("") 'Votre texte
                If Rtb1.Text.Length > 0 Then
                    stamp.Append(ControlChars.Cr + ControlChars.Lf + ControlChars.Cr + ControlChars.Lf) 'add two lines for space
                End If

        End Select 'end select
    End Sub 'OnStamp

    Public Sub UpdateToolbar()

        Dim fnt As Font
        If Not (Rtb1.SelectionFont Is Nothing) Then
            fnt = Rtb1.SelectionFont
        Else
            fnt = Rtb1.Font
        End If
        'Teste les boutons de la toolbar
        tbbnormal.Pushed = Not fnt.Bold 'Bouton normal
        tbbBold.Pushed = fnt.Bold 'Bouton gras
        tbbItalic.Pushed = fnt.Italic 'Bouton italique
        tbbUnderline.Pushed = fnt.Underline 'Bouton souligné
        tbbStrikeout.Pushed = fnt.Strikeout 'Bouton barré
        tbbLeft.Pushed = Rtb1.SelectionAlignment = HorizontalAlignment.Left
        tbbCenter.Pushed = Rtb1.SelectionAlignment = HorizontalAlignment.Center
        tbbRight.Pushed = Rtb1.SelectionAlignment = HorizontalAlignment.Right
        'teste la couleur
        Dim mi As MenuItem
        For Each mi In cmColors.MenuItems
            If Rtb1.SelectionColor.Equals(Color.FromName(mi.Text)) Then
                mi.Checked = True
            Else
                mi.Checked = False
            End If

        Next mi

        'teste la police
        Dim mi2 As MenuItem
        For Each mi2 In cmFonts.MenuItems
            mi2.Checked = fnt.FontFamily.Name = mi2.Text
        Next mi2

        'Teste la taille de la ploice
        Dim mi1 As MenuItem
        For Each mi1 In cmFontSizes.MenuItems
            mi1.Checked = CInt(fnt.SizeInPoints) = Single.Parse(mi1.Text)
        Next mi1
    End Sub 'UpdateToolbar
    Private Sub UpdateToolbarSeperators()
        'Sauve & Ouvre
        If Not tbbSave.Visible And Not tbbOpen.Visible Then
            tbbSeparator3.Visible = False
        Else
            tbbSeparator3.Visible = True
        End If
        'Font & Font Size
        If Not tbbFont.Visible And Not tbbFontSize.Visible Then
            tbbSeparator5.Visible = False
        Else
            tbbSeparator5.Visible = True
        End If
        'Gras, Italique, Souligné, & Barré
        If Not tbbBold.Visible And Not tbbItalic.Visible And Not tbbUnderline.Visible And Not tbbStrikeout.Visible Then
            tbbSeparator1.Visible = False
        Else
            tbbSeparator1.Visible = True
        End If
        'Gauche, Centre, & Droit
        If Not tbbLeft.Visible And Not tbbCenter.Visible And Not tbbRight.Visible Then
            tbbSeparator2.Visible = False
        Else
            tbbSeparator2.Visible = True
        End If
        'Annuler rétablir
        If Not tbbUndo.Visible And Not tbbRedo.Visible Then
            tbbSeparator4.Visible = False
        Else
            tbbSeparator4.Visible = True
        End If
    End Sub 'UpdateToolbarSeperators

#Region "Evenements"
    Private Sub rtb1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        UpdateToolbar() 'Met à jour les boutons de la toolbar 
    End Sub 'rtb1_SelectionChanged
    Private Sub Color_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miBlack.Click, miBlue.Click, miRed.Click, miGreen.Click  ', MiMagenta.Click, Mijaune.Click, Miorange.Click
        If Not (Rtb1.SelectionFont Is Nothing) Then
            'Met la couleur définie dans la toolbar
            ' comme le menu est en français, il faut traduire
            Dim ccolor As String
            Select Case LCase(CType(sender, System.Windows.forms.MenuItem).Text)
                Case "rouge"
                    ccolor = "red"
                Case "noir"
                    ccolor = "black"
                Case "vert"
                    ccolor = "green"
                Case "blanc"
                    ccolor = "white"
                Case "bleu"
                    ccolor = "blue"
                Case "violet"
                    ccolor = "magenta"
                Case "jaune"
                    ccolor = "yellow"
                Case "orange"
                    ccolor = "orange"
                Case "gris"
                    ccolor = "gray"
                Case "rose"
                    ccolor = "pink"
                Case "bleu pale"
                    ccolor = "cyan"
            End Select
            Rtb1.SelectionColor = System.Drawing.Color.FromName(ccolor)
        End If
    End Sub
    Private Sub FontSize_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mi8.Click, mi9.Click, mi10.Click, mi11.Click, mi12.Click, mi14.Click, mi16.Click, mi18.Click, mi20.Click, mi22.Click, mi24.Click, mi26.Click, mi28.Click, mi36.Click, mi48.Click, mi72.Click
        'met la taille de police choisie
        Rtb1.SelectionFont = New Font(Rtb1.SelectionFont.FontFamily.Name, Single.Parse(CType(sender, MenuItem).Text))
    End Sub 'FontSize_Click
    Private Sub Font_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miArial.Click, miCourierNew.Click, miGaramond.Click, miMicrosoftSansSerif.Click, miTahoma.Click, miTimesNewRoman.Click, miVerdana.Click, MiAlgerian.Click
        If Not (Rtb1.SelectionFont Is Nothing) Then
            'met la font choisie
            Rtb1.SelectionFont = New Font(CType(sender, MenuItem).Text, Rtb1.SelectionFont.SizeInPoints)
        End If
    End Sub 'Font_Click

    Private Sub rtb1_LinkClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.LinkClickedEventArgs)
        System.Diagnostics.Process.Start(e.LinkText)
    End Sub 'rtb1_LinkClicked
#End Region
#Region "Les propriétés des icones"
    <Description("Le controle interne de toolbar."), Category("Controles internes")> _
    Public ReadOnly Property Toolbar() As ToolBar
        Get
            Return tb1
        End Get
    End Property
    <Description("Le controle interne de RichText."), Category("Controles internes")> _
        Public ReadOnly Property RichTextBox() As RichTextBox
        Get
            Return Rtb1
        End Get
    End Property
    <Description("Le bouton Sauve est visible ou non."), Category("Apparence")> _
        Public Property ShowSave() As [Boolean]
        Get
            Return tbbSave.Visible
        End Get
        Set(ByVal Value As [Boolean])
            tbbSave.Visible = Value
            UpdateToolbarSeperators()
        End Set
    End Property
    <Description("Le bouton Ouvre est visible ou non."), Category("Apparence")> _
        Public Property ShowOpen() As [Boolean]
        Get
            Return tbbOpen.Visible
        End Get
        Set(ByVal Value As [Boolean])
            tbbOpen.Visible = Value
            UpdateToolbarSeperators()
        End Set
    End Property
    <Description("Le bouton Signature est visible ou non."), Category("Apparence")> _
        Public Property ShowStamp() As [Boolean]
        Get
            Return tbbStamp.Visible
        End Get
        Set(ByVal Value As [Boolean])
            tbbStamp.Visible = Value
        End Set
    End Property
    <Description("Le bouton Couleur est visible ou non."), Category("Apparence")> _
        Public Property ShowColors() As [Boolean]
        Get
            Return tbbColor.Visible
        End Get
        Set(ByVal Value As [Boolean])
            tbbColor.Visible = Value
        End Set
    End Property
    <Description("Le bouton Annuler est visible ou non."), Category("Apparence")> _
        Public Property ShowUndo() As [Boolean]
        Get
            Return tbbUndo.Visible
        End Get
        Set(ByVal Value As [Boolean])
            tbbUndo.Visible = Value
            UpdateToolbarSeperators()
        End Set
    End Property
    <Description("Le bouton Rétablir est visible ou non."), Category("Apparence")> _
        Public Property ShowRedo() As [Boolean]
        Get
            Return tbbRedo.Visible
        End Get
        Set(ByVal Value As [Boolean])
            tbbRedo.Visible = Value
            UpdateToolbarSeperators()
        End Set
    End Property
    <Description("Le bouton Normal est visible ou non."), Category("Apparence")> _
       Public Property Shownormal() As [Boolean]
        Get
            Return tbbnormal.Visible
        End Get
        Set(ByVal Value As [Boolean])
            tbbnormal.Visible = Value
            UpdateToolbarSeperators()
        End Set
    End Property
    <Description("Le bouton Gras est visible ou non."), Category("Apparence")> _
        Public Property ShowBold() As [Boolean]
        Get
            Return tbbBold.Visible
        End Get
        Set(ByVal Value As [Boolean])
            tbbBold.Visible = Value
            UpdateToolbarSeperators()
        End Set
    End Property
    <Description("Le bouton Italic est visible ou non."), Category("Apparence")> _
        Public Property ShowItalic() As [Boolean]
        Get
            Return tbbItalic.Visible
        End Get
        Set(ByVal Value As [Boolean])
            tbbItalic.Visible = Value
            UpdateToolbarSeperators()
        End Set
    End Property
    <Description("Le bouton Soulihné est visible ou non."), Category("Apparence")> _
        Public Property ShowUnderline() As [Boolean]
        Get
            Return tbbUnderline.Visible
        End Get
        Set(ByVal Value As [Boolean])
            tbbUnderline.Visible = Value
            UpdateToolbarSeperators()
        End Set
    End Property
    <Description("Le bouton Barré est visible ou non."), Category("Apparence")> _
        Public Property ShowStrikeout() As [Boolean]
        Get
            Return tbbStrikeout.Visible
        End Get
        Set(ByVal Value As [Boolean])
            tbbStrikeout.Visible = Value
            UpdateToolbarSeperators()
        End Set
    End Property
    <Description("Le bouton Alignement gauche est visible ou non."), Category("Apparence")> _
        Public Property ShowLeftJustify() As [Boolean]
        Get
            Return tbbLeft.Visible
        End Get
        Set(ByVal Value As [Boolean])
            tbbLeft.Visible = Value
            UpdateToolbarSeperators()
        End Set
    End Property
    <Description("Le bouton Alignement droite est visible ou non."), Category("Apparence")> _
        Public Property ShowRightJustify() As [Boolean]
        Get
            Return tbbRight.Visible
        End Get
        Set(ByVal Value As [Boolean])
            tbbRight.Visible = Value
            UpdateToolbarSeperators()
        End Set
    End Property
    <Description("Le bouton Centré est visible ou non."), Category("Apparence")> _
        Public Property ShowCenterJustify() As [Boolean]
        Get
            Return tbbCenter.Visible
        End Get
        Set(ByVal Value As [Boolean])
            tbbCenter.Visible = Value
            UpdateToolbarSeperators()
        End Set
    End Property

    Dim m_StampAction As StampActions = StampActions.Editepar
    <Description("Determine comment le bouton signature réagit."), Category("Divers")> _
        Public Property StampAction() As StampActions
        Get
            Return m_StampAction
        End Get
        Set(ByVal Value As StampActions)
            m_StampAction = Value
        End Set
    End Property
    Dim m_StampColor As Color = Color.Blue
    <Description("Couleur de la signature."), Category("Apparence")> _
    Public Property StampColor() As Color
        Get
            Return m_StampColor
        End Get
        Set(ByVal Value As Color)
            m_StampColor = Value
        End Set
    End Property
    <Description("Le bouton Font est visible ou non."), Category("Apparence")> _
        Public Property ShowFont() As [Boolean]
        Get
            Return tbbFont.Visible
        End Get
        Set(ByVal Value As [Boolean])
            tbbFont.Visible = Value
        End Set
    End Property
    <Description("Le bouton Taille police est visible ou non."), Category("Apparence")> _
        Public Property ShowFontSize() As [Boolean]
        Get
            Return tbbFontSize.Visible
        End Get
        Set(ByVal Value As [Boolean])
            tbbFontSize.Visible = Value
        End Set
    End Property
    <Description("Detecte s'il y a un URLs dans le richtextbox"), Category("Divers")> _
        Public Shadows Property DetectURLs() As [Boolean]
        Get
            Return Rtb1.DetectUrls
        End Get
        Set(ByVal Value As [Boolean])
            Rtb1.DetectUrls = Value
        End Set
    End Property
    <Description("Le bouton aPER9U est visible ou non."), Category("Apparence")> _
        Public Property ShowPreview() As [Boolean]
        Get
            Return tbbprev.Visible
        End Get
        Set(ByVal Value As [Boolean])
            tbbprev.Visible = Value
        End Set
    End Property
    <Description("Le bouton Imprime est visible ou non."), Category("Apparence")> _
            Public Property ShowPrint() As [Boolean]
        Get
            Return tbbprint.Visible
        End Get
        Set(ByVal Value As [Boolean])
            tbbprint.Visible = Value
        End Set
    End Property
    <Description("Le bouton quitter est visible ou non."), Category("Controles internes")> _
        Public Property ShowQuit() As [Boolean]
        Get
            Return tbbQuit.Visible
        End Get
        Set(ByVal Value As [Boolean])
            tbbQuit.Visible = Value
        End Set

    End Property
#End Region ' propriété icones
    '***********************************************************************************

    '***********************************************************************************
    '*****************************************************************************************
    ' pour impression 
    '    '*****************************************************************************************
#Region "structures"
    Private Structure STRUCT_RECT
        Public left As Integer
        Public top As Integer
        Public right As Integer
        Public bottom As Integer
    End Structure

    Private Structure STRUCT_CHAR_RANGE
        Public cpMin As Integer
        Public cpMax As Integer
    End Structure

    Private Structure STRUCT_FORMAT_RANGE
        Public hdc As IntPtr
        Public hdcTarget As IntPtr
        Public rc As RichTextBoxJmb.STRUCT_RECT
        Public rcPage As RichTextBoxJmb.STRUCT_RECT
        Public chrg As RichTextBoxJmb.STRUCT_CHAR_RANGE
    End Structure

    Private Structure STRUCT_CHAR_FORMAT
        Public cbSize As Integer
        Public dwMask As UInt32
        Public dwEffects As UInt32
        Public yHeight As Integer
        Public yOffset As Integer
        Public crTextColor As Integer
        Public bCharSet As Byte
        Public bPitchAndFamily As Byte
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=32)> _
        Public szFaceName As Char()
    End Structure

    Private Const WM_USER As Integer = 1024
    Private Const EM_FORMATRANGE As Integer = 1081
    Private Const EM_GETCHARFORMAT As Integer = 1082
    Private Const EM_SETCHARFORMAT As Integer = 1092
    Private SCF_SELECTION As Integer
    Private SCF_WORD As Integer
    Private SCF_ALL As Integer
    Private Const CFM_BOLD As Long = 1
    Private Const CFM_ITALIC As Long = 2
    Private Const CFM_UNDERLINE As Long = 4
    Private Const CFM_STRIKEOUT As Long = 8
    Private Const CFM_PROTECTED As Long = 16
    Private Const CFM_LINK As Long = 32
    Private Const CFM_SIZE As Long = 2147483648
    Private Const CFM_COLOR As Long = 1073741824
    Private Const CFM_FACE As Long = 536870912
    Private Const CFM_OFFSET As Long = 268435456
    Private Const CFM_CHARSET As Long = 134217728
    Private Const CFE_BOLD As Long = 1
    Private Const CFE_ITALIC As Long = 2
    Private Const CFE_UNDERLINE As Long = 4
    Private Const CFE_STRIKEOUT As Long = 8
    Private Const CFE_PROTECTED As Long = 16
    Private Const CFE_LINK As Long = 32
    Private Const CFE_AUTOCOLOR As Long = 1073741824
#End Region

#Region "Fonctions d'impression"
    <DllImportAttribute("user32.dll")> _
            Private Shared Shadows Function SendMessage(ByVal hWnd As IntPtr, ByVal msg As Integer, ByVal wParam As Integer, ByVal lParam As IntPtr) As Integer
    End Function

    Public Function FormatRange(ByVal measureOnly As Boolean, ByVal e As PrintPageEventArgs, ByVal charFrom As Integer, ByVal charTo As Integer) As Integer
        Dim sTRUCT_CHAR_RANGE As STRUCT_CHAR_RANGE
        Dim sTRUCT_FORMAT_RANGE As STRUCT_FORMAT_RANGE
        Dim sTRUCT_RECT1 As STRUCT_RECT
        Dim sTRUCT_RECT2 As STRUCT_RECT
        Dim j2 As Integer
        Dim rectangle As Rectangle
        sTRUCT_CHAR_RANGE.cpMin = charFrom
        sTRUCT_CHAR_RANGE.cpMax = charTo
        sTRUCT_RECT1.top = HundredthInchToTwips(e.MarginBounds.Top)
        sTRUCT_RECT1.bottom = HundredthInchToTwips(e.MarginBounds.Bottom)
        sTRUCT_RECT1.left = HundredthInchToTwips(e.MarginBounds.Left)
        sTRUCT_RECT1.right = HundredthInchToTwips(e.MarginBounds.Right)
        sTRUCT_RECT2.top = HundredthInchToTwips(e.PageBounds.Top)
        sTRUCT_RECT2.bottom = HundredthInchToTwips(e.PageBounds.Bottom)
        sTRUCT_RECT2.left = HundredthInchToTwips(e.PageBounds.Left)
        sTRUCT_RECT2.right = HundredthInchToTwips(e.PageBounds.Right)
        Dim j1 As IntPtr = e.Graphics.GetHdc()
        sTRUCT_FORMAT_RANGE.chrg = sTRUCT_CHAR_RANGE
        sTRUCT_FORMAT_RANGE.hdc = j1
        sTRUCT_FORMAT_RANGE.hdcTarget = j1
        sTRUCT_FORMAT_RANGE.rc = sTRUCT_RECT1
        sTRUCT_FORMAT_RANGE.rcPage = sTRUCT_RECT2
        If measureOnly Then
            j2 = 0
        Else
            j2 = 1
        End If
        Dim k As IntPtr = Marshal.AllocCoTaskMem(Marshal.SizeOf(sTRUCT_FORMAT_RANGE))
        Marshal.StructureToPtr(sTRUCT_FORMAT_RANGE, k, False)
        Dim i2 As Integer = SendMessage(Rtb1.Handle, 1081, j2, k)
        Marshal.FreeCoTaskMem(k)
        e.Graphics.ReleaseHdc(j1)
        Return i2
    End Function

    Private Function HundredthInchToTwips(ByVal n As Integer) As Integer
        Return Convert.ToInt32(n * 14.4)
    End Function

    Public Sub FormatRangeDone()
        Dim i As IntPtr = New IntPtr(0)
        SendMessage(Rtb1.Handle, 1081, 0, i)
    End Sub

    Public Function SetSelectionFont(ByVal face As String) As Boolean
        Dim sTRUCT_CHAR_FORMAT As STRUCT_CHAR_FORMAT = New STRUCT_CHAR_FORMAT
        sTRUCT_CHAR_FORMAT.cbSize = Marshal.SizeOf(sTRUCT_CHAR_FORMAT)
        sTRUCT_CHAR_FORMAT.dwMask = Convert.ToUInt32(536870912)
        sTRUCT_CHAR_FORMAT.szFaceName = New Char(32) {}
        face.CopyTo(0, sTRUCT_CHAR_FORMAT.szFaceName, 0, Math.Min(31, face.Length))
        Dim i As IntPtr = Marshal.AllocCoTaskMem(Marshal.SizeOf(sTRUCT_CHAR_FORMAT))
        Marshal.StructureToPtr(sTRUCT_CHAR_FORMAT, i, False)
        If SendMessage(Rtb1.Handle, 1092, SCF_SELECTION, i) = 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function SetSelectionSize(ByVal size As Integer) As Boolean
        Dim sTRUCT_CHAR_FORMAT As STRUCT_CHAR_FORMAT = New STRUCT_CHAR_FORMAT
        sTRUCT_CHAR_FORMAT.cbSize = Marshal.SizeOf(sTRUCT_CHAR_FORMAT)
        sTRUCT_CHAR_FORMAT.dwMask = Convert.ToUInt32(2147483648)
        sTRUCT_CHAR_FORMAT.yHeight = size * 20
        Dim i As IntPtr = Marshal.AllocCoTaskMem(Marshal.SizeOf(sTRUCT_CHAR_FORMAT))
        Marshal.StructureToPtr(sTRUCT_CHAR_FORMAT, i, False)
        If SendMessage(Rtb1.Handle, 1092, SCF_SELECTION, i) = 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function SetSelectionBold(ByVal bold As Boolean) As Boolean
        If bold Then
            Return SetSelectionStyle(1, 1)
        Else
            Return SetSelectionStyle(1, 0)
        End If
    End Function

    Public Function SetSelectionItalic(ByVal italic As Boolean) As Boolean
        If italic Then
            Return SetSelectionStyle(2, 2)
        Else
            Return SetSelectionStyle(2, 0)
        End If
    End Function

    Public Function SetSelectionUnderlined(ByVal underlined As Boolean) As Boolean
        If underlined Then
            Return SetSelectionStyle(4, 4)
        Else
            Return SetSelectionStyle(4, 0)
        End If
    End Function

    Private Function SetSelectionStyle(ByVal mask As Integer, ByVal effect As Integer) As Boolean
        Dim sTRUCT_CHAR_FORMAT As New STRUCT_CHAR_FORMAT
        sTRUCT_CHAR_FORMAT.cbSize = Marshal.SizeOf(sTRUCT_CHAR_FORMAT)
        sTRUCT_CHAR_FORMAT.dwMask = Convert.ToUInt32(mask)
        sTRUCT_CHAR_FORMAT.dwEffects = Convert.ToUInt32(effect)
        Dim i As IntPtr = Marshal.AllocCoTaskMem(Marshal.SizeOf(sTRUCT_CHAR_FORMAT))
        Marshal.StructureToPtr(sTRUCT_CHAR_FORMAT, i, False)

        If SendMessage(MyBase.Handle, 1092, SCF_SELECTION, i) = 0 Then
            Return True
        Else
            Return False
        End If
    End Function
#End Region

#Region "Action sur boutons"
    Private Sub tb1_ButtonClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles tb1.ButtonClick

        Select Case e.Button.ToolTipText.ToLower()
            Case "normal"
                If (True) Then
                    If Not (Rtb1.SelectionFont Is Nothing) Then

                        Rtb1.SelectionFont = New Font(Rtb1.SelectionFont, FontStyle.Regular)
                    End If
                End If
            Case "gras"
                If (True) Then
                    If Not (Rtb1.SelectionFont Is Nothing) Then

                        Rtb1.SelectionFont = New Font(Rtb1.SelectionFont, FontStyle.Bold)
                    End If
                End If
            Case "italique"
                If (True) Then
                    If Not (Rtb1.SelectionFont Is Nothing) Then
                        'Modifie la font
                        Rtb1.SelectionFont = New System.Drawing.Font(Rtb1.SelectionFont, System.Drawing.FontStyle.Italic)
                    End If
                End If
            Case "souligné"
                If (True) Then
                    If Not (Rtb1.SelectionFont Is Nothing) Then
                        ' modifie la police
                        Rtb1.SelectionFont = New System.Drawing.Font(Rtb1.SelectionFont, System.Drawing.FontStyle.Underline)
                    End If
                End If
            Case "barré"
                If (True) Then
                    If Not (Rtb1.SelectionFont Is Nothing) Then
                        ' modifie la police
                        Rtb1.SelectionFont = New System.Drawing.Font(Rtb1.SelectionFont, System.Drawing.FontStyle.Strikeout)
                    End If
                End If
            Case "gauche"
                If (True) Then
                    ' modifie l'alignement
                    Rtb1.SelectionAlignment = Windows.Forms.HorizontalAlignment.Left
                End If
            Case "centre"
                If (True) Then
                    ' modifie l'alignement
                    Rtb1.SelectionAlignment = Windows.Forms.HorizontalAlignment.Center
                End If
            Case "droit"
                If (True) Then
                    ' modifie l'alignement
                    Rtb1.SelectionAlignment = Windows.Forms.HorizontalAlignment.Right
                End If
            Case "editer signature"
                If (True) Then
                    OnStamp(New EventArgs)   'Avoie vers l'évenement signature
                End If
            Case "couleur"
                If (True) Then
                    Rtb1.SelectionColor = Color.Black
                End If
            Case "annuler"
                If (True) Then
                    Rtb1.Undo()
                End If
            Case "rétablir"
                If (True) Then
                    Rtb1.Redo()
                End If
            Case "ouvrir"
                If (True) Then
                    Try
                        If ofd1.ShowDialog() = DialogResult.OK And ofd1.FileName.Length > 0 Then
                            If System.IO.Path.GetExtension(ofd1.FileName).ToLower().Equals(".rtf") Then
                                Rtb1.LoadFile(ofd1.FileName, RichTextBoxStreamType.RichText)
                            Else
                                Rtb1.LoadFile(ofd1.FileName, RichTextBoxStreamType.PlainText)
                            End If
                        End If
                    Catch ae As ArgumentException
                        If ae.Message = "Invalid file format." Then
                            MessageBox.Show(("Erreur au chargement du fichier: " + ofd1.FileName))
                        End If
                    End Try
                End If
            Case "sauver"
                If (True) Then
                    If sfd1.ShowDialog() = DialogResult.OK And sfd1.FileName.Length > 0 Then
                        If System.IO.Path.GetExtension(sfd1.FileName).ToLower().Equals(".rtf") Then
                            Rtb1.SaveFile(sfd1.FileName)
                        Else
                            Rtb1.SaveFile(sfd1.FileName, RichTextBoxStreamType.PlainText)
                        End If
                    End If 'end select
                End If 'met à jour la barre d'outils
            Case "aperçu avant impression"
                Dim YPreview As New PrintPreviewDialog
                YPreview.WindowState = FormWindowState.Maximized
                YPreview.Document = PrintDocument1
                YPreview.ShowDialog()
            Case "imprime"
                PrintDocument1.Print()
            Case "ferme"
                If MsgBox("Avez-vous sauvegardé votre saisie ?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    SendKeys.Send("%{F4}")
                End If
        End Select
        UpdateToolbar()

    End Sub
#End Region  ' choix sur icone

#Region "Impression"
    Private Sub PrintDocument1_BeginPrint(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintEventArgs) Handles PrintDocument1.BeginPrint
        m_nFirstCharOnPage = 0
    End Sub

    Private m_nFirstCharOnPage As Integer
    Private ipage As Integer = 1

    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Dim intPrintAreaWidth As Integer
        Dim intPrintAreaHeight As Integer
        Dim font As New Font("Microsoft Sans Serif", 10)
        Dim myPen As New Pen(Color.Blue, 2)
        Dim nWidth As Integer = PrintDocument1.PrinterSettings.DefaultPageSettings.PaperSize.Width
        Dim nHeight As Integer = PrintDocument1.PrinterSettings.DefaultPageSettings.PaperSize.Height
        With PrintDocument1.DefaultPageSettings
            intPrintAreaWidth = .PaperSize.Width - .Margins.Left - .Margins.Right
            intPrintAreaHeight = .PaperSize.Height - .Margins.Top
        End With
        m_nFirstCharOnPage = Me.FormatRange(False, _
                                      e, _
                                    m_nFirstCharOnPage, _
                                  Rtb1.Text.Length)

        ' Vérifie s'il y a d'autres pages à imprimer

        e.Graphics.DrawLine(myPen, 0, intPrintAreaHeight + 30, PrintDocument1.DefaultPageSettings.PaperSize.Width, intPrintAreaHeight + 30)
        e.Graphics.DrawString("page: " & CStr(ipage), font, Brushes.Blue, intPrintAreaWidth + 90, intPrintAreaHeight + 50)
        If (m_nFirstCharOnPage < Rtb1.Text.Length) Then
            ipage = ipage + 1
            e.HasMorePages = True
        Else
            e.HasMorePages = False
        End If
    End Sub
#End Region      ' impression
    ' sauvegarde au cas ou je perde tout afin de ne pas avoir à le recréer, une fois suffit
    ' Font
    '        AddHandler miTimesNewRoman.Click, AddressOf Me.Font_Click
    '        AddHandler MiAlgerian.Click, AddressOf Me.Font_Click
    '        AddHandler miCourierNew.Click, AddressOf Me.Font_Click
    '        AddHandler miArial.Click, AddressOf Me.Font_Click
    '        AddHandler miTimesNewRoman.Click, AddressOf Me.Font_Click
    '        AddHandler miMicrosoftSansSerif.Click, AddressOf Me.Font_Click
    '        AddHandler miGaramond.Click, AddressOf Me.Font_Click
    '        AddHandler miTahoma.Click, AddressOf Me.Font_Click
    '        AddHandler miVerdana.Click, AddressOf Me.Font_Click
    '    ' fontSize
    '        AddHandler mi8.Click, AddressOf Me.FontSize_Click
    '        AddHandler mi9.Click, AddressOf Me.FontSize_Click
    '        AddHandler mi10.Click, AddressOf Me.FontSize_Click
    '        AddHandler mi11.Click, AddressOf Me.FontSize_Click
    '        AddHandler mi12.Click, AddressOf Me.FontSize_Click
    '        AddHandler mi14.Click, AddressOf Me.FontSize_Click
    '        AddHandler mi16.Click, AddressOf Me.FontSize_Click
    '        AddHandler mi18.Click, AddressOf Me.FontSize_Click
    '        AddHandler mi20.Click, AddressOf Me.FontSize_Click
    '        AddHandler mi22.Click, AddressOf Me.FontSize_Click
    '        AddHandler mi24.Click, AddressOf Me.FontSize_Click
    '        AddHandler mi26.Click, AddressOf Me.FontSize_Click
    '        AddHandler mi28.Click, AddressOf Me.FontSize_Click
    '        AddHandler mi36.Click, AddressOf Me.FontSize_Click
    '        AddHandler mi48.Click, AddressOf Me.FontSize_Click
    '        AddHandler mi72.Click, AddressOf Me.FontSize_Click
    '    ' couleurs
    '        AddHandler miBlack.Click, AddressOf Me.Color_Click
    '        AddHandler miBlue.Click, AddressOf Me.Color_Click
    '        AddHandler miRed.Click, AddressOf Me.Color_Click
    '        AddHandler miGreen.Click, AddressOf Me.Color_Click
    '        AddHandler Miblanc.Click, AddressOf Me.Color_Click
    '        AddHandler Mijaune.Click, AddressOf Me.Color_Click
    '        AddHandler Miorange.Click, AddressOf Me.Color_Click
    '        AddHandler MiMagenta.Click, AddressOf Me.Color_Click
End Class
