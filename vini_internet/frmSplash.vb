Option Strict Off
Option Explicit On
Friend Class frmSplash
	Inherits System.Windows.Forms.Form
#Region "Code généré par le Concepteur Windows Form "
	Public Sub New()
		MyBase.New()
		If m_vb6FormDefInstance Is Nothing Then
			If m_InitializingDefInstance Then
				m_vb6FormDefInstance = Me
			Else
				Try 
					'Pour le formulaire de démarrage, la première instance créée est l'instance par défaut.
					If System.Reflection.Assembly.GetExecutingAssembly.EntryPoint.DeclaringType Is Me.GetType Then
						m_vb6FormDefInstance = Me
					End If
				Catch
				End Try
			End If
		End If
		'Cet appel est requis par le Concepteur Windows Form.
		InitializeComponent()
	End Sub
	'La méthode substituée Dispose du formulaire pour nettoyer la liste des composants.
	Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Requis par le Concepteur Windows Form
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents picLogo As System.Windows.Forms.PictureBox
	Public WithEvents lblLicenseTo As System.Windows.Forms.Label
	Public WithEvents lblProductName As System.Windows.Forms.Label
	Public WithEvents lblCompanyProduct As System.Windows.Forms.Label
	Public WithEvents lblPlatform As System.Windows.Forms.Label
	Public WithEvents lblVersion As System.Windows.Forms.Label
	Public WithEvents lblWarning As System.Windows.Forms.Label
	Public WithEvents lblCompany As System.Windows.Forms.Label
	Public WithEvents lblCopyright As System.Windows.Forms.Label
	Public WithEvents fraMainFrame As System.Windows.Forms.GroupBox
	'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
	'Il peut être modifié à l'aide du Concepteur Windows Form.
	'Ne pas le modifier à l'aide de l'éditeur de code.
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmSplash))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.fraMainFrame = New System.Windows.Forms.GroupBox
        Me.picLogo = New System.Windows.Forms.PictureBox
        Me.lblLicenseTo = New System.Windows.Forms.Label
        Me.lblProductName = New System.Windows.Forms.Label
        Me.lblCompanyProduct = New System.Windows.Forms.Label
        Me.lblPlatform = New System.Windows.Forms.Label
        Me.lblVersion = New System.Windows.Forms.Label
        Me.lblWarning = New System.Windows.Forms.Label
        Me.lblCompany = New System.Windows.Forms.Label
        Me.lblCopyright = New System.Windows.Forms.Label
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.fraMainFrame.SuspendLayout()
        Me.SuspendLayout()
        '
        'fraMainFrame
        '
        Me.fraMainFrame.BackColor = System.Drawing.SystemColors.Control
        Me.fraMainFrame.Controls.Add(Me.PictureBox1)
        Me.fraMainFrame.Controls.Add(Me.picLogo)
        Me.fraMainFrame.Controls.Add(Me.lblLicenseTo)
        Me.fraMainFrame.Controls.Add(Me.lblProductName)
        Me.fraMainFrame.Controls.Add(Me.lblCompanyProduct)
        Me.fraMainFrame.Controls.Add(Me.lblPlatform)
        Me.fraMainFrame.Controls.Add(Me.lblVersion)
        Me.fraMainFrame.Controls.Add(Me.lblWarning)
        Me.fraMainFrame.Controls.Add(Me.lblCompany)
        Me.fraMainFrame.Controls.Add(Me.lblCopyright)
        Me.fraMainFrame.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraMainFrame.Location = New System.Drawing.Point(3, -1)
        Me.fraMainFrame.Name = "fraMainFrame"
        Me.fraMainFrame.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraMainFrame.Size = New System.Drawing.Size(492, 281)
        Me.fraMainFrame.TabIndex = 0
        Me.fraMainFrame.TabStop = False
        '
        'picLogo
        '
        Me.picLogo.BackColor = System.Drawing.SystemColors.Control
        Me.picLogo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.picLogo.Cursor = System.Windows.Forms.Cursors.Default
        Me.picLogo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.picLogo.Image = CType(resources.GetObject("picLogo.Image"), System.Drawing.Image)
        Me.picLogo.Location = New System.Drawing.Point(16, 56)
        Me.picLogo.Name = "picLogo"
        Me.picLogo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.picLogo.Size = New System.Drawing.Size(144, 143)
        Me.picLogo.TabIndex = 2
        '
        'lblLicenseTo
        '
        Me.lblLicenseTo.BackColor = System.Drawing.SystemColors.Control
        Me.lblLicenseTo.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblLicenseTo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLicenseTo.Location = New System.Drawing.Point(18, 20)
        Me.lblLicenseTo.Name = "lblLicenseTo"
        Me.lblLicenseTo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblLicenseTo.Size = New System.Drawing.Size(457, 17)
        Me.lblLicenseTo.TabIndex = 1
        Me.lblLicenseTo.Tag = "1055"
        Me.lblLicenseTo.Text = "LicenseTo"
        Me.lblLicenseTo.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblProductName
        '
        Me.lblProductName.AutoSize = True
        Me.lblProductName.BackColor = System.Drawing.SystemColors.Control
        Me.lblProductName.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblProductName.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblProductName.Location = New System.Drawing.Point(178, 80)
        Me.lblProductName.Name = "lblProductName"
        Me.lblProductName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblProductName.Size = New System.Drawing.Size(52, 16)
        Me.lblProductName.TabIndex = 9
        Me.lblProductName.Tag = "1054"
        Me.lblProductName.Text = "GestCom"
        '
        'lblCompanyProduct
        '
        Me.lblCompanyProduct.AutoSize = True
        Me.lblCompanyProduct.BackColor = System.Drawing.SystemColors.Control
        Me.lblCompanyProduct.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCompanyProduct.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCompanyProduct.Location = New System.Drawing.Point(167, 51)
        Me.lblCompanyProduct.Name = "lblCompanyProduct"
        Me.lblCompanyProduct.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCompanyProduct.Size = New System.Drawing.Size(45, 16)
        Me.lblCompanyProduct.TabIndex = 8
        Me.lblCompanyProduct.Tag = "1053"
        Me.lblCompanyProduct.Text = "Vinicom"
        '
        'lblPlatform
        '
        Me.lblPlatform.AutoSize = True
        Me.lblPlatform.BackColor = System.Drawing.SystemColors.Control
        Me.lblPlatform.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblPlatform.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblPlatform.Location = New System.Drawing.Point(391, 160)
        Me.lblPlatform.Name = "lblPlatform"
        Me.lblPlatform.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblPlatform.Size = New System.Drawing.Size(42, 16)
        Me.lblPlatform.TabIndex = 7
        Me.lblPlatform.Tag = "1052"
        Me.lblPlatform.Text = "Win XP"
        Me.lblPlatform.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblVersion
        '
        Me.lblVersion.AutoSize = True
        Me.lblVersion.BackColor = System.Drawing.SystemColors.Control
        Me.lblVersion.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblVersion.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVersion.Location = New System.Drawing.Point(405, 184)
        Me.lblVersion.Name = "lblVersion"
        Me.lblVersion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblVersion.Size = New System.Drawing.Size(20, 16)
        Me.lblVersion.TabIndex = 6
        Me.lblVersion.Tag = "1051"
        Me.lblVersion.Text = "1.0"
        Me.lblVersion.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblWarning
        '
        Me.lblWarning.BackColor = System.Drawing.SystemColors.Control
        Me.lblWarning.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblWarning.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblWarning.Location = New System.Drawing.Point(16, 208)
        Me.lblWarning.Name = "lblWarning"
        Me.lblWarning.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblWarning.Size = New System.Drawing.Size(292, 64)
        Me.lblWarning.TabIndex = 3
        Me.lblWarning.Tag = "1050"
        Me.lblWarning.Text = "Warning"
        '
        'lblCompany
        '
        Me.lblCompany.BackColor = System.Drawing.SystemColors.Control
        Me.lblCompany.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCompany.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCompany.Location = New System.Drawing.Point(314, 224)
        Me.lblCompany.Name = "lblCompany"
        Me.lblCompany.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCompany.Size = New System.Drawing.Size(102, 17)
        Me.lblCompany.TabIndex = 5
        Me.lblCompany.Tag = "1049"
        Me.lblCompany.Text = "MCII"
        '
        'lblCopyright
        '
        Me.lblCopyright.BackColor = System.Drawing.SystemColors.Control
        Me.lblCopyright.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCopyright.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCopyright.Location = New System.Drawing.Point(314, 208)
        Me.lblCopyright.Name = "lblCopyright"
        Me.lblCopyright.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCopyright.Size = New System.Drawing.Size(86, 17)
        Me.lblCopyright.TabIndex = 4
        Me.lblCopyright.Tag = "1048"
        Me.lblCopyright.Text = "Copyright"
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(432, 208)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(48, 48)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 10
        Me.PictureBox1.TabStop = False
        '
        'frmSplash
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(497, 298)
        Me.ControlBox = False
        Me.Controls.Add(Me.fraMainFrame)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(3, 3)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmSplash"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.fraMainFrame.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
#End Region 
#Region "Prise en charge de la mise à niveau "
	Private Shared m_vb6FormDefInstance As frmSplash
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As frmSplash
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New frmSplash()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	
	Private Sub frmSplash_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        'LoadResString(Me)
		'UPGRADE_ISSUE: App propriété App.Revision - Mise à niveau non effectuée Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2069"'
        lblVersion.Text = "Version " & System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).FileMajorPart & "." & System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).FileMinorPart & "."
    End Sub
End Class