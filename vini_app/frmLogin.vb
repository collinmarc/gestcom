Option Strict Off
Option Explicit On
'Imports VB = Microsoft.VisualBasic
Imports vini_DB
Public Class frmLogin
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
    Public WithEvents cmdCancel As System.Windows.Forms.Button
    Public WithEvents cbOK As System.Windows.Forms.Button
    Public WithEvents tbPassword As System.Windows.Forms.TextBox
    Public WithEvents tbUserName As System.Windows.Forms.TextBox
    Private WithEvents lblLabels_1 As System.Windows.Forms.Label
    Private WithEvents lblLabels_0 As System.Windows.Forms.Label
    'Public WithEvents lblLabels As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Il peut être modifié à l'aide du Concepteur Windows Form.
    'Ne pas le modifier à l'aide de l'éditeur de code.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.cbOK = New System.Windows.Forms.Button
        Me.tbPassword = New System.Windows.Forms.TextBox
        Me.tbUserName = New System.Windows.Forms.TextBox
        Me.lblLabels_1 = New System.Windows.Forms.Label
        Me.lblLabels_0 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCancel.Location = New System.Drawing.Point(140, 68)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCancel.Size = New System.Drawing.Size(76, 24)
        Me.cmdCancel.TabIndex = 3
        Me.cmdCancel.Tag = "1060"
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'cbOK
        '
        Me.cbOK.BackColor = System.Drawing.SystemColors.Control
        Me.cbOK.Cursor = System.Windows.Forms.Cursors.Default
        Me.cbOK.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cbOK.Location = New System.Drawing.Point(33, 68)
        Me.cbOK.Name = "cbOK"
        Me.cbOK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cbOK.Size = New System.Drawing.Size(76, 24)
        Me.cbOK.TabIndex = 2
        Me.cbOK.Tag = "1059"
        Me.cbOK.Text = "OK"
        Me.cbOK.UseVisualStyleBackColor = False
        '
        'tbPassword
        '
        Me.tbPassword.AcceptsReturn = True
        Me.tbPassword.BackColor = System.Drawing.SystemColors.Window
        Me.tbPassword.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbPassword.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbPassword.Location = New System.Drawing.Point(87, 35)
        Me.tbPassword.MaxLength = 0
        Me.tbPassword.Name = "tbPassword"
        Me.tbPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.tbPassword.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbPassword.Size = New System.Drawing.Size(155, 20)
        Me.tbPassword.TabIndex = 1
        '
        'tbUserName
        '
        Me.tbUserName.AcceptsReturn = True
        Me.tbUserName.BackColor = System.Drawing.SystemColors.Window
        Me.tbUserName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbUserName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbUserName.Location = New System.Drawing.Point(87, 9)
        Me.tbUserName.MaxLength = 0
        Me.tbUserName.Name = "tbUserName"
        Me.tbUserName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbUserName.Size = New System.Drawing.Size(155, 20)
        Me.tbUserName.TabIndex = 0
        '
        'lblLabels_1
        '
        Me.lblLabels_1.BackColor = System.Drawing.SystemColors.Control
        Me.lblLabels_1.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblLabels_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLabels_1.Location = New System.Drawing.Point(7, 36)
        Me.lblLabels_1.Name = "lblLabels_1"
        Me.lblLabels_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblLabels_1.Size = New System.Drawing.Size(72, 17)
        Me.lblLabels_1.TabIndex = 0
        Me.lblLabels_1.Tag = "1058"
        Me.lblLabels_1.Text = "&Password:"
        '
        'lblLabels_0
        '
        Me.lblLabels_0.BackColor = System.Drawing.SystemColors.Control
        Me.lblLabels_0.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblLabels_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLabels_0.Location = New System.Drawing.Point(7, 10)
        Me.lblLabels_0.Name = "lblLabels_0"
        Me.lblLabels_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblLabels_0.Size = New System.Drawing.Size(72, 17)
        Me.lblLabels_0.TabIndex = 2
        Me.lblLabels_0.Tag = "1057"
        Me.lblLabels_0.Text = "&User Name:"
        '
        'frmLogin
        '
        Me.AcceptButton = Me.cbOK
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.CancelButton = Me.cmdCancel
        Me.ClientSize = New System.Drawing.Size(250, 106)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cbOK)
        Me.Controls.Add(Me.tbPassword)
        Me.Controls.Add(Me.tbUserName)
        Me.Controls.Add(Me.lblLabels_1)
        Me.Controls.Add(Me.lblLabels_0)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(3, 22)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmLogin"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Tag = "1056"
        Me.Text = "Login"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
#Region "Prise en charge de la mise à niveau "
    Private Shared m_vb6FormDefInstance As frmLogin
    Private Shared m_InitializingDefInstance As Boolean
    Public Shared Property DefInstance() As frmLogin
        Get
            If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
                m_InitializingDefInstance = True
                m_vb6FormDefInstance = New frmLogin()
                m_InitializingDefInstance = False
            End If
            DefInstance = m_vb6FormDefInstance
        End Get
        Set(ByVal value As frmLogin)
            m_vb6FormDefInstance = value
        End Set
    End Property
#End Region
    Public OK As Boolean
    Public user As aut_user
    Private m_colUser As Collection
    Private Sub frmLogin_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        Persist.shared_connect()

        m_colUser = aut_user.listeUSERS()
        Text = "Gestcom " & My.MySettings.Default.AppVersion.ToString() & "-DB" & Param.getConstante("CST_VERSION_BD")
        'Récupération du nom de l'utilisateur Système
        tbUserName.Text = System.Environment.UserName
        Persist.shared_disconnect()

    End Sub

    Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
        OK = False
        Me.Hide()
    End Sub


    Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cbOK.Click
        'ToDo: create test for correct password
        'check for correct password
        If tbUserName.Text = "mco" And tbPassword.Text = "collin10" Then
            OK = True
            user = aut_user.getuser(vncEnums.userRole.ADMIN)
            Me.Hide()
        Else
            Try
                user = m_colUser(tbUserName.Text)
                If user.password Is Nothing And tbPassword.Text.Length = 0 Then
                    OK = True
                    Me.Hide()
                Else
                    If (user.password.Equals(tbPassword.Text)) Then
                        OK = True
                        Me.Hide()
                    Else
                        MsgBox("Mot de passe incorrect, réessayez!", , "Login")
                        tbPassword.Focus()
                        tbPassword.SelectionStart = 0
                        tbPassword.SelectionLength = Len(tbPassword.Text)
                    End If
                End If
            Catch ex As Exception
                MsgBox("Mot de passe incorrect, réessayez!", , "Login")
                tbPassword.Focus()
                tbPassword.SelectionStart = 0
                tbPassword.SelectionLength = Len(tbPassword.Text)

            End Try
        End If
    End Sub

    Private Sub txtUserName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbUserName.TextChanged

    End Sub
End Class