<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAppelation
    Inherits FrmDonBase

    'Form remplace la méthode Dispose pour nettoyer la liste des composants.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requise par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer

    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Elle peut être modifiée à l'aide du Concepteur Windows Form.  
    'Ne la modifiez pas à l'aide de l'éditeur de code.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAppelation))
        Me.DsVinicom1 = New vini_DB.dsVinicom()
        Me.m_bsrcAppelation = New System.Windows.Forms.BindingSource(Me.components)
        Me.AppelationTableAdapter1 = New vini_DB.dsVinicomTableAdapters.APPELATIONTableAdapter()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.BindingNavigator1 = New System.Windows.Forms.BindingNavigator(Me.components)
        Me.BindingNavigatorAddNewItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorCountItem = New System.Windows.Forms.ToolStripLabel()
        Me.BindingNavigatorDeleteItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorMoveFirstItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorMovePreviousItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorSeparator = New System.Windows.Forms.ToolStripSeparator()
        Me.BindingNavigatorPositionItem = New System.Windows.Forms.ToolStripTextBox()
        Me.BindingNavigatorSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.BindingNavigatorMoveNextItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorMoveLastItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.ofd_Image = New System.Windows.Forms.OpenFileDialog()
        Me.TextBox4 = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.RichTextBoxJmb1 = New vini_app.RichTextBoxJmb()
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox()
        CType(Me.DsVinicom1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_bsrcAppelation, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.BindingNavigator1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.BindingNavigator1.SuspendLayout()
        Me.SuspendLayout()
        '
        'DsVinicom1
        '
        Me.DsVinicom1.DataSetName = "dsVinicom"
        Me.DsVinicom1.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'm_bsrcAppelation
        '
        Me.m_bsrcAppelation.DataMember = "APPELATION"
        Me.m_bsrcAppelation.DataSource = Me.DsVinicom1
        '
        'AppelationTableAdapter1
        '
        Me.AppelationTableAdapter1.ClearBeforeFill = True
        '
        'TextBox2
        '
        Me.TextBox2.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcAppelation, "APP_CHATEAU", True))
        Me.TextBox2.Location = New System.Drawing.Point(84, 57)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(439, 20)
        Me.TextBox2.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(16, 60)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(47, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Chateau"
        '
        'TextBox3
        '
        Me.TextBox3.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcAppelation, "APP_APPELATION", True))
        Me.TextBox3.Location = New System.Drawing.Point(84, 99)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(438, 20)
        Me.TextBox3.TabIndex = 4
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(20, 103)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(26, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "App"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(17, 153)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(51, 13)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Descriptif"
        '
        'PictureBox1
        '
        Me.PictureBox1.DataBindings.Add(New System.Windows.Forms.Binding("Image", Me.m_bsrcAppelation, "APP_ETIQUETTE", True))
        Me.PictureBox1.Location = New System.Drawing.Point(91, 289)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(208, 205)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 8
        Me.PictureBox1.TabStop = False
        '
        'BindingNavigator1
        '
        Me.BindingNavigator1.AddNewItem = Me.BindingNavigatorAddNewItem
        Me.BindingNavigator1.BindingSource = Me.m_bsrcAppelation
        Me.BindingNavigator1.CountItem = Me.BindingNavigatorCountItem
        Me.BindingNavigator1.DeleteItem = Me.BindingNavigatorDeleteItem
        Me.BindingNavigator1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.BindingNavigator1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BindingNavigatorMoveFirstItem, Me.BindingNavigatorMovePreviousItem, Me.BindingNavigatorSeparator, Me.BindingNavigatorPositionItem, Me.BindingNavigatorCountItem, Me.BindingNavigatorSeparator1, Me.BindingNavigatorMoveNextItem, Me.BindingNavigatorMoveLastItem, Me.BindingNavigatorSeparator2, Me.BindingNavigatorAddNewItem, Me.BindingNavigatorDeleteItem})
        Me.BindingNavigator1.Location = New System.Drawing.Point(0, 515)
        Me.BindingNavigator1.MoveFirstItem = Me.BindingNavigatorMoveFirstItem
        Me.BindingNavigator1.MoveLastItem = Me.BindingNavigatorMoveLastItem
        Me.BindingNavigator1.MoveNextItem = Me.BindingNavigatorMoveNextItem
        Me.BindingNavigator1.MovePreviousItem = Me.BindingNavigatorMovePreviousItem
        Me.BindingNavigator1.Name = "BindingNavigator1"
        Me.BindingNavigator1.PositionItem = Me.BindingNavigatorPositionItem
        Me.BindingNavigator1.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional
        Me.BindingNavigator1.Size = New System.Drawing.Size(1188, 25)
        Me.BindingNavigator1.TabIndex = 9
        Me.BindingNavigator1.Text = "BindingNavigator1"
        '
        'BindingNavigatorAddNewItem
        '
        Me.BindingNavigatorAddNewItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorAddNewItem.Image = CType(resources.GetObject("BindingNavigatorAddNewItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorAddNewItem.Name = "BindingNavigatorAddNewItem"
        Me.BindingNavigatorAddNewItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorAddNewItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorAddNewItem.Text = "Ajouter nouveau"
        '
        'BindingNavigatorCountItem
        '
        Me.BindingNavigatorCountItem.Name = "BindingNavigatorCountItem"
        Me.BindingNavigatorCountItem.Size = New System.Drawing.Size(38, 22)
        Me.BindingNavigatorCountItem.Text = "de {0}"
        Me.BindingNavigatorCountItem.ToolTipText = "Nombre total d'éléments"
        '
        'BindingNavigatorDeleteItem
        '
        Me.BindingNavigatorDeleteItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorDeleteItem.Image = CType(resources.GetObject("BindingNavigatorDeleteItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorDeleteItem.Name = "BindingNavigatorDeleteItem"
        Me.BindingNavigatorDeleteItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorDeleteItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorDeleteItem.Text = "Supprimer"
        '
        'BindingNavigatorMoveFirstItem
        '
        Me.BindingNavigatorMoveFirstItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveFirstItem.Image = CType(resources.GetObject("BindingNavigatorMoveFirstItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveFirstItem.Name = "BindingNavigatorMoveFirstItem"
        Me.BindingNavigatorMoveFirstItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveFirstItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorMoveFirstItem.Text = "Placer en premier"
        '
        'BindingNavigatorMovePreviousItem
        '
        Me.BindingNavigatorMovePreviousItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMovePreviousItem.Image = CType(resources.GetObject("BindingNavigatorMovePreviousItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMovePreviousItem.Name = "BindingNavigatorMovePreviousItem"
        Me.BindingNavigatorMovePreviousItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMovePreviousItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorMovePreviousItem.Text = "Déplacer vers le haut"
        '
        'BindingNavigatorSeparator
        '
        Me.BindingNavigatorSeparator.Name = "BindingNavigatorSeparator"
        Me.BindingNavigatorSeparator.Size = New System.Drawing.Size(6, 25)
        '
        'BindingNavigatorPositionItem
        '
        Me.BindingNavigatorPositionItem.AccessibleName = "Position"
        Me.BindingNavigatorPositionItem.AutoSize = False
        Me.BindingNavigatorPositionItem.Name = "BindingNavigatorPositionItem"
        Me.BindingNavigatorPositionItem.Size = New System.Drawing.Size(50, 21)
        Me.BindingNavigatorPositionItem.Text = "0"
        Me.BindingNavigatorPositionItem.ToolTipText = "Position actuelle"
        '
        'BindingNavigatorSeparator1
        '
        Me.BindingNavigatorSeparator1.Name = "BindingNavigatorSeparator1"
        Me.BindingNavigatorSeparator1.Size = New System.Drawing.Size(6, 25)
        '
        'BindingNavigatorMoveNextItem
        '
        Me.BindingNavigatorMoveNextItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveNextItem.Image = CType(resources.GetObject("BindingNavigatorMoveNextItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveNextItem.Name = "BindingNavigatorMoveNextItem"
        Me.BindingNavigatorMoveNextItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveNextItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorMoveNextItem.Text = "Déplacer vers le bas"
        '
        'BindingNavigatorMoveLastItem
        '
        Me.BindingNavigatorMoveLastItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveLastItem.Image = CType(resources.GetObject("BindingNavigatorMoveLastItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveLastItem.Name = "BindingNavigatorMoveLastItem"
        Me.BindingNavigatorMoveLastItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveLastItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorMoveLastItem.Text = "Placer en dernier"
        '
        'BindingNavigatorSeparator2
        '
        Me.BindingNavigatorSeparator2.Name = "BindingNavigatorSeparator2"
        Me.BindingNavigatorSeparator2.Size = New System.Drawing.Size(6, 25)
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(17, 292)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(36, 13)
        Me.Label5.TabIndex = 10
        Me.Label5.Text = "Image"
        '
        'ofd_Image
        '
        Me.ofd_Image.Filter = "*.jpg|"
        '
        'TextBox4
        '
        Me.TextBox4.Location = New System.Drawing.Point(315, 474)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(208, 20)
        Me.TextBox4.TabIndex = 11
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(529, 474)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(36, 23)
        Me.Button1.TabIndex = 12
        Me.Button1.Text = "..."
        Me.Button1.UseVisualStyleBackColor = True
        '
        'RichTextBoxJmb1
        '
        Me.RichTextBoxJmb1.DetectURLs = True
        Me.RichTextBoxJmb1.Location = New System.Drawing.Point(111, 163)
        Me.RichTextBoxJmb1.Name = "RichTextBoxJmb1"
        Me.RichTextBoxJmb1.ShowBold = True
        Me.RichTextBoxJmb1.ShowCenterJustify = True
        Me.RichTextBoxJmb1.ShowColors = True
        Me.RichTextBoxJmb1.ShowFont = True
        Me.RichTextBoxJmb1.ShowFontSize = True
        Me.RichTextBoxJmb1.ShowItalic = True
        Me.RichTextBoxJmb1.ShowLeftJustify = True
        Me.RichTextBoxJmb1.Shownormal = True
        Me.RichTextBoxJmb1.ShowOpen = True
        Me.RichTextBoxJmb1.ShowPreview = True
        Me.RichTextBoxJmb1.ShowPrint = True
        Me.RichTextBoxJmb1.ShowQuit = True
        Me.RichTextBoxJmb1.ShowRedo = True
        Me.RichTextBoxJmb1.ShowRightJustify = True
        Me.RichTextBoxJmb1.ShowSave = True
        Me.RichTextBoxJmb1.ShowStamp = True
        Me.RichTextBoxJmb1.ShowStrikeout = True
        Me.RichTextBoxJmb1.ShowUnderline = True
        Me.RichTextBoxJmb1.ShowUndo = True
        Me.RichTextBoxJmb1.Size = New System.Drawing.Size(428, 120)
        Me.RichTextBoxJmb1.StampAction = vini_app.RichTextBoxJmb.StampActions.Editepar
        Me.RichTextBoxJmb1.StampColor = System.Drawing.Color.Blue
        Me.RichTextBoxJmb1.TabIndex = 13
        '
        'RichTextBox1
        '
        Me.RichTextBox1.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcAppelation, "APP_DESCRIPTIF", True))
        Me.RichTextBox1.Location = New System.Drawing.Point(572, 204)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.Size = New System.Drawing.Size(100, 96)
        Me.RichTextBox1.TabIndex = 14
        Me.RichTextBox1.Text = ""
        '
        'frmAppelation
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1188, 540)
        Me.Controls.Add(Me.RichTextBox1)
        Me.Controls.Add(Me.RichTextBoxJmb1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.TextBox4)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.BindingNavigator1)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.TextBox3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TextBox2)
        Me.Name = "frmAppelation"
        Me.Text = "Gestion des appelations"
        CType(Me.DsVinicom1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.m_bsrcAppelation, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.BindingNavigator1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.BindingNavigator1.ResumeLayout(False)
        Me.BindingNavigator1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents DsVinicom1 As vini_DB.dsVinicom
    Friend WithEvents AppelationTableAdapter1 As vini_DB.dsVinicomTableAdapters.APPELATIONTableAdapter
    Friend WithEvents m_bsrcAppelation As System.Windows.Forms.BindingSource
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents BindingNavigator1 As System.Windows.Forms.BindingNavigator
    Friend WithEvents BindingNavigatorAddNewItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorCountItem As System.Windows.Forms.ToolStripLabel
    Friend WithEvents BindingNavigatorDeleteItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorMoveFirstItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorMovePreviousItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorSeparator As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents BindingNavigatorPositionItem As System.Windows.Forms.ToolStripTextBox
    Friend WithEvents BindingNavigatorSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents BindingNavigatorMoveNextItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorMoveLastItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents ofd_Image As System.Windows.Forms.OpenFileDialog
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents RichTextBoxJmb1 As RichTextBoxJmb
    Friend WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
End Class
