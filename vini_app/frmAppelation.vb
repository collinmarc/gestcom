Imports vini_DB
Imports Microsoft.Office.Interop
Public Class frmAppelation
    Inherits FrmDonBase

    Protected Overrides Sub setToolbarButtons()
        m_ToolBarNewEnabled = False
        m_ToolBarLoadEnabled = False
        m_ToolBarDelEnabled = False
        m_ToolBarRefreshEnabled = False
        m_ToolBarSaveEnabled = True
    End Sub

    Protected Overrides Sub EnableControls(bEnabled As Boolean)
        MyBase.EnableControls(True)
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        If (ofd_Image.ShowDialog()) Then
            PictureBox1.Image = Image.FromFile(ofd_Image.FileName)
        End If

    End Sub
    Protected Overrides Function frmSave() As Boolean
        AppelationTableAdapter1.Update(DsVinicom1.APPELATION)
    End Function

    Private Sub frmAppelation_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        AppelationTableAdapter1.Fill(DsVinicom1.APPELATION)
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs)
        ' CreateNewDocWord("C:/TEMP/a.doc")
    End Sub
    'Private Sub CreateNewDocWord(sDoc As String)

    '    Dim objWord As Word.Application
    '    Dim docWord As Word.Document

    '    objWord = CreateObject("Word.Application")    '-- ouvrir une session Word
    '    docWord = objWord.Documents.Add     '-- Ajouter un nouveau document à la collection
    '    objWord.Visible = True    '-- montrer l'application Word
    '    docWord.SaveAs(FileName:=sDoc)

    '    docWord = Nothing    '-- détruire l'objet Document
    '    objWord = Nothing    '-- détruire l'objet Word
    'End Sub
End Class