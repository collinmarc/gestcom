Imports FAXCOMLib
Imports vini_DB
Public Class clsFax
    Inherits racine
    Private Shared m_oFaxServer As FaxServer = New FaxServer
    Friend Shared m_strCoverPageName As String
    Friend Shared m_bSendCoverpage As Boolean
    Friend Shared m_SenderCompany As String
    Friend Shared m_SenderAddress As String
    Friend Shared m_SenderName As String
    Friend Shared m_SenderFax As String
    Friend Shared m_SenderTel As String
    Private Shared m_binit As Boolean = False
    Friend Shared m_prefix As String
    Friend Shared m_ServerName As String



    'Public Sub New()
    '    Dim objPersist As Persist

    '    If Not m_binit Then
    '        Persist.shared_connect()
    '        Persist.initFax()
    '        Persist.shared_disconnect()
    '        Try
    '            m_oFaxServer.Connect(m_ServerName)
    '            m_binit = True
    '        Catch ex As Exception
    '            m_binit = False
    '            setError("clsFax.new", ex.ToString)
    '        End Try
    '    End If
    'End Sub
    'Protected Overrides Sub Finalize()
    '    Try
    '        m_oFaxServer.Disconnect()
    '    Catch ex As Exception
    '    End Try
    '    m_binit = False
    '    MyBase.Finalize()
    'End Sub

    Public Function sendFax(ByVal strNomInterlocteur As String, ByVal strTelInterlocuteur As String, ByVal strSubject As String, ByVal strNotes As String, ByVal bSendCoverPage As Boolean, ByVal strfilename As String, ByVal strFaxNumber As String, Optional ByVal objTiers As Tiers = Nothing, Optional ByVal bAdresseLivr As Boolean = True) As Integer
        Dim objFaxDoc As FaxDoc
        Dim oAdresse As Adresse
        Dim nJobId As Integer

        'If Not m_binit Then
        '    Return -1
        'End If


        Try

            'Test : Achaque Fax on se connecte et on se deconnecte
            Persist.shared_connect()
            Persist.initFax()
            Persist.shared_disconnect()
            m_oFaxServer.Connect("")

            objFaxDoc = m_oFaxServer.CreateDocument("Mydoc")
            objFaxDoc.SenderCompany = m_SenderName & " "
            objFaxDoc.SenderName = strNomInterlocteur & " "
            objFaxDoc.SenderOfficePhone = strTelInterlocuteur & " "
            objFaxDoc.SenderFax = m_SenderFax & " "

            objFaxDoc.FaxNumber = m_prefix & strFaxNumber
            objFaxDoc.FileName = strfilename
            If Not objTiers Is Nothing Then
                If bAdresseLivr Then
                    oAdresse = objTiers.AdresseLivraison
                Else
                    oAdresse = objTiers.AdresseFacturation
                End If
                objFaxDoc.RecipientAddress = oAdresse.rue1 & vbCrLf & oAdresse.rue2
                objFaxDoc.RecipientCity = oAdresse.ville
                objFaxDoc.RecipientName = oAdresse.nom
                objFaxDoc.RecipientCompany = objTiers.rs
                objFaxDoc.RecipientZip = oAdresse.cp
                objFaxDoc.RecipientTitle = oAdresse.nom
            End If
            objFaxDoc.SendCoverpage = bSendCoverPage
            objFaxDoc.CoverpageName = m_strCoverPageName
            objFaxDoc.CoverpageSubject = strSubject
            objFaxDoc.CoverpageNote = strNotes
            nJobId = objFaxDoc.Send

            'Test : Achaque Fax on se connecte et on se deconnecte
            m_oFaxServer.Disconnect()

        Catch ex As Exception
            nJobId = -1
            setError("clsFax.SendFax", ex.ToString)
        End Try

        Return nJobId

    End Function 'SendFax


    Public Overrides Function toString() As String
        Return m_SenderCompany
    End Function

    Public Overrides ReadOnly Property shortResume() As String
        Get
            Return String.Empty
        End Get
    End Property
End Class
