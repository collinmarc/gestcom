'===================================================================================================================================
'Projet : Vinicom
'Auteur : Marc Collin 
'Création : 23/06/04
'===================================================================================================================================
' Classe : EtatReglement_Exporte
' Description : Reglement Saisi et exporté
'===================================================================================================================================
'Membres de Classes
'==================
'Public
'Protected
'Private
'Membres d'instances
'==================
'Public
'   toString()      : Rend 'objet sous forme de chaine
'   Equals()        : Test l'équivalence d'instance
'Protected
'Private
'===================================================================================================================================
'Historique :
'============
'
'===================================================================================================================================
Public Class EtatReglement_Exporte
    Inherits EtatReglement

#Region "Accesseurs"

    Public Sub New()
        MyBase.New()
        codeEtat = vncEnums.vncEtatReglement.vncRGLMT_Export
        m_libelle = ETAT_RGLMT_EXPORTE
    End Sub

#End Region
#Region "MéthodesDeClasse"
#End Region
    '=======================================================================
    'Fonction : ToString()
    'Description : Rend l'objet sous forme de String
    'Détails    :  
    'Retour : une Chaine
    '=======================================================================
    Public Overrides Function toString() As String
        Return "ETA =(" & m_codeEtat & ")"
    End Function 'ToString
    '=======================================================================
    'Fonction : Equals()
    'Description : Teste l'équivalence entre 2 objets
    'Détails    :  
    'Retour : Vari si l'objet passé en paramétre est equivalent à 
    '=======================================================================
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        Dim bReturn As Boolean
        bReturn = obj.GetType.Name.Equals(Me.GetType().Name)
        Return bReturn
    End Function

    Public Overrides ReadOnly Property shortResume() As String
        Get
            Return m_codeEtat
        End Get
    End Property
    ''' <summary>
    ''' Rend un état en fonction d'une action
    ''' 
    ''' Action = Annulation Export => etat = Saisi
    ''' </summary>
    ''' <param name="vncAction"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Overrides Function action(ByVal vncAction As vncActionReglement) As vncEtatReglement
        Dim nReturn As vncEtatReglement
        Select Case vncAction
            Case vncEnums.vncActionReglement.vncActionAnnExporter
                nReturn = vncEnums.vncEtatReglement.vncRGLMT_Saisi
            Case Else
                nReturn = m_codeEtat
        End Select
        Return nReturn
    End Function

End Class
