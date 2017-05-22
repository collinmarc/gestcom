'===================================================================================================================================
'Projet : Vinicom
'Auteur : Marc Collin 
'Cr�ation : 23/06/04
'===================================================================================================================================
' Classe : EtatReglement_Saisi
' Description : Reglement Saisi mais non export�
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
'   Equals()        : Test l'�quivalence d'instance
'Protected
'Private
'===================================================================================================================================
'Historique :
'============
'
'===================================================================================================================================
Public Class EtatReglement_Saisi
    Inherits EtatReglement

#Region "Accesseurs"

    Public Sub New()
        MyBase.New()
        codeEtat = vncEnums.vncEtatReglement.vncRGLMT_Saisi
        m_libelle = ETAT_RGLMT_SAISI
    End Sub

#End Region
#Region "M�thodesDeClasse"
#End Region
    '=======================================================================
    'Fonction : ToString()
    'Description : Rend l'objet sous forme de String
    'D�tails    :  
    'Retour : une Chaine
    '=======================================================================
    Public Overrides Function toString() As String
        Return "ETA =(" & m_codeEtat & ")"
    End Function 'ToString
    '=======================================================================
    'Fonction : Equals()
    'Description : Teste l'�quivalence entre 2 objets
    'D�tails    :  
    'Retour : Vari si l'objet pass� en param�tre est equivalent � 
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
    ''' Rend un �tat en fonction d'une action
    ''' 
    ''' Action = Export => etat = Exporte
    ''' </summary>
    ''' <param name="vncAction"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Overrides Function action(ByVal vncAction As vncActionReglement) As vncEtatReglement
        Dim nReturn As vncEtatReglement
        Select Case vncAction
            Case vncEnums.vncActionReglement.vncActionExporter
                nReturn = vncEnums.vncEtatReglement.vncRGLMT_Export
            Case Else
                nReturn = m_codeEtat
        End Select
        Return nReturn
    End Function

End Class
