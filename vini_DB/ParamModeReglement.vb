'===================================================================================================================================
'Projet : Vinicom
'Auteur : Marc Collin 
'Création : 23/06/04
'===================================================================================================================================
' Classe : <<NomClasse >>
' Description : <<Commentaire
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
Public Class ParamModeReglement
    Inherits Param
    Private m_DebutEcheance As String

    Public Sub New()
        MyBase.new()
        type = PAR_MODE_RGLMT
        m_DebutEcheance = String.Empty
    End Sub

    Public Property dDebutEcheance() As String
        Get
            Return m_DebutEcheance
        End Get
        Set(ByVal value As String)
            If Not m_DebutEcheance.Equals(value) Or value Is Nothing Then
                m_DebutEcheance = value
                RaiseUpdated()
            End If
        End Set
    End Property
    '=======================================================================
    'Fonction : ToString()
    'Description : Rend l'objet sous forme de String
    'Détails    :  
    'Retour : une Chaine
    '=======================================================================
    Public Overrides Function toString() As String
        Return MyBase.toString()
    End Function

    '=======================================================================
    'Fonction : Equals()
    'Description : Teste l'équivalence entre 2 objets
    'Détails    :  
    'Retour : Vari si l'objet passé en paramétre est equivalent à 
    '=======================================================================
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        Return MyBase.Equals(obj)
    End Function


    '=======================================================================
    'Fonction : shortResume()
    'Description : Rend un resumé de l'objet sous forme de chaine
    'Détails    :  
    'Retour : Vari si l'objet passé en paramétre est equivalent à 
    '=======================================================================
    Public Overrides ReadOnly Property shortResume() As String
        Get
            Return ""
        End Get
    End Property
    Friend Overrides Function fromRS(ByVal poRS As OleDb.OleDbDataReader) As Boolean
        Dim bReturn As Boolean
        Try

            MyBase.fromRS(poRS)

            Try
                Me.dDebutEcheance = GetString(poRS, "PAR_DDEB_ECHEANCE")
            Catch ex As InvalidCastException
                Me.dDebutEcheance = String.Empty
            End Try

            bReturn = True
        Catch ex As Exception
            setError(ex.Message)
            bReturn = False
        End Try
        Return bReturn
    End Function

    Public Function calDateEcheance(ByVal ddeb As Date) As Date
        Dim dReturn As Date
        dReturn = ddeb
        Try
            Select Case dDebutEcheance
                Case "FDM"
                    'calcul de la Fin du mois
                    '1er du mois suivant
                    Select Case valeur2
                        Case 30
                            dReturn = "01" & "/" & ddeb.AddMonths(2).Month & "/" & ddeb.AddMonths(2).Year
                        Case 60
                            dReturn = "01" & "/" & ddeb.AddMonths(3).Month & "/" & ddeb.AddMonths(3).Year
                        Case 90
                            dReturn = "01" & "/" & ddeb.AddMonths(4).Month & "/" & ddeb.AddMonths(4).Year
                        Case Else
                            dReturn = "01" & "/" & ddeb.AddMonths(1).Month & "/" & ddeb.AddMonths(1).Year
                            dReturn = dReturn.AddDays(valeur2)
                    End Select
                    'Jour Précédant 
                    dReturn = dReturn.AddDays(-1)
                Case "FACT"
                    dReturn = dReturn
                    Select Case valeur2
                        Case 30
                            dReturn = dReturn.AddMonths(1)
                        Case 60
                            dReturn = dReturn.AddMonths(2)
                        Case 90
                            dReturn = dReturn.AddMonths(3)
                        Case Else
                            dReturn = dReturn.AddDays(valeur2)
                    End Select
            End Select

        Catch ex As Exception
            setError("ParamModeReglement", ex.Message)
        End Try
        Return dReturn
    End Function
End Class
