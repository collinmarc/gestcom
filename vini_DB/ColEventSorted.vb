'===================================================================================================================================
'Projet : Vinicom
'Auteur : Marc Collin 
'Création : 23/06/04
'===================================================================================================================================
' Classe : colEvent
' Description : Collection qui gere l'evenment Updated
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

Public Class ColEventSorted
    Inherits racine
    Private m_col As SortedList

    '=======================================================================
    'Fonction : Add()
    'Description : Ajoute un Element à la collection et arme l'Event HandlerUpdated
    'Détails    :  
    'Retour : 
    '=======================================================================
    Public Sub Add(ByVal obj As racine, Optional ByVal sKey As String = "")
        Dim nindex As Integer
        Dim bindexTrouve As Boolean
        Dim skeyorigine As String
        If Trim(sKey) = "" Then
            sKey = CStr(m_col.Count)
        End If

        If m_col.ContainsKey(CStr(sKey)) Then
            skeyorigine = sKey
            bindexTrouve = False
            For nindex = 0 To 100000
                sKey = skeyorigine & "_" & nindex
                If Not m_col.ContainsKey(CStr(sKey)) Then
                    bindexTrouve = True
                    Exit For
                End If
            Next
            Debug.Assert(bindexTrouve, "Aucune clé possible")
        End If
        m_col.Add(sKey, obj)
        AddHandler obj.Updated, AddressOf MyHandler
        RaiseUpdated()
    End Sub
    '=======================================================================
    'Fonction : Count()
    'Description : Rend le nombre d'élement dans la collection
    'Détails    :  
    'Retour : une Chaine
    '=======================================================================
    Public Function Count() As Integer
        Return m_col.Count
    End Function

    '=======================================================================
    'Fonction : Item()
    'Description : Rend l'objet indexé
    'Détails    :  
    'Retour : une objet
    '=======================================================================
    Default Public ReadOnly Property Item(ByVal index As Object) As racine
        Get
            Dim i As Integer = 1
            If index.GetType.Name.Equals(i.GetType.Name) Then
                Return m_col.GetByIndex(index)
            Else
                Return m_col.Item(index)
            End If
        End Get
    End Property

    '=======================================================================
    'Fonction : Remove()
    'Description : Supprime l'objet de la collection et désarme l'EventHandler
    'Détails    :  
    'Retour : une Chaine
    '=======================================================================
    Public Function Remove(ByVal index As Integer) As Boolean
        Dim bReturn As Boolean
        Dim obj As racine
        Try
            obj = m_col.GetByIndex(index)
            RemoveHandler obj.Updated, AddressOf MyHandler
            m_col.Remove(m_col.GetKey(index))
            RaiseUpdated()
            bReturn = True
        Catch ex As Exception
            Throw ex
            bReturn = False
        End Try
        Return bReturn
    End Function
    '=======================================================================
    'Fonction : Remove()
    'Description : Supprime l'objet de la collection et désarme l'EventHandler
    'Détails    :  
    'Retour : une Chaine
    '=======================================================================
    Public Function Remove(ByVal index As String) As Boolean
        Dim bReturn As Boolean
        Dim obj As racine
        Try
            obj = m_col.Item(index)
            RemoveHandler obj.Updated, AddressOf MyHandler
            m_col.Remove(index)
            RaiseUpdated()
            bReturn = True
        Catch ex As Exception
            Throw ex
            bReturn = False
        End Try
        Return bReturn
    End Function
    '=======================================================================
    'Fonction : MyHandler()
    'Description : EventHandler qui déclenche l'événement Updated
    'Détails    :  
    'Retour : 
    '=======================================================================
    Private Sub MyHandler()
        RaiseUpdated()
    End Sub

    Public Sub New()
        m_col = New SortedList
    End Sub
    Public Function GetEnumerator() As IEnumerator
        GetEnumerator = m_col.Values.GetEnumerator
    End Function

    Public Overrides Function toString() As String
        Dim objRacine As racine
        Dim str As String
        str = "{"
        For Each objRacine In m_col
            str = str & "," & objRacine.toString()
        Next
        str = str & "}"
        Return str
    End Function 'ToString

    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        Dim objRacine As racine
        Dim bReturn As Boolean
        bReturn = True
        For Each objRacine In m_col
            bReturn = bReturn & objRacine.Equals(obj)
        Next
        Return bReturn
    End Function 'Equals


    Public Function clear() As Boolean
        Dim bReturn As Boolean
        bReturn = True
        While m_col.Count > 0
            Remove(0)
        End While
    End Function 'Clear

    Public Overrides ReadOnly Property shortResume() As String
        Get
            Return "colEventSorted" & Count() & "items"
        End Get
    End Property

    Public Function keyExists(ByVal strkey As String) As Boolean
        Dim obj As Object
        Try
            obj = m_col(strkey)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public ReadOnly Property Values() As Collection
        Get
            Return m_col.Values
        End Get
    End Property
End Class
