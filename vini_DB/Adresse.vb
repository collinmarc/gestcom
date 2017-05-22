'===================================================================================================================================
'Projet : Vinicom
'Auteur : Marc Collin 
'Création : 23/06/04
'===================================================================================================================================
' Classe : Adresse
' Description : Adresse
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
'===================================================================================================================================Public MustInherit Class Persist
Public Class Adresse
    Inherits racine
    Private m_nom As String
    Private m_Rue1 As String
    Private m_Rue2 As String
    Private m_cp As String
    Private m_Ville As String
    Private m_tel As String
    Private m_fax As String
    Private m_port As String
    Private m_email As String
    Private m_Code As String


    Public Property nom() As String
        Get
            Return m_nom
        End Get
        Set(ByVal Value As String)
            If Not m_nom.Equals(Value) Then
                m_nom = Value
                RaiseUpdated()
            End If
        End Set
    End Property

    Public Property rue1() As String
        Get
            Return m_Rue1
        End Get
        Set(ByVal Value As String)
            If Not Value.Equals(m_Rue1) Then
                m_Rue1 = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property rue2() As String
        Get
            Return m_Rue2
        End Get
        Set(ByVal Value As String)
            If Not m_Rue2.Equals(Value) Then
                m_Rue2 = Value
                RaiseUpdated()
            End If
        End Set
    End Property

    Public Property cp() As String
        Get
            Return m_cp
        End Get
        Set(ByVal Value As String)
            If Not Value.Equals(m_cp) Then
                m_cp = Value
                RaiseUpdated()
            End If
        End Set
    End Property

    Public Property ville() As String
        Get
            Return m_Ville
        End Get
        Set(ByVal Value As String)
            If Not Value.Equals(m_Ville) Then
                m_Ville = Value
                RaiseUpdated()
            End If
        End Set
    End Property

    Public Property tel() As String
        Get
            Return m_tel
        End Get
        Set(ByVal Value As String)
            If Not m_tel.Equals(Value) Then
                'Insertion d'un 0 avant le Numéro de Fax si besoin
                If Value.Length() > 0 Then
                    If Left(Value, 1) <> "0" Then
                        Value = "0" + Value
                    End If
                End If
                m_tel = Value
                RaiseUpdated()
            End If
        End Set
    End Property

    Public Property fax() As String
        Get
            Return m_fax
        End Get
        Set(ByVal Value As String)
            If Not Value.Equals(m_fax) Then
                'Insertion d'un 0 avant le Numéro de Fax si besoin
                If Value.Length() > 0 Then
                    If Left(Value, 1) <> "0" Then
                        Value = "0" + Value
                    End If
                End If
                m_fax = Value
                RaiseUpdated()
            End If
        End Set
    End Property

    Public Property port() As String
        Get
            Return m_port
        End Get
        Set(ByVal Value As String)
            If Not Value.Equals(m_port) Then
                'Insertion d'un 0 avant le Numéro de Fax si besoin
                If Value.Length() > 0 Then
                    If Left(Value, 1) <> "0" Then
                        Value = "0" + Value
                    End If
                End If
                m_port = Value
                RaiseUpdated()
            End If
        End Set
    End Property

    Public Property Email() As String
        Get
            Return m_email
        End Get
        Set(ByVal Value As String)
            If Not Value.Equals(m_email) Then
                m_email = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property Code() As String
        Get
            Return m_Code
        End Get
        Set(ByVal Value As String)
            If Not Value.Equals(m_Code) Then
                m_Code = Value
                RaiseUpdated()
            End If
        End Set
    End Property

    Public Overrides Function toString() As String
        Return "ADR =(" & m_Code & "," & m_nom & "," & m_Rue1 & "," & m_Rue2 & "," & m_cp & "," & m_Ville & "," & m_tel & "," & m_fax & "," & m_port & "," & m_email & ")"
    End Function

    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        Dim objAdr As Adresse
        Dim bReturn As Boolean
        Try
            bReturn = True
            objAdr = CType(obj, Adresse)

            bReturn = bReturn And (Code.Equals(objAdr.Code))
            bReturn = bReturn And (nom.Equals(objAdr.nom))
            bReturn = bReturn And (rue1.Equals(objAdr.rue1))
            bReturn = bReturn And (rue2.Equals(objAdr.rue2))
            bReturn = bReturn And (cp.Equals(objAdr.cp))
            bReturn = bReturn And (ville.Equals(objAdr.ville))
            bReturn = bReturn And (tel.Equals(objAdr.tel))
            bReturn = bReturn And (fax.Equals(objAdr.fax))
            bReturn = bReturn And (port.Equals(objAdr.port))
            bReturn = bReturn And (Email.Equals(objAdr.Email))

        Catch ex As Exception
            bReturn = False
        End Try
        Return bReturn
    End Function 'Equals
    Public Sub New()
        m_Code = ""
        m_nom = ""
        m_Rue1 = ""
        m_Rue2 = ""
        m_cp = ""
        m_Ville = ""
        m_tel = ""
        m_fax = ""
        m_port = ""
        m_email = ""
    End Sub

    Public Overrides ReadOnly Property shortResume() As String
        Get
            Return m_nom & " " & m_Ville
        End Get
    End Property
End Class
