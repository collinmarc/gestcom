'===================================================================================================================================
'Projet : Vinicom
'Auteur : Marc Collin 
'Création : 23/06/04
'===================================================================================================================================
' Classe : Reglement
' Description : Reglement de Facture
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
'===================================================================================================================================
Imports System.IO
Public Class Reglement
    Inherits Persist

    Private m_idFact As Integer                 'ID Facture
    Private m_TypeFact As Integer               'Type de Facture
    Private m_Montant As Decimal                'Montant du reglement
    Private m_Date As Date                      'Date du reglement
    Private m_Reference As String               'Reference du reglement
    Private m_Commentaire As String             'Commentaire
    Private m_Etat As EtatReglement             'Etat du Reglement
    Private m_oFacture As Facture               'Facture Associée

#Region "Accesseurs"
    Public Sub New()
        m_typedonnee = vncEnums.vncTypeDonnee.REGLEMENT
        m_idFact = 0
        m_TypeFact = 0
        m_Montant = 0.0
        m_Date = DATE_DEFAUT
        m_Reference = String.Empty
        m_Commentaire = String.Empty
        m_Etat = New EtatReglement_Saisi()
    End Sub 'New
    Public Sub New(ByVal pFact As Facture)
        Debug.Assert(pFact IsNot Nothing, "pFact must be set")
        m_typedonnee = vncEnums.vncTypeDonnee.REGLEMENT
        m_idFact = pFact.id
        m_TypeFact = pFact.typeDonnee
        m_Montant = 0.0
        m_Date = DATE_DEFAUT
        m_Reference = String.Empty
        m_Commentaire = String.Empty
        m_Etat = New EtatReglement_Saisi()
    End Sub 'New

    Friend Function getData(ByVal pRow As System.Data.DataRow) As Boolean

        Dim oRow As dsVinicom.REGLEMENTRow
        Dim bReturn As Boolean

        Try
            oRow = CType(pRow, dsVinicom.REGLEMENTRow)
            setid(oRow.RGL_ID)
            IdFact = oRow.RGL_IDFACT
            If Not oRow.IsRGL_TYPEFACTNull() Then
                TypeFact = oRow.RGL_TYPEFACT
            End If
            If Not oRow.IsRGL_MONTANTNull() Then
                Montant = oRow.RGL_MONTANT
            End If

            If Not oRow.IsRGL_REFNull() Then
                Reference = oRow.RGL_REF
            End If
            If Not oRow.IsRGL_DATENull() Then
                DateReglement = oRow.RGL_DATE.ToShortDateString()
            End If
            If Not oRow.IsRGL_COMMNull() Then
                Commentaire = oRow.RGL_COMM
            End If
            If Not oRow.IsRGL_ETATNull() Then
                Etat = EtatReglement.createEtat(oRow.RGL_ETAT)
            End If
            resetBooleans()
            bReturn = True

        Catch ex As Exception
            setError(System.Environment.StackTrace, ex.Message)
            bReturn = False
        End Try

        Return bReturn
    End Function



    Public Property IdFact() As Integer
        Get
            Return m_idFact
        End Get
        Set(ByVal Value As Integer)
            If Not m_idFact.Equals(Value) Then
                m_idFact = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property TypeFact() As Integer
        Get
            Return m_TypeFact
        End Get
        Set(ByVal Value As Integer)
            If Not m_TypeFact.Equals(Value) Then
                m_TypeFact = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property DateReglement() As DateTime
        Get
            Return m_Date
        End Get
        Set(ByVal Value As DateTime)
            If Not (Value.Equals(m_Date)) Then
                m_Date = Value.ToShortDateString
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property Reference() As String
        Get
            Return m_Reference
        End Get
        Set(ByVal Value As String)
            If Not Value.Equals(m_Reference) Then
                m_Reference = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property Montant() As Decimal
        Get
            Return m_Montant
        End Get
        Set(ByVal Value As Decimal)
            If Value <> m_Montant Then
                m_Montant = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property Commentaire() As String
        Get
            Return m_Commentaire
        End Get
        Set(ByVal Value As String)
            If Not (Value.Equals(m_Commentaire)) Then
                m_Commentaire = Value
                RaiseUpdated()
            End If
        End Set
    End Property

    Public Property Etat() As EtatReglement
        Get
            Return m_Etat
        End Get
        Private Set(ByVal Value As EtatReglement)
            If Not (Value.Equals(m_Etat)) Then
                m_Etat = Value
                RaiseUpdated()
            End If

        End Set
    End Property


#End Region
#Region "Interface Persist"
    '=======================================================================
    'Fonction : DBLoad()
    'Description : Chargement de l'objet
    'Détails    :  
    'Retour : Vrai di l'opération s'est bien déroulée
    '=======================================================================
    Protected Overrides Function DBLoad(Optional ByVal pid As Integer = 0) As Boolean
        Dim bReturn As Boolean
        If pid <> 0 Then
            m_id = pid
        End If

        Debug.Assert(id <> 0, "idCommande <> 0")
        bReturn = loadReglement()
        Return bReturn
    End Function 'DBLoad
    Public Overrides Function save() As Boolean
        Dim bReturn As Boolean
        shared_connect()
        bReturn = MyBase.Save()
        shared_disconnect()
        Return bReturn
    End Function
    '=======================================================================
    'Fonction : delete()
    'Description : Suppression de l'objet dans la base de l'objet
    'Détails    :  
    'Retour : Vrai si l'opération s'est bien déroulée
    '=======================================================================
    Friend Overrides Function delete() As Boolean
        Debug.Assert(id <> 0, "idCommande <> 0")
        Dim bReturn As Boolean
        bReturn = deleteReglement()
        Return bReturn

    End Function ' delete
    '=======================================================================
    'Fonction : checkFordelete
    'description : Controle si l'élément est supprimable
    '                table commandesClients
    '=======================================================================
    Public Overrides Function checkForDelete() As Boolean
        Dim bReturn As Boolean
        Try
            bReturn = True
        Catch ex As Exception
            bReturn = False
        End Try
        Return bReturn
    End Function 'checkForDelete

    '=======================================================================
    'Fonction : insert()
    'Description : insertion de l'objet dans la base de l'objet
    'Détails    :  
    'Retour : Vrai di l'opération s'est bien déroulée
    '=======================================================================
    Friend Overrides Function insert() As Boolean
        Debug.Assert(m_idFact <> 0, "L'IdFact n'est pas Renseigné")
        Debug.Assert(id = 0, "id = 0")

        Dim bReturn As Boolean
        bReturn = insertREGLEMENT()
        Return bReturn
    End Function 'insert
    '=======================================================================
    'Fonction : Update()
    'Description : Mise à jour de l'objet
    'Détails    :  
    'Retour : Vrai di l'opération s'est bien déroulée
    '=======================================================================
    Friend Overrides Function update() As Boolean
        Debug.Assert(m_idFact <> 0, "L'IDFact  n'est pas Renseigné")
        Debug.Assert(id <> 0, "id <> 0")
        Dim bReturn As Boolean
        bReturn = updateReglement()
        Return bReturn

    End Function 'Update

    '=======================================================================
    '                           METHODE DE CLASSE                          |  
    'Fonction : getListe 
    'Description : Liste des reglements de data à date
    'Retour : Rend une collection de reglements
    '=======================================================================
    Public Shared Function getListe(Optional ByVal pdateDeb As Date = DATE_DEFAUT, Optional ByVal pdateFin As Date = DATE_DEFAUT) As Collection
        Dim colReturn As Collection
        shared_connect()
        colReturn = getListeReglement(pdateDeb, pdateFin)
        shared_disconnect()
        Return colReturn
    End Function

    Public Shared Function getListe(ByVal pIdFact As Integer, ByVal pTypeFact As Integer, Optional ByVal pdateDeb As Date = DATE_DEFAUT, Optional ByVal pdateFin As Date = DATE_DEFAUT) As Collection
        Dim colReturn As Collection
        shared_connect()
        colReturn = getListeReglement(pdateDeb, pdateFin, pIdFact, pTypeFact)
        shared_disconnect()
        Return colReturn
    End Function
    Public Shared Function getListe(ByVal pdateDeb As Date, ByVal pdateFin As Date, ByVal pType As vncEnums.vncTypeDonnee, ByVal pEtat As vncEtatReglement) As Collection
        Dim colReturn As Collection
        shared_connect()
        colReturn = getListeReglement(pdateDeb, pdateFin, pType, pEtat)
        shared_disconnect()
        Return colReturn
    End Function


#End Region

    '=======================================================================
    'Fonction : shortResume()
    'Description : Rend l'objet sous forme de String
    'Détails    :  
    'Retour : une Chaine
    '=======================================================================
    Public Overrides ReadOnly Property shortResume() As String
        Get
            Return "RGLMT(" & m_id & "," & m_idFact & "," & m_Montant & ")"
        End Get
    End Property

    '=======================================================================
    'Fonction : ToString()
    'Description : Rend l'objet sous forme de String
    'Détails    :  
    'Retour : une Chaine
    '=======================================================================
    Public Overrides Function toString() As String
        Return "RGLMT(" & m_id & "," & m_idFact & "," & m_Montant & ")"
    End Function

    '=======================================================================
    'Fonction : Equals()
    'Description : Teste l'équivalence entre 2 objets
    'Détails    :  
    'Retour : Vari si l'objet passé en paramétre est equivalent à 
    '=======================================================================
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        Dim bReturn As Boolean
        Dim objMvt As Reglement
        Try

            bReturn = obj.GetType.Name.Equals(Me.GetType().Name)
            objMvt = obj
            bReturn = MyBase.Equals(obj)

            bReturn = bReturn And (m_idFact.Equals(objMvt.IdFact))
            bReturn = bReturn And (m_Date.ToShortDateString.Equals(objMvt.DateReglement.ToShortDateString))
            bReturn = bReturn And (m_TypeFact.Equals(objMvt.TypeFact))
            bReturn = bReturn And (m_Montant.Equals(objMvt.Montant))
            bReturn = bReturn And (m_Reference.Equals(objMvt.Reference))
            bReturn = bReturn And (m_Commentaire.Equals(objMvt.Commentaire))
            bReturn = bReturn And (m_Etat.Equals(objMvt.Etat))

        Catch ex As Exception
            bReturn = False
        End Try

        Return bReturn
    End Function 'Equals

    Public Sub changeEtat(ByVal pEtat As vncActionReglement)
        m_Etat = m_Etat.changeEtat(pEtat)
        bUpdated = True
    End Sub
    Public Function Exporter(ByVal pstFileName As String) As Boolean
        Dim strLine As String
        Dim oTA As vini_DB.dsVinicomTableAdapters.EXPORTPARAMTableAdapter
        Dim oTAConstantes As vini_DB.dsVinicomTableAdapters.CONSTANTESTableAdapter
        Dim nbLignes As Integer
        Dim nLigne As Integer
        Dim nbChamps As Integer
        Dim nChamps As Integer
        Dim oDT As vini_DB.dsVinicom.EXPORTPARAMDataTable = Nothing
        Dim oRow As vini_DB.dsVinicom.EXPORTPARAMRow = Nothing
        Dim oDTConstantes As vini_DB.dsVinicom.CONSTANTESDataTable = Nothing
        Dim strValeur As String
        Dim strCode As String = String.Empty
        Dim bReturn As Boolean
        Try
            strLine = String.Empty
            oTA = New vini_DB.dsVinicomTableAdapters.EXPORTPARAMTableAdapter()
            oTA.Connection = Persist.oleDBConnection
            oTAConstantes = New vini_DB.dsVinicomTableAdapters.CONSTANTESTableAdapter()
            oTAConstantes.Connection = Persist.oleDBConnection
            oDTConstantes = oTAConstantes.GetData()
            'Recupération du code en fonction du type de facture
            Select Case TypeFact
                Case vncEnums.vncTypeDonnee.FACTCOMM
                    strCode = "RGLMT_COM"
                Case vncEnums.vncTypeDonnee.FACTCOL
                    strCode = "RGLMT_COL"
                Case vncEnums.vncTypeDonnee.FACTTRP
                    strCode = "RGLMT_TRP"
            End Select
            'Recupération du nombre de lignes

            nbLignes = oTA.getNbLignes(strCode)
            For nLigne = 1 To nbLignes
                ' Pour chaque ligne
                'Lecture des paramètres
                oDT = oTA.GetDataBy_TYPE_NLIGNE(strCode, nLigne)
                'Recupéation du nombre de champs
                nbChamps = oTA.getNbChamps(strCode, nLigne)
                For nChamps = 1 To nbChamps
                    'Pour chaque Champs
                    oRow = oDT.FindByEXP_TYPEEXP_NLIGNEEXP_NCHAMPS(strCode, nLigne, nChamps)
                    If oRow IsNot Nothing Then
                        If Trim(oRow.EXP_TYPECHAMPS).Equals("C") Then
                            'Exportation d'une Constante
                            strLine = strLine + Left(oRow.EXP_VALEUR + Space(oRow.EXP_LONGUEUR), oRow.EXP_LONGUEUR)
                        End If
                        If Trim(oRow.EXP_TYPECHAMPS).Equals("V") Then
                            'Exportation d'une Valeur
                            strValeur = getAttributeValue(oRow.EXP_VALEUR, oDTConstantes)
                            strLine = strLine + Left(strValeur + Space(oRow.EXP_LONGUEUR), oRow.EXP_LONGUEUR)
                        End If
                    End If
                Next 'champs
                strLine = strLine + vbCrLf

            Next 'Lignes
            File.AppendAllText(pstFileName, strLine)
            bReturn = True
        Catch ex As Exception
            If oRow IsNot Nothing Then
                Throw New Exception(ex.Message + oRow.EXP_VALEUR)
            Else
                Throw ex
            End If
        End Try
        Return bReturn
    End Function

    Public Function getAttributeValue(ByVal pstrAttributeName As String, ByVal pDTConstantes As dsVinicom.CONSTANTESDataTable) As String
        Dim strReturn As String
        strReturn = String.Empty

        Select Case pstrAttributeName
            Case "CODECOMPTA"
                If (m_oFacture Is Nothing) Then
                    Select Case TypeFact
                        Case vncEnums.vncTypeDonnee.FACTCOL
                            m_oFacture = FactColisageJ.createandload(IdFact)
                        Case vncEnums.vncTypeDonnee.FACTCOMM
                            m_oFacture = FactCom.createandload(IdFact)
                        Case vncEnums.vncTypeDonnee.FACTTRP
                            m_oFacture = FactTRP.createandload(IdFact)
                    End Select
                End If
                If m_oFacture IsNot Nothing Then
                    strReturn = m_oFacture.oTiers.CodeCompta
                End If
            Case "DATEREGLEMENT"
                strReturn = Format(DateReglement, "ddMMyy")
            Case "REFERENCE"
                If (m_oFacture Is Nothing) Then
                    Select Case TypeFact
                        Case vncEnums.vncTypeDonnee.FACTCOL
                            m_oFacture = FactColisageJ.createandload(IdFact)
                        Case vncEnums.vncTypeDonnee.FACTCOMM
                            m_oFacture = FactCom.createandload(IdFact)
                        Case vncEnums.vncTypeDonnee.FACTTRP
                            m_oFacture = FactTRP.createandload(IdFact)
                    End Select
                End If

                If m_oFacture IsNot Nothing Then
                    strReturn = getTXTString(m_oFacture.oTiers.rs)
                End If
            Case "TTC"
                strReturn = Montant.ToString("0000000000.00").Replace(".", "")
            Case "COMPTEBANQUE"
                If pDTConstantes IsNot Nothing Then
                    strReturn = pDTConstantes.Rows(0)("CST_SOC_COMPTEBANQUE")
                End If
            Case "COMPTEBANQUE2"
                If pDTConstantes IsNot Nothing Then
                    strReturn = pDTConstantes.Rows(0)("CST_SOC2_COMPTEBANQUE")
                End If
            Case "FOLIO"
                strReturn = Format(DateTime.Now.Day, "000")

        End Select

        Return strReturn
    End Function

    Private Function getTXTString(ByVal pInputString As String) As String
        Dim strReturn As String = String.Empty
        Dim c As Char
        Dim tab As Char() = pInputString.ToCharArray()

        For Each c In tab
            If Char.IsLetterOrDigit(c) Then
                strReturn = strReturn + c
            Else
                strReturn = strReturn + " "
            End If
        Next
        Return strReturn
    End Function








End Class
