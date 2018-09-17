'===================================================================================================================================
'Projet : Vinicom
'Auteur : Marc Collin 
'Création : 23/06/04
'===================================================================================================================================
' Classe : Persist
' Description : Objet Persistant
'===================================================================================================================================
'Membres de Classes
'==================
'Public
'   Shared_Connect(CS)    : Connection à la base de données
'   Shared_DisConnect()    : Déconnection de la base de données
'   Shared_getLastError()    : Rend la dernière Erreur ADO
'   Shared_isConnected()    : Rend Vrai si la connection est OK
'   getlisteModeReglement() : rend la liste des modes de reglements
'Protected
'Private
'Membres d'instances
'==================
'Public
'Protected
'Private
'   bNew            : L'objet est à insérer dans la base
'   bUpdate         : L'objet est à mettre à jour dans la base
'   bDelete         : L'objet est à supprimer de la base
'   bResume         : L'objet est un résumé (id, code, nom)
'===================================================================================================================================
'Historique :
'============
'08/03/05 : ListeMVTSTK : Tri sur Type de mvt DESC
'10/10/06: ListeMVTSTK : Passage an ADO.net OLEDB
'===================================================================================================================================Public MustInherit Class Persist
Option Explicit On 
'Imports ADODB
Imports System.data.OleDb
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Collections.Generic
Public MustInherit Class Persist
    Inherits racine
    Private Shared WithEvents m_dbconn As dbConnection = New dbConnection
    Private Shared m_ConnectionString As String
    Private Shared m_ReportCnxUser As String
    Private Shared m_ReportCnxPassword As String
    'Sauvegarde des paramètres de connection pour permettre le swith
    Private Shared m_colDBConnection As System.Collections.Stack = New Stack
 
    Protected m_bNew As Boolean
    Private m_bUpdated As Boolean
    Protected m_bDeleted As Boolean
    Protected m_id As Integer
    Protected m_typedonnee As vncTypeDonnee
    Protected m_bResume As Boolean 'True = objet résumé (ID, Code, Nom)
    'Chargement d'un objet en mémoire (doit être redéfini)
    Protected MustOverride Function DBLoad(Optional ByVal pid As Integer = 0) As Boolean
    Friend MustOverride Function delete() As Boolean
    Friend MustOverride Function insert() As Boolean
    Friend MustOverride Function update() As Boolean
    Friend Overridable Function fromRS(ByVal pRS As OleDbDataReader) As Boolean
        Return False
    End Function
    Public MustOverride Function checkForDelete() As Boolean
    Public Shared Event Connected()
    Public Shared Event Disconnected()

    Public Shared Property ConnectionString() As String
        Get
            Return m_ConnectionString
        End Get
        Set(ByVal value As String)
            m_ConnectionString = value
        End Set
    End Property
    Public Shared Sub setReportCnx(ByVal strUSer As String, ByVal strPassword As String)
        m_ReportCnxUser = strUSer
        m_ReportCnxPassword = strPassword
    End Sub


        'Nom : Connect - Méthode de Classe 
        'Description : Connection à la base de Donnée
    Public Shared Function shared_connect() As Boolean
        Debug.Assert(Not String.IsNullOrEmpty(m_ConnectionString), "ConnectionString must ne set before")

        Dim bReturn As Boolean
        bReturn = m_dbconn.connect(m_ConnectionString)
        Return bReturn
    End Function ' Shared_Connect
    'Nom disconnect - Méthode de classe
    'Fonction Deconnection de la base de données
    Public Shared Function shared_disconnect() As Boolean
        Return (m_dbconn.disconnect())
    End Function
    Public Shared Function shared_isConnected() As Boolean
        Return (m_dbconn.isConnected)
    End Function
    Public Shared ReadOnly Property nConnection() As Integer
        Get
            Return m_dbconn.nConnection
        End Get
    End Property
    Public Shared ReadOnly Property oleDBConnection() As OleDbConnection
        Get
            Return m_dbconn.Connection
        End Get
    End Property
    Public Shared ReadOnly Property dbConnection() As dbConnection
        Get
            Return m_dbconn
        End Get
    End Property
    Public Shared Function initFax() As Boolean
        'Dim sqlString As String = " SELECT  " & _
        '                            "CST_SOC_NOMSOC, " & _
        '                            "CST_SOC_ADRESSE_RUE1, " & _
        '                            "CST_SOC_ADRESSE_RUE2, " & _
        '                            "CST_SOC_ADRESSE_CP, " & _
        '                            "CST_SOC_ADRESSE_VILLE, " & _
        '                            "CST_SOC_FAX, " & _
        '                            "CST_FAX_ENVOI_PAGE_GARDE, " & _
        '                            "CST_FAX_SERVERNAME," & _
        '                            "CST_FAX_PREFIX," & _
        '                            "CST_FAX_PAGE_GARDE  " & _
        '                        " FROM CONSTANTES"
        'Dim objOLeDBCommand As OleDbCommand
        'Dim objRS As OleDbDataReader = Nothing
        'Dim bReturn As Boolean

        'objOLeDBCommand = New OleDbCommand
        'objOLeDBCommand.Connection = m_dbconn.Connection
        'objOLeDBCommand.CommandText = sqlString
        'Try
        '    objRS = objOLeDBCommand.ExecuteReader()
        '    If (objRS.Read()) Then
        '        clsFax.m_SenderCompany = GetString(objRS, "CST_SOC_NOMSOC")
        '        clsFax.m_SenderAddress = GetString(objRS, "CST_SOC_ADRESSE_RUE1") & vbCrLf & GetString(objRS, "CST_SOC_ADRESSE_RUE2") & vbCrLf & GetString(objRS, "CST_SOC_ADRESSE_CP") & vbCrLf & GetString(objRS, "CST_SOC_ADRESSE_VILLE")
        '        clsFax.m_SenderFax = GetString(objRS, "CST_SOC_FAX")
        '        clsFax.m_strCoverPageName = GetString(objRS, "CST_FAX_PAGE_GARDE")
        '        clsFax.m_ServerName = GetString(objRS, "CST_FAX_SERVERNAME")
        '        clsFax.m_prefix = GetString(objRS, "CST_FAX_PREFIX")
        '    End If
        '    objRS.Close()
        '    objRS = Nothing
        '    bReturn = True
        'Catch ex As Exception
        '    setError("getListeParam", ex.ToString())
        '    bReturn = False
        'End Try

        'Debug.Assert(bReturn, getErreur())
        'Return bReturn
    End Function


    '=======================================================================
    '                           METHODE DE CLASSE                          |  
    'Fonction : getListeParam 
    'Description : Liste des Paramétres d'un type donné
    'Retour : Rend une collection 
    '=======================================================================
    Public Shared Function getListeParam(Optional ByVal strType As String = "") As Collection
        Dim colReturn As New Collection
        Dim objOLeDBCommand As OleDbCommand
        Dim objParam As Param
        Dim objParamRGLMT As ParamModeReglement
        Dim objRS As OleDbDataReader = Nothing
        Dim strWhere As String = String.Empty
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")



        Dim sqlString As String = "SELECT PAR_ID, PAR_CODE, PAR_TYPE, PAR_VALUE, PAR_VALUE2, PAR_DEFAUT, PAR_DDEB_ECHEANCE FROM PARAMETRE"
        objOLeDBCommand = New OleDbCommand
        objOLeDBCommand.Connection = m_dbconn.connection

        If strType <> "" Then
            strWhere = " PAR_TYPE = ?"
        End If

        If strWhere <> "" Then
            sqlString = sqlString & " WHERE " & strWhere
        End If
        objOLeDBCommand.CommandText = sqlString
        objOLeDBCommand.Parameters.AddWithValue("?", Trim(strType.PadRight(2).Substring(0, 2)))

        Try
            objRS = objOLeDBCommand.ExecuteReader()

            While (objRS.Read())
                objParamRGLMT = Nothing
                If strType.Equals(PAR_MODE_RGLMT) Then
                    objParamRGLMT = New ParamModeReglement
                    objParam = objParamRGLMT
                Else
                    objParam = New Param
                End If
                objParam.fromRS(objRS)
                objParam.resetBooleans()
                colReturn.Add(objParam, objParam.code)
            End While
            objRS.Close()
            objRS = Nothing
            Return colReturn
        Catch ex As Exception
            setError("getListeParam", ex.ToString())
            If Not objRS Is Nothing Then
                objRS.Close()
            End If
            Return New Collection
        End Try

    End Function


    '=======================================================================
    '                           METHODE DE CLASSE                          |  
    'Fonction : getListeContenant 
    'Description : Liste des contenant
    'Retour : Rend une collection 
    '=======================================================================
    Public Shared Function getListeContenants() As Collection
        Dim colReturn As New Collection
        Dim objOLeDBCommand As OleDbCommand
        Dim objParam As contenant
        Dim objRS As OleDbDataReader = Nothing
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")



        Dim sqlString As String = "SELECT CONT_ID, CONT_CODE, CONT_LIBELLE, CONT_CENT, CONT_BOUT , CONT_DEFAUT, CONT_POIDS FROM CONTENANT"
        objOLeDBCommand = New OleDbCommand
        objOLeDBCommand.Connection = m_dbconn.connection

        objOLeDBCommand.CommandText = sqlString

        Try
            objRS = objOLeDBCommand.ExecuteReader()
            While (objRS.Read())
                objParam = New contenant
                objParam.fromRS(objRS)
                objParam.resetBooleans()

                colReturn.Add(objParam, objParam.code)

            End While
            objRS.Close()
            objRS = Nothing
            Return colReturn

        Catch ex As Exception

            setError("getListContenants", ex.ToString())
            objRS.Close()
            objRS = Nothing
            Return New Collection
        End Try

    End Function ' getListeContenant

    '=======================================================================
    '                           METHODE DE CLASSE                          |  
    'Fonction : getListeConstantes 
    'Description : Liste des constantes
    'Retour : Rend une collection 
    '=======================================================================
    Public Shared Function getListeConstantes() As Collection
        Dim colReturn As New Collection
        Dim objOLeDBCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        Dim nColMax As Integer
        Dim nCol As Integer
        Dim strValue As String
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")



        Dim sqlString As String = "SELECT * FROM CONSTANTES"
        objOLeDBCommand = New OleDbCommand
        objOLeDBCommand.Connection = m_dbconn.connection
        objOLeDBCommand.CommandText = sqlString



        Try
            objRS = objOLeDBCommand.ExecuteReader()
            If objRS.Read Then
                nColMax = objRS.FieldCount
                For nCol = 0 To nColMax - 1
                    If objRS.IsDBNull(nCol) Then
                        strValue = ""
                    Else
                        strValue = objRS.GetValue(nCol).ToString()
                    End If
                    colReturn.Add(strValue, objRS.GetName(nCol))
                Next
            End If
            objRS.Close()
            objRS = Nothing
            Return colReturn

        Catch ex As Exception
            setError("getListeConstantes", ex.ToString())
            If Not objRS Is Nothing Then
                objRS.Close()
            End If
            objRS = Nothing
            Return New Collection
        End Try

    End Function ' getListeConstantes

    '=======================================================================
    '                           METHODE DE CLASSE                          |  
    'Fonction : updateCST 
    'Description : Mise à jour d'une constante
    'Retour : True si OK
    '=======================================================================
    Public Shared Function updateCST(ByVal strField As String, ByVal strValue As String) As Boolean
        Dim colReturn As New Collection
        Dim objOLeDBCommand As OleDbCommand
        Dim bReturn As Boolean
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")



        Dim sqlString As String = "UPDATE CONSTANTES SET " & strField & " = ?"
        objOLeDBCommand = New OleDbCommand
        objOLeDBCommand.Connection = m_dbconn.connection

        objOLeDBCommand.CommandText = sqlString



        Try
            objOLeDBCommand.Parameters.AddWithValue("?" , strValue)
            objOLeDBCommand.ExecuteNonQuery()
            bReturn = True
        Catch ex As Exception

            setError("updateCST", ex.ToString())
            bReturn = False
            Throw ex
        End Try
        Return bReturn
    End Function ' getListeConstantes
    Friend Shared Function PurgeLgPrecommande(pdtFin As Date) As Boolean
        Dim bReturn As Boolean
        Try
            Dim sqlString As String = "DELETE FROM PRECOMMANDE " & _
                                   "WHERE PCMD_DATEDERNCMD <= ?"
            Dim objOLeDBCommand As OleDbCommand


            objOLeDBCommand = m_dbconn.Connection.CreateCommand()
            objOLeDBCommand.CommandText = sqlString


            objOLeDBCommand.Parameters.AddWithValue("?", pdtFin)
            m_dbconn.BeginTransaction()
            objOLeDBCommand.Transaction = m_dbconn.transaction
            Try

                objOLeDBCommand.ExecuteNonQuery()
                m_dbconn.transaction.Commit()
                bReturn = True

            Catch ex As Exception
                setError("PurgeLgPrecommande", ex.ToString)
                bReturn = False
                m_dbconn.transaction.Rollback()
            End Try

        Catch ex As Exception
            bReturn = False
        End Try
        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function

    ' Nom : bNew
    ' Fonction : Rend vrai si l'objet est à inserer en base 
    Public Property bNew() As Boolean
        Get
            bNew = m_bNew
        End Get
        Set(ByVal Value As Boolean)
            m_bNew = Value
        End Set
    End Property

    ' Fonction : Rend vrai si l'objet est à detruire
    Public Property bDeleted() As Boolean
        Get
            bDeleted = m_bDeleted
        End Get
        Set(ByVal Value As Boolean)
            m_bDeleted = Value
        End Set
    End Property
    'Rend Vrai si l'objet à été modifié
    Public Property bUpdated() As Boolean
        Get
            bUpdated = m_bUpdated
        End Get
        Set(ByVal Value As Boolean)
            m_bUpdated = Value
        End Set
    End Property
    'Rend L'id unique de l'objet
    Public ReadOnly Property id() As Integer
        Get
            id = m_id
        End Get
    End Property
    Protected Sub setid(ByVal Value As Integer)
        m_id = Value
    End Sub
    Public ReadOnly Property typeDonnee() As vncTypeDonnee
        Get
            Return m_typedonnee
        End Get
    End Property

    '=======================================================================
    'Fonction : Load
    'Description : Chargement d'un Element Persistant
    '           Appel DBLoad qui doit être redéfinis dans la classe de base
    'Retour : 
    '=======================================================================
    Public Function load(Optional ByVal pid As Integer = 0) As Boolean
        Dim bReturn As Boolean
        Dim bResumeOld As Boolean
        shared_connect()

        bResumeOld = m_bResume
        m_bResume = False
        bReturn = DBLoad(pid)
        shared_disconnect()
        'Mise à jour des indicateurs
        If bReturn Then
            resetBooleans()
        Else
            m_bResume = bResumeOld
        End If
        'Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function ' load


    'Sauvegarde d'un Objet
    Public Overridable Function Save() As Boolean
        Dim bReturn As Boolean
        bReturn = True
        If m_bDeleted Then
            If Not m_bNew Then 'on ne supprime pas un nouvel enregistrement
                shared_connect()
                bReturn = delete()
                shared_disconnect()
            End If
        Else
            If m_bNew Then
                shared_connect()
                bReturn = insert()
                shared_disconnect()
            Else
                If m_bUpdated Then
                    shared_connect()
                    bReturn = update()
                    shared_disconnect()
                End If
            End If

        End If
        If bReturn Then
            m_bDeleted = False
            m_bUpdated = False
            m_bNew = False
        End If
        Return bReturn
    End Function 'Save

    'Constructeur de la classe
    Public Sub New()
        m_bNew = True
        m_bUpdated = False
        m_bDeleted = False
        m_id = 0
        m_bResume = False
    End Sub

    Public ReadOnly Property bResume() As Boolean
        Get
            Return m_bResume
        End Get
    End Property

    '=======================================================================
    'Fonction : InsertFRN
    'Description : Insertion d'un client
    '           Appel InsertFRNAdresses
    '                 InsertFRNCommentaires   
    'Retour : Rend Vrai si l'INSERT s'est correctement effectué
    '=======================================================================
    Protected Function insertFRN() As Boolean
        Dim objFRN As Fournisseur
        Dim bReturn As Boolean

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id = 0, "ID=0")
        Debug.Assert(Me.GetType.Name.Equals("Fournisseur"), "Object de type Fournisseur requis")

        Dim sqlString As String = "INSERT INTO FOURNISSEUR( " & _
                                    "FRN_CODE, FRN_NOM, FRN_RS, FRN_RIB1, FRN_RIB2, FRN_RIB3, FRN_RIB4,  FRN_BANQUE, FRN_RGN_ID , FRN_SIRET, FRN_TVAINTRACOM , FRN_RGLMT_ID, FRN_ADR_IDENT,FRN_BEXP_INTERNET, FRN_COMPTA, FRN_ID_MRGLMT1,FRN_ID_MRGLMT2,FRN_ID_MRGLMT3,FRN_IDPRESTASHOP, FRN_BINTERMEDIAIRE, FRN_DOSSIER )" & _
                                  " VALUES (" & _
                                    "?, " & _
                                    "?, " & _
                                    "?, " & _
                                    "?," & _
                                    "?," & _
                                    "?," & _
                                    "?," & _
                                    "?," & _
                                    "?," & _
                                    "?," & _
                                    "?," & _
                                    "?," & _
                                    "?, " & _
                                    "?, " & _
                                    "?, " & _
                                    "?, " & _
                                    "?, " & _
                                    "?, " & _
                                    "?, " & _
                                    "?, " & _
                                    "? " & _
                                    " )"
        Dim objOLeDBCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing



        objFRN = CType(Me, Fournisseur)
        objOLeDBCommand = New OleDbCommand
        objOLeDBCommand.Connection = m_dbconn.connection
        objOLeDBCommand.CommandText = sqlString



        CreateParamP_FRN_CODE(objOLeDBCommand)
        CreateParamP_FRN_NOM(objOLeDBCommand)
        CreateParamP_FRN_RS(objOLeDBCommand)
        CreateParamP_FRN_RIB1(objOLeDBCommand)
        CreateParamP_FRN_RIB2(objOLeDBCommand)
        CreateParamP_FRN_RIB3(objOLeDBCommand)
        CreateParamP_FRN_RIB4(objOLeDBCommand)
        CreateParamP_FRN_BANQUE(objOLeDBCommand)
        CreateParamP_FRN_RGN_ID(objOLeDBCommand)
        CreateParamP_FRN_SIRET(objOLeDBCommand)
        CreateParamP_FRN_TVAINTRACOM(objOLeDBCommand)
        CreateParamP_FRN_RGLMT_ID(objOLeDBCommand)
        CreateParamP_FRN_ADR_IDENT(objOLeDBCommand)
        CreateParamP_FRN_BEXP_INTERNET(objOLeDBCommand)
        CreateParamP_FRN_CODECOMPTA(objOLeDBCommand)
        CreateParamP_FRN_ID_MRGLMT1(objOLeDBCommand)
        CreateParamP_FRN_ID_MRGLMT2(objOLeDBCommand)
        CreateParamP_FRN_ID_MRGLMT3(objOLeDBCommand)
        CreateParamP_FRN_IDPRESTASHOP(objOLeDBCommand)
        CreateParamP_FRN_BINTERMEDIAIRE(objOLeDBCommand)
        CreateParamP_FRN_DOSSIER(objOLeDBCommand)
        m_dbconn.BeginTransaction()
        objOLeDBCommand.Transaction = m_dbconn.transaction
        Try

            objOLeDBCommand.ExecuteNonQuery()
            objOLeDBCommand.CommandText = ("SELECT MAX(FRN_ID) FROM FOURNISSEUR")
            objOLeDBCommand.Transaction = m_dbconn.transaction
            objRS = objOLeDBCommand.ExecuteReader()
            objRS.Read()
            m_id = objRS.GetInt32(0)
            cleanErreur()
            objRS.Close()
            bReturn = UpdateFRNAdresses()   'Insertion des Adresses
            If bReturn Then
                bReturn = UpdateFRNCommentaires()     'Insertion des Commentaires
            End If
            m_dbconn.transaction.Commit()
            bReturn = True

        Catch ex As Exception
            setError("InsertFRN", ex.ToString)
            bReturn = False
            If Not objRS Is Nothing Then
                objRS.Close()
            End If
            m_dbconn.transaction.RollBack()
        End Try

        Debug.Assert(bReturn, getErreur())
        Debug.Assert(m_id <> 0, "ID=0")
        Return bReturn
    End Function
    '=======================================================================
    'Fonction : UpdateFRNAdresses
    'Description : Insertion des Adresses Fournisseur
    'Retour : Rend Vrai si l'UPDATE s'est correctement effectué
    '=======================================================================
    Private Function UpdateFRNAdresses() As Boolean

        Dim objOLeDBCommand As OleDbCommand
        Dim objFRN As Fournisseur


        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "ID<>0")
        Debug.Assert(Me.GetType().Name.Equals("Fournisseur"), "Objet de type Fournisseur Requis")
        'Debug.Assert(objFRN.code <> "")

        Dim sqlString As String = "UPDATE FOURNISSEUR  SET " & _
                                        " FRN_LIV_NOM = ?," & _
                                        " FRN_LIV_RUE1 = ?, " & _
                                        " FRN_LIV_RUE2 = ?, " & _
                                        " FRN_LIV_CP = ?, " & _
                                        " FRN_LIV_VILLE = ?, " & _
                                        " FRN_LIV_TEL = ?, " & _
                                        " FRN_LIV_FAX = ?, " & _
                                        " FRN_LIV_PORT = ?, " & _
                                        " FRN_LIV_EMAIL = ?, " & _
                                        " FRN_FACT_NOM = ?," & _
                                        " FRN_FACT_RUE1 = ?, " & _
                                        " FRN_FACT_RUE2 = ? , " & _
                                        " FRN_FACT_CP = ?, " & _
                                        " FRN_FACT_VILLE = ? , " & _
                                        " FRN_FACT_TEL = ?, " & _
                                        " FRN_FACT_FAX = ? , " & _
                                        " FRN_FACT_PORT = ? , " & _
                                        " FRN_FACT_EMAIL = ? " & _
                                  " WHERE FRN_ID = ?"

        objFRN = CType(Me, Fournisseur)

        objOLeDBCommand = New OleDbCommand(sqlString, m_dbconn.Connection)
        objOLeDBCommand.Transaction = m_dbconn.transaction

        CreateParamP_FRN_LIV_NOM(objOLeDBCommand)
        CreateParamP_FRN_LIV_RUE1(objOLeDBCommand)
        CreateParamP_FRN_LIV_RUE2(objOLeDBCommand)
        CreateParamP_FRN_LIV_CP(objOLeDBCommand)
        CreateParamP_FRN_LIV_VILLE(objOLeDBCommand)
        CreateParamP_FRN_LIV_TEL(objOLeDBCommand)
        CreateParamP_FRN_LIV_FAX(objOLeDBCommand)
        CreateParamP_FRN_LIV_PORT(objOLeDBCommand)
        CreateParamP_FRN_LIV_EMAIL(objOLeDBCommand)
        CreateParamP_FRN_FACT_NOM(objOLeDBCommand)
        CreateParamP_FRN_FACT_RUE1(objOLeDBCommand)
        CreateParamP_FRN_FACT_RUE2(objOLeDBCommand)
        CreateParamP_FRN_FACT_CP(objOLeDBCommand)
        CreateParamP_FRN_FACT_VILLE(objOLeDBCommand)
        CreateParamP_FRN_FACT_TEL(objOLeDBCommand)
        CreateParamP_FRN_FACT_FAX(objOLeDBCommand)
        CreateParamP_FRN_FACT_PORT(objOLeDBCommand)
        CreateParamP_FRN_FACT_EMAIL(objOLeDBCommand)
        CreateParameterP_ID(objOLeDBCommand)


        Try
            objOLeDBCommand.ExecuteNonQuery()
            cleanErreur()
            Return True
        Catch ex As Exception
            setError("UpdateFRNAdresses", ex.ToString())
            Return False
        End Try
    End Function 'UpdateFRNAdresses
    '=======================================================================
    'Fonction : UpdateFRNCommentaires
    'Description : Insertion des commentaires Fournisseur
    'Retour : Rend Vrai si l'UPDATE s'est correctement effectué
    '=======================================================================
    Private Function UpdateFRNCommentaires() As Boolean
        Dim objFRN As Fournisseur
        Debug.Assert(Me.GetType().Name.Equals("Fournisseur"), "Objet de Type Fournisseaur Requis")
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "ID<>0")

        Dim sqlString As String = "UPDATE FOURNISSEUR  SET " & _
                                        " FRN_COM_CMD = ?," & _
                                        " FRN_COM_LIV = ?, " & _
                                        " FRN_COM_FACT = ? , " & _
                                        " FRN_COM_LIBRE = ?  " & _
                                  " WHERE FRN_ID = ?"
        Dim objOLeDBCommand As OleDbCommand

        objFRN = CType(Me, Fournisseur)

        objOLeDBCommand = New OleDbCommand
        objOLeDBCommand.Connection = m_dbconn.Connection
        objOLeDBCommand.CommandText = sqlString
        objOLeDBCommand.Transaction = m_dbconn.transaction



        CreateParamP_FRN_COM_CMD(objOLeDBCommand)
        CreateParamP_FRN_COM_LIV(objOLeDBCommand)
        CreateParamP_FRN_COM_FACT(objOLeDBCommand)
        CreateParamP_FRN_COM_LIBRE(objOLeDBCommand)
        CreateParameterP_ID(objOLeDBCommand)


        Try
            objOLeDBCommand.ExecuteNonQuery()
            cleanErreur()
            Return True
        Catch ex As Exception
            setError("UpdateFRNCommentaires", ex.Message)
            Return False
        End Try
    End Function 'UpdateFRNCommentaires
    '=======================================================================
    'Fonction : InsertCLT
    'Description : Insertion d'un client
    '           Appel InsertCLTAdresses
    '                 InsertCLTCommentaires   
    'Retour : Rend Vrai si l'INSERT s'est correctement effectué
    '=======================================================================
    Protected Function insertCLT() As Boolean
        Dim objCLT As Client
        Dim bReturn As Boolean

        Debug.Assert(Me.GetType().Name.Equals("Client"), "Objet de Type Client Requis")
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id = 0, "ID=0")
        objCLT = CType(Me, Client)

        Dim sqlString As String = "INSERT INTO CLIENT( " & _
                                    "CLT_CODE, CLT_NOM, CLT_RS,  CLT_RIB1, CLT_RIB2, CLT_RIB3, CLT_RIB4,  CLT_BANQUE, CLT_TYPE_ID , CLT_SIRET, CLT_TVAINTRACOM , CLT_RGLMT_ID , CLT_ADR_IDENT, CLT_CODETARIF, CLT_COMPTA, CLT_ID_MRGLMT1,CLT_ID_MRGLMT2,CLT_ID_MRGLMT3, CLT_IDPRESTASHOP, CLT_ORIGINE)" & _
                                  " VALUES (?,?,?,?,?,?,?,?,?,?,?,?, ?, ?, ?,?,?,?,?,?" & _
                                    " )"
        Dim objOLeDBCommand As OleDbCommand
        Dim objOLeDBCommandID As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing


        objOLeDBCommand = New OleDbCommand
        objOLeDBCommandID = New OleDbCommand
        objOLeDBCommand.Connection = m_dbconn.Connection
        objOLeDBCommand.CommandText = sqlString



        CreateParamP_CLT_CODE(objOLeDBCommand)
        CreateParamP_CLT_NOM(objOLeDBCommand)
        CreateParamP_CLT_RS(objOLeDBCommand)
        CreateParamP_CLT_RIB1(objOLeDBCommand)
        CreateParamP_CLT_RIB2(objOLeDBCommand)
        CreateParamP_CLT_RIB3(objOLeDBCommand)
        CreateParamP_CLT_RIB4(objOLeDBCommand)
        CreateParamP_CLT_BANQUE(objOLeDBCommand)
        CreateParamP_CLT_TYPE_ID(objOLeDBCommand)
        CreateParamP_CLT_SIRET(objOLeDBCommand)
        CreateParamP_CLT_TVAINTRACOM(objOLeDBCommand)
        CreateParamP_CLT_RGLMT_ID(objOLeDBCommand)
        CreateParamP_CLT_ADR_IDENT(objOLeDBCommand)
        CreateParamP_CLT_CODETARIF(objOLeDBCommand)
        CreateParamP_CLT_CODECOMPTA(objOLeDBCommand)
        CreateParamP_CLT_ID_MRGLMT1(objOLeDBCommand)
        CreateParamP_CLT_ID_MRGLMT2(objOLeDBCommand)
        CreateParamP_CLT_ID_MRGLMT3(objOLeDBCommand)
        CreateParamP_CLT_IDPRESASHOP(objOLeDBCommand)
        CreateParamP_CLT_ORIGINE(objOLeDBCommand)
        m_dbconn.BeginTransaction()
        objOLeDBCommand.Transaction = m_dbconn.transaction
        Try
            '            m_dbconn.transaction.Begin()
            objOLeDBCommand.ExecuteNonQuery()
            objOLeDBCommandID = New OleDbCommand("SELECT MAX(CLT_ID) FROM CLIENT", m_dbconn.Connection)
            objOLeDBCommandID.Transaction = m_dbconn.transaction
            objRS = objOLeDBCommandID.ExecuteReader()
            objRS.Read()
            m_id = objRS.GetInt32(0)
            cleanErreur()
            objRS.Close()

            bReturn = UpdateCLTAdresses()   'Insertion des Adresses
            If bReturn Then
                bReturn = UpdateCLTCommentaires()     'Insertion des Commentaires
            End If
            objRS = Nothing
            m_dbconn.transaction.commit()

        Catch ex As Exception
            setError("InsertCLT", ex.ToString())
            bReturn = False
            m_dbconn.transaction.RollBack()
        End Try

        '    Debug.Assert(m_id <> 0, "ID=0")
        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'InsertCLT

    '=======================================================================
    'Fonction : UpdateCLTAdresses
    'Description : Insertion des Adresses Client
    'Retour : Rend Vrai si l'UPDATE s'est correctement effectué
    '=======================================================================
    Private Function UpdateCLTAdresses() As Boolean
        Dim objCLT As Client
        Debug.Assert(Me.GetType().Name.Equals("Client"), "Objet de Type Client Requis")
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "ID<>0")
        objCLT = CType(Me, Client)

        Dim sqlString As String = "UPDATE CLIENT  SET " & _
                                        " CLT_LIV_NOM = ?," & _
                                        " CLT_LIV_RUE1 = ?, " & _
                                        " CLT_LIV_RUE2 = ?, " & _
                                        " CLT_LIV_CP = ? , " & _
                                        " CLT_LIV_VILLE = ? , " & _
                                        " CLT_LIV_TEL = ? , " & _
                                        " CLT_LIV_FAX = ? , " & _
                                        " CLT_LIV_PORT = ? , " & _
                                        " CLT_LIV_EMAIL = ? , " & _
                                        " CLT_FACT_NOM = ? ," & _
                                        " CLT_FACT_RUE1 = ? , " & _
                                        " CLT_FACT_RUE2 = ? , " & _
                                        " CLT_FACT_CP = ? , " & _
                                        " CLT_FACT_VILLE = ? , " & _
                                        " CLT_FACT_TEL = ? , " & _
                                        " CLT_FACT_FAX = ? , " & _
                                        " CLT_FACT_PORT = ? , " & _
                                        " CLT_FACT_EMAIL = ?  " & _
                                  " WHERE CLT_ID = ?"
        Dim objOLeDBCommand As OleDbCommand

        objCLT = CType(Me, Client)

        objOLeDBCommand = New OleDbCommand
        objOLeDBCommand.Connection = m_dbconn.Connection
        objOLeDBCommand.Transaction = m_dbconn.transaction
        objOLeDBCommand.CommandText = sqlString



        CreateParamP_CLT_LIV_NOM(objOLeDBCommand)
        CreateParamP_CLT_LIV_RUE1(objOLeDBCommand)
        CreateParamP_CLT_LIV_RUE2(objOLeDBCommand)
        CreateParamP_CLT_LIV_CP(objOLeDBCommand)
        CreateParamP_CLT_LIV_VILLE(objOLeDBCommand)
        CreateParamP_CLT_LIV_TEL(objOLeDBCommand)
        CreateParamP_CLT_LIV_FAX(objOLeDBCommand)
        CreateParamP_CLT_LIV_PORT(objOLeDBCommand)
        CreateParamP_CLT_LIV_EMAIL(objOLeDBCommand)
        CreateParamP_CLT_FACT_NOM(objOLeDBCommand)
        CreateParamP_CLT_FACT_RUE1(objOLeDBCommand)
        CreateParamP_CLT_FACT_RUE2(objOLeDBCommand)
        CreateParamP_CLT_FACT_CP(objOLeDBCommand)
        CreateParamP_CLT_FACT_VILLE(objOLeDBCommand)
        CreateParamP_CLT_FACT_TEL(objOLeDBCommand)
        CreateParamP_CLT_FACT_FAX(objOLeDBCommand)
        CreateParamP_CLT_FACT_PORT(objOLeDBCommand)
        CreateParamP_CLT_FACT_EMAIL(objOLeDBCommand)
        CreateParameterP_ID(objOLeDBCommand)


        Try
            objOLeDBCommand.ExecuteNonQuery()
            cleanErreur()
            Return True
        Catch ex As Exception
            setError("UpdateCLTAdresses", ex.ToString())
            Return False
        End Try
    End Function 'UpdateCLTAdresses
    '=======================================================================
    'Fonction : UpdateCLTCommentaires
    'Description : Insertion des commentaires Client
    'Retour : Rend Vrai si l'UPDATE s'est correctement effectué
    '=======================================================================
    Private Function UpdateCLTCommentaires() As Boolean
        Dim objCLT As Client

        Debug.Assert(Me.GetType().Name.Equals("Client"))
        objCLT = CType(Me, Client)
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "ID<>0")

        Dim sqlString As String = "UPDATE CLIENT  SET " & _
                                        " CLT_COM_CMD = ? ," & _
                                        " CLT_COM_LIV = ? , " & _
                                        " CLT_COM_FACT = ? , " & _
                                        " CLT_COM_LIBRE = ?  " & _
                                  " WHERE CLT_ID = ?"
        Dim objOLeDBCommand As OleDbCommand

        objCLT = CType(Me, Client)

        objOLeDBCommand = New OleDbCommand
        objOLeDBCommand.Connection = m_dbconn.Connection
        objOLeDBCommand.Transaction = m_dbconn.transaction
        objOLeDBCommand.CommandText = sqlString



        CreateParamP_CLT_COM_CMD(objOLeDBCommand)
        CreateParamP_CLT_COM_LIV(objOLeDBCommand)
        CreateParamP_CLT_COM_FACT(objOLeDBCommand)
        CreateParamP_CLT_COM_LIBRE(objOLeDBCommand)
        CreateParameterP_ID(objOLeDBCommand)


        Try
            objOLeDBCommand.ExecuteNonQuery()
            cleanErreur()
            Return True
        Catch ex As Exception
            setError("UpdateCLTCommentaires", ex.Message())
            Return False
        End Try
    End Function 'UpdateCLTCommentaires
    '=======================================================================
    'Fonction : UpdateFRN
    'Description : Insertion d'un Fournisseur
    'Retour : Rend Vrai si l'UPDATE s'est correctement effectué
    '=======================================================================
    Protected Function updateFRN() As Boolean
        Dim objFRN As Fournisseur
        Dim bReturn As Boolean


        Debug.Assert(Me.GetType().Name.Equals("Fournisseur"), "Objet de TypeFournisseaue Requis")
        objFRN = CType(Me, Fournisseur)
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "ID<>0")

        Dim sqlString As String = "UPDATE FOURNISSEUR  SET FRN_CODE = ?," & _
                                        " FRN_NOM = ? , " & _
                                        "FRN_RS = ? , " & _
                                        "FRN_RIB1 = ? , " & _
                                        "FRN_RIB2 = ? , " & _
                                        "FRN_RIB3 = ? , " & _
                                        "FRN_RIB4 = ? , " & _
                                        "FRN_BANQUE = ? , " & _
                                        "FRN_RGN_ID = ? , " & _
                                        "FRN_SIRET = ? , " & _
                                        "FRN_TVAINTRACOM = ? , " & _
                                        "FRN_RGLMT_ID = ? ," & _
                                        "FRN_ADR_IDENT = ? , " & _
                                        "FRN_BEXP_INTERNET = ?, " & _
                                        "FRN_COMPTA = ? ," & _
                                        "FRN_ID_MRGLMT1 = ? ," & _
                                        "FRN_ID_MRGLMT2 = ? ," & _
                                        "FRN_ID_MRGLMT3 = ?, " & _
                                        "FRN_IDPRESTASHOP = ?, " & _
                                        "FRN_BINTERMEDIAIRE = ?, " & _
                                        "FRN_DOSSIER = ? " & _
                                  " WHERE FRN_ID = ?"
        Dim objOLeDBCommand As OleDbCommand

        objFRN = CType(Me, Fournisseur)

        objOLeDBCommand = New OleDbCommand
        objOLeDBCommand.Connection = m_dbconn.Connection
        objOLeDBCommand.CommandText = sqlString



        CreateParamP_FRN_CODE(objOLeDBCommand)
        CreateParamP_FRN_NOM(objOLeDBCommand)
        CreateParamP_FRN_RS(objOLeDBCommand)
        CreateParamP_FRN_RIB1(objOLeDBCommand)
        CreateParamP_FRN_RIB2(objOLeDBCommand)
        CreateParamP_FRN_RIB3(objOLeDBCommand)
        CreateParamP_FRN_RIB4(objOLeDBCommand)
        CreateParamP_FRN_BANQUE(objOLeDBCommand)
        CreateParamP_FRN_RGN_ID(objOLeDBCommand)
        CreateParamP_FRN_SIRET(objOLeDBCommand)
        CreateParamP_FRN_TVAINTRACOM(objOLeDBCommand)
        CreateParamP_FRN_RGLMT_ID(objOLeDBCommand)
        CreateParamP_FRN_ADR_IDENT(objOLeDBCommand)
        CreateParamP_FRN_BEXP_INTERNET(objOLeDBCommand)
        CreateParamP_FRN_CODECOMPTA(objOLeDBCommand)
        CreateParamP_FRN_ID_MRGLMT1(objOLeDBCommand)
        CreateParamP_FRN_ID_MRGLMT2(objOLeDBCommand)
        CreateParamP_FRN_ID_MRGLMT3(objOLeDBCommand)
        CreateParamP_FRN_IDPRESTASHOP(objOLeDBCommand)
        CreateParamP_FRN_BINTERMEDIAIRE(objOLeDBCommand)
        CreateParamP_FRN_DOSSIER(objOLeDBCommand)
        CreateParameterP_ID(objOLeDBCommand)


        Try
            objOLeDBCommand.ExecuteNonQuery()
            cleanErreur()
            bReturn = UpdateFRNAdresses()
            If bReturn Then
                bReturn = UpdateFRNCommentaires()     'Insertion des Commentaires
            End If

            bReturn = True
        Catch ex As Exception
            setError("UpdateFRN", ex.ToString())
            Return False
        End Try
        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'UpdateFRN
    '=======================================================================
    'Fonction : UpdateCLT
    'Description : Insertion d'un Client
    'Retour : Rend Vrai si l'UPDATE s'est correctement effectué
    '=======================================================================
    Protected Function updateCLT() As Boolean
        Dim objCLT As Client
        Dim bReturn As Boolean

        Debug.Assert(Me.GetType.Name.Equals("Client"), "Objet de type Client Requis")
        objCLT = CType(Me, Client)
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "ID<>0")

        Dim sqlString As String = "UPDATE CLIENT  SET CLT_CODE = ? ," & _
                                        " CLT_NOM = ? , " & _
                                        "CLT_RS = ? , " & _
                                        "CLT_RIB1 = ? , " & _
                                        "CLT_RIB2 = ? , " & _
                                        "CLT_RIB3 = ? , " & _
                                        "CLT_RIB4 = ? , " & _
                                        "CLT_BANQUE = ? , " & _
                                        "CLT_TYPE_ID = ? , " & _
                                        "CLT_SIRET = ? , " & _
                                        "CLT_TVAINTRACOM = ? , " & _
                                        "CLT_RGLMT_ID = ? , " & _
                                        "CLT_ADR_IDENT = ? , " & _
                                        "CLT_CODETARIF = ? , " & _
                                        "CLT_COMPTA = ? ," & _
                                        "CLT_ID_MRGLMT1 = ?, " & _
                                        "CLT_ID_MRGLMT2 = ?, " & _
                                        "CLT_ID_MRGLMT3 = ?, " & _
                                        "CLT_IDPRESTASHOP = ?, " & _
                                        "CLT_ORIGINE = ? " & _
                                  " WHERE CLT_ID = ?"
        Dim objOLeDBCommand As OleDbCommand

        objCLT = CType(Me, Client)

        objOLeDBCommand = New OleDbCommand
        objOLeDBCommand.Connection = m_dbconn.Connection
        objOLeDBCommand.CommandText = sqlString



        CreateParamP_CLT_CODE(objOLeDBCommand)
        CreateParamP_CLT_NOM(objOLeDBCommand)
        CreateParamP_CLT_RS(objOLeDBCommand)
        CreateParamP_CLT_RIB1(objOLeDBCommand)
        CreateParamP_CLT_RIB2(objOLeDBCommand)
        CreateParamP_CLT_RIB3(objOLeDBCommand)
        CreateParamP_CLT_RIB4(objOLeDBCommand)
        CreateParamP_CLT_BANQUE(objOLeDBCommand)
        CreateParamP_CLT_TYPE_ID(objOLeDBCommand)
        CreateParamP_CLT_SIRET(objOLeDBCommand)
        CreateParamP_CLT_TVAINTRACOM(objOLeDBCommand)
        CreateParamP_CLT_RGLMT_ID(objOLeDBCommand)
        CreateParamP_CLT_ADR_IDENT(objOLeDBCommand)
        CreateParamP_CLT_CODETARIF(objOLeDBCommand)
        CreateParamP_CLT_CODECOMPTA(objOLeDBCommand)
        CreateParamP_CLT_ID_MRGLMT1(objOLeDBCommand)
        CreateParamP_CLT_ID_MRGLMT2(objOLeDBCommand)
        CreateParamP_CLT_ID_MRGLMT3(objOLeDBCommand)
        CreateParamP_CLT_IDPRESASHOP(objOLeDBCommand)
        CreateParamP_CLT_ORIGINE(objOLeDBCommand)
        CreateParameterP_ID(objOLeDBCommand)


        Try
            objOLeDBCommand.ExecuteNonQuery()
            cleanErreur()
            bReturn = UpdateCLTAdresses()
            If bReturn Then
                bReturn = UpdateCLTCommentaires()     'Insertion des Commentaires
            End If
            bReturn = True

        Catch ex As Exception
            setError("UpdateCLT", ex.ToString())
            bReturn = False
        End Try

        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'UpdateCLT
    ''' <summary>
    ''' Méthode de classe : Renvoie l'ID du fournisseur en Fonction de son Code 
    ''' </summary>
    ''' <returns>ID du Fournisseur ou -1</returns>
    ''' <remarks></remarks>
    Protected Shared Function getFRNIDByKey(ByVal strKey As String) As Integer
        Debug.Assert(Not String.IsNullOrEmpty(strKey), "strKey must ne set")
        Dim strResult As String
        Dim nReturn As Integer

        shared_connect()
        Try
            strResult = executeSQLQuery(" SELECT FRN_ID FROM FOURNISSEUR WHERE FRN_CODE = '" & strKey & "'")
            nReturn = CInt(strResult)
        Catch ex As Exception
            nReturn = -1
        End Try
        shared_disconnect()
        Return nReturn
    End Function
    ''' <summary>
    ''' Méthode de classe : Renvoie l'ID du Client en Fonction de son Code 
    ''' </summary>
    ''' <returns>ID du Client ou -1</returns>
    ''' <remarks></remarks>
    Protected Shared Function getCLTIDByKey(ByVal strKey As String) As Integer
        Debug.Assert(Not String.IsNullOrEmpty(strKey), "strKey must ne set")
        Dim strResult As String
        Dim nReturn As Integer

        shared_connect()
        Try
            strResult = executeSQLQuery(" SELECT CLT_ID FROM CLIENT WHERE CLT_CODE = '" & strKey & "'")
            nReturn = CInt(strResult)
        Catch ex As Exception
            nReturn = -1
        End Try
        shared_disconnect()
        Return nReturn
    End Function
    ''' <summary>
    ''' Méthode de classe : Renvoie l'ID du Client en Fonction de son IdPrestashop
    ''' </summary>
    ''' <returns>ID du Client ou -1</returns>
    ''' <remarks></remarks>
    Protected Shared Function getCLTIDByPrestashopId(ByVal pIdPrestashop As String) As Integer
        Debug.Assert(Not String.IsNullOrEmpty(pIdPrestashop), "strKey must ne set")
        Dim strResult As String
        Dim nReturn As Integer

        shared_connect()
        Try
            strResult = executeSQLQuery(" SELECT CLT_ID FROM CLIENT WHERE CLT_IDPRESTASHOP = " & pIdPrestashop & "")
            nReturn = CInt(strResult)
        Catch ex As Exception
            nReturn = -1
        End Try
        shared_disconnect()
        Return nReturn
    End Function

    Protected Function loadFRN() As Boolean

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "L'id doit être renseigné")

        Dim sqlString As String = "SELECT FRN_ID, FRN_CODE, FRN_NOM, FRN_RS,  FRN_RIB1, " & _
                                    "FRN_RIB2, FRN_RIB3, FRN_RIB4, FRN_RGN_ID, FRN_BANQUE, RQ_REGION.PAR_VALUE AS REGION," & _
                                    "FRN_SIRET, FRN_TVAINTRACOM, FRN_RGLMT_ID, RQ_ModeReglement.PAR_VALUE AS REGLEMENT, " & _
                                    " FRN_LIV_NOM, FRN_LIV_RUE1, FRN_LIV_RUE2, FRN_LIV_CP, FRN_LIV_VILLE, " & _
                                    " FRN_LIV_TEL, FRN_LIV_FAX, FRN_LIV_PORT, FRN_LIV_EMAIL, " & _
                                    " FRN_FACT_NOM, FRN_FACT_RUE1, FRN_FACT_RUE2, FRN_FACT_CP, FRN_FACT_VILLE, " & _
                                    " FRN_FACT_TEL, FRN_FACT_FAX, FRN_FACT_PORT, FRN_FACT_EMAIL, FRN_ADR_IDENT, " & _
                                    " FRN_COM_CMD, FRN_COM_LIV, FRN_COM_FACT, FRN_COM_LIBRE,FRN_BEXP_INTERNET, FRN_COMPTA, FRN_ID_MRGLMT1,FRN_ID_MRGLMT2,FRN_ID_MRGLMT3, FRN_IDPRESTASHOP, FRN_BINTERMEDIAIRE, FRN_DOSSIER " & _
                                  " FROM (FOURNISSEUR LEFT OUTER JOIN RQ_Region ON FOURNISSEUR.FRN_RGN_ID = RQ_Region.PAR_ID) LEFT OUTER JOIN RQ_ModeReglement ON FOURNISSEUR.FRN_RGLMT_ID = RQ_ModeReglement.PAR_ID" & _
                                  " WHERE FRN_ID = ?"
        Dim objOLeDBCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        Dim objFRN As Fournisseur
        Dim bReturn As Boolean

        objOLeDBCommand = New OleDbCommand
        objOLeDBCommand.Connection = m_dbconn.Connection
        objOLeDBCommand.CommandText = sqlString



        'Si le paramètre existe on ne met à jour que sa valeur
        'Autrement on le crée
        CreateParameterP_ID(objOLeDBCommand)

        Try
            objRS = objOLeDBCommand.ExecuteReader()
            If objRS.Read Then

                'm_id = CType(GetString(objRS,"FRN_ID"), Integer)
                objFRN = CType(Me, Fournisseur)
                objFRN.code = GetString(objRS, "FRN_CODE")
                objFRN.nom = GetString(objRS, "FRN_NOM")
                objFRN.rs = GetString(objRS, "FRN_RS")
                objFRN.rib1 = GetString(objRS, "FRN_RIB1")
                objFRN.rib2 = GetString(objRS, "FRN_RIB2")
                objFRN.rib3 = GetString(objRS, "FRN_RIB3")
                objFRN.rib4 = GetString(objRS, "FRN_RIB4")
                objFRN.banque = GetString(objRS, "FRN_BANQUE")
                objFRN.idRegion = getInteger(objRS, "FRN_RGN_ID")
                objFRN.libregion = GetString(objRS, "REGION")
                objFRN.siret = GetString(objRS, "FRN_SIRET")
                objFRN.numtvaintracom = GetString(objRS, "FRN_TVAINTRACOM")
                objFRN.idModeReglement = getInteger(objRS, "FRN_RGLMT_ID")
                objFRN.libModeReglement = GetString(objRS, "REGLEMENT")
                objFRN.AdresseLivraison.nom = GetString(objRS, "FRN_LIV_NOM")
                objFRN.AdresseLivraison.rue1 = GetString(objRS, "FRN_LIV_RUE1")
                objFRN.AdresseLivraison.rue2 = GetString(objRS, "FRN_LIV_RUE2")
                objFRN.AdresseLivraison.cp = GetString(objRS, "FRN_LIV_CP")
                objFRN.AdresseLivraison.ville = GetString(objRS, "FRN_LIV_VILLE")
                objFRN.AdresseLivraison.tel = GetString(objRS, "FRN_LIV_TEL")
                objFRN.AdresseLivraison.fax = GetString(objRS, "FRN_LIV_FAX")
                objFRN.AdresseLivraison.port = GetString(objRS, "FRN_LIV_PORT")
                objFRN.AdresseLivraison.Email = GetString(objRS, "FRN_LIV_EMAIL")
                objFRN.AdresseFacturation.nom = GetString(objRS, "FRN_FACT_NOM")
                objFRN.AdresseFacturation.rue1 = GetString(objRS, "FRN_FACT_RUE1")
                objFRN.AdresseFacturation.rue2 = GetString(objRS, "FRN_FACT_RUE2")
                objFRN.AdresseFacturation.cp = GetString(objRS, "FRN_FACT_CP")
                objFRN.AdresseFacturation.ville = GetString(objRS, "FRN_FACT_VILLE")
                objFRN.AdresseFacturation.tel = GetString(objRS, "FRN_FACT_TEL")
                objFRN.AdresseFacturation.fax = GetString(objRS, "FRN_FACT_FAX")
                objFRN.AdresseFacturation.port = GetString(objRS, "FRN_FACT_PORT")
                objFRN.AdresseFacturation.Email = GetString(objRS, "FRN_FACT_EMAIL")
                objFRN.bAdressesIdentiques = GetBoolean(objRS, "FRN_ADR_IDENT")
                objFRN.CommCommande.comment = GetString(objRS, "FRN_COM_CMD")
                objFRN.CommLivraison.comment = GetString(objRS, "FRN_COM_LIV")
                objFRN.CommFacturation.comment = GetString(objRS, "FRN_COM_FACT")
                objFRN.CommLibre.comment = GetString(objRS, "FRN_COM_LIBRE")
                objFRN.bExportInternet = getInteger(objRS, "FRN_BEXP_INTERNET")
                objFRN.CodeCompta = GetString(objRS, "FRN_COMPTA")
                objFRN.idModeReglement1 = getInteger(objRS, "FRN_ID_MRGLMT1")
                objFRN.idModeReglement2 = getInteger(objRS, "FRN_ID_MRGLMT2")
                objFRN.idModeReglement3 = getInteger(objRS, "FRN_ID_MRGLMT3")
                objFRN.idPrestashop = getInteger(objRS, "FRN_IDPRESTASHOP")
                objFRN.bIntermdiaire = getInteger(objRS, "FRN_BINTERMEDIAIRE")
                objFRN.Dossier = GetString(objRS, "FRN_DOSSIER")
                objRS.Close()
                objRS = Nothing
                cleanErreur()
                bReturn = True
            End If

        Catch ex As Exception
            setError("LoadFRN", ex.ToString())
            bReturn = False
        End Try
        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'Load FRN

    '=========================================================================================
    ' Function : loadFRNLIGHT
    ' description : Chargement d'un résumé du fournisseur
    '=========================================================================================
    Protected Function loadFRNLight() As Boolean

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "L'id doit être renseigné")

        Dim sqlString As String = "SELECT FRN_ID, FRN_CODE, FRN_NOM, FRN_RS, FRN_BEXP_INTERNET, FRN_BINTERMEDIAIRE, FRN_DOSSIER " &
                                  " FROM FOURNISSEUR " &
                                  " WHERE FRN_ID = ?"
        Dim objOLeDBCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        Dim objFRN As Fournisseur
        Dim bReturn As Boolean

        objOLeDBCommand = New OleDbCommand
        objOLeDBCommand.Connection = m_dbconn.Connection
        objOLeDBCommand.CommandText = sqlString



        'Si le paramètre existe on ne met à jour que sa valeur
        'Autrement on le crée
        CreateParameterP_ID(objOLeDBCommand)

        Try
            objRS = objOLeDBCommand.ExecuteReader()
            If objRS.Read() Then
                objFRN = CType(Me, Fournisseur)
                objFRN.code = GetString(objRS, "FRN_CODE")
                objFRN.nom = GetString(objRS, "FRN_NOM")
                objFRN.rs = GetString(objRS, "FRN_RS")
                objFRN.bExportInternet = getInteger(objRS, "FRN_BEXP_INTERNET")
                objFRN.bIntermdiaire = GetBoolean(objRS, "FRN_BINTERMEDIAIRE")
                objFRN.Dossier = GetString(objRS, "FRN_DOSSIER")
                m_bResume = True
            End If

            objRS.Close()
            objRS = Nothing
            cleanErreur()
            bReturn = True

        Catch ex As Exception
            setError("LoadFRNLight", ex.ToString())
            bReturn = False
        End Try
        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'LoadFRNLight
    '=======================================================================
    'Fonction : LoadCLT
    'Description : Chargement de l'objet en base
    'Retour : Rend Vrai si le chargement s'est correctement effectué
    '=======================================================================
    Protected Function loadCLT() As Boolean

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "L'id doit être renseigné")

        Dim sqlString As String = "SELECT CLT_ID, CLT_CODE, CLT_NOM, CLT_RS, CLT_RIB1, " & _
                                    "CLT_RIB2, CLT_RIB3, CLT_RIB4, CLT_TYPE_ID, CLT_BANQUE, RQ_TYPECLIENT.PAR_VALUE as TYPECLIENT," & _
                                    "CLT_SIRET, CLT_TVAINTRACOM, CLT_RGLMT_ID, RQ_ModeReglement.PAR_VALUE as MODEREGLEMENT ,CLT_ADR_IDENT," & _
                                    "CLT_LIV_NOM, CLT_LIV_RUE1, CLT_LIV_RUE2, CLT_LIV_CP, CLT_LIV_VILLE, " & _
                                    " CLT_LIV_TEL, CLT_LIV_FAX, CLT_LIV_PORT, CLT_LIV_EMAIL, " & _
                                    " CLT_FACT_NOM, CLT_FACT_RUE1, CLT_FACT_RUE2, CLT_FACT_CP, CLT_FACT_VILLE, " & _
                                    " CLT_FACT_TEL, CLT_FACT_FAX, CLT_FACT_PORT, CLT_FACT_EMAIL, " & _
                                    " CLT_COM_CMD, CLT_COM_LIV, CLT_COM_FACT, CLT_COM_LIBRE, CLT_CODETARIF, CLT_COMPTA, CLT_ID_MRGLMT1,CLT_ID_MRGLMT2,CLT_ID_MRGLMT3, CLT_IDPRESTASHOP,CLT_ORIGINE " & _
                                   " FROM (RQ_TypeClient LEFT OUTER JOIN CLIENT ON RQ_TypeClient.PAR_ID = CLIENT.CLT_TYPE_ID) LEFT OUTER JOIN RQ_ModeReglement ON CLIENT.CLT_RGLMT_ID = RQ_ModeReglement.PAR_ID" & _
                                  " WHERE CLIENT.CLT_ID=? "
        Dim objOLeDBCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        Dim objCLT As Client
        Dim bReturn As Boolean

        objOLeDBCommand = New OleDbCommand
        objOLeDBCommand.Connection = m_dbconn.Connection
        objOLeDBCommand.CommandText = sqlString



        'Si le paramètre existe on ne met à jour que sa valeur
        'Autrement on le crée
        CreateParameterP_ID(objOLeDBCommand)

        Try
            'objRS = objOLeDBCommand.ExecuteNonQuery()
            objRS = objOLeDBCommand.ExecuteReader()
            objRS.Read()

            If Not objRS.HasRows Then
                objRS.Close()
                Return False
            End If
            objCLT = CType(Me, Client)
            objCLT.code = GetString(objRS, "CLT_CODE")
            objCLT.nom = GetString(objRS, "CLT_NOM")
            objCLT.rs = GetString(objRS, "CLT_RS")
            objCLT.rib1 = GetString(objRS, "CLT_RIB1")
            objCLT.rib2 = GetString(objRS, "CLT_RIB2")
            objCLT.rib3 = GetString(objRS, "CLT_RIB3")
            objCLT.rib4 = GetString(objRS, "CLT_RIB4")
            objCLT.banque = GetString(objRS, "CLT_BANQUE")
            objCLT.idTypeClient = getInteger(objRS, "CLT_TYPE_ID")
            objCLT.siret = GetString(objRS, "CLT_SIRET")
            objCLT.numtvaintracom = GetString(objRS, "CLT_TVAINTRACOM")
            objCLT.idModeReglement = getInteger(objRS, "CLT_RGLMT_ID")
            objCLT.libModeReglement = GetString(objRS, "MODEREGLEMENT")
            objCLT.bAdressesIdentiques = GetBoolean(objRS, "CLT_ADR_IDENT")
            objCLT.AdresseLivraison.nom = GetString(objRS, "CLT_LIV_NOM")
            objCLT.AdresseLivraison.rue1 = GetString(objRS, "CLT_LIV_RUE1")
            objCLT.AdresseLivraison.rue2 = GetString(objRS, "CLT_LIV_RUE2")
            objCLT.AdresseLivraison.cp = GetString(objRS, "CLT_LIV_CP")
            objCLT.AdresseLivraison.ville = GetString(objRS, "CLT_LIV_VILLE")
            objCLT.AdresseLivraison.tel = GetString(objRS, "CLT_LIV_TEL")
            objCLT.AdresseLivraison.fax = GetString(objRS, "CLT_LIV_FAX")
            objCLT.AdresseLivraison.port = GetString(objRS, "CLT_LIV_PORT")
            objCLT.AdresseLivraison.Email = GetString(objRS, "CLT_LIV_EMAIL")
            objCLT.AdresseFacturation.nom = GetString(objRS, "CLT_FACT_NOM")
            objCLT.AdresseFacturation.rue1 = GetString(objRS, "CLT_FACT_RUE1")
            objCLT.AdresseFacturation.rue2 = GetString(objRS, "CLT_FACT_RUE2")
            objCLT.AdresseFacturation.cp = GetString(objRS, "CLT_FACT_CP")
            objCLT.AdresseFacturation.ville = GetString(objRS, "CLT_FACT_VILLE")
            objCLT.AdresseFacturation.tel = GetString(objRS, "CLT_FACT_TEL")
            objCLT.AdresseFacturation.fax = GetString(objRS, "CLT_FACT_FAX")
            objCLT.AdresseFacturation.port = GetString(objRS, "CLT_FACT_PORT")
            objCLT.AdresseFacturation.Email = GetString(objRS, "CLT_FACT_EMAIL")
            objCLT.CommCommande.comment = GetString(objRS, "CLT_COM_CMD")
            objCLT.CommLivraison.comment = GetString(objRS, "CLT_COM_LIV")
            objCLT.CommFacturation.comment = GetString(objRS, "CLT_COM_FACT")
            objCLT.CommLibre.comment = GetString(objRS, "CLT_COM_LIBRE")
            objCLT.CodeTarif = GetString(objRS, "CLT_CODETARIF")
            objCLT.CodeCompta = GetString(objRS, "CLT_COMPTA")
            objCLT.idModeReglement1 = getInteger(objRS, "CLT_ID_MRGLMT1")
            objCLT.idModeReglement2 = getInteger(objRS, "CLT_ID_MRGLMT2")
            objCLT.idModeReglement3 = getInteger(objRS, "CLT_ID_MRGLMT3")
            objCLT.idPrestashop = getInteger(objRS, "CLT_IDPRESTASHOP")
            objCLT.Origine = GetString(objRS, "CLT_ORIGINE")

            objRS.Close()
            objRS = Nothing
            bReturn = True
            cleanErreur()

        Catch ex As Exception
            setError("LoadCLT", ex.ToString())
            bReturn = False
        End Try
        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'LoadCLT

    '=======================================================================
    'Fonction : LoadCLTLight
    'Description : Chargement du Résumé de l'objet
    'Retour : Rend Vrai si le chargement s'est correctement effectué
    '=======================================================================
    Protected Function loadCLTLight() As Boolean

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "L'id doit être renseigné")

        Dim sqlString As String = "SELECT CLT_ID, CLT_CODE, CLT_NOM, CLT_RS,CLT_CODETARIF " & _
                                  " FROM CLIENT " & _
                                  " WHERE CLIENT.CLT_ID = ? "
        Dim objOLeDBCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        Dim objCLT As Client
        Dim bReturn As Boolean

        objOLeDBCommand = New OleDbCommand
        objOLeDBCommand.Connection = m_dbconn.Connection
        objOLeDBCommand.CommandText = sqlString



        'Si le paramètre existe on ne met à jour que sa valeur
        'Autrement on le crée
        CreateParameterP_ID(objOLeDBCommand)

        Try
            'objRS = objOLeDBCommand.ExecuteNonQuery()
            objRS = objOLeDBCommand.ExecuteReader()


            If Not objRS.HasRows Then
                objRS.Close()
                Return False
            End If
            objRS.Read()
            objCLT = CType(Me, Client)
            objCLT.code = GetString(objRS, "CLT_CODE")
            objCLT.nom = GetString(objRS, "CLT_NOM")
            objCLT.rs = GetString(objRS, "CLT_RS")
            objCLT.CodeTarif = GetString(objRS, "CLT_CODETARIF")
            m_bResume = True
            bReturn = True
            objRS.Close()
            objRS = Nothing


            cleanErreur()

        Catch ex As Exception
            setError("LoadCLTLight", ex.ToString())
            bReturn = False
        End Try
        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'LoadCLTLight

    '=======================================================================
    'Fonction : DeleteFRN
    'Description : Suppression d'un client
    'Retour : Rend Vrai si le DELETE s'est correctement effectué
    '=======================================================================
    Protected Function deleteFRN() As Boolean
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "Id <> 0")
        Debug.Assert(Me.GetType().Name.Equals("Fournisseur"))


        Dim sqlString As String = "DELETE FROM Fournisseur WHERE FRN_ID=? "
        Dim objOLeDBCommand As OleDbCommand
        Dim objFRN As Fournisseur

        objFRN = CType(Me, Fournisseur)
        objOLeDBCommand = New OleDbCommand
        objOLeDBCommand.Connection = m_dbconn.Connection
        objOLeDBCommand.CommandText = sqlString



        CreateParameterP_ID(objOLeDBCommand)
        Try
            objOLeDBCommand.ExecuteNonQuery()
            m_id = 0
            objFRN.code = ""
            objFRN.nom = ""
            cleanErreur()
            Return True

        Catch ex As Exception
            setError("deleteFRN", ex.ToString())
            Return False
        End Try
    End Function 'DeleteFRN
    '=======================================================================
    'Fonction : DeleteCLT
    'Description : Suppression d'un client
    'Retour : Rend Vrai si le DELETE s'est correctement effectué
    '=======================================================================
    Protected Function deleteCLT() As Boolean
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "Id <> 0")
        Debug.Assert(Me.GetType().Name.Equals("Client"))


        Dim sqlString As String = "DELETE FROM CLIENT WHERE CLT_ID=? "
        Dim objOLeDBCommand As OleDbCommand
        Dim objCLT As Client

        objCLT = CType(Me, Client)
        objOLeDBCommand = New OleDbCommand
        objOLeDBCommand.Connection = m_dbconn.Connection
        objOLeDBCommand.CommandText = sqlString



        CreateParameterP_ID(objOLeDBCommand)
        Try
            objOLeDBCommand.ExecuteNonQuery()
            m_id = 0
            objCLT.code = ""
            objCLT.nom = ""
            Return True

        Catch ex As Exception
            setError("DeleteCLT", ex.ToString())
            Return False
        End Try
    End Function 'DeleteCLT

    'Renvoie une collection d'objet résumé de Fournisseur
    Protected Shared Function ListeFRN(Optional ByVal strCode As String = "", Optional ByVal strNom As String = "", Optional ByVal strRS As String = "") As Collection
        Dim colReturn As New Collection
        Dim strWhere As String = ""

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")



        Dim sqlString As String = "SELECT FRN_ID, FRN_CODE, FRN_NOM, FRN_RS,FRN_LIV_VILLE,FRN_COMPTA FROM Fournisseur "
        Dim objOLeDBCommand As OleDbCommand
        Dim objFRN As Fournisseur
        Dim objRS As OleDbDataReader = Nothing
        objOLeDBCommand = New OleDbCommand
        objOLeDBCommand.Connection = m_dbconn.Connection

        If strCode <> "" Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & " FRN_CODE LIKE ?"
            objOLeDBCommand.Parameters.AddWithValue("?", strCode)
        End If
        If strNom <> "" Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "FRN_NOM LIKE ?"
            objOLeDBCommand.Parameters.AddWithValue("?", strNom)
        End If

        If strRS <> "" Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "FRN_RS LIKE ?"
            objOLeDBCommand.Parameters.AddWithValue("?", strRS)
        End If


        If strWhere <> "" Then
            sqlString = sqlString & "WHERE " & strWhere
        End If
        sqlString = sqlString & " ORDER BY FRN_CODE"
        objOLeDBCommand.CommandText = sqlString



        Try
            'objRS = objOLeDBCommand.ExecuteNonQuery()
            objRS = objOLeDBCommand.ExecuteReader()

            While (objRS.Read())
                objFRN = New Fournisseur
                objFRN.m_bResume = True ' C'est un objet Résumé
                objFRN.m_id = getInteger(objRS, "FRN_ID")
                objFRN.code = GetString(objRS, "FRN_CODE")
                objFRN.nom = GetString(objRS, "FRN_NOM")
                objFRN.rs = GetString(objRS, "FRN_RS")
                objFRN.AdresseLivraison.ville = GetString(objRS, "FRN_LIV_VILLE")
                objFRN.CodeCompta = GetString(objRS, "FRN_COMPTA")
                objFRN.resetBooleans()
                colReturn.Add(objFRN, objFRN.code)

            End While
            objRS.Close()
            objRS = Nothing
            Return colReturn
        Catch ex As Exception
            setError("ListFRN", ex.ToString())
            Return New Collection
        End Try
    End Function ' LoadFRN

    'Renvoie une collection d'objet résumé de Client
    Protected Shared Function ListeCLT(Optional ByVal strCode As String = "", Optional ByVal strNom As String = "", Optional ByVal strRS As String = "") As Collection
        Dim colReturn As New Collection
        Dim strWhere As String = ""

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")



        Dim sqlString As String = "SELECT CLT_ID, CLT_CODE, CLT_NOM, CLT_RS, CLT_LIV_VILLE, CLT_COMPTA , CLT_IDPRESTASHOP FROM Client "
        Dim objOLeDBCommand As OleDbCommand
        Dim objCLT As Client
        Dim objRS As OleDbDataReader = Nothing
        objOLeDBCommand = New OleDbCommand
        objOLeDBCommand.Connection = m_dbconn.Connection

        If strCode <> "" Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & " CLT_CODE LIKE ?"
            objOLeDBCommand.Parameters.AddWithValue("?", strCode)
        End If
        If strNom <> "" Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "CLT_NOM LIKE ?"
            objOLeDBCommand.Parameters.AddWithValue("?" , strNom)

        End If

        If strRS <> "" Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "CLT_RS LIKE ?"
            objOLeDBCommand.Parameters.AddWithValue("?" , strRS)

        End If


        If strWhere <> "" Then
            sqlString = sqlString & "WHERE " & strWhere
        End If
        sqlString = sqlString & " ORDER BY CLT_CODE"
        objOLeDBCommand.CommandText = sqlString



        Try
            objRS = objOLeDBCommand.ExecuteReader()

            While (objRS.Read())
                objCLT = New Client("", "")
                objCLT.m_bResume = True ' C'est un objet Résumé
                objCLT.m_id = getInteger(objRS, "CLT_ID")
                objCLT.code = GetString(objRS, "CLT_CODE")
                objCLT.nom = GetString(objRS, "CLT_NOM")
                objCLT.rs = GetString(objRS, "CLT_RS")
                objCLT.AdresseLivraison.ville = GetString(objRS, "CLT_LIV_VILLE")
                objCLT.CodeCompta = GetString(objRS, "CLT_COMPTA")
                objCLT.idPrestashop = getInteger(objRS, "CLT_IDPRESTASHOP")
                objCLT.resetBooleans()
                colReturn.Add(objCLT, objCLT.code)

            End While
            objRS.Close()
            objRS = Nothing
            Return colReturn
        Catch ex As Exception
            setError("ListCLT", ex.ToString())
            Return Nothing
        End Try
    End Function 'ListeCLT
    ''' <summary>
    ''' Renvoie une liste des client de type intermédiaire pour une origine
    ''' </summary>
    ''' <param name="strCode"></param>
    ''' <param name="strNom"></param>
    ''' <param name="strRS"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Shared Function ListeCLTIntegmédiairePOurOrigine(pIdTypeClientInterlediaire As Integer, ByVal pOrigine As String) As List(Of Client)
        Dim colReturn As New List(Of Client)
        Dim strWhere As String = ""
        Dim nId As Integer

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")



        Dim strSQL As String
        strSQL = "SELECT CLT_ID FROM client where clt_Origine = '" & pOrigine & "' and CLT_TYPE_id = " & pIdTypeClientInterlediaire
        Dim objOLeDBCommand As OleDbCommand
        Dim objCLT As Client
        Dim objRS As OleDbDataReader = Nothing
        objOLeDBCommand = New OleDbCommand
        objOLeDBCommand.Connection = m_dbconn.Connection

        strSQL = strSQL & " ORDER BY CLT_CODE"
        objOLeDBCommand.CommandText = strSQL



        Try
            objRS = objOLeDBCommand.ExecuteReader()

            While (objRS.Read())
                nId = getInteger(objRS, "CLT_ID")
                objCLT = Client.createandload(nId)
                colReturn.Add(objCLT)

            End While
            objRS.Close()
            objRS = Nothing
        Catch ex As Exception
            setError("ListCLT", ex.ToString())
        End Try
        Return colReturn
    End Function 'ListeCLTOrigine
    '================================================================================================
    '====================== PRODUIT ================================================================
    '=======================================================================
    'Fonction : InsertPRD
    'Description : Insertion d'un PRODUIT
    'Retour : Rend Vrai si l'INSERT s'est correctement effectué
    '=======================================================================
    Protected Function insertPRD() As Boolean
        Dim objPRD As Produit
        Dim bReturn As Boolean


        Debug.Assert(Me.GetType.Name.Equals("Produit"), "Objet de Type Produit Requis")
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id = 0, "ID=0")
        objPRD = CType(Me, Produit)

        Dim sqlString As String = "INSERT INTO PRODUIT( " &
                                    "PRD_CODE," &
                                    "PRD_LIBELLE," &
                                    "PRD_MOT_CLE, " &
                                    "PRD_FRN_ID, " &
                                    "PRD_CONT_ID, " &
                                    "PRD_COND_ID, " &
                                    "PRD_COUL_ID, " &
                                    "PRD_MIL, " &
                                    "PRD_RGN_ID, " &
                                    "PRD_TVA_ID, " &
                                    "PRD_DATE_DERN_INVENT , " &
                                    "PRD_QTE_STK, " &
                                    "PRD_QTE_STOCK_DERN_INVENT, " &
                                    "PRD_DISPO , " &
                                    "PRD_STOCK, " &
                                    "PRD_CODE_STAT," &
                                    "PRD_TARIFA," &
                                    "PRD_TARIFB," &
                                    "PRD_TARIFC, " &
                                    "PRD_DOSSIER " &
                                    " ) VALUES (" &
                                    "? , ? ,?,? , ? ,?,?,?,?,?,?,?,?,?,?,?,?,?,?,? " &
                                    " )"
        Dim objOLeDBCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing



        objOLeDBCommand = New OleDbCommand
        objOLeDBCommand.Connection = m_dbconn.Connection
        objOLeDBCommand.CommandText = sqlString
        m_dbconn.BeginTransaction()
        objOLeDBCommand.Transaction = m_dbconn.transaction


        CreateParamP_PRD_CODE(objOLeDBCommand)
        CreateParamP_PRD_LIBELLE(objOLeDBCommand)
        CreateParamP_PRD_MOT_CLE(objOLeDBCommand)
        CreateParamP_PRD_FRN_ID(objOLeDBCommand)
        CreateParamP_PRD_CONT_ID(objOLeDBCommand)
        CreateParamP_PRD_COND_ID(objOLeDBCommand)
        CreateParamP_PRD_COUL_ID(objOLeDBCommand)
        CreateParamP_PRD_MIL(objOLeDBCommand)
        CreateParamP_PRD_RGN_ID(objOLeDBCommand)
        CreateParamP_PRD_TVA_ID(objOLeDBCommand)
        CreateParamP_PRD_DATE_DERN_INVENT(objOLeDBCommand)
        CreateParamP_PRD_QTE_STK(objOLeDBCommand)
        CreateParamP_PRD_QTE_STOCK_DERN_INVENT(objOLeDBCommand)
        CreateParamP_PRD_DISPO(objOLeDBCommand)
        CreateParamP_PRD_STOCK(objOLeDBCommand)
        CreateParamP_PRD_CODE_STAT(objOLeDBCommand)
        CreateParamP_PRD_TARIFA(objOLeDBCommand)
        CreateParamP_PRD_TARIFB(objOLeDBCommand)
        CreateParamP_PRD_TARIFC(objOLeDBCommand)
        CreateParamP_PRD_DOSSIER(objOLeDBCommand)
        Try
            objOLeDBCommand.ExecuteNonQuery()
            objOLeDBCommand = New OleDbCommand("SELECT MAX(PRD_ID) FROM PRODUIT", m_dbconn.Connection)
            objOLeDBCommand.Transaction = m_dbconn.transaction
            objRS = objOLeDBCommand.ExecuteReader()
            If objRS.Read() Then
                m_id = objRS.GetInt32(0)
                cleanErreur()
                bReturn = True
            Else
                bReturn = False
            End If

        Catch ex As Exception
            setError("InsertPRD", ex.ToString())
            bReturn = False
        End Try
        If Not objRS Is Nothing Then
            objRS.Close()
            objRS = Nothing
        End If
        If bReturn Then
            m_dbconn.transaction.commit()
        Else
            m_dbconn.transaction.RollBack()
        End If

        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'InsertPRD


    '=======================================================================
    'Fonction : UpdatePRD
    'Description : Insertion d'un PRODUIT
    'Retour : Rend Vrai si l'UPDATE s'est correctement effectué
    '=======================================================================
    Protected Function updatePRD() As Boolean
        Dim objPRD As Produit
        Dim bReturn As Boolean

        Debug.Assert(Me.GetType().Name.Equals("Produit"), "Objet de Type Produit Requis")
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "ID<>0")
        objPRD = CType(Me, Produit)

        Dim sqlString As String = "UPDATE PRODUIT  SET " &
                                    "PRD_CODE= ? , " &
                                    "PRD_LIBELLE = ? , " &
                                    "PRD_MOT_CLE = ? , " &
                                    "PRD_FRN_ID = ? , " &
                                    "PRD_CONT_ID = ? , " &
                                    "PRD_COND_ID = ? , " &
                                    "PRD_COUL_ID = ? , " &
                                    "PRD_MIL = ? , " &
                                    "PRD_RGN_ID = ? , " &
                                    "PRD_TVA_ID = ? , " &
                                    "PRD_DATE_DERN_INVENT = ? , " &
                                    "PRD_QTE_STOCK_DERN_INVENT = ? , " &
                                    "PRD_DISPO = ? , " &
                                    "PRD_STOCK = ? , " &
                                    "PRD_QTE_STK = ? ," &
                                    "PRD_CODE_STAT = ? ,  " &
                                    "PRD_TARIFA = ? ,  " &
                                    "PRD_TARIFB = ? ,  " &
                                    "PRD_TARIFC = ? , " &
                                    "PRD_DOSSIER = ?  " &
                                  " WHERE PRD_ID = ?"
        Dim objOLeDBCommand As OleDbCommand

        objPRD = CType(Me, Produit)

        objOLeDBCommand = New OleDbCommand
        objOLeDBCommand.Connection = m_dbconn.Connection
        objOLeDBCommand.CommandText = sqlString



        CreateParamP_PRD_CODE(objOLeDBCommand)
        CreateParamP_PRD_LIBELLE(objOLeDBCommand)
        CreateParamP_PRD_MOT_CLE(objOLeDBCommand)
        CreateParamP_PRD_FRN_ID(objOLeDBCommand)
        CreateParamP_PRD_CONT_ID(objOLeDBCommand)
        CreateParamP_PRD_COND_ID(objOLeDBCommand)
        CreateParamP_PRD_COUL_ID(objOLeDBCommand)
        CreateParamP_PRD_MIL(objOLeDBCommand)
        CreateParamP_PRD_RGN_ID(objOLeDBCommand)
        CreateParamP_PRD_TVA_ID(objOLeDBCommand)
        CreateParamP_PRD_DATE_DERN_INVENT(objOLeDBCommand)
        CreateParamP_PRD_QTE_STOCK_DERN_INVENT(objOLeDBCommand)
        CreateParamP_PRD_DISPO(objOLeDBCommand)
        CreateParamP_PRD_STOCK(objOLeDBCommand)
        CreateParamP_PRD_QTE_STK(objOLeDBCommand)
        CreateParamP_PRD_CODE_STAT(objOLeDBCommand)
        CreateParamP_PRD_TARIFA(objOLeDBCommand)
        CreateParamP_PRD_TARIFB(objOLeDBCommand)
        CreateParamP_PRD_TARIFC(objOLeDBCommand)
        CreateParamP_PRD_DOSSIER(objOLeDBCommand)
        CreateParameterP_ID(objOLeDBCommand)


        Try
            objOLeDBCommand.ExecuteNonQuery()
            cleanErreur()
            bReturn = True

        Catch ex As Exception

            setError("UpdatePRD", ex.ToString())
            bReturn = False
        End Try
        Debug.Assert(bReturn, getErreur())
        Return bReturn

    End Function 'UpdatePRD
    '=======================================================================
    'Fonction : LoadPRD
    'Description : Chargement de l'objet en base
    'Retour : Rend Vrai si le chargement s'est correctement effectué
    '=======================================================================
    Protected Function loadPRD() As Boolean

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "L'id doit être renseigné")

        Dim sqlString As String = " SELECT " &
                                                  " PRODUIT.PRD_ID," _
                                                  & " PRODUIT.PRD_CODE," _
                                                  & " PRODUIT.PRD_LIBELLE, " _
                                                  & "PRODUIT.PRD_MOT_CLE, " _
                                                  & "PRODUIT.PRD_FRN_ID, " _
                                                  & "PRODUIT.PRD_CONT_ID, " _
                                                  & "PRODUIT.PRD_COND_ID, " _
                                                  & "PRODUIT.PRD_COUL_ID, " _
                                                  & "PRODUIT.PRD_MIL, " _
                                                  & "PRODUIT.PRD_RGN_ID, " _
                                                  & "PRODUIT.PRD_TVA_ID, " _
                                                  & "PRODUIT.PRD_DATE_DERN_INVENT, " _
                                                  & "PRODUIT.PRD_QTE_STK, " _
                                                  & "PRODUIT.PRD_QTE_STOCK_DERN_INVENT, " _
                                                  & "PRODUIT.PRD_DISPO, " _
                                                  & "PRODUIT.PRD_CODE_STAT, " _
                                                  & "RQ_Couleur.PAR_VALUE AS COULEUR, " _
                                                  & "RQ_CONDITIONNEMENT.PAR_CODE AS CONDITIONNEMENT, " _
                                                  & "CONTENANT.CONT_LIBELLE AS CONTENANT, " _
                                                  & "RQ_Region.PAR_VALUE AS REGION, " _
                                                  & "FOURNISSEUR.FRN_NOM , " _
                                                  & "RQ_Tva.PAR_VALUE AS TVA, " _
                                                  & "PRODUIT.PRD_STOCK, " _
                                                  & "RQ_QTECMD_PRD.PRD_QTE_COMMANDE, " _
                                                  & "PRD_TARIFA, " _
                                                  & "PRD_TARIFB, " _
                                                  & "PRD_TARIFC, " _
                                                  & "PRD_DOSSIER" &
                                                " FROM ((FOURNISSEUR INNER JOIN (((CONTENANT INNER JOIN (PRODUIT INNER JOIN RQ_Couleur ON PRODUIT.PRD_COUL_ID = RQ_Couleur.PAR_ID) ON CONTENANT.CONT_ID = PRODUIT.PRD_CONT_ID) INNER JOIN RQ_Region ON PRODUIT.PRD_RGN_ID = RQ_Region.PAR_ID) INNER JOIN RQ_Tva ON PRODUIT.PRD_TVA_ID = RQ_Tva.PAR_ID) ON FOURNISSEUR.FRN_ID = PRODUIT.PRD_FRN_ID) INNER JOIN RQ_CONDITIONNEMENT ON PRODUIT.PRD_COND_ID = RQ_CONDITIONNEMENT.PAR_ID) LEFT JOIN RQ_QTECMD_PRD ON PRODUIT.PRD_ID = RQ_QTECMD_PRD.LGCM_PRD_ID " &
                                " WHERE " &
                                " PRD_ID = ?"

        Dim objOLeDBCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        Dim objPRD As Produit
        Dim bReturn As Boolean

        objOLeDBCommand = New OleDbCommand
        objOLeDBCommand.Connection = m_dbconn.Connection
        objOLeDBCommand.CommandText = sqlString



        'Si le paramètre existe on ne met à jour que sa valeur
        'Autrement on le crée
        CreateParameterP_ID(objOLeDBCommand)

        Try
            'objRS = objOLeDBCommand.ExecuteNonQuery()
            objRS = objOLeDBCommand.ExecuteReader()

            If Not objRS.HasRows Then
                objRS.Close()
                setError("LoadPRD", "Produit Id" & m_id & "non trouvé")
                Return False
            End If
            objRS.Read()
            objPRD = CType(Me, Produit)
            objPRD.Fill(objRS)
            objPRD.libConditionnement = GetString(objRS, "CONDITIONNEMENT")
            objPRD.libContenant = GetString(objRS, "CONTENANT")
            objPRD.libCouleur = GetString(objRS, "COULEUR")
            objPRD.libRegion = GetString(objRS, "REGION")
            objPRD.libTVA = CStr(GetString(objRS, "TVA"))
            objPRD.nomFournisseur = GetString(objRS, "FRN_NOM")

            cleanErreur()
            bReturn = True
            objRS.Close()
            objRS = Nothing

        Catch ex As Exception
            setError("LoadPRD", ex.ToString())

            bReturn = False
        End Try
        Debug.Assert(bReturn, getErreur())

        Return bReturn
    End Function 'LoadPRD
    ''' <summary>
    ''' Renvoie l'ID en fonction de la clé 
    ''' </summary>
    ''' <returns>ID  ou -1</returns>
    ''' <remarks></remarks>
    Protected Shared Function getIDPRDByKey(ByVal strKey As String) As Integer
        Debug.Assert(Not String.IsNullOrEmpty(strKey), "strKey must ne set")
        Dim strResult As String
        Dim nReturn As Integer

        shared_connect()
        Try
            strResult = executeSQLQuery(" SELECT PRD_ID FROM PRODUIT WHERE PRD_CODE = '" & strKey & "'")
            nReturn = CInt(strResult)
        Catch ex As Exception
            nReturn = -1
        End Try
        shared_disconnect()
        Return nReturn
    End Function

    '=======================================================================
    'Fonction : LoadPRDLight
    'Description : Chargement de l'objet en base
    'Retour : Rend Vrai si le chargement s'est correctement effectué
    '=======================================================================
    Protected Function loadPRDLight() As Boolean

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "L'id doit être renseigné")

        Dim sqlString As String = " SELECT " & _
                                         " PRODUIT.PRD_ID, " & _
                                         " PRODUIT.PRD_CODE, " & _
                                        " PRODUIT.PRD_LIBELLE, " & _
                                        " PRODUIT.PRD_MOT_CLE, " & _
                                        " PRODUIT.PRD_FRN_ID, " & _
                                        " PRODUIT.PRD_CONT_ID, " & _
                                        " PRODUIT.PRD_COND_ID, " & _
                                        " PRODUIT.PRD_COUL_ID, " & _
                                        " PRODUIT.PRD_MIL, " & _
                                        " PRODUIT.PRD_RGN_ID, " & _
                                        " PRODUIT.PRD_TVA_ID, " & _
                                        " PRODUIT.PRD_DATE_DERN_INVENT, " & _
                                        " PRODUIT.PRD_QTE_STK, " & _
                                        " PRODUIT.PRD_QTE_STOCK_DERN_INVENT, " & _
                                        " PRODUIT.PRD_DISPO, " & _
                                        " PRODUIT.PRD_CODE_STAT, " & _
                                        " PRODUIT.PRD_STOCK, " & _
                                        " PRD_TARIFA, PRD_TARIFB, PRD_TARIFC" & _
                                        " FROM PRODUIT  " & _
                                " WHERE " & _
                                " PRD_ID = ?"

        Dim objOLeDBCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        Dim objPRD As Produit = Nothing
        Dim bReturn As Boolean

        objOLeDBCommand = New OleDbCommand
        objOLeDBCommand.Connection = m_dbconn.Connection
        objOLeDBCommand.CommandText = sqlString



        'Si le paramètre existe on ne met à jour que sa valeur
        'Autrement on le crée
        CreateParameterP_ID(objOLeDBCommand)

        Try
            'objRS = objOLeDBCommand.ExecuteNonQuery()
            objRS = objOLeDBCommand.ExecuteReader()

            If Not objRS.HasRows Then
                objRS.Close()
                Return False
            End If
            objRS.Read()
            objPRD = CType(Me, Produit)
            objPRD.Fill(objRS)
            m_bResume = True
            cleanErreur()
            bReturn = True
            objRS.Close()
            objRS = Nothing

        Catch ex As Exception
            setError("LoadPRDLight", ex.ToString())

            bReturn = False
        End Try
        Debug.Assert(bReturn, getErreur())

        Return bReturn
    End Function 'LoadPRDLight
    '=======================================================================
    'Fonction : DeletePRD
    'Description : Suppression d'un PRODUIT
    'Retour : Rend Vrai si le DELETE s'est correctement effectué
    '=======================================================================
    Protected Function deletePRD() As Boolean
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "Id <> 0")
        Debug.Assert(Me.GetType().Name.Equals("Produit"), "Objet de Type Produit Requis")


        Dim sqlString As String = "DELETE FROM PRODUIT WHERE PRD_ID=? "
        Dim objOLeDBCommand As OleDbCommand
        Dim objPRD As Produit

        objPRD = CType(Me, Produit)
        objOLeDBCommand = New OleDbCommand
        objOLeDBCommand.Connection = m_dbconn.Connection
        objOLeDBCommand.CommandText = sqlString



        CreateParameterP_ID(objOLeDBCommand)
        Try
            objOLeDBCommand.ExecuteNonQuery()
            m_id = 0
            objPRD.code = ""
            objPRD.nom = ""
            Return True

        Catch ex As Exception
            setError("deletePRD", ex.ToString())

            Return False
        End Try
    End Function 'DeletePRD
    '=======================================================================
    'Fonction : ListePRD()
    'Description : Rend une liste de produit suivant les critères 
    '                   Code, Libelle, MotClé, IdFournisseur, idClient (TABLE PRECOMMANDE)
    'Détails    :  
    'Retour : une collection
    '=======================================================================
    Protected Shared Function ListePRD(ByVal pTypeProduit As vncTypeProduit, Optional ByVal strCode As String = "", Optional ByVal strLibelle As String = "", Optional ByVal strMotCle As String = "", Optional ByVal idFournisseur As Integer = 0, Optional ByVal idClient As Integer = 0) As Collection
        Dim colReturn As New Collection
        Dim strWhere As String = ""
        Dim nId As Integer

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")

        Dim sqlString As String = " SELECT PRODUIT.PRD_ID, PRODUIT.PRD_CODE, PRODUIT.PRD_LIBELLE, CONTENANT.CONT_LIBELLE, CONTENANT.CONT_ID, PRODUIT.PRD_MIL, RQ_Couleur.PAR_VALUE, produit.PRD_coul_Id, produit.PRD_cont_id " &
                                    " FROM (CONTENANT INNER JOIN PRODUIT ON CONTENANT.CONT_ID = PRODUIT.PRD_CONT_ID) INNER JOIN RQ_Couleur ON PRODUIT.PRD_COUL_ID = RQ_Couleur.PAR_ID"

        Dim objOLeDBCommand As OleDbCommand
        Dim objPRD As Produit
        Dim objRS As OleDbDataReader = Nothing
        objOLeDBCommand = m_dbconn.Connection.CreateCommand()
        'objOLeDBCommand.Transaction = m_dbconn.transaction

        If pTypeProduit = vncEnums.vncTypeProduit.vncPlateforme Then
            strWhere = " PRD_STOCK = 1 "
        End If
        If pTypeProduit = vncEnums.vncTypeProduit.vncFournisseur Then
            strWhere = " PRD_STOCK = 0 "
        End If

        If strCode <> "" Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & " PRD_CODE LIKE ?"
            objOLeDBCommand.Parameters.AddWithValue("?" , strCode)

        End If
        If strLibelle <> "" Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "PRD_LIBELLE LIKE ?"
            objOLeDBCommand.Parameters.AddWithValue("?" , strLibelle)

        End If

        If strMotCle <> "" Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "PRD_MOT_CLE LIKE ?"
            objOLeDBCommand.Parameters.AddWithValue("?" , strMotCle)

        End If

        If idFournisseur <> 0 Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "PRD_FRN_ID =  ?"
            objOLeDBCommand.Parameters.AddWithValue("?" , idFournisseur)

        End If

        If idClient <> 0 Then
            sqlString = "SELECT PRODUIT.PRD_ID, " &
                                " PRODUIT.PRD_CODE, " &
                                        " PRODUIT.PRD_LIBELLE, " &
                                        " PRODUIT.PRD_MOT_CLE, " &
                                        " PRODUIT.PRD_FRN_ID, " &
                                        " PRODUIT.PRD_CONT_ID, " &
                                        " PRODUIT.PRD_COND_ID, " &
                                        " PRODUIT.PRD_COUL_ID, " &
                                        " PRODUIT.PRD_MIL, " &
                                        " PRODUIT.PRD_RGN_ID, " &
                                        " PRODUIT.PRD_TVA_ID, " &
                                        " PRODUIT.PRD_DATE_DERN_INVENT, " &
                                        " PRODUIT.PRD_QTE_STK, " &
                                        " PRODUIT.PRD_QTE_STOCK_DERN_INVENT, " &
                                        " PRODUIT.PRD_DISPO, " &
                                        " PRODUIT.PRD_CODE_STAT, " &
                                        " PRODUIT.PRD_STOCK, " &
                                        " PRD_TARIFA, PRD_TARIFB, PRD_TARIFC"

            sqlString = sqlString & " FROM CLIENT INNER JOIN PRECOMMANDE ON CLIENT.CLT_ID = PRECOMMANDE.PCMD_CLT_ID INNER JOIN PRODUIT ON PRECOMMANDE.PCMD_PRD_ID = PRODUIT.PRD_ID"
            If strWhere <> "" Then
                strWhere = strWhere & " And "
            End If
            strWhere = strWhere & " PRECOMMANDE.PCMD_CLT_ID = ? "
            objOLeDBCommand.Parameters.AddWithValue("?", idClient)

        End If

        If strWhere <> "" Then
            sqlString = sqlString & " WHERE " & strWhere
        End If
        sqlString = sqlString & " ORDER BY PRD_CODE"


        objOLeDBCommand.CommandText = sqlString



        Try
            Dim Temp As Date = DateTime.Now
            objRS = objOLeDBCommand.ExecuteReader()
            Dim oTimeSpan As TimeSpan = DateTime.Now - Temp
            Debug.WriteLine("ListPrd Time = " & oTimeSpan.TotalMilliseconds())
            While (objRS.Read())
                nId = getInteger(objRS, "PRD_ID")
                objPRD = New Produit()
                objPRD.Fill(objRS)
                ''objPRD.DBLoadLight(nId)
                colReturn.Add(objPRD, objPRD.code)
            End While
            objRS.Close()
            objRS = Nothing
            oTimeSpan = DateTime.Now - Temp
            Debug.WriteLine("ListPrd Time = " & oTimeSpan.TotalMilliseconds())
            Return colReturn
        Catch ex As Exception
            setError("ListPRD", ex.ToString())
            Debug.Assert(False, getErreur())
            If objRS IsNot Nothing Then
                objRS.Close()
            End If
            Return New Collection
        End Try
    End Function 'ListePRD

    Public Overrides Function toString() As String
        Return "Persist :" & m_id & "," & m_bNew & "," & m_bUpdated & "," & m_bDeleted & "," & m_bResume
    End Function

    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        Dim objPersist As Persist
        Dim bReturn As Boolean
        Try
            bReturn = MyBase.Equals(obj)
            objPersist = obj
            bReturn = bReturn And (m_id.Equals(objPersist.id))
        Catch ex As Exception
            bReturn = False
        End Try
        Return bReturn
    End Function 'Equals
    '=======================================================================
    'Fonction : deletePRECOMMANDE
    'Description : Suppression de la PRECOMMANDE d'un client
    'Retour : 
    '=======================================================================
    Protected Function deletePRECOMMANDE() As Boolean
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "Id <> 0")


        Dim sqlString As String = "DELETE FROM PRECOMMANDE WHERE PCMD_CLT_ID=? "
        Dim objOLeDBCommand As OleDbCommand

        objOLeDBCommand = New OleDbCommand
        objOLeDBCommand.Connection = m_dbconn.Connection
        objOLeDBCommand.CommandText = sqlString



        CreateParameterP_ID(objOLeDBCommand)
        Try
            objOLeDBCommand.ExecuteNonQuery()
            Return True

        Catch ex As Exception
            setError("DeletePRECOMMANDE", ex.ToString())
            Return False
        End Try

    End Function 'deletePRECOMMANDE
    '=======================================================================
    'Fonction : insertPRECOMMANDE
    'Description : Sauvegarde des lignes de PRECOMMANDE
    'Retour : 
    '=======================================================================
    Protected Function insertPRECOMMANDE(ByVal colLg As ColEvent) As Boolean
        Dim objPrecommande As preCommande
        Dim bReturn As Boolean
        Dim objLgPrecom As lgPrecomm


        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "Le Client doit être Sauvegardé")
        Debug.Assert(Me.typeDonnee = vncTypeDonnee.PRECOMMANDE, "Objet de type Precommande requis")
        Debug.Assert(Not colLg Is Nothing, "ColLg is Nothing")

        objPrecommande = CType(Me, preCommande)

        Dim sqlString As String = "INSERT INTO PRECOMMANDE( " & _
                                    "PCMD_CLT_ID, " & _
                                    "PCMD_PRD_ID, " & _
                                    "PCMD_QTE_HAB, " & _
                                    "PCMD_QTE_DERN, " & _
                                    "PCMD_PRIXU, " & _
                                    "PCMD_DATEDERNCMD, " & _
                                    "PCMD_REFDERNCMD )" & _
                                  " VALUES (" & _
                                  "? ," & _
                                  "? ," & _
                                  "? ," & _
                                  "? , " & _
                                  "? , " & _
                                  "? , " & _
                                  "?" & _
                                    " )"
        Dim objOLeDBCommand As OleDbCommand
        Dim objOLeDBCommand2 As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing


        objOLeDBCommand = New OleDbCommand
        objOLeDBCommand2 = New OleDbCommand
        objOLeDBCommand.Connection = m_dbconn.Connection
        objOLeDBCommand.CommandText = sqlString
        m_dbconn.BeginTransaction()
        objOLeDBCommand.Transaction = m_dbconn.transaction



        bReturn = True
        CreateParameterP_ID(objOLeDBCommand)
        Dim oParam1 As OleDbParameter = objOLeDBCommand.Parameters.AddWithValue("?" , OleDbType.Integer)
        Dim oParam2 As OleDbParameter = objOLeDBCommand.Parameters.AddWithValue("?" , OleDbType.Double)
        Dim oParam3 As OleDbParameter = objOLeDBCommand.Parameters.AddWithValue("?" , OleDbType.Double)
        Dim oParam4 As OleDbParameter = objOLeDBCommand.Parameters.AddWithValue("?" , OleDbType.Double)
        Dim oParam5 As OleDbParameter = objOLeDBCommand.Parameters.AddWithValue("?" , OleDbType.Date)
        Dim oParam6 As OleDbParameter = objOLeDBCommand.Parameters.AddWithValue("?" , OleDbType.VarChar)
        For Each objLgPrecom In colLg
            Try
                oParam1.Value = objLgPrecom.idProduit
                oParam2.Value = objLgPrecom.qteHab
                oParam3.Value = objLgPrecom.qteDern
                oParam4.Value = objLgPrecom.prixU
                oParam5.Value = objLgPrecom.dateDerniereCommande.ToShortDateString
                oParam6.Value = truncate(objLgPrecom.refDerniereCommande, 50)
                objOLeDBCommand.ExecuteNonQuery()

                objOLeDBCommand2 = New OleDbCommand("SELECT MAX(PCMD_ID) FROM PRECOMMANDE", m_dbconn.Connection)
                objOLeDBCommand2.Transaction = m_dbconn.transaction
                objRS = objOLeDBCommand2.ExecuteReader()
                If objRS.Read() Then
                    objLgPrecom.setid(objRS.GetInt32(0))
                    objRS.Close()

                    objRS = Nothing
                End If

            Catch ex As Exception
                setError("InsertPRECOMMANDE", ex.ToString())
                bReturn = False
                Exit For
            End Try
        Next
        If bReturn Then
            m_dbconn.transaction.commit()
        Else
            m_dbconn.transaction.RollBack()
        End If

        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'InsertPRECOMMANDE
    '=======================================================================
    'Fonction : LoadCLTPRECOMMANDE
    'Description : Chargement des lignes de PRECOMMANDE
    'Retour : Boolean
    '=======================================================================
    Protected Function LoadCLTPRECOMMANDE() As Boolean
        Debug.Assert(m_id <> 0, "Objet Chargé")
        Debug.Assert(Me.GetType().Name.ToUpper().Equals("PreCommande".ToUpper()), "Objet <> PreCommande")

        Dim objOLeDBCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        Dim objPRECOMMANDE As preCommande
        Dim bReturn As Boolean
        Dim idProduit As Integer
        Dim codeProduit As String
        Dim libProduit As String
        Dim millesime As Integer
        Dim libConditionnement As String
        Dim qtehab As Decimal
        Dim qtedern As Decimal
        Dim prixU As Double
        Dim dateDernCommande As Date
        Dim refDernCommande As String
        Dim objLgPrecom As lgPrecomm
        Dim idLg As Integer


        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "L'id doit être renseigné")

        Dim sqlString As String = "SELECT PCMD_ID, PCMD_PRD_ID, PCMD_QTE_HAB, PCMD_QTE_DERN , PCMD_PRIXU, PCMD_DATEDERNCMD,PCMD_REFDERNCMD, PRODUIT.PRD_CODE, PRODUIT.PRD_LIBELLE, PRODUIT.PRD_MIL, RQ_CONDITIONNEMENT.PAR_CODE" & _
                                  " FROM (PRECOMMANDE INNER JOIN PRODUIT ON PRECOMMANDE.PCMD_PRD_ID = PRODUIT.PRD_ID) INNER JOIN RQ_CONDITIONNEMENT ON PRODUIT.PRD_COND_ID = RQ_CONDITIONNEMENT.PAR_ID" & _
                                  " WHERE PRECOMMANDE.PCMD_CLT_ID=? "
        Dim sqlORDERBY As String = " ORDER BY PRODUIT.PRD_CODE"

        sqlString = sqlString & sqlORDERBY

        objOLeDBCommand = New OleDbCommand
        objOLeDBCommand.Connection = m_dbconn.Connection
        objOLeDBCommand.CommandText = sqlString



        'Paramétre ClientID 
        CreateParameterP_ID(objOLeDBCommand)

        Try
            'objRS = objOLeDBCommand.ExecuteNonQuery()
            objRS = objOLeDBCommand.ExecuteReader()

            objPRECOMMANDE = CType(Me, preCommande)
            bReturn = True
            While objRS.Read()

                Try
                    idLg = getInteger(objRS, "PCMD_ID")
                    idProduit = getInteger(objRS, "PCMD_PRD_ID")
                    codeProduit = GetString(objRS, "PRD_CODE")
                    libProduit = GetString(objRS, "PRD_LIBELLE")
                    millesime = GetString(objRS, "PRD_MIL")
                    libConditionnement = GetString(objRS, "PAR_CODE")
                    qtehab = GetString(objRS, "PCMD_QTE_HAB")
                    qtedern = GetString(objRS, "PCMD_QTE_DERN")
                    prixU = GetString(objRS, "PCMD_PRIXU")
                    dateDernCommande = GetString(objRS, "PCMD_DATEDERNCMD")
                    refDernCommande = GetString(objRS, "PCMD_REFDERNCMD")
                    objLgPrecom = objPRECOMMANDE.ajouteLgPrecom(idProduit, codeProduit, libProduit, qtehab, qtedern, prixU, dateDernCommande, refDernCommande)
                    If Not objLgPrecom Is Nothing Then
                        objLgPrecom.setid(idLg)
                        objLgPrecom.millesime = millesime
                        objLgPrecom.libConditionnement = libConditionnement
                    End If
                Catch ex As Exception
                    setError("LoadCLTPRECOMMANDE", "Erreur en ajout de ligne PRECOMMANDE" & ex.Message)
                    bReturn = False
                    Exit While
                End Try


            End While
            objRS.Close()
            objRS = Nothing

        Catch ex As Exception
            setError("LoadCLTPRECOMMANDE", ex.ToString())
            bReturn = False
        End Try

        Debug.Assert(bReturn, getErreur())
        Return bReturn

    End Function 'LoadCLTPRECOMMANDE
    Protected Function UPDATECLTPRECOMMANDE() As Boolean
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "Le Client.id  <> 0")
        Debug.Assert(m_typedonnee = vncEnums.vncTypeDonnee.PRECOMMANDE, "Objet de type PreCommande requis")


        '                                    " PCMD_CLT_ID = P_PCMD_CLT_ID," & _
        '                                    " PCMD_PRD_ID = P_PCMD_PRD_ID," & _
        Dim sqlString As String = "UPDATE PRECOMMANDE SET " & _
                                    " PCMD_QTE_DERN = ?," & _
                                    " PCMD_QTE_HAB = ?," & _
                                    " PCMD_PRIXU = ?," & _
                                    " PCMD_DATEDERNCMD = ?," & _
                                    " PCMD_REFDERNCMD = ?" & _
                                    " WHERE " & _
                                    " PCMD_ID = ?"
        Dim objOLeDBCommand As OleDbCommand
        Dim objPreCommande As preCommande
        Dim bReturn As Boolean
        Dim objlgPrecom As lgPrecomm
        Dim i As Integer
        '        Dim oSW As New Stopwatch()


        objPreCommande = CType(Me, preCommande)
        objOLeDBCommand = New OleDbCommand
        objOLeDBCommand.Connection = m_dbconn.Connection
        objOLeDBCommand.CommandText = sqlString



        bReturn = True
        m_dbconn.BeginTransaction()
        objOLeDBCommand.Transaction = m_dbconn.transaction
        For i = 1 To objPreCommande.getlgPrecomCount
            objlgPrecom = objPreCommande.getLgPrecom(i)
            If (objlgPrecom.bUpdated) Then
                Try
                    objOLeDBCommand.Parameters.Clear()
                    objOLeDBCommand.Parameters.AddWithValue("?", objlgPrecom.qteDern)
                    objOLeDBCommand.Parameters.AddWithValue("?", objlgPrecom.qteHab)
                    objOLeDBCommand.Parameters.AddWithValue("?", objlgPrecom.prixU)
                    objOLeDBCommand.Parameters.AddWithValue("?", objlgPrecom.dateDerniereCommande)
                    objOLeDBCommand.Parameters.AddWithValue("?", truncate(objlgPrecom.refDerniereCommande, 50))
                    objOLeDBCommand.Parameters.AddWithValue("?", objlgPrecom.id)
                    objOLeDBCommand.ExecuteNonQuery()
                Catch ex As Exception
                    setError("UpdatePRECOMMANDE", ex.ToString())
                    bReturn = False
                    Exit For
                End Try
            End If
        Next i
        If bReturn Then
            m_dbconn.transaction.Commit()
        Else
            m_dbconn.transaction.Rollback()
        End If
        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function ' UPDATEPRECOMMANDE

#Region "Mouvement de Stocks"
#Region "Fonction BD"
    '=======================================================================
    'Fonction : ListeMVTSTK()
    'Description : Rend une liste des Mouvements de Stock d'un produit
    '                   Code, Libelle, MotClé, IdFournisseur, idClient (TABLE PRECOMMANDE)
    'Détails    :  
    'Retour : une collection triée
    '=======================================================================
    Protected Shared Function ListeMVTSTK(ByVal pidProduit As Integer, Optional ByVal ptypeMvt As vncEnums.vncTypeMvt = vncEnums.vncTypeMvt.vncmvtRegul, Optional ByVal pIDCmd As Integer = -1) As ColEventSorted
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(pidProduit <> 0, "L'idProduit doit être renseigné")

        Dim objCommand As OLEDBCommand
        Dim objRS As OleDbDataReader = Nothing
        '        Dim objParam As OleDbParameter
        Dim colReturn As New ColEventSorted

        Dim idMvt As Integer
        Dim idPRDmvt As Integer
        Dim datemvt As Date
        Dim typemvt As vncTypeMvt
        Dim refidmvt As Integer
        Dim libmvt As String
        Dim qtemvt As Decimal
        Dim commvt As String
        Dim objmvt As mvtStock



        Dim sqlString As String = " SELECT [STK_ID], [STK_PRD_ID], [STK_DATE], [STK_TYPE], [STK_REF_ID], [STK_LIB], [STK_QTE], [STK_COMM]" & _
                                  " FROM MVT_STOCK "
        Dim strOrder As String = " STK_DATE DESC, MVT_STOCK.STK_TYPE DESC"
        Dim strWhere As String = " MVT_STOCK.STK_PRD_ID = ? "

        If pIDCmd <> -1 Then
            If Len(Trim(strwhere)) <> 0 Then
                strwhere = strwhere & " AND "
            End If
            strWhere = strWhere & " MVT_STOCK.STK_TYPE = ? "
            strwhere = strwhere & " AND "
            strWhere = strWhere & " MVT_STOCK.STK_REF_ID = ? "
        End If


        sqlString = sqlString & " WHERE " & strwhere & " ORDER BY " & strOrder
        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        'Paramétre ProduitID 

        Try
            Dim objParam As OleDbParameter
            objParam = objCommand.Parameters.AddWithValue("?" , pidProduit)
            If pIDCmd <> -1 Then
                objParam = objCommand.Parameters.AddWithValue("?" , ptypeMvt)

                objParam = objCommand.Parameters.AddWithValue("?" , pIDCmd)

            End If

            objRS = objCommand.ExecuteReader()
            colReturn = New ColEventSorted
            While (objRS.Read())
                idMvt = GetString(objRS, "STK_ID")
                idPRDmvt = GetString(objRS, "STK_PRD_ID")
                datemvt = GetString(objRS, "STK_DATE")
                typemvt = GetString(objRS, "STK_TYPE")
                refidmvt = GetString(objRS, "STK_REF_ID")
                libmvt = GetString(objRS, "STK_LIB")
                qtemvt = GetString(objRS, "STK_QTE")
                commvt = GetString(objRS, "STK_COMM")
                'Ajout de la ligne dans le produit
                objmvt = New mvtStock(datemvt, idPRDmvt, typemvt, qtemvt, libmvt)
                objmvt.setid(idMvt)
                objmvt.Commentaire = commvt
                objmvt.idReference = refidmvt
                objmvt.resetBooleans()
                colReturn.Add(objmvt, objmvt.key)
            End While
            objRS.Close()

        Catch ex As Exception
            setError("LoadMVTSTK", ex.ToString())
            Debug.Assert(False, getErreur())
            colReturn = New ColEventSorted
        End Try

        Return colReturn

    End Function 'ListeMVTSTK

    '=======================================================================
    'Fonction : ListeMVTSTK()
    'Description : Rend une liste des Mouvements de Stock d'un Founisseur classé par ordre chronologique
    'Détails    :  
    'Retour : une collection triée
    '=======================================================================
    Protected Shared Function ListeMVTSTK2(ByVal pddeb As Date, ByVal pdfin As Date, Optional ByVal pFourn As Fournisseur = Nothing, Optional ByVal pEtat As vncEtatMVTSTK = vncEtatMVTSTK.vncMVTSTK_Tous) As Collection
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")

        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        Dim colReturn As New Collection

        Dim idMvt As Integer
        Dim idPRDmvt As Integer
        Dim datemvt As Date
        Dim typemvt As vncTypeMvt
        Dim refidmvt As Integer
        Dim libmvt As String
        Dim qtemvt As Decimal
        Dim commvt As String
        Dim objmvt As mvtStock
        Dim etat As Integer
        Dim idFactColisage As Integer



        Dim sqlString As String = " SELECT [STK_ID], [STK_PRD_ID], [STK_DATE], [STK_TYPE], [STK_REF_ID], [STK_LIB], [STK_QTE], [STK_COMM]," & _
                                  " STK_ETAT, STK_IDFACTCOLISAGE " & _
                                  " FROM MVT_STOCK  INNER JOIN PRODUIT ON MVT_STOCK.STK_PRD_ID = PRODUIT.PRD_ID"
        Dim strOrder As String = " STK_DATE ASC, MVT_STOCK.STK_TYPE DESC"
        Dim strWhere As String = " MVT_STOCK.STK_DATE >= ? AND  " _
                                    & " MVT_STOCK.STK_DATE <= ? AND " _
                                    & " PRODUIT.PRD_STOCK = 1"

        If pEtat <> vncEtatMVTSTK.vncMVTSTK_Tous Then
            strWhere = strWhere & " AND  MVT_STOCK.STK_ETAT = ? "
        End If

        If pFourn IsNot Nothing Then
            strWhere = strWhere & " AND PRODUIT.PRD_FRN_ID = ? "
        End If


        sqlString = sqlString & " WHERE " & strWhere & " ORDER BY " & strOrder
        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        'Paramétre ProduitID 

        Try
            Dim objParam As OleDbParameter
            objParam = objCommand.Parameters.AddWithValue("?", pddeb)
            objParam = objCommand.Parameters.AddWithValue("?", pdfin)
            If pEtat <> vncEtatMVTSTK.vncMVTSTK_Tous Then
                objParam = objCommand.Parameters.AddWithValue("?", pEtat)
            End If
            If pFourn IsNot Nothing Then
                objParam = objCommand.Parameters.AddWithValue("?", pFourn.id)
            End If

            objRS = objCommand.ExecuteReader()
            colReturn = New Collection
            While (objRS.Read())
                idMvt = GetString(objRS, "STK_ID")
                idPRDmvt = GetString(objRS, "STK_PRD_ID")
                datemvt = GetString(objRS, "STK_DATE")
                typemvt = GetString(objRS, "STK_TYPE")
                refidmvt = GetString(objRS, "STK_REF_ID")
                libmvt = GetString(objRS, "STK_LIB")
                qtemvt = GetString(objRS, "STK_QTE")
                commvt = GetString(objRS, "STK_COMM")
                etat = GetString(objRS, "STK_ETAT")
                idFactColisage = GetString(objRS, "STK_IDFACTCOLISAGE")
                'Ajout de la ligne dans la collection
                objmvt = New mvtStock(datemvt, idPRDmvt, typemvt, qtemvt, libmvt)
                objmvt.setid(idMvt)
                objmvt.Commentaire = commvt
                objmvt.idReference = refidmvt
                objmvt.Etat = EtatMvtStock.createEtat(etat)
                objmvt.idFactColisage = idFactColisage
                objmvt.resetBooleans()
                colReturn.Add(objmvt)
            End While
            objRS.Close()

        Catch ex As Exception
            setError("ListeMVTSTK", ex.ToString())
            Debug.Assert(False, getErreur())
            colReturn = New Collection
        End Try

        Return colReturn

    End Function 'ListeMVTSTK
    '=======================================================================
    'Fonction : ListeMVTSTK()
    'Description : Rend une liste des Mouvements de Stock d'une facture de colisage
    '                   Code, Libelle, MotClé, IdFournisseur, idClient (TABLE PRECOMMANDE)
    'Détails    :  
    'Retour : une collection triée
    '=======================================================================
    Protected Shared Function ListeMVTSTK_FACTCOL(ByVal pidFactcolisage As Integer) As ColEventSorted
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(pidFactcolisage <> 0, "L'idProduit doit être renseigné")

        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        '        Dim objParam As OleDbParameter
        Dim colReturn As New ColEventSorted

        Dim idMvt As Integer
        Dim idPRDmvt As Integer
        Dim datemvt As Date
        Dim typemvt As vncTypeMvt
        Dim refidmvt As Integer
        Dim libmvt As String
        Dim qtemvt As Decimal
        Dim commvt As String
        Dim objmvt As mvtStock



        Dim sqlString As String = " SELECT [STK_ID], [STK_PRD_ID], [STK_DATE], [STK_TYPE], [STK_REF_ID], [STK_LIB], [STK_QTE], [STK_COMM]" & _
                                  " FROM MVT_STOCK "
        Dim strOrder As String = " STK_DATE DESC, MVT_STOCK.STK_TYPE DESC"
        Dim strWhere As String = ""

        If Len(Trim(strWhere)) <> 0 Then
            strWhere = strWhere & " AND "
        End If
        strWhere = strWhere & " MVT_STOCK.STK_IDFACTCOLISAGE = ?FACTCOLISAGE "


        sqlString = sqlString & " WHERE " & strWhere & " ORDER BY " & strOrder
        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        'Paramétre ProduitID 

        Try
            Dim objParam As OleDbParameter
            objParam = objCommand.Parameters.AddWithValue("?" , pidFactcolisage)

            objRS = objCommand.ExecuteReader()
            colReturn = New ColEventSorted
            While (objRS.Read())
                idMvt = GetString(objRS, "STK_ID")
                idPRDmvt = GetString(objRS, "STK_PRD_ID")
                datemvt = GetString(objRS, "STK_DATE")
                typemvt = GetString(objRS, "STK_TYPE")
                refidmvt = GetString(objRS, "STK_REF_ID")
                libmvt = GetString(objRS, "STK_LIB")
                qtemvt = GetString(objRS, "STK_QTE")
                commvt = GetString(objRS, "STK_COMM")
                'Ajout de la ligne dans le produit
                objmvt = New mvtStock(datemvt, idPRDmvt, typemvt, qtemvt, libmvt)
                objmvt.setid(idMvt)
                objmvt.Commentaire = commvt
                objmvt.idReference = refidmvt
                objmvt.resetBooleans()
                colReturn.Add(objmvt, objmvt.key)
            End While
            objRS.Close()

        Catch ex As Exception
            setError("ListeMVTSTK_FACTCOL", ex.ToString())
            Debug.Assert(False, getErreur())
            colReturn = New ColEventSorted
        End Try

        Return colReturn

    End Function 'ListeMVTSTK_FACTCOL

    '=======================================================================
    'Fonction : deletePRDMVTStock
    'Description : Suppression des Mvts de Stocks d'un produit
    'Retour : 
    '=======================================================================
    Protected Function deletePRDMVTStock() As Boolean
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "Id <> 0")
        Debug.Assert(Me.GetType().Name.Equals("Produit"))


        Dim sqlString As String = "DELETE FROM MVT_STOCK WHERE STK_PRD_ID=? "
        Dim objCommand As OleDbCommand
        Dim obj As Produit

        obj = CType(Me, Produit)
        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString

        CreateParameterP_ID(objCommand)
        Try
            objCommand.ExecuteNonQuery()
            Return True

        Catch ex As Exception
            setError("deletePRDMVTStock", ex.ToString())
            Return False
        End Try

    End Function ' deletePRDMVTStock
    '=======================================================================
    'Fonction : deletePRDPreCommande
    'Description : Suppression des lignes de précommande  d'un produit
    'Retour : 
    '=======================================================================
    Protected Function deletePRDPreCommande() As Boolean
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "Id <> 0")
        Debug.Assert(Me.GetType().Name.Equals("Produit"))


        Dim sqlString As String = "DELETE FROM PRECOMMANDE WHERE PCMD_PRD_ID=?"
        Dim objCommand As OleDbCommand
        Dim obj As Produit

        obj = CType(Me, Produit)
        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        CreateParameterP_ID(objCommand)
        Try
            objCommand.ExecuteNonQuery()
            Return True

        Catch ex As Exception
            setError("deletePRDPreCommande", ex.ToString())
            Return False
        End Try

    End Function ' deletePRDPreCommande
    ''=======================================================================
    ''Fonction : insertPRDMVTStock
    ''Description : Sauvegarde des Lignes de mouvements de stocks
    ''Retour : 
    ''=======================================================================
    'Protected Function insertPRDMVTStock() As Boolean
    '    Dim objPRD As Produit
    '    Dim bReturn As Boolean
    '    Dim objmvtStock As mvtStock

    '    Debug.Assert(shared_isConnected(), "La database doit être ouverte")
    '    Debug.Assert(m_id <> 0, "Le Produit doit être Sauvegardé")
    '    Debug.Assert(Me.GetType().Name.Equals("Produit"))

    '    objPRD = CType(Me, Produit)

    '    bReturn = True
    '    '        objTransaction = m_dbconn.BeginTransaction()
    '    For Each objmvtStock In objPRD.colmvtStock
    '        If Not (objmvtStock.bDeleted) Then

    '            bReturn = bReturn And objmvtStock.insertMVTSTK
    '        End If

    '    Next

    '    Debug.Assert(bReturn, getErreur())
    '    Return bReturn
    'End Function 'insertPRDMVTStock

    '=======================================================================
    'Fonction : LoadMVTSTK
    'Description : Chargement de l'objet en base
    'Retour : Rend Vrai si le chargement s'est correctement effectué
    '=======================================================================
    Protected Function loadMVTSTK() As Boolean

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "L'id doit être renseigné")

        Dim sqlString As String = " SELECT [STK_ID], [STK_PRD_ID], [STK_DATE], [STK_TYPE], [STK_REF_ID], [STK_LIB], [STK_QTE], [STK_COMM], [STK_ETAT], [STK_IDFACTCOLISAGE] " & _
                                  "  FROM MVT_STOCK " & _
                                  " WHERE STK_ID = ?"

        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        Dim objMvt As mvtStock
        Dim bReturn As Boolean

        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        'Si le paramètre existe on ne met à jour que sa valeur
        'Autrement on le crée
        CreateParameterP_ID(objCommand)

        Try
            objRS = objCommand.ExecuteReader
            If Not objRS.HasRows Then
                objRS.Close()
                Return False
            End If
            objRS.Read()
            objMvt = CType(Me, mvtStock)
            objMvt.idProduit = GetString(objRS, "STK_PRD_ID")
            objMvt.datemvt = GetString(objRS, "STK_DATE")
            objMvt.typeMvt = GetString(objRS, "STK_TYPE")
            objMvt.idReference = GetString(objRS, "STK_REF_ID")
            objMvt.libelle = GetString(objRS, "STK_LIB")
            objMvt.qte = GetString(objRS, "STK_QTE")
            objMvt.Commentaire = GetString(objRS, "STK_COMM")
            objMvt.Etat = EtatMvtStock.createEtat(GetValue(objRS, "STK_ETAT"))
            objMvt.idFactColisage = GetValue(objRS, "STK_IDFACTCOLISAGE")
            bReturn = True

        Catch ex As Exception
            setError("LoadMVTSTK", ex.ToString())
            bReturn = False
        End Try
        If (Not objRS Is Nothing) Then
            objRS.Close()

        End If
        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'LoadMVTSTK
    '=======================================================================
    'Fonction : InsertMVTSTK
    'Description : Insertion d'un mouvement de stock
    'Retour : Rend Vrai si l'INSERT s'est correctement effectué
    '=======================================================================
    Protected Function insertMVTSTK() As Boolean
        Dim obj As mvtStock
        Dim bReturn As Boolean

        '        Debug.Assert(Me.GetType().Name.Equals(obj.GetType.Name), "Objet de Type mvtStock Requis")
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        'Debug.Assert(m_id = 0, "ID=0")
        obj = CType(Me, mvtStock)

        Dim sqlString As String = " INSERT INTO MVT_STOCK (" & _
                                    " STK_PRD_ID , " & _
                                    " STK_DATE , " & _
                                    " STK_TYPE , " & _
                                    " STK_REF_ID , " & _
                                    " STK_LIB , " & _
                                    " STK_QTE , " & _
                                    " STK_COMM ,  " & _
                                    " STK_ETAT , " & _
                                    " STK_IDFACTCOLISAGE  " & _
                                    " ) VALUES (" & _
                                    " ? , " & _
                                    " ? , " & _
                                    " ? , " & _
                                    " ? , " & _
                                    " ? , " & _
                                    " ? , " & _
                                    " ? ,  " & _
                                    " ? ,  " & _
                                    " ?  " & _
                                    " )"
        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing


        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString
        m_dbconn.BeginTransaction()
        objCommand.Transaction = m_dbconn.transaction


        CreateParamP_STK_PRD_ID(objCommand)
        CreateParamP_STK_DATE(objCommand)
        CreateParamP_STK_TYPE(objCommand)
        CreateParamP_STK_REF_ID(objCommand)
        CreateParamP_STK_LIB(objCommand)
        CreateParamP_STK_QTE(objCommand)
        CreateParamP_STK_COMM(objCommand)
        CreateParamP_STK_ETAT(objCommand)
        CreateParamP_STK_IDFACTCOLISAGE(objCommand)
        Try
            objCommand.ExecuteNonQuery()
            objCommand = New OleDbCommand("SELECT @@IDENTITY FROM MVT_STOCK", m_dbconn.Connection)
            objCommand.Transaction = m_dbconn.transaction
            objRS = objCommand.ExecuteReader
            If (objRS.Read) Then
                m_id = objRS.GetDecimal(0)
                cleanErreur()
                bReturn = True
            Else
                bReturn = False
            End If

        Catch ex As Exception
            setError("InsertMVTSTK", ex.ToString())
            bReturn = False
        End Try
        If Not objRS Is Nothing Then
            objRS.Close()
        End If
        If bReturn Then
            m_dbconn.transaction.commit()
        Else
            m_dbconn.transaction.RollBack()
        End If

        '    Debug.Assert(m_id <> 0, "ID=0")
        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'InsertMVTSTK

    '=======================================================================
    'Fonction : UpdateMVTSTK
    'Description : Insertion d'un Mouvement de Stock
    'Retour : Rend Vrai si l'UPDATE s'est correctement effectué
    '=======================================================================
    Protected Function updateMVTSTK() As Boolean
        Dim objMVT As mvtStock
        Dim bReturn As Boolean

        'Debug.Assert(Me.GetType.Name.Equals(objMVT.GetType.Name), "Objet de type MvtStock Requis")
        objMVT = CType(Me, mvtStock)
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "ID<>0")

        Dim sqlString As String = " UPDATE MVT_STOCK SET " & _
                                    " STK_PRD_ID = ? , " & _
                                    " STK_DATE =  ? , " & _
                                    " STK_TYPE = ? , " & _
                                    " STK_REF_ID = ? , " & _
                                    " STK_LIB = ? , " & _
                                    " STK_QTE = ? , " & _
                                    " STK_COMM = ? , " & _
                                    " STK_ETAT = ? , " & _
                                    " STK_IDFACTCOLISAGE = ? " & _
                                  " WHERE STK_ID = ?"
        Dim objCommand As OleDbCommand

        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        CreateParamP_STK_PRD_ID(objCommand)
        CreateParamP_STK_DATE(objCommand)
        CreateParamP_STK_TYPE(objCommand)
        CreateParamP_STK_REF_ID(objCommand)
        CreateParamP_STK_LIB(objCommand)
        CreateParamP_STK_QTE(objCommand)
        CreateParamP_STK_COMM(objCommand)
        CreateParamP_STK_ETAT(objCommand)
        CreateParamP_STK_IDFACTCOLISAGE(objCommand)
        CreateParameterP_ID(objCommand)


        Try
            objCommand.ExecuteNonQuery()
            cleanErreur()
            bReturn = True

        Catch ex As Exception
            setError("UpdateMVTSTK", ex.ToString())
            bReturn = False
        End Try

        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'UpdateMVTSTK
    '=======================================================================
    'Fonction : DeleteMVTSTK
    'Description : Suppression d'un Mouvement de Stock
    'Retour : Rend Vrai si le DELETE s'est correctement effectué
    '=======================================================================
    Protected Function deleteMVTSTK() As Boolean
        Dim obj As mvtStock
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "Id <> 0")
        'Debug.Assert(Me.GetType().Name.Equals(obj.GetType.Name))


        Dim sqlString As String = "DELETE FROM MVT_STOCK WHERE STK_ID=? "
        Dim objCommand As OleDbCommand

        obj = CType(Me, mvtStock)
        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        CreateParameterP_ID(objCommand)
        Try
            objCommand.ExecuteNonQuery()
            m_id = 0
            Return True

        Catch ex As Exception
            setError("DeleteMVTSTK", ex.ToString())
            Return False
        End Try
    End Function 'DeleteMVTSTK
    '=======================================================================
    'Fonction : purgePRDMvtStock
    'Description : Purge des Mvts de stocks d'un produit antérieurs à une date
    'Retour : 
    '=======================================================================
    Protected Function purgePRDMVTStock(ByVal pdatePurge As Date) As Boolean
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")

        Dim sqlString As String = "DELETE FROM MVT_STOCK WHERE STK_DATE < ? AND STK_PRD_ID = ?"
        Dim objCommand As OleDbCommand

        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        objCommand.Parameters.AddWithValue("?" , pdatePurge.ToShortDateString)
        CreateParameterP_ID(objCommand)
        Try
            objCommand.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            setError("deletePRDMVTStock", ex.ToString())
            Return False
        End Try

    End Function ' PurgePRDMVTStock

    Protected Function existeMvtSocklib(ByVal strLib As String) As Boolean
        Dim bReturn As Boolean = True
        shared_connect()
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")

        Dim sqlString As String = "SELECT * FROM MVT_STOCK WHERE STK_LIB = ?"
        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        Try

            objCommand = New OleDbCommand
            objCommand.Connection = m_dbconn.Connection
            objCommand.CommandText = sqlString



            objCommand.Parameters.AddWithValue("?" , strLib)

            objRS = objCommand.ExecuteReader
            If Not objRS.HasRows() Then
                bReturn = False
            Else
                bReturn = True
            End If

            objRS.Close()
        Catch ex As Exception
            setError("Persist.existeMvtStockLib()", ex.ToString)
            If Not objRS Is Nothing Then
                objRS.Close()
            End If
            Throw ex
        End Try


        shared_disconnect()
        Return bReturn
    End Function

#End Region

#Region "CreateParametersMVT_STK"
    '============
    'TABLE MVTSTK
    '============
    Private Sub CreateParamP_STK_PRD_ID(ByVal objCommand As OleDbCommand)
        Dim obj As mvtStock
        obj = Me
        objCommand.Parameters.AddWithValue("?", obj.idProduit)
    End Sub
    Private Sub CreateParamP_STK_DATE(ByVal objCommand As OleDbCommand)
        Dim obj As mvtStock
        obj = Me
        objCommand.Parameters.AddWithValue("?", obj.datemvt.ToShortDateString)
    End Sub
    Private Sub CreateParamP_STK_REF_ID(ByVal objCommand As OleDbCommand)
        Dim obj As mvtStock
        obj = Me
        objCommand.Parameters.AddWithValue("?", obj.idReference)
    End Sub

    Private Sub CreateParamP_STK_TYPE(ByVal objCommand As OleDbCommand)
        Dim obj As mvtStock
        obj = Me
        objCommand.Parameters.AddWithValue("?", obj.typeMvt)
    End Sub
    Private Sub CreateParamP_STK_QTE(ByVal objCommand As OleDbCommand)
        Dim obj As mvtStock
        obj = Me
        objCommand.Parameters.AddWithValue("?", CInt(obj.qte))
    End Sub
    Private Sub CreateParamP_STK_LIB(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim obj As mvtStock
        obj = Me
        objCommand.Parameters.AddWithValue("?", truncate(obj.libelle, 50))
    End Sub
    Private Sub CreateParamP_STK_COMM(ByVal objCommand As OleDbCommand)
        Dim obj As mvtStock
        obj = Me
        objCommand.Parameters.AddWithValue("?", truncate(obj.Commentaire, 255))
    End Sub
    Private Sub CreateParamP_STK_ETAT(ByVal objCommand As OleDbCommand)
        Dim obj As mvtStock
        obj = Me
        objCommand.Parameters.AddWithValue("?", obj.Etat.codeEtat)
    End Sub
    Private Sub CreateParamP_STK_IDFACTCOLISAGE(ByVal objCommand As OleDbCommand)
        Dim obj As mvtStock
        obj = Me
        objCommand.Parameters.AddWithValue("?", obj.idFactColisage)
    End Sub

#End Region

#End Region
#Region "Creation des Paramètres ADO"
    '===================================================================================================
    ' CREATION DES PARAMETRES ADO
    '================================================================================================
    '============
    'TABLE FOURNISSEUR
    '============
    Private Sub CreateParamP_FRN_CODE(ByVal objCommand As OleDbCommand)
        Dim objFRN As Fournisseur
        objFRN = Me
        objCommand.Parameters.AddWithValue("?", truncate(objFRN.code, 50))
    End Sub
    Private Sub CreateParamP_FRN_NOM(ByVal objCommand As OleDbCommand)
        Dim objFRN As Fournisseur
        objFRN = Me

        objCommand.Parameters.AddWithValue("?", truncate(objFRN.nom, 50))
    End Sub
    Private Sub CreateParamP_FRN_RS(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objFRN As Fournisseur
        objFRN = Me
        objCommand.Parameters.AddWithValue("?", truncate(objFRN.rs, 50))
    End Sub
    Private Sub CreateParamP_FRN_RIB1(ByVal objCommand As OleDbCommand)
        Dim objFRN As Fournisseur
        objFRN = Me
        objCommand.Parameters.AddWithValue("?", truncate(objFRN.rib1, 5))
    End Sub
    Private Sub CreateParamP_FRN_RIB2(ByVal objCommand As OleDbCommand)
        Dim objFRN As Fournisseur
        objFRN = Me
        objCommand.Parameters.AddWithValue("?", truncate(objFRN.rib2, 5))
    End Sub
    Private Sub CreateParamP_FRN_RIB3(ByVal objCommand As OleDbCommand)
        Dim objFRN As Fournisseur
        objFRN = Me
        objCommand.Parameters.AddWithValue("?", truncate(objFRN.rib3, 11))
    End Sub
    Private Sub CreateParamP_FRN_RIB4(ByVal objCommand As OleDbCommand)
        Dim objFRN As Fournisseur
        objFRN = Me
        objCommand.Parameters.AddWithValue("?", truncate(objFRN.rib4, 2))
    End Sub
    Private Sub CreateParamP_FRN_BANQUE(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objFRN As Fournisseur
        objFRN = Me
        objCommand.Parameters.AddWithValue("?", truncate(objFRN.banque, 50))
    End Sub
    Private Sub CreateParamP_FRN_RGN_ID(ByVal objCommand As OleDbCommand)
        Dim objFRN As Fournisseur
        objFRN = Me
        objCommand.Parameters.AddWithValue("?", objFRN.idRegion)
    End Sub
    Private Sub CreateParamP_FRN_SIRET(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objFRN As Fournisseur
        objFRN = Me
        objCommand.Parameters.AddWithValue("?", truncate(objFRN.siret, 50))
    End Sub
    Private Sub CreateParamP_FRN_TVAINTRACOM(ByVal objCommand As OleDbCommand)
        Dim objFRN As Fournisseur
        objFRN = Me
        objCommand.Parameters.AddWithValue("?", truncate(objFRN.numtvaintracom, 50))
    End Sub
    Private Sub CreateParamP_FRN_CODECOMPTA(ByVal objCommand As OleDbCommand)
        Dim objFRN As Fournisseur
        objFRN = Me
        objCommand.Parameters.AddWithValue("?", truncate(objFRN.CodeCompta, 10))
    End Sub
    Private Sub CreateParamP_FRN_RGLMT_ID(ByVal objCommand As OleDbCommand)
        Dim objFRN As Fournisseur
        objFRN = Me
        objCommand.Parameters.AddWithValue("?", objFRN.idModeReglement)
    End Sub
    Private Sub CreateParamP_FRN_ID_MRGLMT1(ByVal objCommand As OleDbCommand)
        Dim objFRN As Fournisseur
        objFRN = Me
        objCommand.Parameters.AddWithValue("?", objFRN.idModeReglement1)
    End Sub
    Private Sub CreateParamP_FRN_ID_MRGLMT2(ByVal objCommand As OleDbCommand)
        Dim objFRN As Fournisseur
        objFRN = Me
        objCommand.Parameters.AddWithValue("?", objFRN.idModeReglement2)
    End Sub
    Private Sub CreateParamP_FRN_ID_MRGLMT3(ByVal objCommand As OleDbCommand)
        Dim objFRN As Fournisseur
        objFRN = Me
        objCommand.Parameters.AddWithValue("?", objFRN.idModeReglement3)
    End Sub
    Private Sub CreateParamP_FRN_IDPRESTASHOP(ByVal objCommand As OleDbCommand)
        Dim objFRN As Fournisseur
        objFRN = Me
        objCommand.Parameters.AddWithValue("?", objFRN.idPrestashop)
    End Sub
    Private Sub CreateParamP_FRN_BINTERMEDIAIRE(ByVal objCommand As OleDbCommand)
        Dim objFRN As Fournisseur
        objFRN = Me
        objCommand.Parameters.AddWithValue("?", objFRN.bIntermdiaire)
    End Sub
    Private Sub CreateParamP_FRN_DOSSIER(ByVal objCommand As OleDbCommand)
        Dim objFRN As Fournisseur
        objFRN = Me
        objCommand.Parameters.AddWithValue("?", objFRN.Dossier)
    End Sub
    Private Sub CreateParamP_FRN_LIV_NOM(ByVal objCommand As OleDbCommand)
        Dim objFRN As Fournisseur
        objFRN = Me
        objCommand.Parameters.AddWithValue("?", truncate(objFRN.AdresseLivraison.nom, 50))
    End Sub
    Private Sub CreateParamP_FRN_LIV_RUE1(ByVal objCommand As OleDbCommand)
        Dim objFRN As Fournisseur
        objFRN = Me
        objCommand.Parameters.AddWithValue("?", truncate(objFRN.AdresseLivraison.rue1, 50))
    End Sub
    Private Sub CreateParamP_FRN_LIV_RUE2(ByVal objCommand As OleDbCommand)
        Dim objFRN As Fournisseur
        objFRN = Me
        objCommand.Parameters.AddWithValue("?", truncate(objFRN.AdresseLivraison.rue2, 50))
    End Sub
    Private Sub CreateParamP_FRN_LIV_CP(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objFRN As Fournisseur
        objFRN = Me
        objCommand.Parameters.AddWithValue("?", truncate(objFRN.AdresseLivraison.cp, 50))
    End Sub
    Private Sub CreateParamP_FRN_LIV_VILLE(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objFRN As Fournisseur
        objFRN = Me
        objCommand.Parameters.AddWithValue("?", truncate(objFRN.AdresseLivraison.ville, 50))
    End Sub
    Private Sub CreateParamP_FRN_LIV_TEL(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objFRN As Fournisseur
        objFRN = Me
        objCommand.Parameters.AddWithValue("?", truncate(objFRN.AdresseLivraison.tel, 50))
    End Sub
    Private Sub CreateParamP_FRN_LIV_FAX(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objFRN As Fournisseur
        objFRN = Me
        objCommand.Parameters.AddWithValue("?", truncate(objFRN.AdresseLivraison.fax, 50))
    End Sub
    Private Sub CreateParamP_FRN_LIV_PORT(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objFRN As Fournisseur
        objFRN = Me
        objCommand.Parameters.AddWithValue("?", truncate(objFRN.AdresseLivraison.port, 50))
    End Sub
    Private Sub CreateParamP_FRN_LIV_EMAIL(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objFRN As Fournisseur
        objFRN = Me
        objCommand.Parameters.AddWithValue("?", truncate(objFRN.AdresseLivraison.Email, 50))
    End Sub
    Private Sub CreateParamP_FRN_FACT_NOM(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objFRN As Fournisseur
        objFRN = Me
        objCommand.Parameters.AddWithValue("?", truncate(objFRN.AdresseFacturation.nom, 50))
    End Sub
    Private Sub CreateParamP_FRN_FACT_RUE1(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objFRN As Fournisseur
        objFRN = Me
        objCommand.Parameters.AddWithValue("?", truncate(objFRN.AdresseFacturation.rue1, 50))
    End Sub
    Private Sub CreateParamP_FRN_FACT_RUE2(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objFRN As Fournisseur
        objFRN = Me
        objCommand.Parameters.AddWithValue("?", truncate(objFRN.AdresseFacturation.rue2, 50))
    End Sub
    Private Sub CreateParamP_FRN_FACT_CP(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objFRN As Fournisseur
        objFRN = Me
        objCommand.Parameters.AddWithValue("?", truncate(objFRN.AdresseFacturation.cp, 50))
    End Sub
    Private Sub CreateParamP_FRN_FACT_VILLE(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objFRN As Fournisseur
        objFRN = Me
        objCommand.Parameters.AddWithValue("?", truncate(objFRN.AdresseFacturation.ville, 50))
    End Sub
    Private Sub CreateParamP_FRN_FACT_TEL(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objFRN As Fournisseur
        objFRN = Me
        objCommand.Parameters.AddWithValue("?", truncate(objFRN.AdresseFacturation.tel, 50))
    End Sub
    Private Sub CreateParamP_FRN_FACT_FAX(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objFRN As Fournisseur
        objFRN = Me
        objCommand.Parameters.AddWithValue("?", truncate(objFRN.AdresseFacturation.fax, 50))
    End Sub
    Private Sub CreateParamP_FRN_FACT_PORT(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objFRN As Fournisseur
        objFRN = Me
        objCommand.Parameters.AddWithValue("?", truncate(objFRN.AdresseFacturation.port, 50))
    End Sub
    Private Sub CreateParamP_FRN_FACT_EMAIL(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objFRN As Fournisseur
        objFRN = Me
        objCommand.Parameters.AddWithValue("?", truncate(objFRN.AdresseFacturation.Email, 50))
    End Sub
    'Fournisseur - Commentaire de Commande 
    Private Sub CreateParamP_FRN_COM_CMD(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objFRN As Fournisseur
        objFRN = Me
        objCommand.Parameters.AddWithValue("?", objFRN.CommCommande.comment)
    End Sub
    'Fournisseur - Commentaire de Livraison 
    Private Sub CreateParamP_FRN_COM_LIV(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objFRN As Fournisseur
        objFRN = Me
        objCommand.Parameters.AddWithValue("?", objFRN.CommLivraison.comment)
    End Sub
    'Fournisseur - Commentaire de Facturation 
    Private Sub CreateParamP_FRN_COM_FACT(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objFRN As Fournisseur
        objFRN = Me
        objCommand.Parameters.AddWithValue("?", objFRN.CommFacturation.comment)
    End Sub
    'Fournisseur - Commentaire Libre 
    Private Sub CreateParamP_FRN_COM_LIBRE(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objFRN As Fournisseur
        objFRN = Me
        objCommand.Parameters.AddWithValue("?", objFRN.CommLibre.comment)
    End Sub
    'Fournisseur - Adresses identiques
    Private Sub CreateParamP_FRN_ADR_IDENT(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objFRN As Fournisseur
        objFRN = Me
        objCommand.Parameters.AddWithValue("?", objFRN.bAdressesIdentiques)
    End Sub
    'Fournisseur - Export internet
    Private Sub CreateParamP_FRN_BEXP_INTERNET(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objFRN As Fournisseur
        objFRN = Me
        objCommand.Parameters.AddWithValue("?", objFRN.bExportInternet)
    End Sub

    '============
    'TABLE CLIENT
    '============
    Private Sub CreateParameterP_ID(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        objCommand.Parameters.AddWithValue("?", m_id)
    End Sub
    Private Sub CreateParamP_CLT_CODE(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objCLT As Client
        objCLT = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCLT.code, 50))
    End Sub
    Private Sub CreateParamP_CLT_NOM(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objCLT As Client
        objCLT = Me

        objCommand.Parameters.AddWithValue("?", truncate(objCLT.nom, 50))
    End Sub
    Private Sub CreateParamP_CLT_RS(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objCLT As Client
        objCLT = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCLT.rs, 50))
    End Sub
    Private Sub CreateParamP_CLT_RIB1(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objCLT As Client
        objCLT = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCLT.rib1, 5))
    End Sub
    Private Sub CreateParamP_CLT_RIB2(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objCLT As Client
        objCLT = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCLT.rib2, 5))
    End Sub
    Private Sub CreateParamP_CLT_RIB3(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objCLT As Client
        objCLT = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCLT.rib3, 11))
    End Sub
    Private Sub CreateParamP_CLT_RIB4(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objCLT As Client
        objCLT = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCLT.rib4, 2))
    End Sub
    Private Sub CreateParamP_CLT_BANQUE(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objCLT As Client
        objCLT = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCLT.banque, 50))
    End Sub
    Private Sub CreateParamP_CLT_TYPE_ID(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objCLT As Client
        objCLT = Me
        objCommand.Parameters.AddWithValue("?", objCLT.idTypeClient)
    End Sub
    Private Sub CreateParamP_CLT_SIRET(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objCLT As Client
        objCLT = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCLT.siret, 50))
    End Sub
    Private Sub CreateParamP_CLT_TVAINTRACOM(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objCLT As Client
        objCLT = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCLT.numtvaintracom, 50))
    End Sub
    Private Sub CreateParamP_CLT_RGLMT_ID(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objCLT As Client
        objCLT = Me
        objCommand.Parameters.AddWithValue("?", objCLT.idModeReglement)
    End Sub
    Private Sub CreateParamP_CLT_ID_MRGLMT1(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objCLT As Client
        objCLT = Me
        objCommand.Parameters.AddWithValue("?", objCLT.idModeReglement1)
    End Sub
    Private Sub CreateParamP_CLT_ID_MRGLMT2(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objCLT As Client
        objCLT = Me
        objCommand.Parameters.AddWithValue("?", objCLT.idModeReglement2)
    End Sub
    Private Sub CreateParamP_CLT_ID_MRGLMT3(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objCLT As Client
        objCLT = Me
        objCommand.Parameters.AddWithValue("?", objCLT.idModeReglement3)
    End Sub
    Private Sub CreateParamP_CLT_IDPRESASHOP(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objCLT As Client
        objCLT = Me
        objCommand.Parameters.AddWithValue("?", objCLT.idPrestashop)
    End Sub
    Private Sub CreateParamP_CLT_ORIGINE(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objCLT As Client
        objCLT = Me
        objCommand.Parameters.AddWithValue("?", objCLT.Origine)
    End Sub
    Private Sub CreateParamP_CLT_LIV_NOM(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objCLT As Client
        objCLT = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCLT.AdresseLivraison.nom, 50))
    End Sub
    Private Sub CreateParamP_CLT_LIV_RUE1(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objCLT As Client
        objCLT = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCLT.AdresseLivraison.rue1, 50))
    End Sub
    Private Sub CreateParamP_CLT_LIV_RUE2(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objCLT As Client
        objCLT = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCLT.AdresseLivraison.rue2, 50))
    End Sub
    Private Sub CreateParamP_CLT_LIV_CP(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objCLT As Client
        objCLT = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCLT.AdresseLivraison.cp, 50))
    End Sub
    Private Sub CreateParamP_CLT_LIV_VILLE(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objCLT As Client
        objCLT = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCLT.AdresseLivraison.ville, 50))
    End Sub
    Private Sub CreateParamP_CLT_LIV_TEL(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objCLT As Client
        objCLT = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCLT.AdresseLivraison.tel, 50))
    End Sub
    Private Sub CreateParamP_CLT_LIV_FAX(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objCLT As Client
        objCLT = Me
        objCommand.Parameters.AddWithValue("?", Truncate(objCLT.AdresseLivraison.fax,50))
    End Sub
    Private Sub CreateParamP_CLT_LIV_PORT(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objCLT As Client
        objCLT = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCLT.AdresseLivraison.port, 50))
    End Sub
    Private Sub CreateParamP_CLT_LIV_EMAIL(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objCLT As Client
        objCLT = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCLT.AdresseLivraison.Email, 50))
    End Sub
    Private Sub CreateParamP_CLT_FACT_NOM(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objCLT As Client
        objCLT = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCLT.AdresseFacturation.nom, 50))
    End Sub
    Private Sub CreateParamP_CLT_FACT_RUE1(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objCLT As Client
        objCLT = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCLT.AdresseFacturation.rue1, 50))
    End Sub
    Private Sub CreateParamP_CLT_FACT_RUE2(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objCLT As Client
        objCLT = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCLT.AdresseFacturation.rue2, 50))
    End Sub
    Private Sub CreateParamP_CLT_FACT_CP(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objCLT As Client
        objCLT = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCLT.AdresseFacturation.cp, 50))
    End Sub
    Private Sub CreateParamP_CLT_FACT_VILLE(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objCLT As Client
        objCLT = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCLT.AdresseFacturation.ville, 50))
    End Sub
    Private Sub CreateParamP_CLT_FACT_TEL(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objCLT As Client
        objCLT = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCLT.AdresseFacturation.tel, 50))
    End Sub
    Private Sub CreateParamP_CLT_FACT_FAX(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objCLT As Client
        objCLT = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCLT.AdresseFacturation.fax, 50))
    End Sub
    Private Sub CreateParamP_CLT_FACT_PORT(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objCLT As Client
        objCLT = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCLT.AdresseFacturation.port, 50))
    End Sub
    Private Sub CreateParamP_CLT_FACT_EMAIL(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objCLT As Client
        objCLT = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCLT.AdresseFacturation.Email, 50))
    End Sub
    'Client - Commentaire de Commande 
    Private Sub CreateParamP_CLT_COM_CMD(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objCLT As Client
        objCLT = Me
        objCommand.Parameters.AddWithValue("?", objCLT.CommCommande.comment)
    End Sub
    'Client - Commentaire de Livraison 
    Private Sub CreateParamP_CLT_COM_LIV(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objCLT As Client
        objCLT = Me
        objCommand.Parameters.AddWithValue("?", objCLT.CommLivraison.comment)
    End Sub
    'Client - Commentaire de Facturation 
    Private Sub CreateParamP_CLT_COM_FACT(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objCLT As Client
        objCLT = Me
        objCommand.Parameters.AddWithValue("?", objCLT.CommFacturation.comment)
    End Sub
    'Client - Commentaire Libre 
    Private Sub CreateParamP_CLT_COM_LIBRE(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objCLT As Client
        objCLT = Me
        objCommand.Parameters.AddWithValue("?", objCLT.CommLibre.comment)
    End Sub
    'Client - Adresses identiques
    Private Sub CreateParamP_CLT_ADR_IDENT(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objCLT As Client
        objCLT = Me
        objCommand.Parameters.AddWithValue("?", objCLT.bAdressesIdentiques)
    End Sub
    'Client - Code Tarif
    Private Sub CreateParamP_CLT_CODETARIF(ByVal objCommand As OleDbCommand)
        Dim objCLT As Client
        objCLT = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCLT.CodeTarif, 2))
    End Sub
    'Fournisseur - Code Compta
    Private Sub CreateParamP_CLT_CODECOMPTA(ByVal objCommand As OleDbCommand)
        Dim objCLT As Client
        objCLT = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCLT.CodeCompta, 10))
    End Sub

#Region "CreateParameter Produit"
    '=====================================================================================================================
    'TABLE PRODUIT '
    '===============
    Private Sub CreateParamP_PRD_CODE(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objPRD As Produit
        objPRD = Me
        objCommand.Parameters.AddWithValue("?", truncate(objPRD.code, 50))
    End Sub
    Private Sub CreateParamP_PRD_LIBELLE(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objPRD As Produit
        objPRD = Me

        objCommand.Parameters.AddWithValue("?", truncate(objPRD.nom, 50))
    End Sub
    Private Sub CreateParamP_PRD_FRN_ID(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objPRD As Produit
        objPRD = Me
        objCommand.Parameters.AddWithValue("?", objPRD.idFournisseur)
    End Sub
    Private Sub CreateParamP_PRD_MOT_CLE(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objPRD As Produit
        objPRD = Me
        objCommand.Parameters.AddWithValue("?", truncate(objPRD.motcle, 50))
    End Sub
    Private Sub CreateParamP_PRD_CONT_ID(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objPRD As Produit
        objPRD = Me
        objCommand.Parameters.AddWithValue("?", objPRD.idContenant)
    End Sub
    Private Sub CreateParamP_PRD_COND_ID(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objPRD As Produit
        objPRD = Me
        objCommand.Parameters.AddWithValue("?", objPRD.idConditionnement)
    End Sub
    Private Sub CreateParamP_PRD_COUL_ID(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objPRD As Produit
        objPRD = Me
        objCommand.Parameters.AddWithValue("?", objPRD.idCouleur)
    End Sub
    Private Sub CreateParamP_PRD_MIL(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objPRD As Produit
        objPRD = Me
        objCommand.Parameters.AddWithValue("?", objPRD.millesime)
    End Sub
    Private Sub CreateParamP_PRD_RGN_ID(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objPRD As Produit
        objPRD = Me
        objCommand.Parameters.AddWithValue("?", objPRD.idRegion)
    End Sub
    Private Sub CreateParamP_PRD_TVA_ID(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objPRD As Produit
        objPRD = Me
        objCommand.Parameters.AddWithValue("?", objPRD.idTVA)
    End Sub
    Private Sub CreateParamP_PRD_DATE_DERN_INVENT(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objPRD As Produit
        objPRD = Me
        objCommand.Parameters.AddWithValue("?", CDate(objPRD.DateDernInventaire.ToShortDateString))
    End Sub
    Private Sub CreateParamP_PRD_QTE_STK(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objPRD As Produit
        objPRD = Me
        objCommand.Parameters.AddWithValue("?", CInt(objPRD.QteStock))
    End Sub
    Private Sub CreateParamP_PRD_QTE_STOCK_DERN_INVENT(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objPRD As Produit
        objPRD = Me
        objCommand.Parameters.AddWithValue("?", CInt(objPRD.QteStockDernInventaire))
    End Sub
    Private Sub CreateParamP_PRD_DISPO(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objPRD As Produit
        objPRD = Me
        objCommand.Parameters.AddWithValue("?", objPRD.bDisponible)
    End Sub
    Private Sub CreateParamP_PRD_STOCK(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objPRD As Produit
        objPRD = Me
        objCommand.Parameters.AddWithValue("?", objPRD.bStock)
    End Sub
    Private Sub CreateParamP_PRD_CODE_STAT(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objPRD As Produit
        objPRD = Me
        objCommand.Parameters.AddWithValue("?", truncate(objPRD.codeStat, 50))
    End Sub
    Private Sub CreateParamP_PRD_DOSSIER(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objPRD As Produit
        objPRD = Me
        objCommand.Parameters.AddWithValue("?", truncate(objPRD.Dossier, 50))
    End Sub
    Private Sub CreateParamP_PRD_TARIFA(ByVal objCommand As OleDbCommand)
        Dim objPRD As Produit
        objPRD = Me
        objCommand.Parameters.AddWithValue("?", objPRD.TarifA)
    End Sub
    Private Sub CreateParamP_PRD_TARIFB(ByVal objCommand As OleDbCommand)
        Dim objPRD As Produit
        objPRD = Me
        objCommand.Parameters.AddWithValue("?", objPRD.TarifB)
    End Sub
    Private Sub CreateParamP_PRD_TARIFC(ByVal objCommand As OleDbCommand)
        Dim objPRD As Produit
        objPRD = Me
        objCommand.Parameters.AddWithValue("?", objPRD.TarifC)
    End Sub
#End Region
#End Region

#Region "BD Commande Client"
    'Renvoie une collection d'objet résumé de CommandeClient
    Protected Shared Function ListeCMDCLT(ByVal strCode As String, ByVal strRSClient As String, ByVal pEtat As vncEtatCommande, pOrigine As String) As Collection
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Dim colReturn As New Collection
        '        Dim objParam As OleDbParameter
        Dim sqlString As String = "SELECT CMD_ID, CMD_CODE, CMD_DATE, CLT_ID , CLT_CODE, CLT_RS, CMD_ETAT , CMD_TOTAL_HT " & _
                                    " FROM COMMANDE , CLIENT "
        Dim strWhere As String = " COMMANDE.CMD_CLT_ID = CLIENT.CLT_ID "
        Dim objCommand As OleDbCommand
        Dim objCMD As CommandeClient
        Dim objRS As OleDbDataReader = Nothing
        Dim objParam As OleDbParameter
        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection

        If strCode <> "" Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "CMD_CODE LIKE ?"
            objParam = objCommand.Parameters.AddWithValue("?", strCode)
        End If

        If strRSClient <> "" Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "CLIENT.CLT_RS LIKE ?"
            objParam = objCommand.Parameters.AddWithValue("?", strRSClient)

        End If

        If pEtat <> vncEnums.vncEtatCommande.vncRien Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "CMD_ETAT = ?"
            objParam = objCommand.Parameters.AddWithValue("?", pEtat)

        End If

        If pOrigine <> "" Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "CMD_ORIGINE = ?"
            objParam = objCommand.Parameters.AddWithValue("?", pOrigine)

        End If

        If strWhere <> "" Then
            sqlString = sqlString & "WHERE " & strWhere
        End If
        sqlString = sqlString & " ORDER BY CMD_DATE DESC,CMD_CODE DESC"
        objCommand.CommandText = sqlString



        Try
            objRS = objCommand.ExecuteReader
            While (objRS.Read())
                objCMD = New CommandeClient(New Client("", ""))
                objCMD.m_bResume = True ' C'est un objet Résumé
                Try
                    objCMD.m_id = GetString(objRS, "CMD_ID")
                Catch ex As InvalidCastException
                    objCMD.m_id = 0
                End Try
                Try
                    objCMD.code = GetString(objRS, "CMD_CODE")
                Catch ex As InvalidCastException
                    objCMD.code = ""
                End Try
                Try
                    objCMD.oTiers.setid(GetString(objRS, "CLT_ID"))
                Catch ex As InvalidCastException
                    objCMD.oTiers.setid(0)
                End Try
                Try
                    objCMD.oTiers.code = GetString(objRS, "CLT_CODE")
                Catch ex As InvalidCastException
                    objCMD.oTiers.code = ""
                End Try
                Try
                    objCMD.oTiers.rs = GetString(objRS, "CLT_RS")
                Catch ex As InvalidCastException
                    objCMD.oTiers.rs = ""
                End Try

                Try
                    objCMD.dateCommande = GetString(objRS, "CMD_DATE")
                Catch ex As InvalidCastException
                    objCMD.dateCommande = Now()
                End Try

                Try
                    objCMD.totalHT = GetString(objRS, "CMD_TOTAL_HT")
                Catch ex As InvalidCastException
                    objCMD.totalHT = 0
                End Try

                Try
                    objCMD.setEtat(GetString(objRS, "CMD_ETAT"))
                Catch ex As InvalidCastException
                End Try
                objCMD.releasecolLignes()
                objCMD.resetBooleans()
                colReturn.Add(objCMD, objCMD.code)
            End While
            objRS.Close()
            Return colReturn
        Catch ex As Exception
            setError("ListCMDCLT", ex.ToString())
            Return Nothing
        End Try
    End Function 'ListeCMDCLT


    Protected Function updateCMDCLT() As Boolean
        '=======================================================================
        'Fonction : UpdateCMDCLT
        'Description : Mise a jour  d'une Commande Client 

        'Retour : Rend Vrai si l'UPDATE s'est correctement effectué
        '=======================================================================
        Dim bReturn As Boolean
        Dim objCMDCLT As CommandeClient
        bReturn = False

        Debug.Assert(Me.GetType().Name.Equals("CommandeClient"), "Objet de Type Commande Client Requis")
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "ID<>0")
        objCMDCLT = CType(Me, CommandeClient)

        Dim sqlString As String = "UPDATE COMMANDE SET " & _
                                    "CMD_CLT_ID = ? , " & _
                                    "CMD_DATE = ? , " & _
                                    "CMD_DATE_VALID = ? , " & _
                                    "CMD_DATE_LIV = ? , " & _
                                    "CMD_DATE_ENLEV = ? , " & _
                                    "CMD_REF_LIV = ? , " & _
                                    "CMD_ETAT = ? , " & _
                                    "CMD_TOTAL_HT= ? ," & _
                                    "CMD_TOTAL_TTC= ? ," & _
                                    "CMD_TYPE= ? ," & _
                                    "CMD_TYPE_TRANSPORT=  ? ," & _
                                    "CMD_CODE = ? ," & _
                                    "CMD_CLT_LIV_NOM = ? ," & _
                                    "CMD_CLT_LIV_RUE1= ? ," & _
                                    "CMD_CLT_LIV_RUE2= ? ," & _
                                    "CMD_CLT_LIV_CP= ? ," & _
                                    "CMD_CLT_LIV_VILLE= ? ," & _
                                    "CMD_CLT_LIV_TEL= ? ," & _
                                    "CMD_CLT_LIV_FAX= ? ," & _
                                    "CMD_CLT_LIV_PORT= ? ," & _
                                    "CMD_CLT_LIV_EMAIL= ? ," & _
                                    "CMD_CLT_ADR_IDENT= ? ," & _
                                    "CMD_CLT_FACT_NOM= ? ," & _
                                    "CMD_CLT_FACT_RUE1= ? ," & _
                                    "CMD_CLT_FACT_RUE2= ? ," & _
                                    "CMD_CLT_FACT_CP= ? ," & _
                                    "CMD_CLT_FACT_VILLE= ? ," & _
                                    "CMD_CLT_FACT_TEL= ? ," & _
                                    "CMD_CLT_FACT_FAX= ? ," & _
                                    "CMD_CLT_FACT_PORT= ? ," & _
                                    "CMD_CLT_FACT_EMAIL= ? ," & _
                                    "CMD_CLT_RGLMT_ID= ? ," & _
                                    "CMD_CLT_BANQUE= ? ," & _
                                    "CMD_CLT_RIB1= ? ," & _
                                    "CMD_CLT_RIB2= ? ," & _
                                    "CMD_CLT_RIB3= ? ," & _
                                    "CMD_CLT_RIB4= ? ," & _
                                    "CMD_COM_LIBRE= ? ," & _
                                    "CMD_COM_COM= ? ," & _
                                    "CMD_COM_LIV= ? ," & _
                                    "CMD_COM_FACT= ? ," & _
                                    "CMD_TRP_ID= ? ," & _
                                    "CMD_TRP_CODE= ? ," & _
                                    "CMD_TRP_NOM= ? ," & _
                                    "CMD_TRP_RUE1= ? ," & _
                                    "CMD_TRP_RUE2= ? ," & _
                                    "CMD_TRP_CP= ? ," & _
                                    "CMD_TRP_VILLE= ? ," & _
                                    "CMD_TRP_TEL= ? ," & _
                                    "CMD_TRP_FAX= ? ," & _
                                    "CMD_TRP_PORT= ? ," & _
                                    "CMD_TRP_EMAIL= ? ," & _
                                    "CMD_MT_TRANSPORT= ? ," & _
                                    "CMD_QTE_COLIS= ? ," & _
                                    "CMD_QTE_PAL_PREP= ? ," & _
                                    "CMD_QTE_PAL_NONPREP= ? ," & _
                                    "CMD_POIDS= ? , " & _
                                    "CMD_PU_PAL_PREP= ? ," & _
                                    "CMD_PU_PAL_NONPREP= ? ," & _
                                    "CMD_BFACTTRP= ? , " & _
                                    "CMD_IDFACTTRP= ? , " & _
                                    "CMD_LETTREVOITURE= ? , " & _
                                    "CMD_COUT_TRANSPORT= ? , " & _
                                    "CMD_REFFACT_TRP= ? ," & _
                                    "CMD_CLT_NOM= ? ," & _
                                    "CMD_CLT_RS= ? ," & _
                                    "CMD_CLT_NOMLIVRAISON= ? ," & _
                                    "CMD_CLT_RSLIVRAISON= ?, " & _
                                    "CMD_IDPRESTASHOP= ?, " & _
                                    "CMD_NAMEPRESTASHOP= ?, " & _
                                    "CMD_ORIGINE= ? " & _
                                    " WHERE CMD_ID = ?"
        Dim objCommand As OleDbCommand
        '        Dim objParam As OleDbParameter


        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        CreateParameterP_CMD_CLT_ID(objCommand)
        CreateParameterP_CMD_DATE(objCommand)
        CreateParameterP_CMD_DATE_VALID(objCommand)
        CreateParameterP_CMD_DATE_LIV(objCommand)
        CreateParameterP_CMD_DATE_ENLEV(objCommand)
        CreateParameterP_CMD_REF_LIV(objCommand)
        CreateParameterP_CMD_ETAT(objCommand)
        CreateParameterP_CMD_TOTAL_HT(objCommand)
        CreateParameterP_CMD_TOTAL_TTC(objCommand)
        CreateParameterP_CMD_TYPE(objCommand)
        CreateParameterP_CMD_TYPE_TRANSPORT(objCommand)
        CreateParameterP_CMD_CODE(objCommand)
        CreateParameterP_CMD_CLT_LIV_NOM(objCommand)
        CreateParameterP_CMD_CLT_LIV_RUE1(objCommand)
        CreateParameterP_CMD_CLT_LIV_RUE2(objCommand)
        CreateParameterP_CMD_CLT_LIV_CP(objCommand)
        CreateParameterP_CMD_CLT_LIV_VILLE(objCommand)
        CreateParameterP_CMD_CLT_LIV_TEL(objCommand)
        CreateParameterP_CMD_CLT_LIV_FAX(objCommand)
        CreateParameterP_CMD_CLT_LIV_PORT(objCommand)
        CreateParameterP_CMD_CLT_LIV_EMAIL(objCommand)
        CreateParameterP_CMD_CLT_ADR_IDENT(objCommand)
        CreateParameterP_CMD_CLT_FACT_NOM(objCommand)
        CreateParameterP_CMD_CLT_FACT_RUE1(objCommand)
        CreateParameterP_CMD_CLT_FACT_RUE2(objCommand)
        CreateParameterP_CMD_CLT_FACT_CP(objCommand)
        CreateParameterP_CMD_CLT_FACT_VILLE(objCommand)
        CreateParameterP_CMD_CLT_FACT_TEL(objCommand)
        CreateParameterP_CMD_CLT_FACT_FAX(objCommand)
        CreateParameterP_CMD_CLT_FACT_PORT(objCommand)
        CreateParameterP_CMD_CLT_FACT_EMAIL(objCommand)
        CreateParameterP_CMD_CLT_RGLMT_ID(objCommand)
        CreateParameterP_CMD_CLT_BANQUE(objCommand)
        CreateParameterP_CMD_CLT_RIB1(objCommand)
        CreateParameterP_CMD_CLT_RIB2(objCommand)
        CreateParameterP_CMD_CLT_RIB3(objCommand)
        CreateParameterP_CMD_CLT_RIB4(objCommand)
        CreateParameterP_CMD_COM_LIBRE(objCommand)
        CreateParameterP_CMD_COM_COM(objCommand)
        CreateParameterP_CMD_COM_LIV(objCommand)
        CreateParameterP_CMD_COM_FACT(objCommand)
        CreateParameterP_CMD_TRP_ID(objCommand)
        CreateParameterP_CMD_TRP_CODE(objCommand)
        CreateParameterP_CMD_TRP_NOM(objCommand)
        CreateParameterP_CMD_TRP_RUE1(objCommand)
        CreateParameterP_CMD_TRP_RUE2(objCommand)
        CreateParameterP_CMD_TRP_CP(objCommand)
        CreateParameterP_CMD_TRP_VILLE(objCommand)
        CreateParameterP_CMD_TRP_TEL(objCommand)
        CreateParameterP_CMD_TRP_FAX(objCommand)
        CreateParameterP_CMD_TRP_PORT(objCommand)
        CreateParameterP_CMD_TRP_EMAIL(objCommand)
        CreateParameterP_CMD_MT_TRANSPORT(objCommand)
        CreateParameterP_CMD_QTE_COLIS(objCommand)
        CreateParameterP_CMD_QTE_PAL_PREP(objCommand)
        CreateParameterP_CMD_QTE_PAL_NONPREP(objCommand)
        CreateParameterP_CMD_POIDS(objCommand)
        CreateParameterP_CMD_PU_PAL_PREP(objCommand)
        CreateParameterP_CMD_PU_PAL_NONPREP(objCommand)
        CreateParameterP_CMD_BFACTTRP(objCommand)
        CreateParameterP_CMD_IDFACTTRP(objCommand)
        CreateParameterP_CMD_LETTREVOITURE(objCommand)
        CreateParameterP_CMD_COUT_TRANSPORT(objCommand)
        CreateParameterP_CMD_REFFACT_TRP(objCommand)
        CreateParameterP_CMD_CLT_NOM(objCommand)
        CreateParameterP_CMD_CLT_RS(objCommand)
        CreateParameterP_CMD_CLT_NOMLIVRAISON(objCommand)
        CreateParameterP_CMD_CLT_RSLIVRAISON(objCommand)
        CreateParameterP_CMD_IDPRESTASHOP(objCommand)
        CreateParameterP_CMD_NAMEPRESTASHOP(objCommand)
        CreateParameterP_CMD_Origine(objCommand)

        CreateParameterP_ID(objCommand)
        Try
            '            m_dbconn.BeginTransaction()
            objCommand.ExecuteNonQuery()
            '           m_dbconn.transaction.commit()
            bReturn = True

        Catch ex As Exception
            setError("UpdateCMDCLT", ex.ToString())
            bReturn = False
        End Try

        Debug.Assert(bReturn, "UpdateCMDCLT" & getErreur())
        Return bReturn
    End Function 'UpdateCMDCLT
    Protected Function insertCMDCLT() As Boolean
        '=======================================================================
        'Fonction : InsertCMDCLT
        'Description : Insertion d'une Commande Client 

        'Retour : Rend Vrai si l'INSERT s'est correctement effectué
        '=======================================================================
        Dim bReturn As Boolean
        Dim objCMDCLT As CommandeClient
        bReturn = False

        Debug.Assert(Me.GetType().Name.Equals("CommandeClient"), "Objet de Type Client Requis")
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id = 0, "ID=0")
        objCMDCLT = CType(Me, CommandeClient)
        Dim sqlString As String = "INSERT INTO COMMANDE( " & _
                                    "CMD_CLT_ID," & _
                                    "CMD_DATE," & _
                                    "CMD_DATE_VALID," & _
                                    "CMD_DATE_LIV," & _
                                    "CMD_DATE_ENLEV," & _
                                    "CMD_REF_LIV," & _
                                    "CMD_ETAT," & _
                                    "CMD_TOTAL_HT," & _
                                    "CMD_TOTAL_TTC," & _
                                    "CMD_TYPE," & _
                                    "CMD_TYPE_TRANSPORT," & _
                                    "CMD_CODE," & _
                                    "CMD_CLT_LIV_NOM," & _
                                    "CMD_CLT_LIV_RUE1," & _
                                    "CMD_CLT_LIV_RUE2," & _
                                    "CMD_CLT_LIV_CP," & _
                                    "CMD_CLT_LIV_VILLE," & _
                                    "CMD_CLT_LIV_TEL," & _
                                    "CMD_CLT_LIV_FAX," & _
                                    "CMD_CLT_LIV_PORT," & _
                                    "CMD_CLT_LIV_EMAIL," & _
                                    "CMD_CLT_ADR_IDENT," & _
                                    "CMD_CLT_FACT_NOM," & _
                                    "CMD_CLT_FACT_RUE1," & _
                                    "CMD_CLT_FACT_RUE2," & _
                                    "CMD_CLT_FACT_CP," & _
                                    "CMD_CLT_FACT_VILLE," & _
                                    "CMD_CLT_FACT_TEL," & _
                                    "CMD_CLT_FACT_FAX," & _
                                    "CMD_CLT_FACT_PORT," & _
                                    "CMD_CLT_FACT_EMAIL," & _
                                    "CMD_CLT_RGLMT_ID," & _
                                    "CMD_CLT_BANQUE," & _
                                    "CMD_CLT_RIB1," & _
                                    "CMD_CLT_RIB2," & _
                                    "CMD_CLT_RIB3," & _
                                    "CMD_CLT_RIB4," & _
                                    "CMD_COM_LIBRE," & _
                                    "CMD_COM_COM," & _
                                    "CMD_COM_LIV," & _
                                    "CMD_COM_FACT," & _
                                    "CMD_TRP_ID," & _
                                    "CMD_TRP_CODE," & _
                                    "CMD_TRP_NOM," & _
                                    "CMD_TRP_RUE1," & _
                                    "CMD_TRP_RUE2," & _
                                    "CMD_TRP_CP," & _
                                    "CMD_TRP_VILLE," & _
                                    "CMD_TRP_TEL," & _
                                    "CMD_TRP_FAX," & _
                                    "CMD_TRP_PORT," & _
                                    "CMD_TRP_EMAIL," & _
                                    "CMD_MT_TRANSPORT," & _
                                    "CMD_QTE_COLIS," & _
                                    "CMD_QTE_PAL_PREP," & _
                                    "CMD_QTE_PAL_NONPREP," & _
                                    "CMD_POIDS," & _
                                    "CMD_PU_PAL_PREP," & _
                                    "CMD_PU_PAL_NONPREP," & _
                                    "CMD_BFACTTRP," & _
                                    "CMD_IDFACTTRP," & _
                                    "CMD_LETTREVOITURE," & _
                                    "CMD_COUT_TRANSPORT," & _
                                    "CMD_REFFACT_TRP," & _
                                    "CMD_CLT_NOM," & _
                                    "CMD_CLT_RS," & _
                                    "CMD_CLT_NOMLIVRAISON," & _
                                    "CMD_CLT_RSLIVRAISON," & _
                                    "CMD_IDPRESTASHOP," & _
                                    "CMD_NAMEPRESTASHOP," & _
                                    "CMD_ORIGINE" & _
                                          " ) VALUES ( " & _
                                            "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? , " & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? " & _
                                    " )"
        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        '        Dim objParam As OleDbParameter


        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString
        m_dbconn.BeginTransaction()
        objCommand.Transaction = m_dbconn.transaction


        CreateParameterP_CMD_CLT_ID(objCommand)
        CreateParameterP_CMD_DATE(objCommand)
        CreateParameterP_CMD_DATE_VALID(objCommand)
        CreateParameterP_CMD_DATE_LIV(objCommand)
        CreateParameterP_CMD_DATE_ENLEV(objCommand)
        CreateParameterP_CMD_REF_LIV(objCommand)
        CreateParameterP_CMD_ETAT(objCommand)
        CreateParameterP_CMD_TOTAL_HT(objCommand)
        CreateParameterP_CMD_TOTAL_TTC(objCommand)
        CreateParameterP_CMD_TYPE(objCommand)
        CreateParameterP_CMD_TYPE_TRANSPORT(objCommand)
        CreateParameterP_CMD_CODE(objCommand)
        CreateParameterP_CMD_CLT_LIV_NOM(objCommand)
        CreateParameterP_CMD_CLT_LIV_RUE1(objCommand)
        CreateParameterP_CMD_CLT_LIV_RUE2(objCommand)
        CreateParameterP_CMD_CLT_LIV_CP(objCommand)
        CreateParameterP_CMD_CLT_LIV_VILLE(objCommand)
        CreateParameterP_CMD_CLT_LIV_TEL(objCommand)
        CreateParameterP_CMD_CLT_LIV_FAX(objCommand)
        CreateParameterP_CMD_CLT_LIV_PORT(objCommand)
        CreateParameterP_CMD_CLT_LIV_EMAIL(objCommand)
        CreateParameterP_CMD_CLT_ADR_IDENT(objCommand)
        CreateParameterP_CMD_CLT_FACT_NOM(objCommand)
        CreateParameterP_CMD_CLT_FACT_RUE1(objCommand)
        CreateParameterP_CMD_CLT_FACT_RUE2(objCommand)
        CreateParameterP_CMD_CLT_FACT_CP(objCommand)
        CreateParameterP_CMD_CLT_FACT_VILLE(objCommand)
        CreateParameterP_CMD_CLT_FACT_TEL(objCommand)
        CreateParameterP_CMD_CLT_FACT_FAX(objCommand)
        CreateParameterP_CMD_CLT_FACT_PORT(objCommand)
        CreateParameterP_CMD_CLT_FACT_EMAIL(objCommand)
        CreateParameterP_CMD_CLT_RGLMT_ID(objCommand)
        CreateParameterP_CMD_CLT_BANQUE(objCommand)
        CreateParameterP_CMD_CLT_RIB1(objCommand)
        CreateParameterP_CMD_CLT_RIB2(objCommand)
        CreateParameterP_CMD_CLT_RIB3(objCommand)
        CreateParameterP_CMD_CLT_RIB4(objCommand)
        CreateParameterP_CMD_COM_LIBRE(objCommand)
        CreateParameterP_CMD_COM_COM(objCommand)
        CreateParameterP_CMD_COM_LIV(objCommand)
        CreateParameterP_CMD_COM_FACT(objCommand)
        CreateParameterP_CMD_TRP_ID(objCommand)
        CreateParameterP_CMD_TRP_CODE(objCommand)
        CreateParameterP_CMD_TRP_NOM(objCommand)
        CreateParameterP_CMD_TRP_RUE1(objCommand)
        CreateParameterP_CMD_TRP_RUE2(objCommand)
        CreateParameterP_CMD_TRP_CP(objCommand)
        CreateParameterP_CMD_TRP_VILLE(objCommand)
        CreateParameterP_CMD_TRP_TEL(objCommand)
        CreateParameterP_CMD_TRP_FAX(objCommand)
        CreateParameterP_CMD_TRP_PORT(objCommand)
        CreateParameterP_CMD_TRP_EMAIL(objCommand)
        CreateParameterP_CMD_MT_TRANSPORT(objCommand)
        CreateParameterP_CMD_QTE_COLIS(objCommand)
        CreateParameterP_CMD_QTE_PAL_PREP(objCommand)
        CreateParameterP_CMD_QTE_PAL_NONPREP(objCommand)
        CreateParameterP_CMD_POIDS(objCommand)
        CreateParameterP_CMD_PU_PAL_PREP(objCommand)
        CreateParameterP_CMD_PU_PAL_NONPREP(objCommand)
        CreateParameterP_CMD_BFACTTRP(objCommand)
        CreateParameterP_CMD_IDFACTTRP(objCommand)
        CreateParameterP_CMD_LETTREVOITURE(objCommand)
        CreateParameterP_CMD_COUT_TRANSPORT(objCommand)
        CreateParameterP_CMD_REFFACT_TRP(objCommand)
        CreateParameterP_CMD_CLT_NOM(objCommand)
        CreateParameterP_CMD_CLT_RS(objCommand)
        CreateParameterP_CMD_CLT_NOMLIVRAISON(objCommand)
        CreateParameterP_CMD_CLT_RSLIVRAISON(objCommand)
        CreateParameterP_CMD_IDPRESTASHOP(objCommand)
        CreateParameterP_CMD_NAMEPRESTASHOP(objCommand)
        CreateParameterP_CMD_Origine(objCommand)
        Try

            objCommand.ExecuteNonQuery()
            objCommand = New OleDbCommand("SELECT MAX(CMD_ID) FROM COMMANDE", m_dbconn.Connection)
            objCommand.Transaction = m_dbconn.transaction
            objRS = objCommand.ExecuteReader
            If (objRS.HasRows) Then
                objRS.Read()
                m_id = objRS.GetInt32(0)
            End If
            cleanErreur()
            objRS.Close()
            objRS = Nothing
            bReturn = True

        Catch ex As Exception
            setError("InsertCMDCLT", ex.ToString())
            If Not objRS Is Nothing Then
                objRS.Close()
            End If
            bReturn = False
        End Try

        If bReturn Then
            m_dbconn.transaction.Commit()
        Else
            m_dbconn.transaction.Rollback()
        End If

        Debug.Assert(bReturn, "InsertCMDCLT: " & getErreur())
        Return bReturn
    End Function 'insertCMDCLT
    Protected Function getNumeroCommandeClient() As Integer
        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        Dim nReturn As Integer = -1
        Dim bReturn As Boolean
        Dim strStringLect As String

        Try
            strStringLect = "DECLARE @return_value int;EXEC	@return_value = GETNEXTNUM_CMD_CLT; SELECT 'CODE' = @return_value"
            objCommand = New OleDbCommand
            objCommand.Connection = m_dbconn.Connection
            objCommand = New OleDbCommand(strStringLect, m_dbconn.Connection)
            objRS = objCommand.ExecuteReader
            If objRS.HasRows Then
                objRS.Read()
                nReturn = objRS.GetInt32(0)
                objRS.Close()
                objRS = Nothing
                bReturn = True
            Else
                objRS.Close()
                bReturn = False
            End If
        Catch ex As Exception
            setError("Persist.getNumeroCommandeClient", ex.ToString())
            nReturn = -1
            bReturn = False
        End Try
        Debug.Assert(bReturn, "getNumeroCommandeClient: " & getErreur())
        Return nReturn
    End Function 'GetNumeroCommandeClient
    Protected Function getNumeroBonAppro() As Integer
        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        Dim nReturn As Integer = -1
        Dim bReturn As Boolean
        Dim strStringLect As String

        Try
            strStringLect = "DECLARE @return_value int;EXEC	@return_value = GETNEXTNUM_BA; SELECT 'CODE' = @return_value"
            objCommand = New OleDbCommand
            objCommand.Connection = m_dbconn.Connection
            objCommand = New OleDbCommand(strStringLect, m_dbconn.Connection)
            objRS = objCommand.ExecuteReader
            If objRS.HasRows Then
                objRS.Read()
                nReturn = objRS.GetInt32(0)
                objRS.Close()
                objRS = Nothing
                bReturn = True
            Else
                objRS.Close()
                bReturn = False
            End If
        Catch ex As Exception
            setError("Persist.getNumeroBonAppro", ex.ToString())
            nReturn = -1
            bReturn = False
        End Try
        Debug.Assert(bReturn, "getNumeroBonAppro: " & getErreur())
        Return nReturn
    End Function 'GetNumeroBonAppro
    Protected Function existeCommandeClient() As Boolean
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "L'id doit être renseigné")
        Debug.Assert(Me.GetType.Name.Equals("Client"), "Object de Type Client Requis")

        Dim sqlString As String = "SELECT " & _
                                    "COUNT(*)" & _
                                    " FROM COMMANDE WHERE " & _
                                   " CMD_CLT_ID = ?"

        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        Dim bReturn As Boolean
        Dim nNbre As Integer

        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        'Si le paramètre existe on ne met à jour que sa valeur
        'Autrement on le crée
        CreateParameterP_ID(objCommand)

        Try

            objRS = objCommand.ExecuteReader
            If Not objRS.HasRows Then
                bReturn = False
            Else
                objRS.Read()
                nNbre = objRS.GetInt32(0)
                If nNbre > 0 Then
                    bReturn = True
                Else
                    bReturn = False
                End If
            End If

            objRS.Close()
            objRS = Nothing

        Catch ex As Exception
            If Not objRS Is Nothing Then
                objRS.Close()
            End If
            Throw ex
        End Try
        Return bReturn
    End Function 'existeCommandeClient
    '==========================================================================
    'Function : existeLgCommandeProduit
    'Description : rend vrai s'il existe une ligne dans LG_COMMANDE Correspondant au produit
    Protected Function existeLgCommandeProduit() As Boolean
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "L'id doit être renseigné")
        Debug.Assert(Me.GetType.Name.Equals("Produit"), "Object de Type Produit Requis")

        Dim sqlString As String = "SELECT " & _
                                    "COUNT(*)" & _
                                    " FROM LIGNE_COMMANDE WHERE " & _
                                   " LGCM_PRD_ID = ?"

        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        Dim bReturn As Boolean
        Dim nNbre As Integer

        objCommand = New OleDbCommand(sqlString, m_dbconn.Connection)


        'Si le paramètre existe on ne met à jour que sa valeur
        'Autrement on le crée
        CreateParameterP_ID(objCommand)

        Try

            objRS = objCommand.ExecuteReader
            If Not objRS.HasRows Then
                objRS.Close()
                Return False
            End If
            objRS.Read()
            nNbre = objRS.GetInt32(0)
            If nNbre > 0 Then
                bReturn = True
            Else
                bReturn = False
            End If
            objRS.Close()

        Catch ex As Exception
            If Not objRS Is Nothing Then
                objRS.Close()
            End If
            Throw ex
        End Try
        Return bReturn
    End Function 'existeLgCommandeProduit

    Protected Function existeProduitFournisseur() As Boolean
        '==========================================================================
        'Function : existeProduitFournisseur
        'Description : rend vrai s'il existe un Produit correspondant au Fournisseur
        '============================================================================
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "L'id doit être renseigné")
        Debug.Assert(Me.GetType.Name.Equals("Fournisseur"), "Object de Type Fournisseur Requis")

        Dim sqlString As String = "SELECT " & _
                                    "COUNT(*)" & _
                                    " FROM PRODUIT WHERE " & _
                                   " PRD_FRN_ID = ?"

        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        Dim bReturn As Boolean
        Dim nNbre As Integer

        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        'Si le paramètre existe on ne met à jour que sa valeur
        'Autrement on le crée
        CreateParameterP_ID(objCommand)

        Try

            objRS = objCommand.ExecuteReader
            If Not objRS.HasRows Then
                objRS.Close()
                Return False
            End If
            objRS.Read()
            nNbre = objRS.GetInt32(0)
            If nNbre > 0 Then
                bReturn = True
            Else
                bReturn = False
            End If
            objRS.Close()

        Catch ex As Exception
            If Not objRS Is Nothing Then
                objRS.Close()
            End If
            Throw ex
        End Try
        Return bReturn
    End Function 'existeProduitFournisseur
    Protected Function existeBonApproFournisseur() As Boolean
        '==========================================================================
        'Function : existeBonApproFournisseur
        'Description : rend vrai s'il existe un Bon Appro correspondant au Fournisseur
        '============================================================================
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "L'id doit être renseigné")
        Debug.Assert(Me.GetType.Name.Equals("Fournisseur"), "Object de Type Fournisseur Requis")

        Dim sqlString As String = "SELECT " & _
                                    "COUNT(*)" & _
                                    " FROM BONAPPRO WHERE " & _
                                   " CMD_FRN_ID = ?"

        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        Dim bReturn As Boolean
        Dim nNbre As Integer

        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        'Si le paramètre existe on ne met à jour que sa valeur
        'Autrement on le crée
        CreateParameterP_ID(objCommand)

        Try

            objRS = objCommand.ExecuteReader
            If Not objRS.HasRows Then
                objRS.Close()
                Return False
            End If
            objRS.Read()
            nNbre = objRS.GetInt32(0)
            If nNbre > 0 Then
                bReturn = True
            Else
                bReturn = False
            End If
            objRS.Close()

        Catch ex As Exception
            If Not objRS Is Nothing Then
                objRS.Close()
            End If
            Throw ex

        End Try
        Return bReturn
    End Function 'existeBonApproFournisseur
    Protected Function existeSousCommandeCommande() As Boolean
        '==========================================================================
        'Function : existeSousCommandeCommande
        'Description : rend vrai s'il existe une Sous Commande Correspondant à la commande
        '===========================================================================
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "L'id doit être renseigné")
        Debug.Assert(Me.GetType.Name.Equals("CommandeClient"), "Object de Type CommandeClient Requis")

        Dim sqlString As String = "SELECT " & _
                                    " COUNT(*) " & _
                                    " FROM SOUSCOMMANDE WHERE " & _
                                   " SCMD_CMD_ID = ?"

        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        Dim bReturn As Boolean
        Dim nNbre As Integer

        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        'Si le paramètre existe on ne met à jour que sa valeur
        'Autrement on le crée
        CreateParameterP_ID(objCommand)

        Try

            objRS = objCommand.ExecuteReader
            If Not objRS.HasRows Then
                objRS.Close()
                Return False
            End If
            objRS.Read()
            nNbre = objRS.GetInt32(0)
            If nNbre > 0 Then
                bReturn = True
            Else
                bReturn = False
            End If
            objRS.Close()

        Catch ex As Exception
            If Not objRS Is Nothing Then
                objRS.Close()
            End If
            Throw ex
        End Try
        Return bReturn
    End Function 'existeSousCommandeCommande

    Protected Function loadCMDCLT() As Boolean

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "L'id doit être renseigné")

        Dim sqlString As String = "SELECT " & _
                                    "CMD_CLT_ID," & _
                                    "CMD_DATE," & _
                                    "CMD_DATE_VALID," & _
                                    "CMD_DATE_LIV," & _
                                    "CMD_DATE_ENLEV," & _
                                    "CMD_REF_LIV," & _
                                    "CMD_ETAT," & _
                                    "CMD_TOTAL_HT," & _
                                    "CMD_TOTAL_TTC," & _
                                    "CMD_TYPE," & _
                                    "CMD_TYPE_TRANSPORT," & _
                                    "CMD_CODE," & _
                                    "CMD_CLT_NOM," & _
                                    "CMD_CLT_RS," & _
                                    "CMD_CLT_LIV_NOM," & _
                                    "CMD_CLT_LIV_RUE1," & _
                                    "CMD_CLT_LIV_RUE2," & _
                                    "CMD_CLT_LIV_CP," & _
                                    "CMD_CLT_LIV_VILLE," & _
                                    "CMD_CLT_LIV_TEL," & _
                                    "CMD_CLT_LIV_FAX," & _
                                    "CMD_CLT_LIV_PORT," & _
                                    "CMD_CLT_LIV_EMAIL," & _
                                    "CMD_CLT_ADR_IDENT," & _
                                    "CMD_CLT_FACT_NOM," & _
                                    "CMD_CLT_FACT_RUE1," & _
                                    "CMD_CLT_FACT_RUE2," & _
                                    "CMD_CLT_FACT_CP," & _
                                    "CMD_CLT_FACT_VILLE," & _
                                    "CMD_CLT_FACT_TEL," & _
                                    "CMD_CLT_FACT_FAX," & _
                                    "CMD_CLT_FACT_PORT," & _
                                    "CMD_CLT_FACT_EMAIL," & _
                                    "CMD_CLT_RGLMT_ID," & _
                                    "CMD_CLT_BANQUE," & _
                                    "CMD_CLT_RIB1," & _
                                    "CMD_CLT_RIB2," & _
                                    "CMD_CLT_RIB3," & _
                                    "CMD_CLT_RIB4," & _
                                    "CMD_COM_LIBRE," & _
                                    "CMD_COM_COM," & _
                                    "CMD_COM_LIV," & _
                                    "CMD_COM_FACT," & _
                                    "CMD_TRP_ID," & _
                                    "CMD_TRP_CODE," & _
                                    "CMD_TRP_NOM," & _
                                    "CMD_TRP_RUE1," & _
                                    "CMD_TRP_RUE2," & _
                                    "CMD_TRP_CP," & _
                                    "CMD_TRP_VILLE," & _
                                    "CMD_TRP_TEL," & _
                                    "CMD_TRP_FAX," & _
                                    "CMD_TRP_PORT," & _
                                    "CMD_TRP_EMAIL, " & _
                                    "CMD_MT_TRANSPORT, " & _
                                    "CMD_QTE_COLIS, " & _
                                    "CMD_QTE_PAL_PREP, " & _
                                    "CMD_QTE_PAL_NONPREP, " & _
                                    "CMD_POIDS, " & _
                                    "CMD_PU_PAL_PREP, " & _
                                    "CMD_PU_PAL_NONPREP, " & _
                                    "CMD_BFACTTRP, " & _
                                    "CMD_IDFACTTRP, " & _
                                    "CMD_LETTREVOITURE, " & _
                                    "CMD_COUT_TRANSPORT, " & _
                                    "CMD_REFFACT_TRP, " & _
                                    "CMD_CLT_NOMLIVRAISON, " & _
                                    "CMD_CLT_RSLIVRAISON, " & _
                                    "CMD_IDPRESTASHOP, " & _
                                    "CMD_NAMEPRESTASHOP, " & _
                                    "CMD_ORIGINE, " & _
                                    "RQ_ModeReglement.PAR_VALUE" & _
                                    " FROM COMMANDE LEFT OUTER JOIN RQ_ModeReglement ON COMMANDE.CMD_CLT_RGLMT_ID = RQ_ModeReglement.PAR_ID WHERE " & _
                                   " CMD_ID = ?"

        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        Dim objCMDCLT As CommandeClient
        '        Dim objParam As OleDbParameter
        Dim bReturn As Boolean
        Dim nId As Integer

        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        'Si le paramètre existe on ne met à jour que sa valeur
        'Autrement on le crée
        CreateParameterP_ID(objCommand)

        Try
            objRS = objCommand.ExecuteReader
            If Not objRS.HasRows Then
                objRS.Close()
                Return False
            End If
            objRS.Read()
            objCMDCLT = CType(Me, CommandeClient)
            objCMDCLT.code = GetString(objRS, "CMD_CODE")
            objCMDCLT.setEtat(GetString(objRS, "CMD_ETAT"))
            objCMDCLT.caracteristiqueTiers.rib1 = GetString(objRS, "CMD_CLT_RIB1")
            objCMDCLT.caracteristiqueTiers.rib2 = GetString(objRS, "CMD_CLT_RIB2")
            objCMDCLT.caracteristiqueTiers.rib3 = GetString(objRS, "CMD_CLT_RIB3")
            objCMDCLT.caracteristiqueTiers.rib4 = GetString(objRS, "CMD_CLT_RIB4")
            objCMDCLT.caracteristiqueTiers.banque = GetString(objRS, "CMD_CLT_BANQUE")
            objCMDCLT.caracteristiqueTiers.idModeReglement = GetString(objRS, "CMD_CLT_RGLMT_ID")
            objCMDCLT.caracteristiqueTiers.libModeReglement = GetString(objRS, "PAR_VALUE")
            objCMDCLT.setTiers(New Client)
            nId = GetString(objRS, "CMD_CLT_ID")
            objCMDCLT.oTiers.setid(nId)
            objCMDCLT.dateCommande = GetString(objRS, "CMD_DATE")
            objCMDCLT.dateValidation = GetString(objRS, "CMD_DATE_VALID")
            objCMDCLT.dateLivraison = GetString(objRS, "CMD_DATE_LIV")
            objCMDCLT.dateEnlevement = GetString(objRS, "CMD_DATE_ENLEV")
            objCMDCLT.refLivraison = GetString(objRS, "CMD_REF_LIV")
            objCMDCLT.totalHT = GetString(objRS, "CMD_TOTAL_HT")
            objCMDCLT.totalTTC = GetString(objRS, "CMD_TOTAL_TTC")
            objCMDCLT.typeCommande = GetString(objRS, "CMD_TYPE")
            objCMDCLT.typeTransport = GetString(objRS, "CMD_TYPE_TRANSPORT")
            objCMDCLT.caracteristiqueTiers.nom = GetString(objRS, "CMD_CLT_NOM")
            objCMDCLT.caracteristiqueTiers.rs = GetString(objRS, "CMD_CLT_RS")
            objCMDCLT.caracteristiqueTiers.AdresseLivraison.nom = GetString(objRS, "CMD_CLT_LIV_NOM")
            objCMDCLT.caracteristiqueTiers.AdresseLivraison.rue1 = GetString(objRS, "CMD_CLT_LIV_RUE1")
            objCMDCLT.caracteristiqueTiers.AdresseLivraison.rue2 = GetString(objRS, "CMD_CLT_LIV_RUE2")
            objCMDCLT.caracteristiqueTiers.AdresseLivraison.cp = GetString(objRS, "CMD_CLT_LIV_CP")
            objCMDCLT.caracteristiqueTiers.AdresseLivraison.ville = GetString(objRS, "CMD_CLT_LIV_VILLE")
            objCMDCLT.caracteristiqueTiers.AdresseLivraison.tel = GetString(objRS, "CMD_CLT_LIV_TEL")
            objCMDCLT.caracteristiqueTiers.AdresseLivraison.fax = GetString(objRS, "CMD_CLT_LIV_FAX")
            objCMDCLT.caracteristiqueTiers.AdresseLivraison.port = GetString(objRS, "CMD_CLT_LIV_PORT")
            objCMDCLT.caracteristiqueTiers.AdresseLivraison.Email = GetString(objRS, "CMD_CLT_LIV_EMAIL")
            objCMDCLT.caracteristiqueTiers.bAdressesIdentiques = GetString(objRS, "CMD_CLT_ADR_IDENT")
            objCMDCLT.caracteristiqueTiers.AdresseFacturation.nom = GetString(objRS, "CMD_CLT_FACT_NOM")
            objCMDCLT.caracteristiqueTiers.AdresseFacturation.rue1 = GetString(objRS, "CMD_CLT_FACT_RUE1")
            objCMDCLT.caracteristiqueTiers.AdresseFacturation.rue2 = GetString(objRS, "CMD_CLT_FACT_RUE2")
            objCMDCLT.caracteristiqueTiers.AdresseFacturation.cp = GetString(objRS, "CMD_CLT_FACT_CP")
            objCMDCLT.caracteristiqueTiers.AdresseFacturation.ville = GetString(objRS, "CMD_CLT_FACT_VILLE")
            objCMDCLT.caracteristiqueTiers.AdresseFacturation.tel = GetString(objRS, "CMD_CLT_FACT_TEL")
            objCMDCLT.caracteristiqueTiers.AdresseFacturation.fax = GetString(objRS, "CMD_CLT_FACT_FAX")
            objCMDCLT.caracteristiqueTiers.AdresseFacturation.port = GetString(objRS, "CMD_CLT_FACT_PORT")
            objCMDCLT.caracteristiqueTiers.AdresseFacturation.Email = GetString(objRS, "CMD_CLT_FACT_EMAIL")
            objCMDCLT.oTransporteur = New Transporteur
            If (IsNumeric(GetString(objRS, "CMD_TRP_ID"))) Then
                objCMDCLT.oTransporteur.setid(GetString(objRS, "CMD_TRP_ID"))
                objCMDCLT.oTransporteur.code = GetString(objRS, "CMD_TRP_CODE")
                objCMDCLT.oTransporteur.nom = GetString(objRS, "CMD_TRP_NOM")
                objCMDCLT.oTransporteur.rs = GetString(objRS, "CMD_TRP_NOM")
                objCMDCLT.oTransporteur.AdresseLivraison.nom = GetString(objRS, "CMD_TRP_NOM")
                objCMDCLT.oTransporteur.AdresseLivraison.rue1 = GetString(objRS, "CMD_TRP_RUE1")
                objCMDCLT.oTransporteur.AdresseLivraison.rue2 = GetString(objRS, "CMD_TRP_RUE2")
                objCMDCLT.oTransporteur.AdresseLivraison.cp = GetString(objRS, "CMD_TRP_CP")
                objCMDCLT.oTransporteur.AdresseLivraison.ville = GetString(objRS, "CMD_TRP_VILLE")
                objCMDCLT.oTransporteur.AdresseLivraison.tel = GetString(objRS, "CMD_TRP_TEL")
                objCMDCLT.oTransporteur.AdresseLivraison.fax = GetString(objRS, "CMD_TRP_FAX")
                objCMDCLT.oTransporteur.AdresseLivraison.port = GetString(objRS, "CMD_TRP_PORT")
                objCMDCLT.oTransporteur.AdresseLivraison.Email = GetString(objRS, "CMD_TRP_EMAIL")
            End If
            objCMDCLT.CommLibre.comment = GetString(objRS, "CMD_COM_LIBRE")
            objCMDCLT.CommCommande.comment = GetString(objRS, "CMD_COM_COM")
            objCMDCLT.CommLivraison.comment = GetString(objRS, "CMD_COM_LIV")
            objCMDCLT.CommFacturation.comment = GetString(objRS, "CMD_COM_FACT")
            objCMDCLT.montantTransport = GetValue(objRS, "CMD_MT_TRANSPORT")
            objCMDCLT.qteColis = GetValue(objRS, "CMD_QTE_COLIS")
            objCMDCLT.qtePalettesNonPreparees = GetValue(objRS, "CMD_QTE_PAL_NONPREP")
            objCMDCLT.qtePalettesPreparees = GetValue(objRS, "CMD_QTE_PAL_PREP")
            objCMDCLT.poids = GetValue(objRS, "CMD_POIDS")
            objCMDCLT.puPalettesNonPreparees = GetValue(objRS, "CMD_PU_PAL_NONPREP")
            objCMDCLT.puPalettesPreparees = GetValue(objRS, "CMD_PU_PAL_PREP")
            objCMDCLT.bFactTransport = GetValue(objRS, "CMD_BFACTTRP")
            objCMDCLT.idFactTransport = GetValue(objRS, "CMD_IDFACTTRP")
            'Lettre Voiture
            objCMDCLT.lettreVoiture = GetString(objRS, "CMD_LETTREVOITURE")
            'CoutTranport
            objCMDCLT.coutTransport = GetValue(objRS, "CMD_COUT_TRANSPORT")
            'Ref Facture Transporteur
            objCMDCLT.refFactTRP = GetString(objRS, "CMD_REFFACT_TRP")
            'Nom et raison Solciale de livraison
            objCMDCLT.NomLivraison = GetString(objRS, "CMD_CLT_NOMLIVRAISON")
            objCMDCLT.RaisonSocialeLivraison = GetString(objRS, "CMD_CLT_RSLIVRAISON")
            'Prestashop
            objCMDCLT.IDPrestashop = getLong(objRS, "CMD_IDPRESTASHOP")
            objCMDCLT.NamePrestashop = GetString(objRS, "CMD_NAMEPRESTASHOP")
            objCMDCLT.Origine = Trim(GetString(objRS, "CMD_ORIGINE"))
            cleanErreur()
            objRS.Close()
            bReturn = True
        Catch ex As Exception
            setError("LoadCMDCLT", ex.ToString())
            If Not objRS Is Nothing Then
                objRS.Close()
            End If
            bReturn = False
        End Try
        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'LoadCMDCLT
    Protected Function deleteCMDCLT() As Boolean
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "Id <> 0")
        Debug.Assert(Me.GetType().Name.Equals("CommandeClient"))


        Dim sqlString As String = "DELETE FROM COMMANDE WHERE CMD_ID=? "
        Dim objCommand As OleDbCommand
        Dim objCMDCLT As CommandeClient
        '        Dim objParam As OleDbParameter

        objCMDCLT = CType(Me, CommandeClient)
        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        m_dbconn.BeginTransaction()
        objCommand.CommandText = sqlString
        objCommand.Transaction = m_dbconn.transaction


        CreateParameterP_ID(objCommand)
        Try
            objCommand.ExecuteNonQuery()
            m_id = 0
            objCMDCLT.resetCode()
            m_dbconn.transaction.Commit()
            Return True

        Catch ex As Exception
            setError("DeleteCMDCLT", ex.ToString())
            m_dbconn.transaction.Rollback()
            Return False
        End Try
    End Function 'DeleteCMDCLT
    Protected Function deleteCMDCLT_SCMD() As Boolean
        '=====================================================================================
        ' Finction deleteCMDCLT_SCMD
        ' Description : Suppression des Sous Commandes d'une Commande
        '   ATTENTION : Les Sous Commande qui ont été affectée à une facture de Commission ne sont pas Supprimées
        '=====================================================================================
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "Id <> 0")
        Debug.Assert(Me.GetType().Name.Equals("CommandeClient"))


        Dim sqlString As String = "DELETE FROM SOUSCOMMANDE WHERE SCMD_CMD_ID=? " & _
                                    " AND SCMD_FACT_ID is NULL"
        Dim sqlString2 As String = "UPDATE LIGNE_COMMANDE SET LGCM_SCMD_ID = NULL WHERE LGCM_CMD_ID=? "
        Dim objCommand As OleDbCommand
        Dim objCMDCLT As CommandeClient
        '        Dim objParam As OleDbParameter

        objCMDCLT = CType(Me, CommandeClient)
        m_dbconn.BeginTransaction()

        Try
            objCommand = New OleDbCommand
            objCommand.Connection = m_dbconn.Connection
            objCommand.Transaction = m_dbconn.transaction
            objCommand.CommandText = sqlString2
            CreateParameterP_ID(objCommand)
            objCommand.ExecuteNonQuery()

            objCommand = New OleDbCommand
            objCommand.Connection = m_dbconn.Connection
            objCommand.Transaction = m_dbconn.transaction
            objCommand.CommandText = sqlString
            CreateParameterP_ID(objCommand)
            objCommand.ExecuteNonQuery()
            m_dbconn.transaction.Commit()
            Return True

        Catch ex As Exception
            m_dbconn.transaction.Rollback()
            setError("DeleteCMDCLT_SCMD", ex.ToString())
            Return False
        End Try
    End Function 'DeleteCMDCLT_SCMD
    '=======================================================================
    'Fonction : DeleteCMDCLT_MVTSTK
    'Description : Suppression des mouvements de stocks d'une commande
    'Retour : Rend Vrai si le DELETE s'est correctement effectué
    '=======================================================================
    Protected Function deleteCMDCLT_MVTSTK() As Boolean
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "Id <> 0")


        Dim sqlString As String = "DELETE FROM MVT_STOCK WHERE STK_TYPE=2 and STK_REF_ID=? "
        Dim objCommand As OleDbCommand
        '        Dim objParam As OleDbParameter

        objCommand = New OleDbCommand
        m_dbconn.BeginTransaction()
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString
        objCommand.Transaction = m_dbconn.transaction



        CreateParameterP_ID(objCommand)
        Try
            objCommand.ExecuteNonQuery()
            m_dbconn.transaction.Commit()
            Return True

        Catch ex As Exception
            setError("DeleteCMDCLT_MVTSTK", ex.ToString())
            m_dbconn.transaction.Rollback()
            Return False
        End Try
    End Function 'DeleteCMDCLT_MVTSTK
#End Region

#Region "CreateParameter CommandeClient"
    Private Sub CreateParameterP_CMD_CLT_ID(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        Debug.Assert(objCMD.oTiers.id <> 0)
        objOLeDBCommand.Parameters.AddWithValue("?", objCMD.oTiers.id)
    End Sub
    Private Sub CreateParameterP_CMD_DATE(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", objCMD.dateCommande.ToShortDateString)
    End Sub
    Private Sub CreateParameterP_CMD_DATE_VALID(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", objCMD.dateValidation.ToShortDateString)
    End Sub
    Private Sub CreateParameterP_CMD_DATE_LIV(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", objCMD.dateLivraison.ToShortDateString)
    End Sub
    Private Sub CreateParameterP_CMD_DATE_ENLEV(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", objCMD.dateEnlevement.ToShortDateString)
    End Sub
    Private Sub CreateParameterP_CMD_REF_LIV(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", truncate(objCMD.refLivraison, 50))
    End Sub

    Private Sub CreateParameterP_CMD_ETAT(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", objCMD.etat.codeEtat)
    End Sub
    Private Sub CreateParameterP_CMD_TOTAL_HT(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", CDbl(objCMD.totalHT))
    End Sub
    Private Sub CreateParameterP_CMD_TOTAL_TTC(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", CDbl(objCMD.totalTTC))
    End Sub
    Private Sub CreateParameterP_CMD_TYPE(ByVal objOLeDBCommand As OleDbCommand)
        If m_typedonnee = vncEnums.vncTypeDonnee.COMMANDECLIENT Then
            Dim objCMD As CommandeClient
            objCMD = Me
            objOLeDBCommand.Parameters.AddWithValue("?", objCMD.typeCommande)
        Else
            objOLeDBCommand.Parameters.AddWithValue("?", vbNull)
        End If

    End Sub
    Private Sub CreateParameterP_CMD_TYPE_TRANSPORT(ByVal objOLeDBCommand As OleDbCommand)
        If m_typedonnee = vncEnums.vncTypeDonnee.COMMANDECLIENT Then
            Dim objCMD As CommandeClient
            objCMD = Me
            objOLeDBCommand.Parameters.AddWithValue("?", objCMD.typeTransport)
        Else
            objOLeDBCommand.Parameters.AddWithValue("?", vbNull)
        End If

    End Sub
    Private Sub CreateParameterP_CMD_CODE(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", truncate(objCMD.code, 50))
    End Sub
    Private Sub CreateParameterP_CMD_CLT_NOM(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", truncate(objCMD.caracteristiqueTiers.nom, 50))
    End Sub
    Private Sub CreateParameterP_CMD_CLT_RS(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", truncate(objCMD.caracteristiqueTiers.rs, 50))
    End Sub
    Private Sub CreateParameterP_CMD_CLT_LIV_NOM(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", truncate(objCMD.caracteristiqueTiers.AdresseLivraison.nom, 50))
    End Sub
    Private Sub CreateParameterP_CMD_CLT_LIV_RUE1(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", truncate(objCMD.caracteristiqueTiers.AdresseLivraison.rue1, 50))

    End Sub
    Private Sub CreateParameterP_CMD_CLT_LIV_RUE2(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", truncate(objCMD.caracteristiqueTiers.AdresseLivraison.rue2, 50))
    End Sub
    Private Sub CreateParameterP_CMD_CLT_LIV_CP(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", truncate(objCMD.caracteristiqueTiers.AdresseLivraison.cp, 50))
    End Sub
    Private Sub CreateParameterP_CMD_CLT_LIV_VILLE(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", truncate(objCMD.caracteristiqueTiers.AdresseLivraison.ville, 50))
    End Sub
    Private Sub CreateParameterP_CMD_CLT_LIV_TEL(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", truncate(objCMD.caracteristiqueTiers.AdresseLivraison.tel, 50))

    End Sub
    Private Sub CreateParameterP_CMD_CLT_LIV_FAX(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", truncate(objCMD.caracteristiqueTiers.AdresseLivraison.fax, 50))

    End Sub
    Private Sub CreateParameterP_CMD_CLT_LIV_PORT(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", truncate(objCMD.caracteristiqueTiers.AdresseLivraison.port, 50))

    End Sub
    Private Sub CreateParameterP_CMD_CLT_LIV_EMAIL(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", truncate(objCMD.caracteristiqueTiers.AdresseLivraison.Email, 50))

    End Sub
    Private Sub CreateParameterP_CMD_CLT_ADR_IDENT(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", objCMD.caracteristiqueTiers.bAdressesIdentiques)
    End Sub
    Private Sub CreateParameterP_CMD_CLT_FACT_NOM(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", truncate(objCMD.caracteristiqueTiers.AdresseFacturation.nom, 50))
    End Sub
    Private Sub CreateParameterP_CMD_CLT_FACT_RUE1(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", truncate(objCMD.caracteristiqueTiers.AdresseFacturation.rue1, 50))

    End Sub
    Private Sub CreateParameterP_CMD_CLT_FACT_RUE2(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", truncate(objCMD.caracteristiqueTiers.AdresseFacturation.rue2, 50))
    End Sub
    Private Sub CreateParameterP_CMD_CLT_FACT_CP(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", truncate(objCMD.caracteristiqueTiers.AdresseFacturation.cp, 50))
    End Sub
    Private Sub CreateParameterP_CMD_CLT_FACT_VILLE(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", truncate(objCMD.caracteristiqueTiers.AdresseFacturation.ville, 50))
    End Sub
    Private Sub CreateParameterP_CMD_CLT_FACT_TEL(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", truncate(objCMD.caracteristiqueTiers.AdresseFacturation.tel, 50))

    End Sub
    Private Sub CreateParameterP_CMD_CLT_FACT_FAX(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", truncate(objCMD.caracteristiqueTiers.AdresseFacturation.fax, 50))

    End Sub
    Private Sub CreateParameterP_CMD_CLT_FACT_PORT(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", truncate(objCMD.caracteristiqueTiers.AdresseFacturation.port, 50))

    End Sub
    Private Sub CreateParameterP_CMD_CLT_FACT_EMAIL(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", truncate(objCMD.caracteristiqueTiers.AdresseFacturation.Email, 50))
    End Sub
    Private Sub CreateParameterP_CMD_CLT_RGLMT_ID(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", objCMD.caracteristiqueTiers.idModeReglement)
    End Sub
    Private Sub CreateParameterP_CMD_CLT_BANQUE(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", truncate(objCMD.caracteristiqueTiers.banque, 50))

    End Sub
    Private Sub CreateParameterP_CMD_CLT_RIB1(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", truncate(objCMD.caracteristiqueTiers.rib1, 5))
    End Sub
    Private Sub CreateParameterP_CMD_CLT_RIB2(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", truncate(objCMD.caracteristiqueTiers.rib2, 5))

    End Sub
    Private Sub CreateParameterP_CMD_CLT_RIB3(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", truncate(objCMD.caracteristiqueTiers.rib3, 11))

    End Sub
    Private Sub CreateParameterP_CMD_CLT_RIB4(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", truncate(objCMD.caracteristiqueTiers.rib4, 2))

    End Sub
    Private Sub CreateParameterP_CMD_COM_LIBRE(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", objCMD.CommLibre.comment)

    End Sub
    Private Sub CreateParameterP_CMD_COM_COM(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", objCMD.CommCommande.comment)

    End Sub
    Private Sub CreateParameterP_CMD_COM_LIV(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", objCMD.CommLivraison.comment)

    End Sub
    Private Sub CreateParameterP_CMD_COM_FACT(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", objCMD.CommFacturation.comment)

    End Sub
    Private Sub CreateParameterP_CMD_TRP_ID(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", objCMD.oTransporteur.id)
    End Sub
    Private Sub CreateParameterP_CMD_TRP_CODE(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", objCMD.oTransporteur.code)
    End Sub
    Private Sub CreateParameterP_CMD_TRP_NOM(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", truncate(objCMD.oTransporteur.nom, 50))
    End Sub
    Private Sub CreateParameterP_CMD_TRP_RUE1(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", truncate(objCMD.oTransporteur.AdresseLivraison.rue1, 50))
    End Sub
    Private Sub CreateParameterP_CMD_TRP_RUE2(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", truncate(objCMD.oTransporteur.AdresseLivraison.rue2, 50))
    End Sub
    Private Sub CreateParameterP_CMD_TRP_CP(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", truncate(objCMD.oTransporteur.AdresseLivraison.cp, 50))
    End Sub
    Private Sub CreateParameterP_CMD_TRP_VILLE(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", truncate(objCMD.oTransporteur.AdresseLivraison.ville, 50))
    End Sub
    Private Sub CreateParameterP_CMD_TRP_TEL(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", truncate(objCMD.oTransporteur.AdresseLivraison.tel, 50))

    End Sub
    Private Sub CreateParameterP_CMD_TRP_FAX(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", truncate(objCMD.oTransporteur.AdresseLivraison.fax, 50))

    End Sub
    Private Sub CreateParameterP_CMD_TRP_PORT(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", truncate(objCMD.oTransporteur.AdresseLivraison.port, 50))

    End Sub
    Private Sub CreateParameterP_CMD_TRP_EMAIL(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", truncate(objCMD.oTransporteur.AdresseLivraison.Email, 50))
    End Sub
    Private Sub CreateParameterP_CMD_MT_TRANSPORT(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", CDbl(objCMD.montantTransport))
    End Sub
    Private Sub CreateParameterP_CMD_COUT_TRANSPORT(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", CDbl(objCMD.coutTransport))
    End Sub
    Private Sub CreateParameterP_CMD_LETTREVOITURE(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", truncate(objCMD.lettreVoiture, 50))
    End Sub
    Private Sub CreateParameterP_CMD_REFFACT_TRP(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", truncate(objCMD.refFactTRP, 50))
    End Sub
    Private Sub CreateParameterP_CMD_QTE_COLIS(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", CDbl(objCMD.qteColis))
    End Sub
    Private Sub CreateParameterP_CMD_QTE_PAL_PREP(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", CDbl(objCMD.qtePalettesPreparees))
    End Sub
    Private Sub CreateParameterP_CMD_QTE_PAL_NONPREP(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", CDbl(objCMD.qtePalettesNonPreparees))
    End Sub
    Private Sub CreateParameterP_CMD_PU_PAL_PREP(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", CDbl(objCMD.puPalettesPreparees))
    End Sub
    Private Sub CreateParameterP_CMD_PU_PAL_NONPREP(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", CDbl(objCMD.puPalettesNonPreparees))
    End Sub
    Private Sub CreateParameterP_CMD_POIDS(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", CDbl(objCMD.poids))
    End Sub
    Private Sub CreateParameterP_CMD_BFACTTRP(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As CommandeClient
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", objCMD.bFactTransport)
    End Sub
    Private Sub CreateParameterP_CMD_IDFACTTRP(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As CommandeClient
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", objCMD.idFactTransport)
    End Sub
    Private Sub CreateParameterP_CMD_CLT_NOMLIVRAISON(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As CommandeClient
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", truncate(objCMD.NomLivraison, 50))
    End Sub
    Private Sub CreateParameterP_CMD_CLT_RSLIVRAISON(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As CommandeClient
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", truncate(objCMD.RaisonSocialeLivraison, 50))
    End Sub
    Private Sub CreateParameterP_CMD_IDPRESTASHOP(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As CommandeClient
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", objCMD.IDPrestashop)
    End Sub
    Private Sub CreateParameterP_CMD_NAMEPRESTASHOP(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As CommandeClient
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", truncate(objCMD.NamePrestashop, 50))
    End Sub
    Private Sub CreateParameterP_CMD_Origine(ByVal objOLeDBCommand As OleDbCommand)
        Dim objCMD As CommandeClient
        objCMD = Me
        objOLeDBCommand.Parameters.AddWithValue("?", truncate(objCMD.Origine, 50))
    End Sub
#End Region
#Region "Fonction Database SousCommande"
    Protected Function loadSCMD(Optional ByVal pstrCode As String = "NULL") As Boolean

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")

        Dim sqlString As String = "SELECT " &
                                    "SCMD_ID," &
                                    "SCMD_CODE," &
                                    "SCMD_ETAT," &
                                    "SCMD_CMD_ID," &
                                    "SCMD_FACT_ID," &
                                    "SCMD_FRN_ID," &
                                    "SCMD_DATE," &
                                    "SCMD_TOTAL_HT," &
                                    "SCMD_TOTAL_TTC," &
                                    "SCMD_DATE_LIV," &
                                    "SCMD_TRP_NOM," &
                                    "SCMD_TRP_RUE1," &
                                    "SCMD_TRP_RUE2," &
                                    "SCMD_TRP_CP," &
                                    "SCMD_TRP_VILLE," &
                                    "SCMD_TRP_TEL," &
                                    "SCMD_TRP_FAX," &
                                    "sCMD_TRP_PORT," &
                                    "sCMD_TRP_EMAIL," &
                                    "SCMD_COM_COM," &
                                    "SCMD_COM_LIV," &
                                    "SCMD_COM_FACT," &
                                    "sCMD_COM_LIBRE," &
                                    "sCMD_REF_LIV," &
                                    "SCMD_FACT_REF," &
                                    "SCMD_FACT_DATE," &
                                    "SCMD_FACT_TOTAL_HT," &
                                    "SCMD_FACT_TOTAL_TTC," &
                                    "SCMD_COM_BASE," &
                                    "SCMD_COM_TAUX," &
                                    "SCMD_COM_MONTANT," &
                                    "SCMD_CLT_ID, " &
                                    "SCMD_BEXPORT, " &
                                    "SCMD_BEXPORTQUADRA, " &
                                    "CMD_CODE " &
                                    " FROM SOUSCOMMANDE , COMMANDE "
        Dim strWhere As String = ""
        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        Dim objSCMD As SousCommande
        '        Dim objParam As OleDbParameter
        Dim bReturn As Boolean



        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection

        strWhere = "WHERE SCMD_CMD_ID = CMD_ID AND "
        If pstrCode.Equals("NULL") Then
            strWhere = strWhere + " SCMD_ID = " & m_id & " "
        Else
            strWhere = strWhere + " SCMD_CODE LIKE '" & pstrCode & "' "
        End If

        objCommand.CommandText = sqlString + strWhere
        Try
            objRS = objCommand.ExecuteReader
            If Not objRS.HasRows Then
                objRS.Close()
                Return False
            End If
            objRS.Read()
            objSCMD = CType(Me, SousCommande)
            objSCMD.setid(GetValue(objRS, "SCMD_ID"))
            objSCMD.code = GetString(objRS, "SCMD_CODE")
            objSCMD.setEtat(GetString(objRS, "SCMD_ETAT"))
            objSCMD.idCommandeClient = GetString(objRS, "SCMD_CMD_ID")
            objSCMD.idFactCom = getInteger(objRS, "SCMD_FACT_ID")
            objSCMD.oTiers.setid(GetString(objRS, "SCMD_CLT_ID"))
            objSCMD.oFournisseur.setid(GetString(objRS, "SCMD_FRN_ID"))
            objSCMD.dateCommande = GetString(objRS, "SCMD_DATE")
            objSCMD.totalHT = GetString(objRS, "SCMD_TOTAL_HT")
            objSCMD.totalTTC = GetString(objRS, "SCMD_TOTAL_TTC")
            objSCMD.dateLivraison = GetString(objRS, "SCMD_DATE_LIV")
            objSCMD.oTransporteur.AdresseLivraison.nom = GetString(objRS, "SCMD_TRP_NOM")
            objSCMD.oTransporteur.AdresseLivraison.rue1 = GetString(objRS, "SCMD_TRP_RUE1")
            objSCMD.oTransporteur.AdresseLivraison.rue2 = GetString(objRS, "SCMD_TRP_RUE2")
            objSCMD.oTransporteur.AdresseLivraison.cp = GetString(objRS, "SCMD_TRP_CP")
            objSCMD.oTransporteur.AdresseLivraison.ville = GetString(objRS, "SCMD_TRP_VILLE")
            objSCMD.oTransporteur.AdresseLivraison.tel = GetString(objRS, "SCMD_TRP_TEL")
            objSCMD.oTransporteur.AdresseLivraison.fax = GetString(objRS, "SCMD_TRP_FAX")
            objSCMD.oTransporteur.AdresseLivraison.port = GetString(objRS, "SCMD_TRP_PORT")
            objSCMD.oTransporteur.AdresseLivraison.Email = GetString(objRS, "SCMD_TRP_EMAIL")
            objSCMD.CommCommande.comment = (GetString(objRS, "SCMD_COM_COM"))
            objSCMD.CommLivraison.comment = GetString(objRS, "SCMD_COM_LIV")
            objSCMD.CommFacturation.comment = GetString(objRS, "SCMD_COM_FACT")
            objSCMD.CommLibre.comment = GetString(objRS, "SCMD_COM_LIBRE")
            objSCMD.refLivraison = GetString(objRS, "SCMD_REF_LIV")
            objSCMD.refFactFournisseur = GetString(objRS, "SCMD_FACT_REF")
            objSCMD.dateFactFournisseur = GetString(objRS, "SCMD_FACT_DATE")
            objSCMD.totalHTFacture = GetString(objRS, "SCMD_FACT_TOTAL_HT")
            objSCMD.totalTTCFacture = GetString(objRS, "SCMD_FACT_TOTAL_TTC")
            objSCMD.baseCommission = GetString(objRS, "SCMD_COM_BASE")
            objSCMD.tauxCommission = GetString(objRS, "SCMD_COM_TAUX")
            objSCMD.MontantCommission = GetString(objRS, "SCMD_COM_MONTANT")
            objSCMD.codeCommandeClient = GetString(objRS, "CMD_CODE")
            objSCMD.bExportInternet = GetValue(objRS, "SCMD_BEXPORT")
            If Not objRS.IsDBNull(objRS.GetOrdinal("SCMD_BEXPORTQUADRA")) Then
                objSCMD.bExportQuadra = GetValue(objRS, "SCMD_BEXPORTQUADRA")
            End If
            cleanErreur()
            bReturn = True
        Catch ex As Exception
            setError("LoadSCMD", ex.ToString())
            bReturn = False
        End Try
        If Not objRS Is Nothing Then
            objRS.Close()
        End If
        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'LoadSCMD
    Protected Function getNumeroSousCommandeClient() As Integer
        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        Dim nReturn As Integer = -1
        Dim bReturn As Boolean
        Dim strStringLect As String

        Try
            strStringLect = "DECLARE @return_value int;EXEC	@return_value = GETNEXTNUM_SCMD; SELECT 'CODE' = @return_value"
            objCommand = New OleDbCommand
            objCommand.Connection = m_dbconn.Connection
            objCommand = New OleDbCommand(strStringLect, m_dbconn.Connection)
            objRS = objCommand.ExecuteReader
            If objRS.HasRows Then
                objRS.Read()
                nReturn = objRS.GetInt32(0)
                objRS.Close()
                objRS = Nothing
                bReturn = True
            Else
                objRS.Close()
                bReturn = False
            End If
        Catch ex As Exception
            setError("Persist.getNumeroSousCommandeClient", ex.ToString())
            nReturn = -1
            bReturn = False
        End Try
        Debug.Assert(bReturn, "getNumeroSousCommandeClient: " & getErreur())
        Return nReturn
    End Function 'GetNumeroSousCommandeClient
    Protected Function GetNumeroFactureCommission() As Integer
        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        Dim nReturn As Integer = -1
        Dim bReturn As Boolean
        Dim strStringLect As String

        Try
            strStringLect = "DECLARE @return_value int;EXEC	@return_value = GETNEXTNUM_FACTCOM; SELECT 'CODE' = @return_value"
            objCommand = New OleDbCommand
            objCommand.Connection = m_dbconn.Connection
            objCommand = New OleDbCommand(strStringLect, m_dbconn.Connection)
            objRS = objCommand.ExecuteReader
            If objRS.HasRows Then
                objRS.Read()
                nReturn = objRS.GetInt32(0)
                objRS.Close()
                objRS = Nothing
                bReturn = True
            Else
                objRS.Close()
                bReturn = False
            End If
        Catch ex As Exception
            setError("Persist.GetNumeroFactureCommission", ex.ToString())
            nReturn = -1
            bReturn = False
        End Try
        Debug.Assert(bReturn, "GetNumeroFactureCommission: " & getErreur())
        Return nReturn
    End Function 'GetNumeroFactCommission
    Protected Function GetNumeroFactureTransport() As Integer
        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        Dim nReturn As Integer = -1
        Dim bReturn As Boolean
        Dim strStringLect As String

        Try
            strStringLect = "DECLARE @return_value int;EXEC	@return_value = GETNEXTNUM_FACTTRP; SELECT 'CODE' = @return_value"
            objCommand = New OleDbCommand
            objCommand.Connection = m_dbconn.Connection
            objCommand = New OleDbCommand(strStringLect, m_dbconn.Connection)
            objRS = objCommand.ExecuteReader
            If objRS.HasRows Then
                objRS.Read()
                nReturn = objRS.GetInt32(0)
                objRS.Close()
                objRS = Nothing
                bReturn = True
            Else
                objRS.Close()
                bReturn = False
            End If
        Catch ex As Exception
            setError("Persist.GetNumeroFactureTransport", ex.ToString())
            nReturn = -1
            bReturn = False
        End Try
        Debug.Assert(bReturn, "GetNumeroFactureTransport: " & getErreur())
        Return nReturn
    End Function 'GetNumeroFactureTransport
    ''' <summary>
    ''' Rend le prochain numéro de facture HOBIVIN
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Function GetNumeroFactureHBV() As Integer
        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        Dim nReturn As Integer = -1
        Dim bReturn As Boolean
        Dim strStringLect As String

        Try
            strStringLect = "DECLARE @return_value int;EXEC	@return_value = GETNEXTNUM_FACTHBV; SELECT 'CODE' = @return_value"
            objCommand = New OleDbCommand
            objCommand.Connection = m_dbconn.Connection
            objCommand = New OleDbCommand(strStringLect, m_dbconn.Connection)
            objRS = objCommand.ExecuteReader
            If objRS.HasRows Then
                objRS.Read()
                nReturn = objRS.GetInt32(0)
                objRS.Close()
                objRS = Nothing
                bReturn = True
            Else
                objRS.Close()
                bReturn = False
            End If
        Catch ex As Exception
            setError("Persist.GetNumeroFactureHBV", ex.ToString())
            nReturn = -1
            bReturn = False
        End Try
        Debug.Assert(bReturn, "GetNumeroFactureHBV: " & getErreur())
        Return nReturn
    End Function 'GetNumeroFactureHBV
    Protected Function GetNumeroFactureColisage() As Integer
        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        Dim nReturn As Integer = -1
        Dim bReturn As Boolean
        Dim strStringLect As String

        Try
            strStringLect = "DECLARE @return_value int;EXEC	@return_value = GETNEXTNUM_FACTCOL; SELECT 'CODE' = @return_value"
            objCommand = New OleDbCommand
            objCommand.Connection = m_dbconn.Connection
            objCommand = New OleDbCommand(strStringLect, m_dbconn.Connection)
            objRS = objCommand.ExecuteReader
            If objRS.HasRows Then
                objRS.Read()
                nReturn = objRS.GetInt32(0)
                objRS.Close()
                objRS = Nothing
                bReturn = True
            Else
                objRS.Close()
                bReturn = False
            End If
        Catch ex As Exception
            setError("Persist.GetNumeroFactureColisage", ex.ToString())
            nReturn = -1
            bReturn = False
        End Try
        Debug.Assert(bReturn, "GetNumeroFactureColisage: " & getErreur())
        Return nReturn
    End Function 'GetNumeroFactColisageisage
    Protected Function insertSCMD() As Boolean
        '=======================================================================
        'Fonction : InsertSCMD
        'Description : Insertion d'une Sous Commande Client 

        'Retour : Rend Vrai si l'INSERT s'est correctement effectué
        '=======================================================================
        Dim bReturn As Boolean
        Dim objCMDCLT As SousCommande
        bReturn = False

        Debug.Assert(Me.GetType().Name.Equals("SousCommande"), "Objet de Type 'S²ousCommande' Requis")
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id = 0, "ID=0")
        objCMDCLT = CType(Me, SousCommande)

        Dim sqlString As String = "INSERT INTO SOUSCOMMANDE( " &
                                    "SCMD_CODE," &
                                    "SCMD_ETAT," &
                                    "SCMD_CMD_ID," &
                                    "SCMD_FACT_ID," &
                                    "SCMD_FRN_ID," &
                                    "SCMD_DATE," &
                                    "SCMD_TOTAL_HT," &
                                    "SCMD_TOTAL_TTC," &
                                    "SCMD_DATE_LIV," &
                                    "SCMD_COM_COM," &
                                    "SCMD_COM_LIV," &
                                    "SCMD_COM_FACT," &
                                    "SCMD_COM_LIBRE," &
                                    "SCMD_TRP_NOM," &
                                    "SCMD_TRP_RUE1," &
                                    "SCMD_TRP_RUE2," &
                                    "SCMD_TRP_CP," &
                                    "SCMD_TRP_VILLE," &
                                    "SCMD_TRP_TEL," &
                                    "SCMD_TRP_FAX," &
                                    "sCMD_TRP_PORT," &
                                    "sCMD_TRP_EMAIL," &
                                    "sCMD_REF_LIV," &
                                    "SCMD_FACT_REF," &
                                    "SCMD_FACT_DATE," &
                                    "SCMD_FACT_TOTAL_HT," &
                                    "SCMD_FACT_TOTAL_TTC," &
                                    "SCMD_COM_BASE," &
                                    "SCMD_COM_TAUX," &
                                    "SCMD_COM_MONTANT," &
                                    "SCMD_CLT_ID," &
                                    "SCMD_BEXPORT," &
                                    "SCMD_BEXPORTQUADRA" &
                                  " ) VALUES ( " &
                                    "? ," &
                                    "? ," &
                                    "? ," &
                                    "? ," &
                                    "? ," &
                                    "? ," &
                                    "? ," &
                                    "? ," &
                                    "? ," &
                                    "? ," &
                                    "? ," &
                                    "? ," &
                                    "? ," &
                                    "? ," &
                                    "? ," &
                                    "? ," &
                                    "? ," &
                                    "? ," &
                                    "? ," &
                                    "? ," &
                                    "? ," &
                                    "? ," &
                                    "? ," &
                                    "? ," &
                                    "? ," &
                                    "? ," &
                                    "? ," &
                                    "? ," &
                                    "? ," &
                                    "? ," &
                                    "? ," &
                                    "? ," &
                                    "?" &
                                    " )"
        Dim objCommand As OleDbCommand
        Dim objCommand2 As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        '        Dim objParam As OleDbParameter

        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString
        m_dbconn.BeginTransaction()
        objCommand.Transaction = m_dbconn.transaction

        CreateParameterP_sCMD_CODE(objCommand)
        CreateParameterP_SCMD_ETAT(objCommand)
        CreateParameterP_SCMD_CMD_ID(objCommand)
        CreateParameterP_SCMD_FACT_ID(objCommand)
        CreateParameterP_SCMD_FRN_ID(objCommand)
        CreateParameterP_SCMD_DATE(objCommand)
        CreateParameterP_SCMD_TOTAL_HT(objCommand)
        CreateParameterP_SCMD_TOTAL_TTC(objCommand)
        CreateParameterP_SCMD_DATE_LIV(objCommand)
        CreateParameterP_SCMD_COM_COM(objCommand)
        CreateParameterP_SCMD_COM_LIV(objCommand)
        CreateParameterP_SCMD_COM_FACT(objCommand)
        CreateParameterP_SCMD_COM_LIBRE(objCommand)
        CreateParameterP_SCMD_TRP_NOM(objCommand)
        CreateParameterP_SCMD_TRP_RUE1(objCommand)
        CreateParameterP_SCMD_TRP_RUE2(objCommand)
        CreateParameterP_SCMD_TRP_CP(objCommand)
        CreateParameterP_SCMD_TRP_VILLE(objCommand)
        CreateParameterP_SCMD_TRP_TEL(objCommand)
        CreateParameterP_SCMD_TRP_FAX(objCommand)
        CreateParameterP_SCMD_TRP_PORT(objCommand)
        CreateParameterP_SCMD_TRP_EMAIL(objCommand)
        CreateParameterP_SCMD_REF_LIV(objCommand)
        CreateParameterP_SCMD_FACT_REF(objCommand)
        CreateParameterP_SCMD_FACT_DATE(objCommand)
        CreateParameterP_SCMD_FACT_TOTAL_HT(objCommand)
        CreateParameterP_SCMD_FACT_TOTAL_TTC(objCommand)
        CreateParameterP_SCMD_COM_BASE(objCommand)
        CreateParameterP_SCMD_COM_TAUX(objCommand)
        CreateParameterP_SCMD_COM_MONTANT(objCommand)
        CreateParameterP_SCMD_CLT_ID(objCommand)
        CreateParameterP_SCMD_BEXPORT(objCommand)
        CreateParameterP_SCMD_BEXPORTQUADRA(objCommand)
        Try
            objCommand.ExecuteNonQuery()
            objCommand2 = New OleDbCommand("SELECT MAX(SCMD_ID) FROM SOUSCOMMANDE", m_dbconn.Connection)
            objCommand2.Transaction = m_dbconn.transaction
            objRS = objCommand2.ExecuteReader
            If objRS.Read() Then
                m_id = objRS.GetInt32(0)
                cleanErreur()
                bReturn = True
            End If

        Catch ex As Exception
            setError("InsertSCMD", ex.ToString())
            bReturn = False
        End Try
        If Not objRS Is Nothing Then
            objRS.Close()
        End If

        '    Debug.Assert(m_id <> 0, "ID=0")
        Debug.Assert(bReturn, "InsertSCMD: " & getErreur())
        If bReturn Then
            m_dbconn.transaction.Commit()
        Else
            m_dbconn.transaction.Rollback()
        End If
        Return bReturn
    End Function 'insertSCMD
    Protected Function updateSCMD() As Boolean
        '=======================================================================
        'Fonction : UpdateSCMD
        'Description : Mise à jour  d'une sous-Commande 

        'Retour : Rend Vrai si l'UPDATE s'est correctement effectué
        '=======================================================================
        Dim bReturn As Boolean
        Dim objSCMD As SousCommande
        bReturn = False

        Debug.Assert(Me.GetType().Name.Equals("SousCommande"), "Objet de Type SousCommande Requis")
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "ID<>0")
        objSCMD = CType(Me, SousCommande)

        Dim sqlString As String = "UPDATE SOUSCOMMANDE SET " &
                                    "SCMD_CODE = ? ," &
                                    "SCMD_ETAT = ? ," &
                                    "SCMD_CMD_ID = ? ," &
                                    "SCMD_FRN_ID = ? ," &
                                    "SCMD_FACT_ID = ? ," &
                                    "SCMD_DATE = ? ," &
                                    "SCMD_TOTAL_HT = ? ," &
                                    "SCMD_TOTAL_TTC= ? ," &
                                    "SCMD_DATE_LIV= ? ," &
                                    "SCMD_COM_COM = ? ," &
                                    "SCMD_COM_LIV = ? ," &
                                    "SCMD_COM_FACT = ? ," &
                                    "SCMD_COM_LIBRE = ? ," &
                                    "SCMD_TRP_NOM = ? ," &
                                    "SCMD_TRP_RUE1 = ? ," &
                                    "SCMD_TRP_RUE2 = ? ," &
                                    "SCMD_TRP_CP = ? ," &
                                    "SCMD_TRP_VILLE = ? ," &
                                    "SCMD_TRP_TEL = ? ," &
                                    "SCMD_TRP_FAX = ? ," &
                                    "sCMD_TRP_PORT = ? ," &
                                    "sCMD_TRP_EMAIL = ? ," &
                                    "sCMD_REF_LIV = ? ," &
                                    "SCMD_FACT_REF = ? ," &
                                    "SCMD_FACT_DATE = ? ," &
                                    "SCMD_FACT_TOTAL_HT = ? ," &
                                    "SCMD_FACT_TOTAL_TTC = ? ," &
                                    "SCMD_COM_BASE = ? ," &
                                    "SCMD_COM_TAUX = ? ," &
                                    "SCMD_COM_MONTANT = ? ," &
                                    "SCMD_CLT_ID= ? , " &
                                    "SCMD_BEXPORT= ? ," &
                                    "SCMD_BEXPORTQUADRA= ? " &
                                    " WHERE SCMD_ID = ?"

        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        '        Dim objParam As OleDbParameter


        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString
        m_dbconn.BeginTransaction()
        objCommand.Transaction = m_dbconn.transaction



        CreateParameterP_sCMD_CODE(objCommand)
        CreateParameterP_SCMD_ETAT(objCommand)
        CreateParameterP_SCMD_CMD_ID(objCommand)
        CreateParameterP_SCMD_FRN_ID(objCommand)
        CreateParameterP_SCMD_FACT_ID(objCommand)
        CreateParameterP_SCMD_DATE(objCommand)
        CreateParameterP_SCMD_TOTAL_HT(objCommand)
        CreateParameterP_SCMD_TOTAL_TTC(objCommand)
        CreateParameterP_SCMD_DATE_LIV(objCommand)
        CreateParameterP_SCMD_COM_COM(objCommand)
        CreateParameterP_SCMD_COM_LIV(objCommand)
        CreateParameterP_SCMD_COM_FACT(objCommand)
        CreateParameterP_SCMD_COM_LIBRE(objCommand)
        CreateParameterP_SCMD_TRP_NOM(objCommand)
        CreateParameterP_SCMD_TRP_RUE1(objCommand)
        CreateParameterP_SCMD_TRP_RUE2(objCommand)
        CreateParameterP_SCMD_TRP_CP(objCommand)
        CreateParameterP_SCMD_TRP_VILLE(objCommand)
        CreateParameterP_SCMD_TRP_TEL(objCommand)
        CreateParameterP_SCMD_TRP_FAX(objCommand)
        CreateParameterP_SCMD_TRP_PORT(objCommand)
        CreateParameterP_SCMD_TRP_EMAIL(objCommand)
        CreateParameterP_SCMD_REF_LIV(objCommand)
        CreateParameterP_SCMD_FACT_REF(objCommand)
        CreateParameterP_SCMD_FACT_DATE(objCommand)
        CreateParameterP_SCMD_FACT_TOTAL_HT(objCommand)
        CreateParameterP_SCMD_FACT_TOTAL_TTC(objCommand)
        CreateParameterP_SCMD_COM_BASE(objCommand)
        CreateParameterP_SCMD_COM_TAUX(objCommand)
        CreateParameterP_SCMD_COM_MONTANT(objCommand)
        CreateParameterP_SCMD_CLT_ID(objCommand)
        CreateParameterP_SCMD_BEXPORT(objCommand)
        CreateParameterP_SCMD_BEXPORTQUADRA(objCommand)
        CreateParameterP_ID(objCommand)
        Try
            objCommand.ExecuteNonQuery()
            m_dbconn.transaction.Commit()
            bReturn = True

        Catch ex As Exception
            setError("UpdateSCMD", ex.ToString())
            m_dbconn.transaction.Rollback()
            bReturn = False
        End Try

        Debug.Assert(bReturn, "UpdateSCMD" & getErreur())
        Return bReturn
    End Function 'UpdateSCMD
    Protected Function deleteSCMD() As Boolean
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "Id <> 0")
        Debug.Assert(Me.GetType().Name.Equals("SousCommande"))


        Dim sqlString As String = "DELETE FROM SOUSCOMMANDE WHERE SCMD_ID=? "
        Dim objCommand As OleDbCommand
        Dim objSCMD As SousCommande
        '        Dim objParam As OleDbParameter

        objSCMD = CType(Me, SousCommande)
        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        CreateParameterP_ID(objCommand)
        Try
            objCommand.ExecuteNonQuery()
            m_id = 0
            objSCMD.resetCode()
            Return True

        Catch ex As Exception
            setError("DeleteSCMD", ex.ToString())
            Return False
        End Try
    End Function 'DeleteSCMD
    Protected Shared Function ListeSCMD(Optional ByVal pddeb As Date = DATE_DEFAUT, Optional ByVal pdfin As Date = DATE_DEFAUT, Optional ByVal pCodeFourn As String = "", Optional ByVal pEtat As vncEtatCommande = vncEnums.vncEtatCommande.vncRien) As List(Of SousCommande)
        '============================================================================
        'Function : listeSCMD
        'Description : Rend une liste de sous commandes
        '============================================================================
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Dim colReturn As New List(Of SousCommande)
        '        Dim objParam As OleDbParameter
        Dim sqlString As String = "SELECT " &
                                "SOUSCOMMANDE.SCMD_ID, " &
                                "SOUSCOMMANDE.SCMD_CODE, " &
                                "SOUSCOMMANDE.SCMD_DATE, " &
                                "COMMANDE.CMD_ID, " &
                                "COMMANDE.CMD_CODE, " &
                                "SOUSCOMMANDE.SCMD_TOTAL_HT, " &
                                "SOUSCOMMANDE.SCMD_ETAT, " &
                                "SOUSCOMMANDE.SCMD_FACT_REF, " &
                                "FOURNISSEUR.FRN_ID, " &
                                "FOURNISSEUR.FRN_CODE, " &
                                "FOURNISSEUR.FRN_RS, " &
                                "CLIENT.CLT_ID, " &
                                "CLIENT.CLT_CODE, " &
                                "CLIENT.CLT_RS, " &
                                "SOUSCOMMANDE.SCMD_FACT_DATE, " &
                                "SOUSCOMMANDE.SCMD_COM_MONTANT, " &
                                "SOUSCOMMANDE.SCMD_COM_TAUX, " &
                                "SOUSCOMMANDE.SCMD_COM_BASE, " &
                                "SOUSCOMMANDE.SCMD_FACT_TOTAL_TTC," &
                                "SOUSCOMMANDE.SCMD_FACT_TOTAL_HT," &
                                "SOUSCOMMANDE.SCMD_FACT_ID, " &
                                "SOUSCOMMANDE.SCMD_BEXPORT, " &
                                "SOUSCOMMANDE.SCMD_BEXPORTQUADRA " &
                                "FROM " &
                                "((SOUSCOMMANDE INNER JOIN CLIENT ON SOUSCOMMANDE.SCMD_CLT_ID = CLIENT.CLT_ID) " &
                                " INNER JOIN COMMANDE ON SOUSCOMMANDE.SCMD_CMD_ID = COMMANDE.CMD_ID) " &
                                " INNER JOIN FOURNISSEUR ON SOUSCOMMANDE.SCMD_FRN_ID = FOURNISSEUR.FRN_ID "


        Dim strWhere As String = ""
        Dim objCommand As OleDbCommand
        '       Dim objSCMD As SousCommande
        Dim objRS As OleDbDataReader = Nothing
        '      Dim nId As Integer
        Dim strChampDate As String
        Dim objParam As OleDbParameter



        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection

        'Détermination du champ date
        'Si la demande porte sur les sous-commandes a facturer alors le champs date est SCMD_FACT_DATE
        strChampDate = "SCMD_DATE"
        'If pEtat = vncEnums.vncEtatCommande.vncSCMDRapprochee Or pEtat = vncEnums.vncEtatCommande.vncSCMDRapprocheeInt Then
        '    strChampDate = "SCMD_FACT_DATE"
        'End If

        If pddeb <> DATE_DEFAUT Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & " " & strChampDate & " >=  ?"
            objParam = objCommand.Parameters.AddWithValue("?", pddeb)

        End If
        If pdfin <> DATE_DEFAUT Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & " " & strChampDate & " <=  ?"
            objParam = objCommand.Parameters.AddWithValue("?", pdfin)

        End If
        If pCodeFourn <> "" Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "FOURNISSEUR.FRN_CODE LIKE ?"
            objParam = objCommand.Parameters.AddWithValue("?", pCodeFourn)

        End If

        If pEtat <> vncEnums.vncEtatCommande.vncRien Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "SOUSCOMMANDE.SCMD_ETAT = ?"
            objParam = objCommand.Parameters.AddWithValue("?", pEtat)

        End If



        If strWhere <> "" Then
            sqlString = sqlString & " WHERE " & strWhere
        End If
        sqlString = sqlString & " ORDER BY FOURNISSEUR.FRN_CODE ASC, SOUSCOMMANDE.SCMD_CODE ASC"
        objCommand.CommandText = sqlString


        colReturn = selectSCMD(objCommand)

        Return colReturn
    End Function 'ListeSCMD
    ''' <summary>
    ''' Liste des sous commandes à facturer en Commission 
    ''' </summary>
    ''' <param name="pddeb">Date de facture Fournisseur DEBUT</param>
    ''' <param name="pdfin">Date de facture Fournisseur FIN</param>
    ''' <param name="pCodeFourn">code Fournisseur ou ""</param>
    ''' <returns>une Collection de sous commandes</returns>
    ''' <remarks></remarks>
    Protected Shared Function ListeSCMDAFacturerCom(Optional ByVal pddeb As Date = DATE_DEFAUT, Optional ByVal pdfin As Date = DATE_DEFAUT, Optional ByVal pCodeFourn As String = "") As List(Of SousCommande)
        '============================================================================
        'Function : listeSCMD
        'Description : Rend une liste de sous commandes
        '============================================================================
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Dim colReturn As New List(Of SousCommande)
        Dim sqlString As String = "SELECT " &
                                "SOUSCOMMANDE.SCMD_ID, " &
                                "SOUSCOMMANDE.SCMD_CODE, " &
                                "SOUSCOMMANDE.SCMD_DATE, " &
                                "COMMANDE.CMD_ID, " &
                                "COMMANDE.CMD_CODE, " &
                                "SOUSCOMMANDE.SCMD_TOTAL_HT, " &
                                "SOUSCOMMANDE.SCMD_ETAT, " &
                                "SOUSCOMMANDE.SCMD_FACT_REF, " &
                                "FOURNISSEUR.FRN_ID, " &
                                "FOURNISSEUR.FRN_CODE, " &
                                "FOURNISSEUR.FRN_RS, " &
                                "CLIENT.CLT_ID, " &
                                "CLIENT.CLT_CODE, " &
                                "CLIENT.CLT_RS, " &
                                "SOUSCOMMANDE.SCMD_FACT_DATE, " &
                                "SOUSCOMMANDE.SCMD_COM_MONTANT, " &
                                "SOUSCOMMANDE.SCMD_COM_TAUX, " &
                                "SOUSCOMMANDE.SCMD_COM_BASE, " &
                                "SOUSCOMMANDE.SCMD_FACT_TOTAL_TTC," &
                                "SOUSCOMMANDE.SCMD_FACT_TOTAL_HT," &
                                "SOUSCOMMANDE.SCMD_FACT_ID, " &
                                "SOUSCOMMANDE.SCMD_BEXPORT, " &
                                "SOUSCOMMANDE.SCMD_BEXPORTQUADRA " &
                                "FROM " &
                                "((SOUSCOMMANDE INNER JOIN CLIENT ON SOUSCOMMANDE.SCMD_CLT_ID = CLIENT.CLT_ID) " &
                                " INNER JOIN COMMANDE ON SOUSCOMMANDE.SCMD_CMD_ID = COMMANDE.CMD_ID) " &
                                " INNER JOIN FOURNISSEUR ON SOUSCOMMANDE.SCMD_FRN_ID = FOURNISSEUR.FRN_ID "


        Dim strWhere As String = ""
        Dim objCommand As OleDbCommand
        Dim strChampDate As String
        Dim objParam As OleDbParameter



        'Creation de la commande
        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection

        'Pour les Souscommandes à Facture on recherche 
        ' LEs souscommandes avec un Tx de commissions
        ' Les sous commandes Rapprochées et RapprochéInternet
        strWhere = "SCMD_COM_TAUX <> 0 AND "
        strWhere = strWhere & "(SOUSCOMMANDE.SCMD_ETAT = " & vncEnums.vncEtatCommande.vncSCMDRapprochee & " or SOUSCOMMANDE.SCMD_ETAT = " & vncEnums.vncEtatCommande.vncSCMDRapprocheeInt & ")"

        'Détermination du champ date
        '23/01/09 : Filtre sur la date de facture Fournisseur
        strChampDate = "SCMD_FACT_DATE"

        If pddeb <> DATE_DEFAUT Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & " " & strChampDate & " >=  ?"
            objParam = objCommand.Parameters.AddWithValue("?", pddeb)

        End If
        If pdfin <> DATE_DEFAUT Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & " " & strChampDate & " <=  ?"
            objParam = objCommand.Parameters.AddWithValue("?", pdfin)

        End If
        If pCodeFourn <> "" Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "FOURNISSEUR.FRN_CODE LIKE ?"
            objParam = objCommand.Parameters.AddWithValue("?", pCodeFourn)

        End If


        If strWhere <> "" Then
            sqlString = sqlString & " WHERE " & strWhere
        End If
        sqlString = sqlString & " ORDER BY FOURNISSEUR.FRN_CODE ASC, SOUSCOMMANDE.SCMD_CODE ASC"

        objCommand.CommandText = sqlString

        'Lecture de la command
        colReturn = selectSCMD(objCommand)

        Return colReturn
    End Function 'ListeSCMDAFacturerCom

    Private Shared Function selectSCMD(ByVal objCommand As OleDbCommand) As List (of SousCommande)
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Dim colReturn As New List(Of SousCommande)
        Dim objSCMD As SousCommande
        Dim objRS As OleDbDataReader = Nothing
        Dim nId As Integer



        Try
            objRS = objCommand.ExecuteReader
            While objRS.Read()
                Try
                    nId = GetString(objRS, "SCMD_ID")
                    objSCMD = New SousCommande
                    objSCMD.setid(nId)
                    Try
                        objSCMD.code = GetString(objRS, "SCMD_CODE")
                    Catch ex As InvalidCastException
                        objSCMD.code = ""
                    End Try
                    Try
                        objSCMD.dateCommande = GetString(objRS, "SCMD_DATE")
                    Catch ex As InvalidCastException
                        objSCMD.dateCommande = DATE_DEFAUT
                    End Try

                    Try
                        objSCMD.oFournisseur.setid(GetString(objRS, "FRN_ID"))
                    Catch ex As InvalidCastException
                        objSCMD.oFournisseur.setid(0)
                    End Try
                    Try
                        objSCMD.oFournisseur.code = GetString(objRS, "FRN_CODE")
                    Catch ex As InvalidCastException
                        objSCMD.oFournisseur.code = ""
                    End Try
                    Try
                        objSCMD.oFournisseur.rs = GetString(objRS, "FRN_RS")
                    Catch ex As InvalidCastException
                        objSCMD.oFournisseur.rs = ""
                    End Try
                    Try
                        objSCMD.oTiers.setid(GetString(objRS, "CLT_ID"))
                    Catch ex As InvalidCastException
                        objSCMD.oTiers.setid(0)
                    End Try
                    Try
                        objSCMD.oTiers.code = GetString(objRS, "CLT_CODE")
                    Catch ex As InvalidCastException
                        objSCMD.oTiers.code = ""
                    End Try
                    Try
                        objSCMD.oTiers.rs = GetString(objRS, "CLT_RS")
                    Catch ex As InvalidCastException
                        objSCMD.oTiers.rs = ""
                    End Try
                    Try
                        objSCMD.idCommandeClient = GetString(objRS, "CMD_ID")
                    Catch ex As InvalidCastException
                        objSCMD.idCommandeClient = 0
                    End Try
                    Try
                        objSCMD.codeCommandeClient = GetString(objRS, "CMD_CODE")
                    Catch ex As InvalidCastException
                        objSCMD.codeCommandeClient = ""
                    End Try
                    Try
                        objSCMD.totalHT = GetString(objRS, "SCMD_TOTAL_HT")
                    Catch ex As InvalidCastException
                        objSCMD.totalHT = 0
                    End Try
                    Try
                        objSCMD.setEtat(GetString(objRS, "SCMD_ETAT"))
                    Catch ex As InvalidCastException
                        objSCMD.setEtat(vncEnums.vncEtatCommande.vncSCMDGeneree)
                    End Try
                    Try
                        objSCMD.refFactFournisseur = GetString(objRS, "SCMD_FACT_REF")
                    Catch ex As InvalidCastException
                        objSCMD.refFactFournisseur = ""
                    End Try
                    Try
                        objSCMD.dateFactFournisseur = GetString(objRS, "SCMD_FACT_DATE")
                    Catch ex As InvalidCastException
                        objSCMD.dateFactFournisseur = DATE_DEFAUT
                    End Try
                    Try
                        objSCMD.totalTTCFacture = GetString(objRS, "SCMD_FACT_TOTAL_TTC")
                    Catch ex As InvalidCastException
                        objSCMD.totalTTCFacture = 0
                    End Try
                    Try
                        objSCMD.totalHTFacture = GetString(objRS, "SCMD_FACT_TOTAL_HT")
                    Catch ex As InvalidCastException
                        objSCMD.totalHTFacture = 0
                    End Try
                    Try
                        objSCMD.MontantCommission = GetString(objRS, "SCMD_COM_MONTANT")
                    Catch ex As InvalidCastException
                        objSCMD.MontantCommission = 0
                    End Try
                    Try
                        objSCMD.tauxCommission = GetString(objRS, "SCMD_COM_TAUX")
                    Catch ex As InvalidCastException
                        objSCMD.tauxCommission = 0
                    End Try
                    Try
                        objSCMD.baseCommission = GetString(objRS, "SCMD_COM_BASE")
                    Catch ex As InvalidCastException
                        objSCMD.baseCommission = 0
                    End Try
                    Try
                        objSCMD.idFactCom = getInteger(objRS, "SCMD_FACT_ID")
                    Catch ex As InvalidCastException
                        objSCMD.idFactCom = 0
                    End Try
                    Try
                        objSCMD.bExportInternet = GetValue(objRS, "SCMD_BEXPORT")
                    Catch ex As InvalidCastException
                        objSCMD.bExportInternet = False
                    End Try
                    Try
                        objSCMD.bExportQuadra = GetValue(objRS, "SCMD_BEXPORTQUADRA")
                    Catch ex As InvalidCastException
                        objSCMD.bExportQuadra = False
                    End Try
                    objSCMD.m_bResume = True
                    objSCMD.resetBooleans()
                    colReturn.Add(objSCMD)
                Catch ex As InvalidCastException
                    colReturn = Nothing
                    Exit While
                End Try
            End While
            objRS.Close()
            Return colReturn
        Catch ex As Exception
            setError("SelectCMD", ex.ToString() & objCommand.CommandText)
            colReturn = Nothing
        End Try
        Debug.Assert(Not colReturn Is Nothing, "SelectCMD: colReturn is nothing" & SousCommande.getErreur())
        Return colReturn

    End Function ' selectSCMD
    ''' <summary>
    ''' Rend une liste de Sous Commandes a exporter
    ''' Si pTypeExportDemande = Quadra => filtre sur le bExportQuadra
    ''' Si pTypeExportDemande = Internet => filtre sur le Etat = generée
    ''' </summary>
    ''' <param name="pOrigine">Origine de la commande</param>
    ''' <param name="pddeb">date de debut</param>
    ''' <param name="pdfin">date de fin </param>
    ''' <param name="pCodeFourn">Code fournisseur </param>
    ''' <param name="pFRNTypeExport">Type d'export pour le fournisseur</param>
    ''' <param name="pTypeExportDemande">Type d'export Demandé</param>
    ''' <returns></returns>
    Protected Shared Function ListeSCMDAExporter(pOrigine As vncOrigineCmd,
                                                         ByVal pddeb As Date,
                                                         ByVal pdfin As Date,
                                                         ByVal pCodeFourn As String,
                                                         pFRNTypeExport As vncTypeExportScmd,
                                                         pTypeExportDemande As vncTypeExportScmd) _
                                                         As List(Of SousCommande)
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Dim colReturn As New List(Of SousCommande)
        '        Dim objParam As OleDbParameter
        Dim sqlString As String = "SELECT " &
                                "SOUSCOMMANDE.SCMD_ID, " &
                                "SOUSCOMMANDE.SCMD_CODE, " &
                                "SOUSCOMMANDE.SCMD_DATE, " &
                                "COMMANDE.CMD_ID, " &
                                "COMMANDE.CMD_CODE, " &
                                "SOUSCOMMANDE.SCMD_TOTAL_HT, " &
                                "SOUSCOMMANDE.SCMD_ETAT, " &
                                "SOUSCOMMANDE.SCMD_FACT_REF, " &
                                "FOURNISSEUR.FRN_ID, " &
                                "FOURNISSEUR.FRN_CODE, " &
                                "FOURNISSEUR.FRN_RS, " &
                                "CLIENT.CLT_ID, " &
                                "CLIENT.CLT_CODE, " &
                                "CLIENT.CLT_RS, " &
                                "SOUSCOMMANDE.SCMD_FACT_DATE, " &
                                "SOUSCOMMANDE.SCMD_COM_MONTANT, " &
                                "SOUSCOMMANDE.SCMD_COM_TAUX, " &
                                "SOUSCOMMANDE.SCMD_COM_BASE, " &
                                "SOUSCOMMANDE.SCMD_FACT_TOTAL_TTC," &
                                "SOUSCOMMANDE.SCMD_FACT_TOTAL_HT," &
                                "SOUSCOMMANDE.SCMD_FACT_ID, " &
                                "SOUSCOMMANDE.SCMD_BEXPORT, " &
                                "SOUSCOMMANDE.SCMD_BEXPORTQUADRA " &
                                "FROM " &
                                "((SOUSCOMMANDE INNER JOIN CLIENT ON SOUSCOMMANDE.SCMD_CLT_ID = CLIENT.CLT_ID) " &
                                " INNER JOIN COMMANDE ON SOUSCOMMANDE.SCMD_CMD_ID = COMMANDE.CMD_ID) " &
                                " INNER JOIN FOURNISSEUR ON SOUSCOMMANDE.SCMD_FRN_ID = FOURNISSEUR.FRN_ID "


        Dim strWhere As String = ""
        Dim objCommand As OleDbCommand
        Dim strChampDate As String
        Dim objParam As OleDbParameter



        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection

        'Détermination du champ date
        strChampDate = "SCMD_DATE"

        'Préparation du critere 

        If strWhere <> "" Then
            strWhere = strWhere & " AND "
        End If
        strWhere = strWhere & " " & strChampDate & " >=  ?"
        objParam = objCommand.Parameters.AddWithValue("?", pddeb)

        If strWhere <> "" Then
            strWhere = strWhere & " AND "
        End If
        strWhere = strWhere & " " & strChampDate & " <=  ?"
        objParam = objCommand.Parameters.AddWithValue("?", pdfin)

        If pCodeFourn <> "" Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "FOURNISSEUR.FRN_CODE LIKE ?"
            objParam = objCommand.Parameters.AddWithValue("?", pCodeFourn)

        End If

        'pour l'export internet On ne prend que les sous commande Générée
        If pTypeExportDemande = vncTypeExportScmd.vncExportInternet Then
            strWhere = strWhere & " AND SOUSCOMMANDE.SCMD_ETAT = " & vncEnums.vncEtatCommande.vncSCMDGeneree
        End If
        'pour l'export QUADRA On ne prend que les sous commande non Exportée (qqsoit l'état)
        If pTypeExportDemande = vncTypeExportScmd.vncExportQuadra Then
            strWhere = strWhere & " AND SOUSCOMMANDE.SCMD_BEXPORTQUADRA = 'false'"
        End If

        'On choisi l'origine des commandes
        If pOrigine = vncOrigineCmd.vncVinicom Then
            strWhere = strWhere & " AND COMMANDE.CMD_ORIGINE = '" & Dossier.VINICOM & "'"
        End If
        If pOrigine = vncOrigineCmd.vncHOBIVIN Then
            strWhere = strWhere & " AND COMMANDE.CMD_ORIGINE = '" & Dossier.HOBIVIN & "'"
        End If


        'Filtre sur le Type d'export du Fournisseur (Internet ou Quadra)
        strWhere = strWhere & " AND FOURNISSEUR.FRN_BEXP_INTERNET = " & pFRNTypeExport

        If strWhere <> "" Then
            sqlString = sqlString & " WHERE " & strWhere
        End If
        sqlString = sqlString & " ORDER BY FOURNISSEUR.FRN_CODE ASC, SOUSCOMMANDE.SCMD_CODE ASC"
        objCommand.CommandText = sqlString

        colReturn = selectSCMD(objCommand)

        Return colReturn
    End Function 'ListeSCMDAExporterInternet


    Protected Shared Function ListeFACTCOMEtat(Optional ByVal pddeb As Date = DATE_DEFAUT, Optional ByVal pdfin As Date = DATE_DEFAUT, Optional ByVal pCodeFourn As String = "", Optional ByVal pEtat As vncEtatCommande = vncEnums.vncEtatCommande.vncRien) As Collection
        '============================================================================
        'Function : ListeFACTCOMEtat
        'Description : Rend une liste de Facture de commission en fonction leur état
        '============================================================================
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Dim colReturn As New Collection
        Dim colTemp As New Collection
        '        Dim objParam As OleDbParameter
        Dim sqlString As String = "SELECT " & _
                                    "FCT_ID, FRN_ID " & _
                                  " FROM FACTCOM, FOURNISSEUR "

        Dim strWhere As String = " FACTCOM.FCT_FRN_ID = FOURNISSEUR.FRN_ID "
        Dim objCommand As OleDbCommand
        Dim objFCT As FactCom
        Dim objRS As OleDbDataReader = Nothing
        Dim nId As Integer
        Dim objParam As OleDbParameter



        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection


        If pddeb <> DATE_DEFAUT Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & " FCT_DATE >=  ?"
            objParam = objCommand.Parameters.AddWithValue("?", pddeb)

        End If
        If pdfin <> DATE_DEFAUT Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & " FCT_DATE <=  ?"
            objParam = objCommand.Parameters.AddWithValue("?", pdfin)

        End If
        If pCodeFourn <> "" Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "FOURNISSEUR.FRN_CODE LIKE ?"
            objParam = objCommand.Parameters.AddWithValue("?", pCodeFourn)

        End If

        If pEtat <> vncEnums.vncEtatCommande.vncRien Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "FCT_ETAT = ?"
            objParam = objCommand.Parameters.AddWithValue("?", pEtat)

        End If



        If strWhere <> "" Then
            sqlString = sqlString & "WHERE " & strWhere
        End If
        sqlString = sqlString & " ORDER BY FOURNISSEUR.FRN_CODE ASC"
        objCommand.CommandText = sqlString



        Try
            objRS = objCommand.ExecuteReader
            While objRS.Read()
                Try
                    nId = GetString(objRS, "FCT_ID")
                    objFCT = New FactCom
                    objFCT.setid(nId)
                    colTemp.Add(objFCT)
                Catch ex As InvalidCastException
                    colReturn = Nothing
                    Exit While
                End Try
            End While
            objRS.Close()
            'Parcours de la collection temporaire
            For Each objFCT In colTemp
                objFCT.load()
                objFCT.resetBooleans()
                If objFCT.id <> 0 Then
                    colReturn.Add(objFCT, objFCT.code)
                End If
            Next
            Return colReturn
        Catch ex As Exception
            setError("ListeFACTCOM", ex.ToString() & sqlString)
            colReturn = Nothing
        End Try
        Debug.Assert(Not colReturn Is Nothing, "ListeFACTCOMEtat: colReturn is nothing" & FactCom.getErreur())
        Return colReturn
    End Function 'ListeFactCom
    Protected Shared Function ListeCMDCLTEtat(Optional ByVal pddeb As Date = DATE_DEFAUT, Optional ByVal pdfin As Date = DATE_DEFAUT, Optional ByVal pRSClient As String = "", Optional ByVal pEtat As vncEtatCommande = vncEnums.vncEtatCommande.vncRien, Optional pOrigine As String = "") As Collection
        '============================================================================
        'Function : ListeFACTCOMEtat
        'Description : Rend une liste de Facture de commission en fonction leur état
        '============================================================================
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Dim colReturn As New Collection
        Dim colTemp As New Collection
        Dim colId As New Collection
        '        Dim objParam As OleDbParameter
        Dim sqlString As String = "SELECT " & _
                                    "CMD_ID, CLT_ID " & _
                                  " FROM COMMANDE, CLIENT "

        Dim strWhere As String = " COMMANDE.CMD_CLT_ID = CLIENT.CLT_ID "
        Dim objCommand As OleDbCommand
        Dim objCMD As CommandeClient
        Dim objRS As OleDbDataReader = Nothing
        Dim strId As String
        Dim objParam As OleDbParameter



        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection


        If pddeb <> DATE_DEFAUT Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & " CMD_DATE >=  ?"
            objParam = objCommand.Parameters.AddWithValue("?", pddeb)

        End If
        If pdfin <> DATE_DEFAUT Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & " CMD_DATE <=  ?"
            objParam = objCommand.Parameters.AddWithValue("?", pdfin)

        End If
        If pRSClient <> "" Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "CLIENT.CLT_RS LIKE ?"
            objParam = objCommand.Parameters.AddWithValue("?", pRSClient)

        End If

        If pEtat <> vncEnums.vncEtatCommande.vncRien Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "CMD_ETAT = ?"
            objParam = objCommand.Parameters.AddWithValue("?", pEtat)

        End If

        If pOrigine <> "" Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "CMD_ORIGINE = ?"
            objParam = objCommand.Parameters.AddWithValue("?", pOrigine)

        End If


        If strWhere <> "" Then
            sqlString = sqlString & "WHERE " & strWhere
        End If
        sqlString = sqlString & " ORDER BY CLIENT.CLT_CODE ASC"
        objCommand.CommandText = sqlString



        Try
            objRS = objCommand.ExecuteReader
            While objRS.Read
                Try
                    strId = GetString(objRS, "CMD_ID")
                    colTemp.Add(strId)
                Catch ex As InvalidCastException
                    colReturn = Nothing
                    Exit While
                End Try
            End While
            objRS.Close()
            For Each strId In colTemp
                objCMD = CommandeClient.createandload(strId)
                If objCMD.id <> 0 Then
                    colReturn.Add(objCMD, objCMD.code)
                End If
            Next
            Return colReturn
        Catch ex As Exception
            setError("ListeCMDCLTEtat", ex.ToString() & sqlString)
            colReturn = Nothing
        End Try
        Debug.Assert(Not colReturn Is Nothing, "ListeCMDCLTEtat: colReturn is nothing" & Commande.getErreur())
        Return colReturn
    End Function 'ListeCMDCLTEtat
    Protected Shared Function ListeCMDCLT_TRP(Optional ByVal pddeb As Date = DATE_DEFAUT, Optional ByVal pdfin As Date = DATE_DEFAUT, Optional ByVal pRSClient As String = "", Optional ByVal pCodeClient As String = "") As Collection
        '============================================================================
        'Function : ListeCMDCLT_TRP
        'Description : Rend une liste de Commande Client avec l'indicateur de facture de transport à YES et pas de factures de transport attachée
        '============================================================================
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Dim colReturn As New Collection
        Dim colTemp As New Collection
        '        Dim objParam As OleDbParameter
        Dim sqlString As String = "SELECT " & _
                                    "CMD_ID, CLT_ID " & _
                                  " FROM COMMANDE, CLIENT "

        Dim strWhere As String = " COMMANDE.CMD_CLT_ID = CLIENT.CLT_ID "
        strWhere = strWhere & " And COMMANDE.CMD_BFACTTRP=1 "
        strWhere = strWhere & " And COMMANDE.CMD_IDFACTTRP=0 "
        strWhere = strWhere & " And COMMANDE.CMD_ETAT in (" & vncEtatCommande.vncLivree & "," _
                                                            & vncEtatCommande.vncEclatee & "," _
                                                            & vncEtatCommande.vncRapprochee & "," _
                                                            & vncEtatCommande.vncTransmiseQuadra & ") "
        Dim objCommand As OleDbCommand
        Dim objCMD As CommandeClient
        Dim objRS As OleDbDataReader = Nothing
        Dim strId As String
        Dim objParam As OleDbParameter



        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection


        If pddeb <> DATE_DEFAUT Then
            If strWhere <> "" Then
                strWhere = strWhere & " And "
            End If
            strWhere = strWhere & " CMD_DATE_LIV >=  ?"
            objParam = objCommand.Parameters.AddWithValue("?", pddeb)

        End If
        If pdfin <> DATE_DEFAUT Then
            If strWhere <> "" Then
                strWhere = strWhere & " And "
            End If
            strWhere = strWhere & " CMD_DATE_LIV <=  ?"
            objParam = objCommand.Parameters.AddWithValue("?", pdfin)

        End If
        If pRSClient <> "" Then
            If strWhere <> "" Then
                strWhere = strWhere & " And "
            End If
            strWhere = strWhere & "CLIENT.CLT_RS Like ?"
            objParam = objCommand.Parameters.AddWithValue("?", pRSClient)

        End If

        If pCodeClient <> "" Then
            If strWhere <> "" Then
                strWhere = strWhere & " And "
            End If
            strWhere = strWhere & "CLIENT.CLT_CODE = ?"
            objParam = objCommand.Parameters.AddWithValue("?", pCodeClient)

        End If


        If strWhere <> "" Then
            sqlString = sqlString & "WHERE " & strWhere
        End If
        sqlString = sqlString & " ORDER BY CLIENT.CLT_CODE ASC"
        objCommand.CommandText = sqlString



        Try
            objRS = objCommand.ExecuteReader
            While objRS.Read
                Try
                    strId = GetString(objRS, "CMD_ID")
                    colTemp.Add(strId)
                Catch ex As InvalidCastException
                    colReturn = Nothing
                    Exit While
                End Try
            End While
            objRS.Close()
            For Each strId In colTemp
                objCMD = CommandeClient.createandload(strId)
                objCMD.resetBooleans()
                If objCMD.id <> 0 Then
                    colReturn.Add(objCMD, objCMD.code)
                End If

            Next
        Catch ex As Exception
            setError("ListeCMDCLT_TRP", ex.ToString() & sqlString)
            colReturn = Nothing
        End Try
        objRS = Nothing
        Debug.Assert(Not colReturn Is Nothing, "ListeCMDCLTEtat: colReturn is nothing" & Commande.getErreur())
        Return colReturn
    End Function 'ListeCMDCLT_TRP

    '================================================================================
    'Fonction : LoadColSCMD
    'Description : Chargement de la liste des Sous Commandes d'une commande
    '================================================================================
    Protected Function LoadColSCMD() As Boolean
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "Id <> 0")
        Debug.Assert(Me.GetType().Name.Equals("CommandeClient"))

        Dim bReturn As Boolean
        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        Dim objCMD As CommandeClient
        Dim objSCMD As SousCommande

        Dim sqlString As String = "SELECT " & _
                                    "SCMD_ID," & _
                                    "SCMD_CODE," & _
                                    "SCMD_ETAT," & _
                                    "SCMD_CMD_ID," & _
                                    "SCMD_FACT_ID," & _
                                    "SCMD_FRN_ID," & _
                                    "SCMD_DATE," & _
                                    "SCMD_TOTAL_HT," & _
                                    "SCMD_TOTAL_TTC," & _
                                    "SCMD_DATE_LIV," & _
                                    "SCMD_TRP_NOM," & _
                                    "SCMD_TRP_RUE1," & _
                                    "SCMD_TRP_RUE2," & _
                                    "SCMD_TRP_CP," & _
                                    "SCMD_TRP_VILLE," & _
                                    "SCMD_TRP_TEL," & _
                                    "SCMD_TRP_FAX," & _
                                    "sCMD_TRP_PORT," & _
                                    "sCMD_TRP_EMAIL," & _
                                    "SCMD_COM_COM," & _
                                    "SCMD_COM_LIV," & _
                                    "SCMD_COM_FACT," & _
                                    "sCMD_COM_LIBRE," & _
                                    "sCMD_REF_LIV," & _
                                    "SCMD_FACT_REF," & _
                                    "SCMD_FACT_DATE," & _
                                    "SCMD_FACT_TOTAL_HT," & _
                                    "SCMD_FACT_TOTAL_TTC," & _
                                    "SCMD_COM_BASE," & _
                                    "SCMD_COM_TAUX," & _
                                    "SCMD_COM_MONTANT," & _
                                    "SCMD_CLT_ID, " & _
                                    "CMD_CODE " & _
                                    " FROM SOUSCOMMANDE , COMMANDE " & _
                                    "WHERE " & _
                                    "SCMD_CMD_ID = CMD_ID AND " & _
                                    "SCMD_CMD_ID = ? "

        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        'Paramétre Commande ID 
        CreateParameterP_ID(objCommand)

        Try
            objCMD = CType(Me, CommandeClient)
            objRS = objCommand.ExecuteReader
            While objRS.Read
                Try
                    objSCMD = New SousCommande(Me, New Fournisseur)
                    objSCMD.setid(getInteger(objRS, "SCMD_ID"))
                    objSCMD.code = GetString(objRS, "SCMD_CODE")
                    objSCMD.setEtat(GetString(objRS, "SCMD_ETAT"))
                    objSCMD.idCommandeClient = GetString(objRS, "SCMD_CMD_ID")
                    objSCMD.idFactCom = getInteger(objRS, "SCMD_FACT_ID")
                    objSCMD.oTiers.setid(GetString(objRS, "SCMD_CLT_ID"))
                    objSCMD.oFournisseur.setid(GetString(objRS, "SCMD_FRN_ID"))
                    objSCMD.dateCommande = GetString(objRS, "SCMD_DATE")
                    objSCMD.totalHT = GetString(objRS, "SCMD_TOTAL_HT")
                    objSCMD.totalTTC = GetString(objRS, "SCMD_TOTAL_TTC")
                    objSCMD.dateLivraison = GetString(objRS, "SCMD_DATE_LIV")
                    objSCMD.oTransporteur.AdresseLivraison.nom = GetString(objRS, "SCMD_TRP_NOM")
                    objSCMD.oTransporteur.AdresseLivraison.rue1 = GetString(objRS, "SCMD_TRP_RUE1")
                    objSCMD.oTransporteur.AdresseLivraison.rue2 = GetString(objRS, "SCMD_TRP_RUE2")
                    objSCMD.oTransporteur.AdresseLivraison.cp = GetString(objRS, "SCMD_TRP_CP")
                    objSCMD.oTransporteur.AdresseLivraison.ville = GetString(objRS, "SCMD_TRP_VILLE")
                    objSCMD.oTransporteur.AdresseLivraison.tel = GetString(objRS, "SCMD_TRP_TEL")
                    objSCMD.oTransporteur.AdresseLivraison.fax = GetString(objRS, "SCMD_TRP_FAX")
                    objSCMD.oTransporteur.AdresseLivraison.port = GetString(objRS, "SCMD_TRP_PORT")
                    objSCMD.oTransporteur.AdresseLivraison.Email = GetString(objRS, "SCMD_TRP_EMAIL")
                    objSCMD.CommCommande.comment = (GetString(objRS, "SCMD_COM_COM"))
                    objSCMD.CommLivraison.comment = GetString(objRS, "SCMD_COM_LIV")
                    objSCMD.CommFacturation.comment = GetString(objRS, "SCMD_COM_FACT")
                    objSCMD.CommLibre.comment = GetString(objRS, "SCMD_COM_LIBRE")
                    objSCMD.refLivraison = GetString(objRS, "SCMD_REF_LIV")
                    objSCMD.refFactFournisseur = GetString(objRS, "SCMD_FACT_REF")
                    objSCMD.dateFactFournisseur = GetString(objRS, "SCMD_FACT_DATE")
                    objSCMD.totalHTFacture = GetString(objRS, "SCMD_FACT_TOTAL_HT")
                    objSCMD.totalTTCFacture = GetString(objRS, "SCMD_FACT_TOTAL_TTC")
                    objSCMD.baseCommission = GetString(objRS, "SCMD_COM_BASE")
                    objSCMD.tauxCommission = GetString(objRS, "SCMD_COM_TAUX")
                    objSCMD.MontantCommission = GetString(objRS, "SCMD_COM_MONTANT")
                    objSCMD.codeCommandeClient = GetString(objRS, "CMD_CODE")

                    objSCMD.resetBooleans()
                    objCMD.colSousCommandes.Add(objSCMD, objSCMD.code)
                Catch ex As InvalidCastException
                    bReturn = False
                    Exit While
                End Try
            End While
            objRS.Close()
            'Chargement des fournisseurs
            For Each objSCMD In objCMD.colSousCommandes
                objSCMD.oFournisseur.loadFRNLight()
            Next
            bReturn = True
        Catch ex As Exception
            setError("LoadcolSousCommande", ex.ToString())
            bReturn = False
        End Try

        Debug.Assert(bReturn, "LoadcolSousCommande" & getErreur())
        Return bReturn

    End Function ' LoadColSCMD

    '==========================================================================================
    ' Méthode : LoadcolLGCMD 
    'Description : Chargement des lignes de commandes (CommandeClient ou SousCommande)
    '               La clause where est choisie en fonction du type de l'objet courant
    '==========================================================================================
    Protected Function LoadcolLGCMD() As Boolean

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "L'id doit être renseigné")
        Debug.Assert(m_typedonnee = vncEnums.vncTypeDonnee.BA Or _
                     m_typedonnee = vncEnums.vncTypeDonnee.COMMANDECLIENT Or _
                     m_typedonnee = vncEnums.vncTypeDonnee.SSCOMMANDE _
                        , "Objet de type CommandeClient ou SousCommande ou bonAppro requis")
        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        '        Dim objParam As OleDbParameter
        Dim bReturn As Boolean
        Dim objCMD As Commande
        Dim objLGCMD As LgCommande
        Dim oProduit As Produit = Nothing
        Dim lgnum As String
        Dim qtecommande As Double
        Dim qteLivr As Decimal
        Dim qteFact As Decimal
        Dim prixU As Double
        Dim prixHT As Double
        Dim prixTTC As Double
        Dim bGratuit As Boolean
        Dim bEclatee As Boolean
        Dim idCmd As Integer
        Dim idscmd As Integer
        Dim idlgcmd As Integer
        Dim idBA As Integer
        Dim poids As Decimal
        Dim qteColis As Decimal
        Dim txComm As Decimal
        Dim MtComm As Decimal


        Dim sqlString As String = "SELECT " & _
                                    "LGCM_ID," & _
                                    "LGCM_NUM," & _
                                    "LGCM_CMD_ID," & _
                                    "LGCM_SCMD_ID," & _
                                    "LGCM_BA_ID," & _
                                    "LGCM_NUM," & _
                                    "LGCM_PRD_ID," & _
                                    "LGCM_QTE_COMMANDE," & _
                                    "LGCM_QTE_LIV," & _
                                    "LGCM_QTE_FACT," & _
                                    "LGCM_PRIX_UNITAIRE," & _
                                    "LGCM_PRIX_HT," & _
                                    "LGCM_PRIX_TTC," & _
                                    "LGCM_BGRATUIT, " & _
                                    "LGCM_BECLATEE, " & _
                                    "LGCM_POIDS, " & _
                                    "LGCM_QTE_COLIS, " & _
                                    "LGCM_TXCOMM, " & _
                                    "LGCM_MTCOMM, " & _
                                    "PRD_CODE" & _
                                    " FROM PRODUIT INNER JOIN LIGNE_COMMANDE ON (PRODUIT.PRD_ID = LIGNE_COMMANDE.LGCM_PRD_ID) "
        '                                  " FROM LIGNE_COMMANDE  "
        Dim strClauseWhereCMDCLT As String = " WHERE LIGNE_COMMANDE.LGCM_CMD_ID = ? "
        Dim strClauseWhereSCMD As String = " WHERE LIGNE_COMMANDE.LGCM_SCMD_ID = ? "
        Dim strClauseWhereBA As String = " WHERE LIGNE_COMMANDE.LGCM_BA_ID = ? "

        'Clause where en Fonction du type de l'objet courant
        Select Case m_typedonnee
            Case vncEnums.vncTypeDonnee.COMMANDECLIENT
                sqlString = sqlString & strClauseWhereCMDCLT
            Case vncEnums.vncTypeDonnee.SSCOMMANDE
                sqlString = sqlString & strClauseWhereSCMD
            Case vncEnums.vncTypeDonnee.BA
                sqlString = sqlString & strClauseWhereBA
        End Select

        'TRi par code Produit
        'sqlString = sqlString & " ORDER BY PRD_CODE ASC, LGCM_NUM ASC"
        'sqlString = sqlString & " ORDER BY LGCM_PRD_ID ASC, LGCM_NUM ASC"
        sqlString = sqlString & " ORDER BY PRD_CODE ASC, LIGNE_COMMANDE.LGCM_NUM ASC"

        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString

        'Paramétre ClientID 
        CreateParameterP_ID(objCommand)

        Try
            objCMD = CType(Me, Commande)
            objRS = objCommand.ExecuteReader
            bReturn = True
            While objRS.Read
                'Private m_num As Integer                'Numéro d'ordre de ligne
                'Private m_oProduit As Produit           'Produit Commandé
                'Private m_qteCommande As decimal 'Quantite commandée
                'Private m_prixU As Double               'Prix Unitaire
                'Private m_prixT As Double               'Prix Total
                'Private m_bGratuit As Boolean           'Ligne Gratuite

                Try
                    lgnum = GetString(objRS, "LGCM_NUM")
                Catch ex As InvalidCastException
                    lgnum = 0
                End Try

                Try 'Produit
                    '                    pushConnection()
                    oProduit = New Produit
                    oProduit.setid(getInteger(objRS, "LGCM_PRD_ID"))
                    '                   popConnection()
                Catch ex As InvalidCastException
                    oProduit.setid(0)
                End Try

                Try 'ID de la ligne commande
                    idlgcmd = Getinteger(objRS, "LGCM_ID")
                Catch ex As InvalidCastException
                    idlgcmd = 0
                End Try

                Try 'ID de la commande
                    idCmd = getInteger(objRS, "LGCM_CMD_ID")
                Catch ex As InvalidCastException
                    idCmd = 0
                End Try

                Try 'ID de la Souscommande
                    idscmd = getInteger(objRS, "LGCM_SCMD_ID")
                    If idscmd = -1 Then
                        idscmd = 0
                    End If
                Catch ex As InvalidCastException
                    idscmd = 0
                End Try
                Try 'ID Bon Appro
                    idBA = getInteger(objRS, "LGCM_BA_ID")
                Catch ex As InvalidCastException
                    idBA = 0
                End Try
                Try 'Qte Commandée
                    qtecommande = GetString(objRS, "LGCM_QTE_COMMANDE")
                Catch ex As InvalidCastException
                    qtecommande = 0
                End Try
                Try 'Qte Livrée
                    qteLivr = GetString(objRS, "LGCM_QTE_LIV")
                Catch ex As InvalidCastException
                    qteLivr = 0
                End Try
                Try 'Qte Facturée
                    qteFact = GetString(objRS, "LGCM_QTE_FACT")
                Catch ex As InvalidCastException
                    qteFact = 0
                End Try
                Try 'Prix unitaire
                    prixU = GetString(objRS, "LGCM_PRIX_UNITAIRE")
                Catch ex As InvalidCastException
                    prixU = 0
                End Try
                Try 'Prix HT
                    prixHT = GetString(objRS, "LGCM_PRIX_HT")
                Catch ex As InvalidCastException
                    prixHT = 0
                End Try
                Try 'Montant TTC
                    prixTTC = GetString(objRS, "LGCM_PRIX_TTC")
                Catch ex As InvalidCastException
                    prixTTC = 0
                End Try
                Try 'Ligne gratuite ou pas
                    bGratuit = GetString(objRS, "LGCM_BGRATUIT")
                Catch ex As InvalidCastException
                    bGratuit = False
                End Try
                Try 'Ligne Eclatée en sousCommande
                    bEclatee = GetBoolean(objRS, "LGCM_BECLATEE")
                Catch ex As InvalidCastException
                    bEclatee = False
                End Try
                Try 'Poids de la ligne
                    poids = GetString(objRS, "LGCM_POIDS")
                Catch ex As InvalidCastException
                    poids = 0
                End Try
                Try 'Nombre de colis
                    qteColis = GetValue(objRS, "LGCM_QTE_COLIS")
                Catch ex As InvalidCastException
                    qteColis = 0
                End Try
                Try 'Taux de commission
                    txComm = GetValue(objRS, "LGCM_TXCOMM")
                Catch ex As InvalidCastException
                    txComm = 0
                End Try
                Try 'Montant de la commission
                    MtComm = GetValue(objRS, "LGCM_MTCOMM")
                Catch ex As InvalidCastException
                    MtComm = 0
                End Try

                'Ajout de la ligne dans le client, le prix de la commande n'est pas recalculé!!!
                Try
                    objLGCMD = New LgCommande(idCmd, idscmd)
                    objLGCMD.num = lgnum
                    objLGCMD.oProduit = oProduit
                    objLGCMD.qteCommande = qtecommande
                    objLGCMD.qteLiv = qteLivr
                    objLGCMD.qteFact = qteFact
                    objLGCMD.prixU = prixU
                    objLGCMD.prixHT = prixHT
                    objLGCMD.prixTTC = prixTTC
                    objLGCMD.idCmd = idCmd
                    objLGCMD.idSCmd = idscmd
                    objLGCMD.idBA = idBA
                    objLGCMD.bGratuit = bGratuit
                    objLGCMD.bLigneEclatee = bEclatee
                    objLGCMD.poids = poids
                    objLGCMD.qteColis = qteColis
                    objLGCMD.TxComm = txComm
                    objLGCMD.MtComm = MtComm
                    objLGCMD.setid(idlgcmd)
                    objLGCMD.resetBooleans()
                    objLGCMD = objCMD.AjouteLigne(objLGCMD, False)
                    If objLGCMD Is Nothing Then
                        Throw New Exception("Impossible d'ajouter la ligne de sous commande" & lgnum)
                    End If

                Catch ex As Exception
                    setError("LoadLGCMDColLignes", "Erreur en ajout de ligne Commande" & ex.Message)
                    bReturn = False
                    Exit While
                End Try

            End While
            objRS.Close()
            shared_connect()
            For Each objLGCMD In objCMD.colLignes
                '                objLGCMD.oProduit.load()
                objLGCMD.oProduit.DBLoadLight()
            Next
            shared_disconnect()

        Catch ex As Exception
            setError("LoadCMDColLignes", ex.ToString())
            bReturn = False
        End Try

        If Not objRS Is Nothing Then
            objRS.Close()
        End If

        Debug.Assert(bReturn, getErreur())
        Return bReturn

    End Function 'LoadcolLGCMD
    Protected Function UPDATEcolLGCMD() As Boolean
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "La Commande.id  <> 0")
        Debug.Assert(m_typedonnee = vncEnums.vncTypeDonnee.COMMANDECLIENT Or _
                    m_typedonnee = vncEnums.vncTypeDonnee.SSCOMMANDE Or _
                    m_typedonnee = vncEnums.vncTypeDonnee.BA _
                    , "Objet de type CommandeClient ou SousCommande requis")


        Dim sqlString As String = "UPDATE LIGNE_COMMANDE SET " & _
                                    " LGCM_CMD_ID = ? ," & _
                                    " LGCM_SCMD_ID= ? ," & _
                                    " LGCM_BA_ID= ? ," & _
                                    " LGCM_NUM= ? ," & _
                                    " LGCM_PRD_ID = ? ," & _
                                    " LGCM_QTE_COMMANDE = ? ," & _
                                    " LGCM_QTE_LIV = ? ," & _
                                    " LGCM_QTE_FACT = ? ," & _
                                    " LGCM_PRIX_UNITAIRE = ? ," & _
                                    " LGCM_PRIX_HT = ? ," & _
                                    " LGCM_PRIX_TTC= ? ," & _
                                    " LGCM_POIDS = ? , " & _
                                    " LGCM_QTE_COLIS = ? , " & _
                                    " LGCM_TXCOMM = ? , " & _
                                    " LGCM_MTCOMM = ? , " & _
                                    " LGCM_BGRATUIT = ? , " & _
                                    " LGCM_BECLATEE = ?" & _
                                    " WHERE " & _
                                    " LGCM_ID = ?"
        Dim objCommand As OleDbCommand
        Dim objCMD As Commande
        '        Dim objParam As OleDbParameter
        Dim bReturn As Boolean
        Dim objlgCMD As LgCommande

        objCMD = CType(Me, Commande)
        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString
        m_dbconn.BeginTransaction()
        objCommand.Transaction = m_dbconn.transaction



        bReturn = True
        Dim oParamIDCMD As System.Data.OleDb.OleDbParameter = objCommand.Parameters.Add("?", OleDbType.Integer)
        Dim oParamIDSCMD As System.Data.OleDb.OleDbParameter = objCommand.Parameters.Add("?", OleDbType.Integer)
        Dim oParamIDBA As System.Data.OleDb.OleDbParameter = objCommand.Parameters.Add("?", OleDbType.Integer)
        Dim oParam4 As System.Data.OleDb.OleDbParameter = objCommand.Parameters.Add("?", OleDbType.Integer)
        Dim oParam5 As System.Data.OleDb.OleDbParameter = objCommand.Parameters.Add("?", OleDbType.Integer)
        Dim oParam6 As System.Data.OleDb.OleDbParameter = objCommand.Parameters.Add("?", OleDbType.Decimal)
        Dim oParam7 As System.Data.OleDb.OleDbParameter = objCommand.Parameters.Add("?", OleDbType.Decimal)
        Dim oParam8 As System.Data.OleDb.OleDbParameter = objCommand.Parameters.Add("?", OleDbType.Decimal)
        Dim oParam9 As System.Data.OleDb.OleDbParameter = objCommand.Parameters.Add("?", OleDbType.Decimal)
        Dim oParam10 As System.Data.OleDb.OleDbParameter = objCommand.Parameters.Add("?", OleDbType.Decimal)
        Dim oParam11 As System.Data.OleDb.OleDbParameter = objCommand.Parameters.Add("?", OleDbType.Decimal)
        Dim oParam12 As System.Data.OleDb.OleDbParameter = objCommand.Parameters.Add("?", OleDbType.Decimal)
        Dim oParam13 As System.Data.OleDb.OleDbParameter = objCommand.Parameters.Add("?", OleDbType.Decimal)
        'Taux de commission
        Dim oParamTxComm As System.Data.OleDb.OleDbParameter = objCommand.Parameters.Add("?", OleDbType.Decimal)
        'Montant de commssion
        Dim oParamMtcomm As System.Data.OleDb.OleDbParameter = objCommand.Parameters.Add("?", OleDbType.Decimal)

        Dim oParam14 As System.Data.OleDb.OleDbParameter = objCommand.Parameters.Add("?", OleDbType.Boolean)
        Dim oParam15 As System.Data.OleDb.OleDbParameter = objCommand.Parameters.Add("?", OleDbType.Boolean)
        Dim oParam16 As System.Data.OleDb.OleDbParameter = objCommand.Parameters.Add("?", OleDbType.Integer)
        For Each objlgCMD In objCMD.colLignes
            Try
                If (objlgCMD.idCmd = 0) Then
                    oParamIDCMD.Value = DBNull.Value
                Else
                    oParamIDCMD.Value = (objlgCMD.idCmd)
                End If
                If objlgCMD.idSCmd = 0 Then
                    oParamIDSCMD.Value = DBNull.Value
                Else
                    oParamIDSCMD.Value = (objlgCMD.idSCmd)
                End If
                If (objlgCMD.idBA = 0) Then
                    oParamIDBA.Value = DBNull.Value
                Else
                    oParamIDBA.Value = (objlgCMD.idBA)
                End If

                oParam4.Value = (objlgCMD.num)
                oParam5.Value = (objlgCMD.oProduit.id)
                oParam6.Value = (objlgCMD.qteCommande)
                oParam7.Value = (objlgCMD.qteLiv)
                oParam8.Value = (objlgCMD.qteFact)
                oParam9.Value = (objlgCMD.prixU)
                oParam10.Value = (objlgCMD.prixHT)
                oParam11.Value = (objlgCMD.prixTTC)
                oParam12.Value = (objlgCMD.poids)
                oParam13.Value = (objlgCMD.qteColis)
                oParam14.Value = (objlgCMD.bGratuit)
                oParam15.Value = (objlgCMD.bLigneEclatee)
                oParam16.Value = (objlgCMD.id)
                oParamTxComm.Value = (objlgCMD.TxComm)
                oParamMtcomm.Value = (objlgCMD.MtComm)
                objCommand.ExecuteNonQuery()
                'objRS.Close()
            Catch ex As Exception
                setError("UpdateSCMDColLignes", ex.ToString())
                bReturn = False
                Exit For
            End Try
        Next objlgCMD

        If bReturn Then
            m_dbconn.transaction.Commit()
        Else
            m_dbconn.transaction.Rollback()
        End If

        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function ' UPDATEcolLGCMD
    Protected Function INSERTcolLGCMD() As Boolean
        Dim objCMD As Commande
        Dim bReturn As Boolean
        Dim objLgCMD As LgCommande

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "Le Client doit être Sauvegardé")
        Debug.Assert(m_typedonnee = vncEnums.vncTypeDonnee.BA Or _
                     m_typedonnee = vncEnums.vncTypeDonnee.COMMANDECLIENT Or _
                     m_typedonnee = vncEnums.vncTypeDonnee.SSCOMMANDE _
                        , "Objet de type CommandeClient ou SousCommande ou bonAppro requis")
        objCMD = Me
        Debug.Assert(Not objCMD.colLignes Is Nothing, "ColLignes is Nothing")


        Dim sqlString As String = "INSERT INTO LIGNE_COMMANDE (" & _
                                    "LGCM_NUM," & _
                                    "LGCM_CODE," & _
                                    "LGCM_CMD_ID," & _
                                    "LGCM_SCMD_ID," & _
                                    "LGCM_PRD_ID," & _
                                    "LGCM_QTE_COMMANDE," & _
                                    "LGCM_QTE_LIV," & _
                                    "LGCM_QTE_FACT," & _
                                    "LGCM_PRIX_UNITAIRE," & _
                                    "LGCM_PRIX_HT," & _
                                    "LGCM_PRIX_TTC," & _
                                    "LGCM_BGRATUIT," & _
                                    "LGCM_BECLATEE, " & _
                                    "LGCM_POIDS, " & _
                                    "LGCM_QTE_COLIS, " & _
                                    "LGCM_TXCOMM, " & _
                                    "LGCM_MTCOMM, " & _
                                    "LGCM_BA_ID " & _
                                    ") VALUES ( " & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "?" & _
                                  " )"
        Dim objCommand As OleDbCommand
        Dim objCommand2 As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        '        Dim objParam As OleDbParameter
        Dim bTypeCommandeClient As Boolean

        ' Vrai si l'objet courant est de type CommandeClient
        bTypeCommandeClient = Me.GetType.Name.Equals("CommandeClient")
        'Clause where en Fonction du type de l'objet courant

        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString
        m_dbconn.BeginTransaction()
        objCommand.Transaction = m_dbconn.transaction
        'Commande de lecture de l'id
        objCommand2 = New OleDbCommand("SELECT MAX(LGCM_ID) FROM LIGNE_COMMANDE", m_dbconn.Connection)
        objCommand2.Transaction = m_dbconn.transaction


        bReturn = True
        Dim oParam1 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Integer)
        Dim oParam2 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.VarChar)
        Dim oParamIDCmd As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Integer)
        oParamIDCmd.IsNullable = True
        Dim oParamIDSCMD As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Integer)
        oParamIDSCMD.IsNullable = True
        Dim oParam5 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Integer)
        Dim oParam6 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Double)
        Dim oParam7 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Double)
        Dim oParam8 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Double)
        Dim oParam9 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Double)
        Dim oParam10 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Double)
        Dim oParam11 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Double)
        Dim oParam12 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Boolean)
        Dim oParam13 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Boolean)
        Dim oParam14 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Double)
        Dim oParam15 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Double)
        Dim oParamTxComm As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Decimal)
        Dim oParamMtComm As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Decimal)
        Dim oParamIDBa As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Integer)
        oParamIDBa.IsNullable = True
        For Each objLgCMD In objCMD.colLignes
            Try
                'Affectation de id de la commande ou sous-commande en fonction du type l'objet
                If m_typedonnee = vncEnums.vncTypeDonnee.BA Then
                    objLgCMD.idBA = id
                    objLgCMD.idCmd = 0
                    objLgCMD.idSCmd = 0
                End If
                If m_typedonnee = vncEnums.vncTypeDonnee.SSCOMMANDE Then
                    objLgCMD.idBA = 0
                    objLgCMD.idCmd = 0
                    objLgCMD.idSCmd = id
                End If
                If m_typedonnee = vncEnums.vncTypeDonnee.COMMANDECLIENT Then
                    objLgCMD.idBA = 0
                    objLgCMD.idCmd = id
                    'objLgCMD.idSCmd = id
                End If

                oParam1.Value = objLgCMD.num
                oParam2.Value = (objCMD.code & objLgCMD.num)
                If objLgCMD.idCmd = 0 Or objLgCMD.idCmd = -1 Then
                    oParamIDCmd.Value = DBNull.Value
                Else
                    oParamIDCmd.Value = (objLgCMD.idCmd)
                End If
                If objLgCMD.idSCmd = 0 Or objLgCMD.idSCmd = -1 Then
                    oParamIDSCMD.Value = DBNull.Value
                Else
                    oParamIDSCMD.Value = (objLgCMD.idSCmd)
                End If
                oParam5.Value = (objLgCMD.oProduit.id)
                oParam6.Value = (CDbl(objLgCMD.qteCommande))
                oParam7.Value = (CDbl(objLgCMD.qteLiv))
                oParam8.Value = (CDbl(objLgCMD.qteFact))
                oParam9.Value = (CDbl(objLgCMD.prixU))
                oParam10.Value = (CDbl(objLgCMD.prixHT))
                oParam11.Value = (CDbl(objLgCMD.prixTTC))
                oParam12.Value = (objLgCMD.bGratuit)
                oParam13.Value = (objLgCMD.bLigneEclatee)
                oParam14.Value = (CDbl(objLgCMD.poids))
                oParam15.Value = (CDbl(objLgCMD.qteColis))
                oParamTxComm.Value = objLgCMD.TxComm
                oParamMtComm.Value = objLgCMD.MtComm
                If objLgCMD.idBA = 0 Or objLgCMD.idBA = -1 Then
                    oParamIDBa.Value = DBNull.Value
                Else
                    oParamIDBa.Value = (objLgCMD.idBA)
                End If
                objCommand.ExecuteNonQuery()
                objRS = objCommand2.ExecuteReader
                If objRS.Read() Then
                    objLgCMD.setid(objRS.GetInt32(0))
                End If
                cleanErreur()
                objRS.Close()
            Catch ex As Exception
                setError("InsertcolLGCmd", ex.ToString())
                bReturn = False
                Exit For
            End Try
        Next
        If bReturn Then
            m_dbconn.transaction.Commit()
        Else
            m_dbconn.transaction.Rollback()
        End If

        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function ' insertcolLGCMD
    '==========================================================================
    'Methode : deletecolLGCMD
    'Description : Suppression des lignes d'une commande
    '               la clé de destrcution est choisie en fonction de la classe de l'objet
    Protected Function deletecolLgCMD() As Boolean
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "Id <> 0")
        Debug.Assert(m_typedonnee = vncEnums.vncTypeDonnee.BA Or _
                     m_typedonnee = vncEnums.vncTypeDonnee.COMMANDECLIENT Or _
                     m_typedonnee = vncEnums.vncTypeDonnee.SSCOMMANDE _
                        , "Objet de type CommandeClient ou SousCommande ou bonAppro requis")


        Dim sqlString As String = "DELETE FROM LIGNE_COMMANDE "
        Dim strClauseWhereCMDCLT As String = " WHERE LIGNE_COMMANDE.LGCM_CMD_ID = ? "
        Dim strClauseWhereSCMD As String = " WHERE LIGNE_COMMANDE.LGCM_SCMD_ID = ? "
        Dim strClauseWhereBA As String = " WHERE LIGNE_COMMANDE.LGCM_BA_ID = ? "
        Dim objCMD As Commande
        Dim objCommand As OleDbCommand
        '        Dim objParam As OleDbParameter
        Dim bReturn As Boolean

        'Clause where en Fonction du type de l'objet courant
        Select Case m_typedonnee
            Case vncEnums.vncTypeDonnee.COMMANDECLIENT
                sqlString = sqlString & strClauseWhereCMDCLT
            Case vncEnums.vncTypeDonnee.SSCOMMANDE
                sqlString = sqlString & strClauseWhereSCMD
            Case vncEnums.vncTypeDonnee.BA
                sqlString = sqlString & strClauseWhereBA
        End Select

        objCMD = CType(Me, Commande)
        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        CreateParameterP_ID(objCommand)
        Try
            objCommand.ExecuteNonQuery()

            bReturn = True

        Catch ex As Exception
            setError("DeletecolLGCMD", ex.ToString())
            bReturn = False
        End Try
        Debug.Assert(bReturn, "deletecolLgCMD" & getErreur())
        Return bReturn
    End Function 'DeletecolLGCMD

#Region "CreateParameter SpousCommande"
    Private Sub CreateParameterP_sCMD_CODE(ByVal objCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCMD.code, 50))
    End Sub
    Private Sub CreateParameterP_SCMD_ETAT(ByVal objCommand As OleDbCommand)
        Dim objSCMD As SousCommande
        objSCMD = Me
        objCommand.Parameters.AddWithValue("?", objSCMD.etat.codeEtat)
    End Sub
    Private Sub CreateParameterP_SCMD_CMD_ID(ByVal objCommand As OleDbCommand)
        Dim objCMD As SousCommande
        objCMD = Me
        Debug.Assert(objCMD.idCommandeClient <> 0)
        objCommand.Parameters.AddWithValue("?", objCMD.idCommandeClient)
    End Sub
    Private Sub CreateParameterP_SCMD_FRN_ID(ByVal objCommand As OleDbCommand)
        Dim objCMD As SousCommande
        objCMD = Me
        Debug.Assert(objCMD.oFournisseur.id <> 0)
        objCommand.Parameters.AddWithValue("?", objCMD.oFournisseur.id)
    End Sub
    Private Sub CreateParameterP_SCMD_FACT_ID(ByVal objCommand As OleDbCommand)
        Dim objCMD As SousCommande
        objCMD = Me
        If (objCMD.idFactCom = 0) Or (objCMD.idFactCom = -1) Then
            objCommand.Parameters.AddWithValue("?", DBNull.Value)
        Else
            objCommand.Parameters.AddWithValue("?", objCMD.idFactCom)
        End If
    End Sub
    Private Sub CreateParameterP_SCMD_DATE(ByVal objCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", objCMD.dateCommande.ToShortDateString)
    End Sub
    Private Sub CreateParameterP_SCMD_TOTAL_HT(ByVal objCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", CDbl(objCMD.totalHT))
    End Sub
    Private Sub CreateParameterP_SCMD_TOTAL_TTC(ByVal objCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", CDbl(objCMD.totalTTC))
    End Sub
    Private Sub CreateParameterP_SCMD_DATE_LIV(ByVal objCommand As OleDbCommand)
        Dim objCMD As SousCommande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", objCMD.dateLivraison.ToShortDateString)
    End Sub
    Private Sub CreateParameterP_SCMD_COM_COM(ByVal objCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", objCMD.CommCommande.comment)
    End Sub
    Private Sub CreateParameterP_SCMD_COM_LIV(ByVal objCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", objCMD.CommLivraison.comment)
    End Sub
    Private Sub CreateParameterP_SCMD_COM_FACT(ByVal objCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", objCMD.CommFacturation.comment)
    End Sub
    Private Sub CreateParameterP_SCMD_COM_LIBRE(ByVal objCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", objCMD.CommLibre.comment)
    End Sub
    Private Sub CreateParameterP_SCMD_TRP_NOM(ByVal objCommand As OleDbCommand)
        Dim objCMD As SousCommande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCMD.oTransporteur.AdresseLivraison.nom, 50))
    End Sub
    Private Sub CreateParameterP_SCMD_TRP_RUE1(ByVal objCommand As OleDbCommand)
        Dim objCMD As SousCommande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCMD.oTransporteur.AdresseLivraison.rue1, 50))
    End Sub
    Private Sub CreateParameterP_SCMD_TRP_RUE2(ByVal objCommand As OleDbCommand)
        Dim objCMD As SousCommande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCMD.oTransporteur.AdresseLivraison.rue2, 50))
    End Sub
    Private Sub CreateParameterP_SCMD_TRP_CP(ByVal objCommand As OleDbCommand)
        Dim objCMD As SousCommande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCMD.oTransporteur.AdresseLivraison.cp, 50))
    End Sub
    Private Sub CreateParameterP_SCMD_TRP_VILLE(ByVal objCommand As OleDbCommand)
        Dim objCMD As SousCommande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCMD.oTransporteur.AdresseLivraison.ville, 50))
    End Sub
    Private Sub CreateParameterP_SCMD_TRP_TEL(ByVal objCommand As OleDbCommand)
        Dim objCMD As SousCommande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCMD.oTransporteur.AdresseLivraison.tel, 50))

    End Sub
    Private Sub CreateParameterP_SCMD_TRP_FAX(ByVal objCommand As OleDbCommand)
        Dim objCMD As SousCommande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCMD.oTransporteur.AdresseLivraison.fax, 50))

    End Sub
    Private Sub CreateParameterP_SCMD_TRP_PORT(ByVal objCommand As OleDbCommand)
        Dim objCMD As SousCommande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCMD.oTransporteur.AdresseLivraison.port, 50))

    End Sub
    Private Sub CreateParameterP_SCMD_TRP_EMAIL(ByVal objCommand As OleDbCommand)
        Dim objCMD As SousCommande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCMD.oTransporteur.AdresseLivraison.Email, 50))
    End Sub
    Private Sub CreateParameterP_SCMD_REF_LIV(ByVal objCommand As OleDbCommand)
        Dim objCMD As SousCommande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCMD.refLivraison, 50))
    End Sub
    Private Sub CreateParameterP_SCMD_FACT_REF(ByVal objCommand As OleDbCommand)
        Dim objCMD As SousCommande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCMD.refFactFournisseur, 50))
    End Sub
    Private Sub CreateParameterP_SCMD_FACT_DATE(ByVal objCommand As OleDbCommand)
        Dim objCMD As SousCommande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", objCMD.dateFactFournisseur.ToShortDateString)
    End Sub
    Private Sub CreateParameterP_SCMD_FACT_TOTAL_HT(ByVal objCommand As OleDbCommand)
        Dim objCMD As SousCommande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", CDbl(objCMD.totalHTFacture))
    End Sub
    Private Sub CreateParameterP_SCMD_FACT_TOTAL_TTC(ByVal objCommand As OleDbCommand)
        Dim objCMD As SousCommande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", CDbl(objCMD.totalTTCFacture))
    End Sub
    Private Sub CreateParameterP_SCMD_COM_BASE(ByVal objCommand As OleDbCommand)
        Dim objCMD As SousCommande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", CDbl(objCMD.baseCommission))
    End Sub
    Private Sub CreateParameterP_SCMD_COM_TAUX(ByVal objCommand As OleDbCommand)
        Dim objCMD As SousCommande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", CDbl(objCMD.tauxCommission))
    End Sub
    Private Sub CreateParameterP_SCMD_COM_MONTANT(ByVal objCommand As OleDbCommand)
        Dim objCMD As SousCommande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", CDbl(objCMD.MontantCommission))
    End Sub
    Private Sub CreateParameterP_SCMD_CLT_ID(ByVal objCommand As OleDbCommand)
        Dim objCMD As SousCommande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", objCMD.oTiers.id)
    End Sub
    Private Sub CreateParameterP_SCMD_BEXPORT(ByVal objCommand As OleDbCommand)
        Dim objCMD As SousCommande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", objCMD.bExportInternet)
    End Sub
    Private Sub CreateParameterP_SCMD_BEXPORTQUADRA(ByVal objCommand As OleDbCommand)
        Dim objCMD As SousCommande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", objCMD.bExportQuadra)
    End Sub
#End Region
#End Region

#Region "BonAppro"
    Protected Shared Function ListeBA(ByVal strCode As String, ByVal strRSFourn As String, ByVal pEtat As vncEtatCommande) As Collection
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Dim colReturn As New Collection
        '        Dim objParam As OleDbParameter
        Dim sqlString As String = "SELECT CMD_ID, CMD_CODE, CMD_DATE, FRN_ID , FRN_CODE, FRN_RS, CMD_ETAT " & _
                                    "FROM BONAPPRO , FOURNISSEUR "
        Dim strWhere As String = " BONAPPRO.CMD_FRN_ID = FOURNISSEUR.FRN_ID "
        Dim objCommand As OleDbCommand
        Dim objBA As BonAppro
        Dim objRS As OleDbDataReader = Nothing
        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        Dim objParam As OleDbParameter

        If strCode <> "" Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "CMD_CODE LIKE ?"
            objParam = objCommand.Parameters.AddWithValue("?", strCode)

        End If

        If strRSFourn <> "" Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "FOURNISSEUR.FRN_RS LIKE ?"
            objParam = objCommand.Parameters.AddWithValue("?", strRSFourn)

        End If

        If pEtat <> vncEnums.vncEtatCommande.vncRien Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "CMD_ETAT = ?"
            objParam = objCommand.Parameters.AddWithValue("?", pEtat)

        End If


        If strWhere <> "" Then
            sqlString = sqlString & "WHERE " & strWhere
        End If
        sqlString = sqlString & " ORDER BY CMD_DATE DESC"
        objCommand.CommandText = sqlString



        Try
            objRS = objCommand.ExecuteReader
            While (objRS.Read())
                objBA = New BonAppro(New Fournisseur)
                objBA.m_bResume = True ' C'est un objet Résumé
                Try
                    objBA.m_id = GetString(objRS, "CMD_ID")
                Catch ex As InvalidCastException
                    Debug.Assert(False, "Erreur de cast value = " & GetString(objRS, "CMD_ID"))
                    objBA.m_id = 0
                End Try
                Try
                    objBA.code = GetString(objRS, "CMD_CODE")
                Catch ex As InvalidCastException
                    Debug.Assert(False, "Erreur de cast value = " & GetString(objRS, "CMD_CODE"))
                    objBA.code = ""
                End Try
                Try
                    objBA.oTiers.setid(GetString(objRS, "FRN_ID"))
                Catch ex As InvalidCastException
                    Debug.Assert(False, "Erreur de cast value = " & GetString(objRS, "FRN_ID"))
                    objBA.oTiers.setid(0)
                End Try
                Try
                    objBA.oTiers.code = GetString(objRS, "FRN_CODE")
                Catch ex As InvalidCastException
                    Debug.Assert(False, "Erreur de cast value = " & GetString(objRS, "FRN_CODE"))
                    objBA.oTiers.code = ""
                End Try
                Try
                    objBA.oTiers.rs = GetString(objRS, "FRN_RS")
                Catch ex As InvalidCastException
                    Debug.Assert(False, "Erreur de cast value = " & GetString(objRS, "FRN_NOM"))
                    objBA.oTiers.rs = ""
                End Try

                Try
                    objBA.dateCommande = GetString(objRS, "CMD_DATE")
                Catch ex As InvalidCastException
                    Debug.Assert(False, "Erreur de cast value = " & GetString(objRS, "CMD_DATE"))
                    objBA.dateCommande = DATE_DEFAUT
                End Try

                Try
                    objBA.setEtat(GetString(objRS, "CMD_ETAT"))
                Catch ex As InvalidCastException
                    Debug.Assert(False, "Erreur de cast value = " & GetString(objRS, "CMD_ETAT"))
                End Try

                objBA.resetBooleans()
                colReturn.Add(objBA, objBA.code)
            End While
            objRS.Close()
            objRS = Nothing
            Return colReturn
        Catch ex As Exception
            setError("ListBA", ex.ToString())
            Return Nothing
        End Try
    End Function 'ListeBA
    Protected Shared Function ListeBAEtat(Optional ByVal pddeb As Date = DATE_DEFAUT, Optional ByVal pdfin As Date = DATE_DEFAUT, Optional ByVal pRSClient As String = "", Optional ByVal pEtat As vncEtatCommande = vncEnums.vncEtatCommande.vncRien) As Collection
        '============================================================================
        'Function : ListeBAEtat
        'Description : Rend une liste de BA en fonction leur état
        '============================================================================
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Dim colReturn As New Collection
        Dim colTemp As New Collection
        '        Dim objParam As OleDbParameter
        Dim sqlString As String = "SELECT " & _
                                    "CMD_ID, FRN_ID " & _
                                  " FROM BONAPPRO, FOURNISSEUR "

        Dim strWhere As String = " BONAPPRO.CMD_FRN_ID = FOURNISSEUR.FRN_ID "
        Dim objCommand As OleDbCommand
        Dim objBA As BonAppro
        Dim objRS As OleDbDataReader = Nothing
        Dim strId As String
        Dim objParam As OleDbParameter

        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection

        If pddeb <> DATE_DEFAUT Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & " CMD_DATE >=  ?"
            objParam = objCommand.Parameters.AddWithValue("?", pddeb)

        End If

        If pdfin <> DATE_DEFAUT Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & " CMD_DATE <=  ?"
            objParam = objCommand.Parameters.AddWithValue("?", pdfin)

        End If

        If pRSClient <> "" Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "FOURNISSEUR.FRN_RS LIKE ?"
            objParam = objCommand.Parameters.AddWithValue("?", pRSClient)

        End If

        If pEtat <> vncEnums.vncEtatCommande.vncRien Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "CMD_ETAT = ?"
            objParam = objCommand.Parameters.AddWithValue("?", pEtat)

        End If



        If strWhere <> "" Then
            sqlString = sqlString & "WHERE " & strWhere
        End If
        sqlString = sqlString & " ORDER BY FOURNISSEUR.FRN_CODE ASC"
        objCommand.CommandText = sqlString



        Try
            objRS = objCommand.ExecuteReader
            While objRS.Read()
                Try
                    strId = GetString(objRS, "CMD_ID")
                    colTemp.Add(strId)
                Catch ex As InvalidCastException
                    colReturn = Nothing
                    Exit While
                End Try
            End While
            objRS.Close()
            objRS = Nothing
            For Each strId In colTemp
                objBA = BonAppro.createandload(strId)
                If objBA.id <> 0 Then
                    colReturn.Add(objBA, objBA.code)
                End If

            Next
            Return colReturn
        Catch ex As Exception
            setError("ListeBAEtat", ex.ToString() & sqlString)
            colReturn = Nothing
        End Try

        Debug.Assert(Not colReturn Is Nothing, "ListeBAEtat: colReturn is nothing" & getErreur())
        Return colReturn
    End Function 'ListeBA

    Protected Function deleteBA() As Boolean
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "Id <> 0")
        Debug.Assert(Me.GetType().Name.Equals("BonAppro"))


        Dim sqlString As String = "DELETE FROM BONAPPRO WHERE CMD_ID=? "
        Dim objCommand As OleDbCommand
        Dim objBA As BonAppro
        '        Dim objParam As OleDbParameter

        objBA = CType(Me, BonAppro)
        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        CreateParameterP_ID(objCommand)
        Try
            objCommand.ExecuteNonQuery()
            m_id = 0
            objBA.resetCode()
            Return True

        Catch ex As Exception
            setError("DeleteBA", ex.ToString())
            Return False
        End Try
    End Function 'DeleteBA
    '=======================================================================
    'Fonction : DeleteCMDCLT_MVTSTK
    'Description : Suppression des mouvements de stocks d'une commande
    'Retour : Rend Vrai si le DELETE s'est correctement effectué
    '=======================================================================
    Protected Function deleteBA_MVTSTK() As Boolean
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "Id <> 0")


        Dim sqlString As String = "DELETE FROM MVT_STOCK WHERE STK_TYPE=3 and STK_REF_ID=? "
        Dim objCommand As OleDbCommand
        '        Dim objParam As OleDbParameter

        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        CreateParameterP_ID(objCommand)
        Try
            objCommand.ExecuteNonQuery()
            Return True

        Catch ex As Exception
            setError("DeleteBA_MVTSTK", ex.ToString())
            Return False
        End Try
    End Function 'DeleteBA_MVTSTK
    Protected Function deleteBAcolLgCMD() As Boolean
        '==========================================================================
        'Methode : deleteBAcolLGCMD
        'Description : Suppression des lignes d'une commande
        '               la clé de destrcution est choisie en fonction de la classe de l'objet
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "Id <> 0")
        Debug.Assert(Me.GetType().Name.Equals("BonAppro"), "Objet de type CommandeClient ou SousCommande requis")


        Dim sqlString As String = "DELETE FROM LIGNE_COMMANDE "
        Dim strClauseWhereBA As String = " WHERE LIGNE_COMMANDE.LGCM_BA_ID = ? "
        Dim objBA As BonAppro
        Dim objCommand As OleDbCommand
        '        Dim objParam As OleDbParameter
        Dim bReturn As Boolean

        sqlString = sqlString & strClauseWhereBA

        objBA = CType(Me, BonAppro)
        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        CreateParameterP_ID(objCommand)
        Try
            objCommand.ExecuteNonQuery()
            bReturn = True

        Catch ex As Exception
            setError("DeleteBAcolLGCMD", ex.ToString())
            bReturn = False
        End Try
        Debug.Assert(bReturn, "deleteBAcolLgCMD" & getErreur())
        Return bReturn
    End Function 'DeleteBAcolLGCMD
    Protected Function insertBA() As Boolean
        '=======================================================================
        'Fonction : InsertBA
        'Description : Insertion d'un Bon Appro

        'Retour : Rend Vrai si l'INSERT s'est correctement effectué
        '=======================================================================
        Dim bReturn As Boolean
        Dim objBA As BonAppro
        bReturn = False

        Debug.Assert(m_typedonnee = vncEnums.vncTypeDonnee.BA, "Objet de Type BonAppro Requis")
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id = 0, "ID=0")
        objBA = CType(Me, BonAppro)

        Dim sqlString As String = "INSERT INTO BONAPPRO( " & _
                                    "CMD_FRN_ID," & _
                                    "CMD_DATE," & _
                                    "CMD_DATE_VALID," & _
                                    "CMD_DATE_LIV," & _
                                    "CMD_DATE_ENLEV," & _
                                    "CMD_REF_LIV," & _
                                    "CMD_ETAT," & _
                                    "CMD_TOTAL_HT," & _
                                    "CMD_TOTAL_TTC," & _
                                    "CMD_TYPE," & _
                                    "CMD_TYPE_TRANSPORT," & _
                                    "CMD_CODE," & _
                                    "CMD_CLT_LIV_NOM," & _
                                    "CMD_CLT_LIV_RUE1," & _
                                    "CMD_CLT_LIV_RUE2," & _
                                    "CMD_CLT_LIV_CP," & _
                                    "CMD_CLT_LIV_VILLE," & _
                                    "CMD_CLT_LIV_TEL," & _
                                    "CMD_CLT_LIV_FAX," & _
                                    "CMD_CLT_LIV_PORT," & _
                                    "CMD_CLT_LIV_EMAIL," & _
                                    "CMD_CLT_ADR_IDENT," & _
                                    "CMD_CLT_FACT_NOM," & _
                                    "CMD_CLT_FACT_RUE1," & _
                                    "CMD_CLT_FACT_RUE2," & _
                                    "CMD_CLT_FACT_CP," & _
                                    "CMD_CLT_FACT_VILLE," & _
                                    "CMD_CLT_FACT_TEL," & _
                                    "CMD_CLT_FACT_FAX," & _
                                    "CMD_CLT_FACT_PORT," & _
                                    "CMD_CLT_FACT_EMAIL," & _
                                    "CMD_CLT_RGLMT_ID," & _
                                    "CMD_CLT_BANQUE," & _
                                    "CMD_CLT_RIB1," & _
                                    "CMD_CLT_RIB2," & _
                                    "CMD_CLT_RIB3," & _
                                    "CMD_CLT_RIB4," & _
                                    "CMD_COM_LIBRE," & _
                                    "CMD_COM_COM," & _
                                    "CMD_COM_LIV," & _
                                    "CMD_COM_FACT," & _
                                    "CMD_TRP_ID," & _
                                    "CMD_TRP_CODE," & _
                                    "CMD_TRP_NOM," & _
                                    "CMD_TRP_RUE1," & _
                                    "CMD_TRP_RUE2," & _
                                    "CMD_TRP_CP," & _
                                    "CMD_TRP_VILLE," & _
                                    "CMD_TRP_TEL," & _
                                    "CMD_TRP_FAX," & _
                                    "CMD_TRP_PORT," & _
                                    "CMD_TRP_EMAIL," & _
                                    "CMD_QTE_COLIS," & _
                                    "CMD_QTE_PAL_PREP," & _
                                    "CMD_QTE_PAL_NONPREP," & _
                                    "CMD_POIDS," & _
                                    "CMD_PU_PAL_PREP," & _
                                    "CMD_PU_PAL_NONPREP," & _
                                    "CMD_MT_TRANSPORT," & _
                                    "CMD_LETTREVOITURE," & _
                                    "CMD_COUT_TRANSPORT," & _
                                    "CMD_REFFACT_TRP," & _
                                    "CMD_CLT_NOM," & _
                                    "CMD_CLT_RS" & _
                                  " ) VALUES ( " & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? , " & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "?" & _
                                    " )"
        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        '        Dim objParam As OleDbParameter


        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString
        m_dbconn.BeginTransaction()
        objCommand.Transaction = m_dbconn.transaction



        CreateParameterP_CMD_FRN_ID(objCommand)
        CreateParameterP_CMD_DATE(objCommand)
        CreateParameterP_CMD_DATE_VALID(objCommand)
        CreateParameterP_CMD_DATE_LIV(objCommand)
        CreateParameterP_CMD_DATE_ENLEV(objCommand)
        CreateParameterP_CMD_REF_LIV(objCommand)
        CreateParameterP_CMD_ETAT(objCommand)
        CreateParameterP_CMD_TOTAL_HT(objCommand)
        CreateParameterP_CMD_TOTAL_TTC(objCommand)
        CreateParameterP_CMD_TYPE(objCommand)
        CreateParameterP_CMD_TYPE_TRANSPORT(objCommand)
        CreateParameterP_CMD_CODE(objCommand)
        CreateParameterP_CMD_CLT_LIV_NOM(objCommand)
        CreateParameterP_CMD_CLT_LIV_RUE1(objCommand)
        CreateParameterP_CMD_CLT_LIV_RUE2(objCommand)
        CreateParameterP_CMD_CLT_LIV_CP(objCommand)
        CreateParameterP_CMD_CLT_LIV_VILLE(objCommand)
        CreateParameterP_CMD_CLT_LIV_TEL(objCommand)
        CreateParameterP_CMD_CLT_LIV_FAX(objCommand)
        CreateParameterP_CMD_CLT_LIV_PORT(objCommand)
        CreateParameterP_CMD_CLT_LIV_EMAIL(objCommand)
        CreateParameterP_CMD_CLT_ADR_IDENT(objCommand)
        CreateParameterP_CMD_CLT_FACT_NOM(objCommand)
        CreateParameterP_CMD_CLT_FACT_RUE1(objCommand)
        CreateParameterP_CMD_CLT_FACT_RUE2(objCommand)
        CreateParameterP_CMD_CLT_FACT_CP(objCommand)
        CreateParameterP_CMD_CLT_FACT_VILLE(objCommand)
        CreateParameterP_CMD_CLT_FACT_TEL(objCommand)
        CreateParameterP_CMD_CLT_FACT_FAX(objCommand)
        CreateParameterP_CMD_CLT_FACT_PORT(objCommand)
        CreateParameterP_CMD_CLT_FACT_EMAIL(objCommand)
        CreateParameterP_CMD_CLT_RGLMT_ID(objCommand)
        CreateParameterP_CMD_CLT_BANQUE(objCommand)
        CreateParameterP_CMD_CLT_RIB1(objCommand)
        CreateParameterP_CMD_CLT_RIB2(objCommand)
        CreateParameterP_CMD_CLT_RIB3(objCommand)
        CreateParameterP_CMD_CLT_RIB4(objCommand)
        CreateParameterP_CMD_COM_LIBRE(objCommand)
        CreateParameterP_CMD_COM_COM(objCommand)
        CreateParameterP_CMD_COM_LIV(objCommand)
        CreateParameterP_CMD_COM_FACT(objCommand)
        CreateParameterP_CMD_TRP_ID(objCommand)
        CreateParameterP_CMD_TRP_CODE(objCommand)
        CreateParameterP_CMD_TRP_NOM(objCommand)
        CreateParameterP_CMD_TRP_RUE1(objCommand)
        CreateParameterP_CMD_TRP_RUE2(objCommand)
        CreateParameterP_CMD_TRP_CP(objCommand)
        CreateParameterP_CMD_TRP_VILLE(objCommand)
        CreateParameterP_CMD_TRP_TEL(objCommand)
        CreateParameterP_CMD_TRP_FAX(objCommand)
        CreateParameterP_CMD_TRP_PORT(objCommand)
        CreateParameterP_CMD_TRP_EMAIL(objCommand)
        CreateParameterP_CMD_QTE_COLIS(objCommand)
        CreateParameterP_CMD_QTE_PAL_PREP(objCommand)
        CreateParameterP_CMD_QTE_PAL_NONPREP(objCommand)
        CreateParameterP_CMD_POIDS(objCommand)
        CreateParameterP_CMD_PU_PAL_PREP(objCommand)
        CreateParameterP_CMD_PU_PAL_NONPREP(objCommand)
        CreateParameterP_CMD_MT_TRANSPORT(objCommand)
        CreateParameterP_CMD_LETTREVOITURE(objCommand)
        CreateParameterP_CMD_COUT_TRANSPORT(objCommand)
        CreateParameterP_CMD_REFFACT_TRP(objCommand)
        CreateParameterP_CMD_CLT_NOM(objCommand)
        CreateParameterP_CMD_CLT_RS(objCommand)
        Try
            objCommand.ExecuteNonQuery()
            objCommand = New OleDbCommand("SELECT MAX(CMD_ID) FROM BONAPPRO", m_dbconn.Connection)
            objCommand.Transaction = m_dbconn.transaction
            objRS = objCommand.ExecuteReader
            If (objRS.Read) Then
                m_id = objRS.GetInt32(0)
                cleanErreur()
                bReturn = True
            Else
                setError("InsertBA", "NO BAID")
                bReturn = False
            End If
            objRS.Close()
            objRS = Nothing

        Catch ex As Exception
            setError("InsertBA", ex.ToString())
            bReturn = False
        End Try
        If bReturn Then
            m_dbconn.transaction.Commit()
        Else
            m_dbconn.transaction.Rollback()
        End If
        '    Debug.Assert(m_id <> 0, "ID=0")
        Debug.Assert(bReturn, "InsertBA: " & getErreur())
        Return bReturn
    End Function 'insertBA
    Protected Function loadBA() As Boolean

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "L'id doit être renseigné")

        Dim sqlString As String = "SELECT " & _
                                    "CMD_FRN_ID," & _
                                    "CMD_DATE," & _
                                    "CMD_DATE_VALID," & _
                                    "CMD_DATE_LIV," & _
                                    "CMD_DATE_ENLEV," & _
                                    "CMD_REF_LIV," & _
                                    "CMD_ETAT," & _
                                    "CMD_TOTAL_HT," & _
                                    "CMD_TOTAL_TTC," & _
                                    "CMD_TYPE," & _
                                    "CMD_TYPE_TRANSPORT," & _
                                    "CMD_CODE," & _
                                    "CMD_CLT_NOM," & _
                                    "CMD_CLT_RS," & _
                                    "CMD_CLT_LIV_NOM," & _
                                    "CMD_CLT_LIV_RUE1," & _
                                    "CMD_CLT_LIV_RUE2," & _
                                    "CMD_CLT_LIV_CP," & _
                                    "CMD_CLT_LIV_VILLE," & _
                                    "CMD_CLT_LIV_TEL," & _
                                    "CMD_CLT_LIV_FAX," & _
                                    "CMD_CLT_LIV_PORT," & _
                                    "CMD_CLT_LIV_EMAIL," & _
                                    "CMD_CLT_ADR_IDENT," & _
                                    "CMD_CLT_FACT_NOM," & _
                                    "CMD_CLT_FACT_RUE1," & _
                                    "CMD_CLT_FACT_RUE2," & _
                                    "CMD_CLT_FACT_CP," & _
                                    "CMD_CLT_FACT_VILLE," & _
                                    "CMD_CLT_FACT_TEL," & _
                                    "CMD_CLT_FACT_FAX," & _
                                    "CMD_CLT_FACT_PORT," & _
                                    "CMD_CLT_FACT_EMAIL," & _
                                    "CMD_CLT_RGLMT_ID," & _
                                    "CMD_CLT_BANQUE," & _
                                    "CMD_CLT_RIB1," & _
                                    "CMD_CLT_RIB2," & _
                                    "CMD_CLT_RIB3," & _
                                    "CMD_CLT_RIB4," & _
                                    "CMD_COM_LIBRE," & _
                                    "CMD_COM_COM," & _
                                    "CMD_COM_LIV," & _
                                    "CMD_COM_FACT," & _
                                    "CMD_TRP_ID," & _
                                    "CMD_TRP_CODE," & _
                                    "CMD_TRP_NOM," & _
                                    "CMD_TRP_RUE1," & _
                                    "CMD_TRP_RUE2," & _
                                    "CMD_TRP_CP," & _
                                    "CMD_TRP_VILLE," & _
                                    "CMD_TRP_TEL," & _
                                    "CMD_TRP_FAX," & _
                                    "CMD_TRP_PORT," & _
                                    "CMD_TRP_EMAIL, " & _
                                    "CMD_QTE_COLIS, " & _
                                    "CMD_QTE_PAL_PREP, " & _
                                    "CMD_QTE_PAL_NONPREP, " & _
                                    "CMD_POIDS, " & _
                                    "CMD_PU_PAL_PREP, " & _
                                    "CMD_PU_PAL_NONPREP, " & _
                                    "CMD_MT_TRANSPORT, " & _
                                    "CMD_LETTREVOITURE, " & _
                                    "CMD_COUT_TRANSPORT, " & _
                                    "CMD_REFFACT_TRP, " & _
                                    "RQ_MODEREGLEMENT.PAR_VALUE " & _
                                    " FROM BONAPPRO  LEFT OUTER JOIN RQ_ModeReglement ON BONAPPRO.CMD_CLT_RGLMT_ID = RQ_ModeReglement.PAR_ID WHERE  " & _
                                   " CMD_ID = ?"

        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        Dim objBA As BonAppro
        '        Dim objParam As OleDbParameter
        Dim bReturn As Boolean
        Dim nId As Integer

        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        'Si le paramètre existe on ne met à jour que sa valeur
        'Autrement on le crée
        CreateParameterP_ID(objCommand)

        Try
            objRS = objCommand.ExecuteReader
            If Not objRS.HasRows Then
                objRS.Close()
                Return False
            End If
            objRS.Read()
            objBA = CType(Me, BonAppro)
            Try
                objBA.code = GetString(objRS, "CMD_CODE")
            Catch ex As InvalidCastException
                objBA.code = ""
            End Try

            Try 'Etat de la Commande
                objBA.setEtat(GetString(objRS, "CMD_ETAT"))
            Catch ex As InvalidCastException
                objBA.setEtat(vncEnums.vncEtatCommande.vncEnCoursSaisie)
            End Try
            Try
                objBA.caracteristiqueTiers.rib1 = GetString(objRS, "CMD_CLT_RIB1")
            Catch ex As InvalidCastException
                objBA.caracteristiqueTiers.rib1 = ""
            End Try

            Try
                objBA.caracteristiqueTiers.rib2 = GetString(objRS, "CMD_CLT_RIB2")
            Catch ex As InvalidCastException
                objBA.caracteristiqueTiers.rib2 = ""
            End Try

            Try
                objBA.caracteristiqueTiers.rib3 = GetString(objRS, "CMD_CLT_RIB3")
            Catch ex As InvalidCastException
                objBA.caracteristiqueTiers.rib3 = ""
            End Try

            Try
                objBA.caracteristiqueTiers.rib4 = GetString(objRS, "CMD_CLT_RIB4")
            Catch ex As InvalidCastException
                objBA.caracteristiqueTiers.rib4 = ""
            End Try

            Try
                objBA.caracteristiqueTiers.banque = GetString(objRS, "CMD_CLT_BANQUE")
            Catch ex As InvalidCastException
                objBA.caracteristiqueTiers.banque = ""
            End Try

            Try
                objBA.caracteristiqueTiers.idModeReglement = GetString(objRS, "CMD_CLT_RGLMT_ID")
            Catch ex As InvalidCastException
                objBA.caracteristiqueTiers.idModeReglement = 0
            End Try
            Try
                objBA.caracteristiqueTiers.libModeReglement = GetString(objRS, "PAR_VALUE")
            Catch ex As InvalidCastException
                objBA.caracteristiqueTiers.libModeReglement = ""
            End Try

            Try 'IDFournisseur
                objBA.setTiers(New Fournisseur)
                nId = GetString(objRS, "CMD_FRN_ID")
                objBA.oTiers.setid(nId)
            Catch ex As InvalidCastException
                objBA.oTiers.setid(0)
            End Try

            Try 'Date de Commande
                objBA.dateCommande = GetString(objRS, "CMD_DATE")
            Catch ex As InvalidCastException
                objBA.dateCommande = DATE_DEFAUT
            End Try

            Try 'Date de Validation
                objBA.dateValidation = GetString(objRS, "CMD_DATE_VALID")
            Catch ex As InvalidCastException
                objBA.dateValidation = DATE_DEFAUT
            End Try

            Try 'Date de Livraison
                objBA.dateLivraison = GetString(objRS, "CMD_DATE_LIV")
            Catch ex As InvalidCastException
                objBA.dateLivraison = DATE_DEFAUT
            End Try

            Try 'Date de Enlevement
                objBA.dateEnlevement = GetString(objRS, "CMD_DATE_ENLEV")
            Catch ex As InvalidCastException
                objBA.dateEnlevement = DATE_DEFAUT
            End Try

            Try 'Reference de Livraison
                objBA.refLivraison = GetString(objRS, "CMD_REF_LIV")
            Catch ex As InvalidCastException
                objBA.refLivraison = ""
            End Try

            Try 'totalHT
                objBA.totalHT = GetString(objRS, "CMD_TOTAL_HT")
            Catch ex As InvalidCastException
                objBA.totalHT = 0
            End Try

            Try 'totalTTC
                objBA.totalTTC = GetString(objRS, "CMD_TOTAL_TTC")
            Catch ex As InvalidCastException
                objBA.totalTTC = 0
            End Try


            Try 'Nom du tiers
                objBA.caracteristiqueTiers.nom = GetString(objRS, "CMD_CLT_NOM")
            Catch ex As InvalidCastException
                objBA.caracteristiqueTiers.nom = ""
            End Try

            Try 'Raison sociale du tiers
                objBA.caracteristiqueTiers.rs = GetString(objRS, "CMD_CLT_RS")
            Catch ex As InvalidCastException
                objBA.caracteristiqueTiers.rs = ""
            End Try

            Try 'Adresse Livraison Nom
                objBA.caracteristiqueTiers.AdresseLivraison.nom = GetString(objRS, "CMD_CLT_LIV_NOM")
            Catch ex As InvalidCastException
                objBA.caracteristiqueTiers.AdresseLivraison.nom = ""
            End Try
            Try 'Adresse Livraison Rue1
                objBA.caracteristiqueTiers.AdresseLivraison.rue1 = GetString(objRS, "CMD_CLT_LIV_RUE1")
            Catch ex As InvalidCastException
                objBA.caracteristiqueTiers.AdresseLivraison.rue1 = ""
            End Try
            Try 'Adresse Livraison rue2
                objBA.caracteristiqueTiers.AdresseLivraison.rue2 = GetString(objRS, "CMD_CLT_LIV_RUE2")
            Catch ex As InvalidCastException
                objBA.caracteristiqueTiers.AdresseLivraison.rue2 = ""
            End Try
            Try 'Adresse Livraison cp
                objBA.caracteristiqueTiers.AdresseLivraison.cp = GetString(objRS, "CMD_CLT_LIV_CP")
            Catch ex As InvalidCastException
                objBA.caracteristiqueTiers.AdresseLivraison.cp = ""
            End Try
            Try 'Adresse Livraison Ville
                objBA.caracteristiqueTiers.AdresseLivraison.ville = GetString(objRS, "CMD_CLT_LIV_VILLE")
            Catch ex As InvalidCastException
                objBA.caracteristiqueTiers.AdresseLivraison.ville = ""
            End Try
            Try 'Adresse Livraison Tel
                objBA.caracteristiqueTiers.AdresseLivraison.tel = GetString(objRS, "CMD_CLT_LIV_TEL")
            Catch ex As InvalidCastException
                objBA.caracteristiqueTiers.AdresseLivraison.tel = ""
            End Try
            Try 'Adresse Livraison fax
                objBA.caracteristiqueTiers.AdresseLivraison.fax = GetString(objRS, "CMD_CLT_LIV_FAX")
            Catch ex As InvalidCastException
                objBA.caracteristiqueTiers.AdresseLivraison.fax = ""
            End Try

            Try 'Adresse Livraison port
                objBA.caracteristiqueTiers.AdresseLivraison.port = GetString(objRS, "CMD_CLT_LIV_PORT")
            Catch ex As InvalidCastException
                objBA.caracteristiqueTiers.AdresseLivraison.port = ""
            End Try
            Try 'Adresse Livraison Email
                objBA.caracteristiqueTiers.AdresseLivraison.Email = GetString(objRS, "CMD_CLT_LIV_EMAIL")
            Catch ex As InvalidCastException
                objBA.caracteristiqueTiers.AdresseLivraison.Email = ""
            End Try

            Try 'Adresses identiques
                objBA.caracteristiqueTiers.bAdressesIdentiques = GetString(objRS, "CMD_CLT_ADR_IDENT")
            Catch ex As InvalidCastException
                objBA.caracteristiqueTiers.bAdressesIdentiques = False
            End Try

            Try 'Adresse Facturation Nom
                objBA.caracteristiqueTiers.AdresseFacturation.nom = GetString(objRS, "CMD_CLT_FACT_NOM")
            Catch ex As InvalidCastException
                objBA.caracteristiqueTiers.AdresseFacturation.nom = ""
            End Try
            Try 'Adresse Facturation Rue1
                objBA.caracteristiqueTiers.AdresseFacturation.rue1 = GetString(objRS, "CMD_CLT_FACT_RUE1")
            Catch ex As InvalidCastException
                objBA.caracteristiqueTiers.AdresseFacturation.rue1 = ""
            End Try
            Try 'Adresse Facturation rue2
                objBA.caracteristiqueTiers.AdresseFacturation.rue2 = GetString(objRS, "CMD_CLT_FACT_RUE2")
            Catch ex As InvalidCastException
                objBA.caracteristiqueTiers.AdresseFacturation.rue2 = ""
            End Try
            Try 'Adresse Facturation cp
                objBA.caracteristiqueTiers.AdresseFacturation.cp = GetString(objRS, "CMD_CLT_FACT_CP")
            Catch ex As InvalidCastException
                objBA.caracteristiqueTiers.AdresseFacturation.cp = ""
            End Try
            Try 'Adresse Facturation Ville
                objBA.caracteristiqueTiers.AdresseFacturation.ville = GetString(objRS, "CMD_CLT_FACT_VILLE")
            Catch ex As InvalidCastException
                objBA.caracteristiqueTiers.AdresseFacturation.ville = ""
            End Try
            Try 'Adresse Facturation Tel
                objBA.caracteristiqueTiers.AdresseFacturation.tel = GetString(objRS, "CMD_CLT_FACT_TEL")
            Catch ex As InvalidCastException
                objBA.caracteristiqueTiers.AdresseFacturation.tel = ""
            End Try
            Try 'Adresse Facturation fax
                objBA.caracteristiqueTiers.AdresseFacturation.fax = GetString(objRS, "CMD_CLT_FACT_FAX")
            Catch ex As InvalidCastException
                objBA.caracteristiqueTiers.AdresseFacturation.fax = ""
            End Try

            Try 'Adresse Facturation port
                objBA.caracteristiqueTiers.AdresseFacturation.port = GetString(objRS, "CMD_CLT_FACT_PORT")
            Catch ex As InvalidCastException
                objBA.caracteristiqueTiers.AdresseFacturation.port = ""
            End Try
            Try 'Adresse Facturation Email
                objBA.caracteristiqueTiers.AdresseFacturation.Email = GetString(objRS, "CMD_CLT_FACT_EMAIL")
            Catch ex As InvalidCastException
                objBA.caracteristiqueTiers.AdresseFacturation.Email = ""
            End Try

            Try 'Id du Transporteur 
                objBA.oTransporteur.setid(GetString(objRS, "CMD_TRP_ID"))
            Catch ex As InvalidCastException
                objBA.oTransporteur.setid(0)
            End Try
            Try 'Adresse Transporteur Code
                objBA.oTransporteur.code = GetString(objRS, "CMD_TRP_CODE")
            Catch ex As InvalidCastException
                objBA.oTransporteur.AdresseLivraison.Code = ""
            End Try
            Try 'Adresse Transporteur Nom
                objBA.oTransporteur.nom = GetString(objRS, "CMD_TRP_NOM")
                objBA.oTransporteur.rs = GetString(objRS, "CMD_TRP_NOM")
                objBA.oTransporteur.AdresseLivraison.nom = GetString(objRS, "CMD_TRP_NOM")
            Catch ex As InvalidCastException
                objBA.oTransporteur.AdresseLivraison.nom = ""
            End Try
            Try 'Adresse Transporteur Rue1
                objBA.oTransporteur.AdresseLivraison.rue1 = GetString(objRS, "CMD_TRP_RUE1")
            Catch ex As InvalidCastException
                objBA.oTransporteur.AdresseLivraison.rue1 = ""
            End Try
            Try 'Adresse Transporteur rue2
                objBA.oTransporteur.AdresseLivraison.rue2 = GetString(objRS, "CMD_TRP_RUE2")
            Catch ex As InvalidCastException
                objBA.oTransporteur.AdresseLivraison.rue2 = ""
            End Try
            Try 'Adresse Transporteur cp
                objBA.oTransporteur.AdresseLivraison.cp = GetString(objRS, "CMD_TRP_CP")
            Catch ex As InvalidCastException
                objBA.oTransporteur.AdresseLivraison.cp = ""
            End Try
            Try 'Adresse Transporteur Ville
                objBA.oTransporteur.AdresseLivraison.ville = GetString(objRS, "CMD_TRP_VILLE")
            Catch ex As InvalidCastException
                objBA.oTransporteur.AdresseLivraison.ville = ""
            End Try
            Try 'Adresse Transporteur Tel
                objBA.oTransporteur.AdresseLivraison.tel = GetString(objRS, "CMD_TRP_TEL")
            Catch ex As InvalidCastException
                objBA.oTransporteur.AdresseLivraison.tel = ""
            End Try
            Try 'Adresse Transporteur fax
                objBA.oTransporteur.AdresseLivraison.fax = GetString(objRS, "CMD_TRP_FAX")
            Catch ex As InvalidCastException
                objBA.oTransporteur.AdresseLivraison.fax = ""
            End Try

            Try 'Adresse Transporteur port
                objBA.oTransporteur.AdresseLivraison.port = GetString(objRS, "CMD_TRP_PORT")
            Catch ex As InvalidCastException
                objBA.oTransporteur.AdresseLivraison.port = ""
            End Try
            Try 'Adresse Transporteur Email
                objBA.oTransporteur.AdresseLivraison.Email = GetString(objRS, "CMD_TRP_EMAIL")
            Catch ex As InvalidCastException
                objBA.oTransporteur.AdresseLivraison.Email = ""
            End Try

            Try 'Commentaires Libre
                objBA.CommLibre.comment = GetString(objRS, "CMD_COM_LIBRE")
            Catch ex As InvalidCastException
                objBA.CommLibre.comment = ""
            End Try

            Try 'Commentaires Commande
                objBA.CommCommande.comment = GetString(objRS, "CMD_COM_COM")
            Catch ex As InvalidCastException
                objBA.CommCommande.comment = ""
            End Try

            Try 'Commentaires Livraison
                objBA.CommLivraison.comment = GetString(objRS, "CMD_COM_LIV")
            Catch ex As InvalidCastException
                objBA.CommLivraison.comment = ""
            End Try

            Try 'Commentaires Facturation
                objBA.CommFacturation.comment = GetString(objRS, "CMD_COM_FACT")
            Catch ex As InvalidCastException
                objBA.CommFacturation.comment = ""
            End Try
            Try 'Qte Colis
                objBA.qteColis = GetString(objRS, "CMD_QTE_COLIS")
            Catch ex As InvalidCastException
                objBA.qteColis = 0
            End Try
            Try 'Qte Palettes Non Preparées
                objBA.qtePalettesNonPreparees = GetString(objRS, "CMD_QTE_PAL_NONPREP")
            Catch ex As InvalidCastException
                objBA.qtePalettesNonPreparees = 0
            End Try
            Try 'Qte Palettes Preparées
                objBA.qtePalettesPreparees = GetString(objRS, "CMD_QTE_PAL_PREP")
            Catch ex As InvalidCastException
                objBA.qtePalettesPreparees = 0
            End Try
            Try 'Poids
                objBA.poids = GetString(objRS, "CMD_POIDS")
            Catch ex As InvalidCastException
                objBA.poids = 0
            End Try
            Try 'PU Palettes Non Preparées
                objBA.puPalettesNonPreparees = GetString(objRS, "CMD_PU_PAL_NONPREP")
            Catch ex As InvalidCastException
                objBA.puPalettesNonPreparees = 0
            End Try
            Try 'PU Palettes Preparées
                objBA.puPalettesPreparees = GetString(objRS, "CMD_PU_PAL_PREP")
            Catch ex As InvalidCastException
                objBA.puPalettesPreparees = 0
            End Try
            Try 'Montant du transport
                objBA.montantTransport = GetString(objRS, "CMD_MT_TRANSPORT")
            Catch ex As InvalidCastException
                objBA.montantTransport = 0
            End Try
            Try 'Lettre Voiture
                objBA.lettreVoiture = GetString(objRS, "CMD_LETTREVOITURE")
            Catch ex As InvalidCastException
                objBA.lettreVoiture = String.Empty
            End Try
            Try 'CoutTranport
                objBA.coutTransport = GetString(objRS, "CMD_COUT_TRANSPORT")
            Catch ex As InvalidCastException
                objBA.coutTransport = 0
            End Try
            Try 'Ref Facture Transporteur
                objBA.refFactTRP = GetString(objRS, "CMD_REFFACT_TRP")
            Catch ex As InvalidCastException
                objBA.refFactTRP = String.Empty
            End Try


            cleanErreur()
            objRS.Close()
            objRS = Nothing
            bReturn = True
        Catch ex As Exception
            setError("LoadBAT", ex.ToString())
            If Not objRS Is Nothing Then
                objRS.Close()
            End If
            bReturn = False
        End Try
            Debug.Assert(bReturn, "LoadBA:" & getErreur())
            Return bReturn
    End Function 'LoadBA
    Protected Function updateBA() As Boolean
        '=======================================================================
        'Fonction : UpdateBA
        'Description : Mise a jour  d'un bon D'appro
        'Retour : Rend Vrai si l'UPDATE s'est correctement effectué
        '=======================================================================
        Dim bReturn As Boolean
        bReturn = False

        Debug.Assert(m_typedonnee = vncEnums.vncTypeDonnee.BA, "Objet de Type BonAppro Requis")
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "ID<>0")

        Dim sqlString As String = "UPDATE BONAPPRO SET " & _
                                    "CMD_FRN_ID = ? , " & _
                                    "CMD_DATE = ? , " & _
                                    "CMD_DATE_VALID = ? , " & _
                                    "CMD_DATE_LIV = ? , " & _
                                    "CMD_DATE_ENLEV = ? , " & _
                                    "CMD_REF_LIV = ? , " & _
                                    "CMD_ETAT = ? , " & _
                                    "CMD_TOTAL_HT= ? ," & _
                                    "CMD_TOTAL_TTC= ? ," & _
                                    "CMD_TYPE= ? ," & _
                                    "CMD_TYPE_TRANSPORT= ? ," & _
                                    "CMD_CODE = ? ," & _
                                    "CMD_CLT_LIV_NOM = ? ," & _
                                    "CMD_CLT_LIV_RUE1= ? ," & _
                                    "CMD_CLT_LIV_RUE2= ? ," & _
                                    "CMD_CLT_LIV_CP= ? ," & _
                                    "CMD_CLT_LIV_VILLE= ? ," & _
                                    "CMD_CLT_LIV_TEL= ? ," & _
                                    "CMD_CLT_LIV_FAX= ? ," & _
                                    "CMD_CLT_LIV_PORT= ? ," & _
                                    "CMD_CLT_LIV_EMAIL= ? ," & _
                                    "CMD_CLT_ADR_IDENT= ? ," & _
                                    "CMD_CLT_FACT_NOM= ? ," & _
                                    "CMD_CLT_FACT_RUE1= ? ," & _
                                    "CMD_CLT_FACT_RUE2= ? ," & _
                                    "CMD_CLT_FACT_CP= ? ," & _
                                    "CMD_CLT_FACT_VILLE= ? ," & _
                                    "CMD_CLT_FACT_TEL= ? ," & _
                                    "CMD_CLT_FACT_FAX= ? ," & _
                                    "CMD_CLT_FACT_PORT= ? ," & _
                                    "CMD_CLT_FACT_EMAIL= ? ," & _
                                    "CMD_CLT_RGLMT_ID= ? ," & _
                                    "CMD_CLT_BANQUE= ? ," & _
                                    "CMD_CLT_RIB1= ? ," & _
                                    "CMD_CLT_RIB2= ? ," & _
                                    "CMD_CLT_RIB3= ? ," & _
                                    "CMD_CLT_RIB4= ? ," & _
                                    "CMD_COM_LIBRE= ? ," & _
                                    "CMD_COM_COM= ? ," & _
                                    "CMD_COM_LIV= ? ," & _
                                    "CMD_COM_FACT= ? ," & _
                                    "CMD_TRP_ID= ? ," & _
                                    "CMD_TRP_CODE= ? ," & _
                                    "CMD_TRP_NOM= ? ," & _
                                    "CMD_TRP_RUE1= ? ," & _
                                    "CMD_TRP_RUE2= ? ," & _
                                    "CMD_TRP_CP= ? ," & _
                                    "CMD_TRP_VILLE= ? ," & _
                                    "CMD_TRP_TEL= ? ," & _
                                    "CMD_TRP_FAX= ? ," & _
                                    "CMD_TRP_PORT= ? ," & _
                                    "CMD_TRP_EMAIL= ? ," & _
                                    "CMD_QTE_COLIS= ? ," & _
                                    "CMD_QTE_PAL_PREP= ? ," & _
                                    "CMD_QTE_PAL_NONPREP= ? ," & _
                                    "CMD_PU_PAL_PREP= ? ," & _
                                    "CMD_PU_PAL_NONPREP= ? ," & _
                                    "CMD_POIDS= ? , " & _
                                    "CMD_MT_TRANSPORT= ? , " & _
                                    "CMD_LETTREVOITURE= ? , " & _
                                    "CMD_COUT_TRANSPORT= ? , " & _
                                    "CMD_REFFACT_TRP= ? , " & _
                                    "CMD_CLT_NOM= ? ," & _
                                    "CMD_CLT_RS= ? " & _
                                    " WHERE CMD_ID = ?"
        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        '        Dim objParam As OleDbParameter


        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        CreateParameterP_CMD_FRN_ID(objCommand)
        CreateParameterP_CMD_DATE(objCommand)
        CreateParameterP_CMD_DATE_VALID(objCommand)
        CreateParameterP_CMD_DATE_LIV(objCommand)
        CreateParameterP_CMD_DATE_ENLEV(objCommand)
        CreateParameterP_CMD_REF_LIV(objCommand)
        CreateParameterP_CMD_ETAT(objCommand)
        CreateParameterP_CMD_TOTAL_HT(objCommand)
        CreateParameterP_CMD_TOTAL_TTC(objCommand)
        CreateParameterP_CMD_TYPE(objCommand)
        CreateParameterP_CMD_TYPE_TRANSPORT(objCommand)
        CreateParameterP_CMD_CODE(objCommand)
        CreateParameterP_CMD_CLT_LIV_NOM(objCommand)
        CreateParameterP_CMD_CLT_LIV_RUE1(objCommand)
        CreateParameterP_CMD_CLT_LIV_RUE2(objCommand)
        CreateParameterP_CMD_CLT_LIV_CP(objCommand)
        CreateParameterP_CMD_CLT_LIV_VILLE(objCommand)
        CreateParameterP_CMD_CLT_LIV_TEL(objCommand)
        CreateParameterP_CMD_CLT_LIV_FAX(objCommand)
        CreateParameterP_CMD_CLT_LIV_PORT(objCommand)
        CreateParameterP_CMD_CLT_LIV_EMAIL(objCommand)
        CreateParameterP_CMD_CLT_ADR_IDENT(objCommand)
        CreateParameterP_CMD_CLT_FACT_NOM(objCommand)
        CreateParameterP_CMD_CLT_FACT_RUE1(objCommand)
        CreateParameterP_CMD_CLT_FACT_RUE2(objCommand)
        CreateParameterP_CMD_CLT_FACT_CP(objCommand)
        CreateParameterP_CMD_CLT_FACT_VILLE(objCommand)
        CreateParameterP_CMD_CLT_FACT_TEL(objCommand)
        CreateParameterP_CMD_CLT_FACT_FAX(objCommand)
        CreateParameterP_CMD_CLT_FACT_PORT(objCommand)
        CreateParameterP_CMD_CLT_FACT_EMAIL(objCommand)
        CreateParameterP_CMD_CLT_RGLMT_ID(objCommand)
        CreateParameterP_CMD_CLT_BANQUE(objCommand)
        CreateParameterP_CMD_CLT_RIB1(objCommand)
        CreateParameterP_CMD_CLT_RIB2(objCommand)
        CreateParameterP_CMD_CLT_RIB3(objCommand)
        CreateParameterP_CMD_CLT_RIB4(objCommand)
        CreateParameterP_CMD_COM_LIBRE(objCommand)
        CreateParameterP_CMD_COM_COM(objCommand)
        CreateParameterP_CMD_COM_LIV(objCommand)
        CreateParameterP_CMD_COM_FACT(objCommand)
        CreateParameterP_CMD_TRP_ID(objCommand)
        CreateParameterP_CMD_TRP_CODE(objCommand)
        CreateParameterP_CMD_TRP_NOM(objCommand)
        CreateParameterP_CMD_TRP_RUE1(objCommand)
        CreateParameterP_CMD_TRP_RUE2(objCommand)
        CreateParameterP_CMD_TRP_CP(objCommand)
        CreateParameterP_CMD_TRP_VILLE(objCommand)
        CreateParameterP_CMD_TRP_TEL(objCommand)
        CreateParameterP_CMD_TRP_FAX(objCommand)
        CreateParameterP_CMD_TRP_PORT(objCommand)
        CreateParameterP_CMD_TRP_EMAIL(objCommand)
        CreateParameterP_CMD_QTE_COLIS(objCommand)
        CreateParameterP_CMD_QTE_PAL_PREP(objCommand)
        CreateParameterP_CMD_QTE_PAL_NONPREP(objCommand)
        CreateParameterP_CMD_PU_PAL_PREP(objCommand)
        CreateParameterP_CMD_PU_PAL_NONPREP(objCommand)
        CreateParameterP_CMD_POIDS(objCommand)
        CreateParameterP_CMD_MT_TRANSPORT(objCommand)
        CreateParameterP_CMD_LETTREVOITURE(objCommand)
        CreateParameterP_CMD_COUT_TRANSPORT(objCommand)
        CreateParameterP_CMD_REFFACT_TRP(objCommand)
        CreateParameterP_CMD_CLT_NOM(objCommand)
        CreateParameterP_CMD_CLT_RS(objCommand)
        CreateParameterP_ID(objCommand)
        Try
            objCommand.ExecuteNonQuery()
            bReturn = True

        Catch ex As Exception
            setError("updateBA", ex.ToString())
            bReturn = False
        End Try

        Debug.Assert(bReturn, "UpdateBA" & getErreur())
        Return bReturn
    End Function 'UpdateBA
    Private Sub CreateParameterP_CMD_FRN_ID(ByVal objCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        Debug.Assert(objCMD.oTiers.id <> 0)
        objCommand.Parameters.AddWithValue("?", objCMD.oTiers.id)
    End Sub

#End Region

    '========================================================================================

#Region "Facture de Commission"
    Protected Function loadFACTCOM() As Boolean

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "L'id doit être renseigné")

        Dim sqlString As String = "SELECT " & _
                                    "FCT_CODE," & _
                                    "FCT_ETAT," & _
                                    "FCT_FRN_ID," & _
                                    "FCT_DATE," & _
                                    "FCT_DATE_STAT," & _
                                    "FCT_TOTAL_HT," & _
                                    "FCT_TOTAL_TTC," & _
                                    "FCT_PERIODE," & _
                                    "FCT_COM_FACT," & _
                                    "FCT_MONTANT_REGLEMENT," & _
                                    "FCT_DATE_REGLEMENT," & _
                                    "FCT_REF_REGLEMENT," & _
                                    "FCT_IDMODEREGLEMENT," & _
                                    "FCT_DECHEANCE," & _
                                    "FCT_BINTERNET" & _
                                    " FROM FACTCOM WHERE " & _
                                   " FCT_ID = ?"

        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        Dim objFACT As FactCom
        '        Dim objParam As OleDbParameter
        Dim bReturn As Boolean

        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        'Si le paramètre existe on ne met à jour que sa valeur
        'Autrement on le crée
        CreateParameterP_ID(objCommand)

        Try
            objRS = objCommand.ExecuteReader
            If Not objRS.HasRows Then
                objRS.Close()
                m_id = 0
                Return False
            End If
            objRS.Read()
            objFACT = CType(Me, FactCom)
            Try
                objFACT.code = GetString(objRS, "FCT_CODE")
            Catch ex As InvalidCastException
                objFACT.code = ""
            End Try
            Try
                objFACT.setEtat(GetString(objRS, "FCT_ETAT"))
            Catch ex As InvalidCastException
                objFACT.setEtat(vncEnums.vncEtatCommande.vncFactComGeneree)
            End Try
            Try
                objFACT.dateCommande = GetString(objRS, "FCT_DATE")
            Catch ex As InvalidCastException
                objFACT.dateCommande = DATE_DEFAUT
            End Try
            Try 'Date Statistique
                objFACT.dateStatistique = GetString(objRS, "FCT_DATE_STAT")
            Catch ex As InvalidCastException
                objFACT.dateStatistique = DATE_DEFAUT
            End Try
            Try 'PERIODE
                objFACT.periode = GetString(objRS, "FCT_PERIODE")
            Catch ex As InvalidCastException
                objFACT.periode = ""
            End Try
            Try 'Total HT
                objFACT.totalHT = GetString(objRS, "FCT_TOTAL_HT")
            Catch ex As InvalidCastException
                objFACT.totalHT = 0
            End Try
            Try 'Total TTC
                objFACT.totalTTC = GetString(objRS, "FCT_TOTAL_TTC")
            Catch ex As InvalidCastException
                objFACT.totalTTC = 0
            End Try
            Try
                objFACT.montantReglement = GetString(objRS, "FCT_MONTANT_REGLEMENT")
            Catch ex As InvalidCastException
                objFACT.montantReglement = 0
            End Try
            Try
                objFACT.dateReglement = GetString(objRS, "FCT_DATE_REGLEMENT")
            Catch ex As InvalidCastException
                objFACT.dateReglement = DATE_DEFAUT
            End Try
            Try
                objFACT.refReglement = GetString(objRS, "FCT_REF_REGLEMENT")
            Catch ex As InvalidCastException
                objFACT.refReglement = ""
            End Try
            Try
                objFACT.bExportInternet = GetValue(objRS, "FCT_BINTERNET")
            Catch ex As InvalidCastException
                objFACT.bExportInternet = False
            End Try
            Try
                objFACT.oTiers.setid(GetString(objRS, "FCT_FRN_ID"))
            Catch ex As InvalidCastException
                objFACT.oTiers.setid(0)
            End Try
            Try
                objFACT.CommFacturation.comment = GetString(objRS, "FCT_COM_FACT")
            Catch ex As InvalidCastException
                objFACT.CommFacturation.comment = ""
            End Try
            Try
                objFACT.idModeReglement = GetString(objRS, "FCT_IDMODEREGLEMENT")
            Catch ex As InvalidCastException
                objFACT.idModeReglement = 0
            End Try
            'Chargement de la date d'échéance après le mode de reglement pour éviter le recalcul
            Try
                objFACT.dEcheance = GetString(objRS, "FCT_DECHEANCE")
            Catch ex As InvalidCastException
                objFACT.dEcheance = DATE_DEFAUT
            End Try

            cleanErreur()
            bReturn = True
        Catch ex As Exception
            setError("LoadFACTCOM", ex.ToString())
            bReturn = False
        End Try
        If Not objRS Is Nothing Then
            objRS.Close()
        End If
        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'LoadFACTCOM
    Protected Function LoadcolLgFactCom() As Boolean
        '================================================================================
        'Fonction : LoadcolLgFactCom
        'Description : Chargement de la liste des Sous Commandes d'une Facture de commission
        '               Parcours de la table des SOUSCOMMANDES et pour Chaque sous-commande
        '               Création d'une ligne de commande 
        '================================================================================
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "Id <> 0")
        Debug.Assert(Me.GetType().Name.Equals("FactCom"))

        Dim bReturn As Boolean
        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        Dim objCMD As FactCom
        Dim nid As Integer
        Dim objSCMD As SousCommande
        Dim oCol As Collection = New Collection

        Dim sqlString As String = "SELECT " & _
                                    "SCMD_ID, SCMD_FRN_ID" & _
                                  " FROM SOUSCOMMANDE" & _
                                  " WHERE SCMD_FACT_ID = ? "

        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        'Paramétre ClientID 
        CreateParameterP_ID(objCommand)

        Try
            objCMD = CType(Me, FactCom)
            objRS = objCommand.ExecuteReader()
            While objRS.Read()
                Try
                    nid = GetString(objRS, "SCMD_ID")

                    objSCMD = New SousCommande
                    objSCMD.setid(nid)
                    objSCMD.oFournisseur.setid(getInteger(objRS, "SCMD_FRN_ID"))
                    oCol.Add(objSCMD)
                Catch ex As InvalidCastException
                    bReturn = False
                    Exit While
                End Try
            End While
            objRS.Close()
            'Chargement des sous commandes et creation d'une ligne de facture de commision
            For Each objSCMD In oCol
                objSCMD.load()
                objCMD.AjouteLigneFactCom(objSCMD, False)
            Next
            bReturn = True
        Catch ex As Exception
            setError("LoadcolLgFactCom", ex.ToString())
            bReturn = False
        End Try

        Debug.Assert(bReturn, "LoadcolSousCommande" & getErreur())
        Return bReturn

    End Function ' LoadColLGFACTCOM


    Protected Function insertFACTCOM() As Boolean
        '=======================================================================
        'Fonction : insertFACTCOM
        'Description : Insertion d'une Facture de commission
        'Retour : Rend Vrai si l'INSERT s'est correctement effectué
        '=======================================================================
        Dim bReturn As Boolean
        Dim objFACT As FactCom

        bReturn = False

        Debug.Assert(Me.GetType().Name.Equals("FactCom"), "Objet de Type 'FactCom' Requis")
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id = 0, "ID=0")
        objFACT = CType(Me, FactCom)

        Dim sqlString As String = "INSERT INTO FACTCOM( " & _
                                    "FCT_CODE," & _
                                    "FCT_ETAT," & _
                                    "FCT_FRN_ID," & _
                                    "FCT_DATE," & _
                                    "FCT_DATE_STAT," & _
                                    "FCT_TOTAL_HT," & _
                                    "FCT_TOTAL_TTC," & _
                                    "FCT_PERIODE," & _
                                    "FCT_COM_FACT," & _
                                    "FCT_MONTANT_REGLEMENT," & _
                                    "FCT_DATE_REGLEMENT," & _
                                    "FCT_REF_REGLEMENT," & _
                                    "FCT_IDMODEREGLEMENT," & _
                                    "FCT_DECHEANCE," & _
                                    "FCT_BINTERNET" & _
                                  " ) VALUES ( " & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? " & _
                                    " )"
        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        '        Dim objParam As OleDbParameter


        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString
        m_dbconn.BeginTransaction()
        objCommand.Transaction = m_dbconn.transaction


        CreateParameterP_FCT_CODE(objCommand)
        CreateParameterP_FCT_ETAT(objCommand)
        CreateParameterP_FCT_FRN_ID(objCommand)
        CreateParameterP_FCT_DATE(objCommand)
        CreateParameterP_FCT_DATE_STAT(objCommand)
        CreateParameterP_FCT_TOTAL_HT(objCommand)
        CreateParameterP_FCT_TOTAL_TTC(objCommand)
        CreateParameterP_FCT_PERIODE(objCommand)
        CreateParameterP_FCT_COM_FACT(objCommand)
        CreateParameterP_FCT_MONTANT_REGLEMENT(objCommand)
        CreateParameterP_FCT_DATE_REGLEMENT(objCommand)
        CreateParameterP_FCT_REF_REGLEMENT(objCommand)
        CreateParameterP_FCT_IDMODEREGLEMENT(objCommand)
        CreateParameterP_FCT_DECHEANCE(objCommand)
        CreateParameterP_FCT_BINTERNET(objCommand)
        Try
            objCommand.ExecuteNonQuery()
            objCommand.CommandText = "SELECT MAX(FCT_ID) FROM FACTCOM"
            objCommand.Transaction = m_dbconn.transaction
            objRS = objCommand.ExecuteReader
            If (objRS.Read) Then
                m_id = objRS.GetInt32(0)
                cleanErreur()
                bReturn = True
            Else
                bReturn = False
            End If
            objRS.Close()
            bReturn = True

        Catch ex As Exception
            setError("InsertFACTCOM", ex.ToString())
            bReturn = False
        End Try

        If bReturn Then
            m_dbconn.transaction.Commit()
        Else
            m_dbconn.transaction.Rollback()
        End If
        '    Debug.Assert(m_id <> 0, "ID=0")
        Debug.Assert(bReturn, "InsertFACTCOM: " & getErreur())
        Return bReturn
    End Function 'insertFACTCOM
    Protected Function UpdateFACTCOM() As Boolean
        '=======================================================================
        'Fonction : UpdateFACTCOM
        'Description : Mise à jour  d'une Facture de commission
        'Retour : Rend Vrai si l'UPDATE s'est correctement effectué
        '=======================================================================
        Dim bReturn As Boolean
        Dim objSCMD As FactCom
        bReturn = False

        Debug.Assert(Me.GetType().Name.Equals("FactCom"), "Objet de Type FactCom Requis")
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "ID<>0")
        objSCMD = CType(Me, FactCom)

        Dim sqlString As String = "UPDATE FACTCOM SET " & _
                                    "FCT_CODE = ? ," & _
                                    "FCT_ETAT = ? ," & _
                                    "FCT_FRN_ID = ? ," & _
                                    "FCT_DATE = ? ," & _
                                    "FCT_DATE_STAT = ? ," & _
                                    "FCT_TOTAL_HT = ? ," & _
                                    "FCT_TOTAL_TTC = ? ," & _
                                    "FCT_PERIODE = ? ," & _
                                    "FCT_COM_FACT = ? ," & _
                                    "FCT_MONTANT_REGLEMENT = ? ," & _
                                    "FCT_DATE_REGLEMENT = ? ," & _
                                    "FCT_REF_REGLEMENT = ? , " & _
                                    "FCT_IDMODEREGLEMENT = ? , " & _
                                    "FCT_DECHEANCE = ? , " & _
                                    "FCT_BINTERNET = ? " & _
                                    " WHERE FCT_ID = ?"

        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        '        Dim objParam As OleDbParameter



        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString

        m_dbconn.BeginTransaction()
        objCommand.Transaction = m_dbconn.transaction



        CreateParameterP_FCT_CODE(objCommand)
        CreateParameterP_FCT_ETAT(objCommand)
        CreateParameterP_FCT_FRN_ID(objCommand)
        CreateParameterP_FCT_DATE(objCommand)
        CreateParameterP_FCT_DATE_STAT(objCommand)
        CreateParameterP_FCT_TOTAL_HT(objCommand)
        CreateParameterP_FCT_TOTAL_TTC(objCommand)
        CreateParameterP_FCT_PERIODE(objCommand)
        CreateParameterP_FCT_COM_FACT(objCommand)
        CreateParameterP_FCT_MONTANT_REGLEMENT(objCommand)
        CreateParameterP_FCT_DATE_REGLEMENT(objCommand)
        CreateParameterP_FCT_REF_REGLEMENT(objCommand)
        CreateParameterP_FCT_IDMODEREGLEMENT(objCommand)
        CreateParameterP_FCT_DECHEANCE(objCommand)
        CreateParameterP_FCT_BINTERNET(objCommand)
        CreateParameterP_ID(objCommand)
        Try
            objCommand.ExecuteNonQuery()
            m_dbconn.transaction.Commit()
            bReturn = True

        Catch ex As Exception
            setError("UpdateFACTCOM", ex.ToString())
            m_dbconn.transaction.Rollback()
            bReturn = False
        End Try

        Debug.Assert(bReturn, "UpdateFACTCOM" & getErreur())
        Return bReturn
    End Function 'UpdateFACTCOM
    Protected Function deleteFACTCOM() As Boolean
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "Id <> 0")
        Debug.Assert(Me.GetType().Name.Equals("FactCom"))


        Dim sqlString As String = "DELETE FROM FACTCOM WHERE FCT_ID=? "
        Dim objCommand As OleDbCommand
        Dim objFACT As FactCom
        '        Dim objParam As OleDbParameter

        objFACT = CType(Me, FactCom)
        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        CreateParameterP_ID(objCommand)
        Try
            objCommand.ExecuteNonQuery()
            m_id = 0
            objFACT.resetCode()
            Return True

        Catch ex As Exception
            setError("DeleteFACTCOM", ex.ToString())
            Return False
        End Try
    End Function 'DeleteFACTCOM
    Protected Overloads Shared Function ListeFACTCOM(ByVal strCode As String, ByVal strRSFournisseur As String, ByVal pEtat As vncEtatCommande) As Collection
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Dim colReturn As New Collection
        '        Dim objParam As OleDbParameter
        Dim sqlString As String = "SELECT FCT_ID, FCT_CODE, FCT_DATE, FRN_ID , FRN_CODE, FRN_RS, FCT_ETAT, FCT_TOTAL_HT , FCT_TOTAL_TTC, FCT_BINTERNET " & _
                                    "FROM FACTCOM , FOURNISSEUR "
        Dim strWhere As String = " FACTCOM.FCT_FRN_ID = FOURNISSEUR.FRN_ID "
        Dim objCommand As OleDbCommand
        Dim objFACT As FactCom
        Dim objRS As OleDbDataReader = Nothing
        Dim objParam As OleDbParameter
        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection

        If strCode <> "" Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "FCT_CODE LIKE ?"
            objParam = objCommand.Parameters.AddWithValue("?", strCode)

        End If

        If strRSFournisseur <> "" Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "FOURNISSEUR.FRN_RS LIKE ?"
            objParam = objCommand.Parameters.AddWithValue("?", strRSFournisseur)

        End If

        If pEtat <> vncEnums.vncEtatCommande.vncRien Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "FCT_ETAT = ?"
            objParam = objCommand.Parameters.AddWithValue("?", pEtat)

        End If


        If strWhere <> "" Then
            sqlString = sqlString & " WHERE " & strWhere
        End If
        sqlString = sqlString & " ORDER BY FCT_DATE DESC, FRN_CODE ASC"
        'sqlString = sqlString & " ORDER BY FRN_CODE ASC"
        objCommand.CommandText = sqlString



        Try
            objRS = objCommand.ExecuteReader
            While (objRS.Read())
                objFACT = New FactCom
                objFACT.m_bResume = True ' C'est un objet Résumé
                Try
                    objFACT.m_id = GetString(objRS, "FCT_ID")
                Catch ex As InvalidCastException
                    objFACT.m_id = 0
                End Try
                Try
                    objFACT.code = GetString(objRS, "FCT_CODE")
                Catch ex As InvalidCastException
                    objFACT.code = ""
                End Try
                Try
                    objFACT.oTiers.setid(GetString(objRS, "FRN_ID"))
                Catch ex As InvalidCastException
                    objFACT.oTiers.setid(0)
                End Try
                Try
                    objFACT.oTiers.code = GetString(objRS, "FRN_CODE")
                Catch ex As InvalidCastException
                    objFACT.oTiers.code = ""
                End Try
                Try
                    objFACT.oTiers.rs = GetString(objRS, "FRN_RS")
                Catch ex As InvalidCastException
                    objFACT.oTiers.nom = ""
                End Try

                Try
                    objFACT.dateCommande = GetString(objRS, "FCT_DATE")
                Catch ex As InvalidCastException
                    objFACT.dateCommande = Now()
                End Try

                Try
                    objFACT.setEtat(GetString(objRS, "FCT_ETAT"))
                Catch ex As InvalidCastException
                End Try
                Try
                    objFACT.totalHT = GetString(objRS, "FCT_TOTAL_HT")
                Catch ex As InvalidCastException
                    objFACT.totalHT = 0
                End Try
                Try
                    objFACT.totalTTC = GetString(objRS, "FCT_TOTAL_TTC")
                Catch ex As InvalidCastException
                    objFACT.totalTTC = 0
                End Try
                Try
                    objFACT.bExportInternet = GetValue(objRS, "FCT_BINTERNET")
                Catch ex As InvalidCastException
                    objFACT.bExportInternet = False
                End Try
                objFACT.resetBooleans()
                colReturn.Add(objFACT, CStr(objFACT.code))
            End While
            objRS.Close()
            Return colReturn
        Catch ex As Exception
            setError("ListeFACTCOM", ex.ToString())
            Return Nothing
        End Try
    End Function 'ListeFACTCOM

    ''' <summary>
    ''' Liste des facture de commission non reglee
    ''' </summary>
    ''' <param name="strCode"></param>
    ''' <param name="strRSFournisseur"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Overloads Shared Function ListeFACTCOMNonReglee(ByVal strCode As String, ByVal strRSFournisseur As String) As Collection
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Dim colReturn As New Collection
        '        Dim objParam As OleDbParameter
        Dim sqlString As String = "SELECT FCT_ID, FCT_CODE, FCT_DATE, FRN_ID , FRN_CODE, FRN_RS, FCT_TOTAL_HT , FCT_TOTAL_TTC, FCT_BINTERNET " & _
                                    "FROM RQ_FACTURES, FACTCOM , FOURNISSEUR "
        Dim strWhere As String = " FACTCOM.FCT_FRN_ID = FOURNISSEUR.FRN_ID AND RQ_FACTURES.FACT_ID = FCT_ID AND RQ_FACTURES.FACT_TYPEFACT = " & vncEnums.vncTypeDonnee.FACTCOMM & _
        " AND (RQ_FACTURES.FACT_TOTAL_REGLEMENT IS NULL OR RQ_FACTURES.FACT_TOTAL_REGLEMENT < RQ_FACTURES.FACT_TOTAL_TTC)"
        Dim objCommand As OleDbCommand
        Dim objFACT As FactCom
        Dim objRS As OleDbDataReader = Nothing
        Dim objParam As OleDbParameter
        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection

        If strCode <> "" Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "FCT_CODE LIKE ?"
            objParam = objCommand.Parameters.AddWithValue("?", strCode)

        End If

        If strRSFournisseur <> "" Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "FOURNISSEUR.FRN_RS LIKE ?"
            objParam = objCommand.Parameters.AddWithValue("?", strRSFournisseur)

        End If


        If strWhere <> "" Then
            sqlString = sqlString & " WHERE " & strWhere
        End If
        sqlString = sqlString & " ORDER BY FCT_DATE DESC, FRN_CODE ASC"
        'sqlString = sqlString & " ORDER BY FRN_CODE ASC"
        objCommand.CommandText = sqlString



        Try
            objRS = objCommand.ExecuteReader
            While (objRS.Read())
                objFACT = New FactCom
                objFACT.m_bResume = True ' C'est un objet Résumé
                Try
                    objFACT.m_id = GetString(objRS, "FCT_ID")
                Catch ex As InvalidCastException
                    objFACT.m_id = 0
                End Try
                Try
                    objFACT.code = GetString(objRS, "FCT_CODE")
                Catch ex As InvalidCastException
                    objFACT.code = ""
                End Try
                Try
                    objFACT.oTiers.setid(GetString(objRS, "FRN_ID"))
                Catch ex As InvalidCastException
                    objFACT.oTiers.setid(0)
                End Try
                Try
                    objFACT.oTiers.code = GetString(objRS, "FRN_CODE")
                Catch ex As InvalidCastException
                    objFACT.oTiers.code = ""
                End Try
                Try
                    objFACT.oTiers.rs = GetString(objRS, "FRN_RS")
                Catch ex As InvalidCastException
                    objFACT.oTiers.nom = ""
                End Try

                Try
                    objFACT.dateCommande = GetString(objRS, "FCT_DATE")
                Catch ex As InvalidCastException
                    objFACT.dateCommande = Now()
                End Try

                Try
                    objFACT.totalHT = GetString(objRS, "FCT_TOTAL_HT")
                Catch ex As InvalidCastException
                    objFACT.totalHT = 0
                End Try
                Try
                    objFACT.totalTTC = GetString(objRS, "FCT_TOTAL_TTC")
                Catch ex As InvalidCastException
                    objFACT.totalTTC = 0
                End Try
                Try
                    objFACT.bExportInternet = GetValue(objRS, "FCT_BINTERNET")
                Catch ex As InvalidCastException
                    objFACT.bExportInternet = False
                End Try
                objFACT.resetBooleans()
                colReturn.Add(objFACT, CStr(objFACT.code))
            End While
            objRS.Close()
            Return colReturn
        Catch ex As Exception
            setError("ListeFACTCOMNonReglee", ex.ToString())
            Return Nothing
        End Try
    End Function 'ListeFACTCOMNonReglee

#Region "Paramètres FACTCOM"
    Private Sub CreateParameterP_FCT_CODE(ByVal objCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", objCMD.code)

    End Sub
    Private Sub CreateParameterP_FCT_ETAT(ByVal objCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", objCMD.etat.codeEtat)
    End Sub
    Private Sub CreateParameterP_FCT_FRN_ID(ByVal objCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", objCMD.oTiers.id)

    End Sub
    Private Sub CreateParameterP_FCT_DATE(ByVal objCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", objCMD.dateCommande.ToShortDateString)

    End Sub
    Private Sub CreateParameterP_FCT_DATE_STAT(ByVal objCommand As OleDbCommand)
        Dim objCMD As FactCom
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", objCMD.dateStatistique.ToShortDateString)

    End Sub
    Private Sub CreateParameterP_FCT_TOTAL_HT(ByVal objCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", CDbl(objCMD.totalHT))
    End Sub

    Private Sub CreateParameterP_FCT_TOTAL_TTC(ByVal objCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", CDbl(objCMD.totalTTC))
    End Sub
    Private Sub CreateParameterP_FCT_COM_FACT(ByVal objCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", objCMD.CommFacturation.comment)
    End Sub
    Private Sub CreateParameterP_FCT_PERIODE(ByVal objCommand As OleDbCommand)
        Dim objCMD As FactCom
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCMD.periode, 50))
    End Sub
    Private Sub CreateParameterP_FCT_MONTANT_REGLEMENT(ByVal objCommand As OleDbCommand)
        Dim objCMD As FactCom
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", CDbl(objCMD.montantReglement))

    End Sub
    Private Sub CreateParameterP_FCT_DATE_REGLEMENT(ByVal objCommand As OleDbCommand)
        Dim objCMD As FactCom
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", objCMD.dateReglement)

    End Sub
    Private Sub CreateParameterP_FCT_REF_REGLEMENT(ByVal objCommand As OleDbCommand)
        Dim objCMD As FactCom
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCMD.refReglement, 50))
    End Sub
    Private Sub CreateParameterP_FCT_IDMODEREGLEMENT(ByVal objCommand As OleDbCommand)
        Dim objCMD As FactCom
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", objCMD.idModeReglement)
    End Sub
    Private Sub CreateParameterP_FCT_DECHEANCE(ByVal objCommand As OleDbCommand)
        Dim objCMD As FactCom
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", objCMD.dEcheance)
    End Sub
    Private Sub CreateParameterP_FCT_BINTERNET(ByVal objCommand As OleDbCommand)
        Dim objCMD As FactCom
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", objCMD.bExportInternet)
    End Sub
#End Region
#End Region
#Region "FactTRP"
    Protected Overloads Shared Function ListeFACTTRP(ByVal strCode As String, ByVal strRSClient As String, ByVal pEtat As vncEtatCommande) As Collection
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Dim colReturn As New Collection
        '        Dim objParam As OleDbParameter
        Dim sqlString As String = "SELECT FTRP_ID, FTRP_CODE, FTRP_DATE, FTRP_CLT_ID , CLT_CODE, CLT_RS, FTRP_ETAT, FTRP_TOTAL_HT, FTRP_TOTAL_TTC, FTRP_BINTERNET " & _
                                    "FROM FACTTRP , CLIENT "
        Dim strWhere As String = " FACTTRP.FTRP_CLT_ID = CLIENT.CLT_ID "
        Dim objCommand As OleDbCommand
        Dim objFACT As FactTRP
        Dim objRS As OleDbDataReader = Nothing
        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        Dim bReturn As Boolean
        Dim objParam As OleDbParameter

        If strCode <> "" Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "FTRP_CODE LIKE ?"
            objParam = objCommand.Parameters.AddWithValue("?", strCode)

        End If

        If strRSClient <> "" Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "CLIENT.CLT_RS LIKE ?"
            objParam = objCommand.Parameters.AddWithValue("?", strRSClient)

        End If

        If pEtat <> vncEnums.vncEtatCommande.vncRien Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "FTRP_ETAT = ?"
            objParam = objCommand.Parameters.AddWithValue("?", pEtat)

        End If


        If strWhere <> "" Then
            sqlString = sqlString & " WHERE " & strWhere
        End If
        sqlString = sqlString & " ORDER BY FTRP_DATE DESC, CLT_CODE ASC"
        'sqlString = sqlString & " ORDER BY FRN_CODE ASC"
        objCommand.CommandText = sqlString



        Try
            objRS = objCommand.ExecuteReader
            While (objRS.Read())
                objFACT = New FactTRP
                objFACT.m_bResume = True ' C'est un objet Résumé
                Try
                    objFACT.m_id = GetString(objRS, "FTRP_ID")
                Catch ex As InvalidCastException
                    objFACT.m_id = 0
                End Try
                Try
                    objFACT.code = GetString(objRS, "FTRP_CODE")
                Catch ex As InvalidCastException
                    objFACT.code = ""
                End Try
                Try
                    objFACT.oTiers.setid(GetString(objRS, "FTRP_CLT_ID"))
                Catch ex As InvalidCastException
                    objFACT.oTiers.setid(0)
                End Try
                Try
                    objFACT.oTiers.code = GetString(objRS, "CLT_CODE")
                Catch ex As InvalidCastException
                    objFACT.oTiers.code = ""
                End Try
                Try
                    objFACT.oTiers.rs = GetString(objRS, "CLT_RS")
                Catch ex As InvalidCastException
                    objFACT.oTiers.nom = ""
                End Try

                Try
                    objFACT.dateCommande = GetString(objRS, "FTRP_DATE")
                Catch ex As InvalidCastException
                    objFACT.dateCommande = Now()
                End Try

                Try
                    objFACT.setEtat(GetString(objRS, "FTRP_ETAT"))
                Catch ex As InvalidCastException
                End Try
                Try
                    objFACT.totalHT = GetString(objRS, "FTRP_TOTAL_HT")
                Catch ex As InvalidCastException
                    objFACT.totalHT = 0
                End Try

                Try
                    objFACT.totalTTC = GetString(objRS, "FTRP_TOTAL_TTC")
                Catch ex As InvalidCastException
                    objFACT.totalTTC = 0
                End Try
                Try
                    objFACT.bExportInternet = GetValue(objRS, "FTRP_BINTERNET")
                Catch ex As InvalidCastException
                    objFACT.bExportInternet = False
                End Try
                objFACT.resetBooleans()
                colReturn.Add(objFACT, CStr(objFACT.code))
            End While
            objRS.Close()
            objRS = Nothing
            bReturn = True
        Catch ex As Exception
            setError("ListeFACTTRP", sqlString & vbCrLf & ex.ToString())
            bReturn = False
            colReturn = Nothing
        End Try

        Debug.Assert(bReturn, getErreur())

        Return colReturn

    End Function 'ListeFACTTRP
    ''' <summary>
    ''' Liste des facture de colisage non reglee
    ''' </summary>
    ''' <param name="strCode"></param>
    ''' <param name="strRSClient"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Overloads Shared Function ListeFACTTRPNonReglee(ByVal strCode As String, ByVal strRSClient As String) As Collection
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Dim colReturn As New Collection
        '        Dim objParam As OleDbParameter
        Dim sqlString As String = "SELECT FTRP_ID, FTRP_CODE, FTRP_DATE, CLT_ID , CLT_CODE, CLT_RS, FTRP_TOTAL_HT , FTRP_TOTAL_TTC, FTRP_BINTERNET " & _
                                    "FROM RQ_FACTURES, FACTTRP , CLIENT"
        Dim strWhere As String = " FACTTRP.FTRP_CLT_ID = CLIENT.CLT_ID AND RQ_FACTURES.FACT_ID = FTRP_ID AND RQ_FACTURES.FACT_TYPEFACT = " & vncEnums.vncTypeDonnee.FACTTRP & _
        " AND (RQ_FACTURES.FACT_TOTAL_REGLEMENT IS NULL OR RQ_FACTURES.FACT_TOTAL_REGLEMENT < RQ_FACTURES.FACT_TOTAL_TTC)"
        Dim objCommand As OleDbCommand
        Dim objFACT As FactTRP
        Dim objRS As OleDbDataReader = Nothing
        Dim objParam As OleDbParameter
        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection

        If strCode <> "" Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "FTRP_CODE LIKE ?"
            objParam = objCommand.Parameters.AddWithValue("?", strCode)

        End If

        If strRSClient <> "" Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "CLIENT.CLT_RS LIKE ?"
            objParam = objCommand.Parameters.AddWithValue("?", strRSClient)

        End If


        If strWhere <> "" Then
            sqlString = sqlString & " WHERE " & strWhere
        End If
        sqlString = sqlString & " ORDER BY FTRP_DATE DESC, CLT_CODE ASC"
        'sqlString = sqlString & " ORDER BY FRN_CODE ASC"
        objCommand.CommandText = sqlString



        Try
            objRS = objCommand.ExecuteReader
            While (objRS.Read())
                objFACT = New FactTRP
                objFACT.m_bResume = True ' C'est un objet Résumé
                Try
                    objFACT.m_id = GetString(objRS, "FTRP_ID")
                Catch ex As InvalidCastException
                    objFACT.m_id = 0
                End Try
                Try
                    objFACT.code = GetString(objRS, "FTRP_CODE")
                Catch ex As InvalidCastException
                    objFACT.code = ""
                End Try
                Try
                    objFACT.oTiers.setid(GetString(objRS, "CLT_ID"))
                Catch ex As InvalidCastException
                    objFACT.oTiers.setid(0)
                End Try
                Try
                    objFACT.oTiers.code = GetString(objRS, "CLT_CODE")
                Catch ex As InvalidCastException
                    objFACT.oTiers.code = ""
                End Try
                Try
                    objFACT.oTiers.rs = GetString(objRS, "CLT_RS")
                Catch ex As InvalidCastException
                    objFACT.oTiers.nom = ""
                End Try

                Try
                    objFACT.dateCommande = GetString(objRS, "FTRP_DATE")
                Catch ex As InvalidCastException
                    objFACT.dateCommande = Now()
                End Try

                Try
                    objFACT.totalHT = GetString(objRS, "FTRP_TOTAL_HT")
                Catch ex As InvalidCastException
                    objFACT.totalHT = 0
                End Try
                Try
                    objFACT.totalTTC = GetString(objRS, "FTRP_TOTAL_TTC")
                Catch ex As InvalidCastException
                    objFACT.totalTTC = 0
                End Try
                Try
                    objFACT.bExportInternet = GetValue(objRS, "FTRP_BINTERNET")
                Catch ex As InvalidCastException
                    objFACT.bExportInternet = False
                End Try
                objFACT.resetBooleans()
                colReturn.Add(objFACT, CStr(objFACT.code))
            End While
            objRS.Close()
            Return colReturn
        Catch ex As Exception
            setError("ListeFACTCOMNonReglee", ex.ToString())
            Return Nothing
        End Try
    End Function 'ListeFACTTRPNonReglee

    Protected Shared Function ListeFACTTRPEtat(Optional ByVal pddeb As Date = DATE_DEFAUT, Optional ByVal pdfin As Date = DATE_DEFAUT, Optional ByVal pCodeClient As String = "", Optional ByVal pEtat As vncEtatCommande = vncEnums.vncEtatCommande.vncRien) As Collection
        '============================================================================
        'Function : ListeFACTTRPEtat
        'Description : Rend une liste de Facture de transport en fonction leur état
        '============================================================================
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Dim colReturn As New Collection
        Dim colTemp As New Collection
        '        Dim objParam As OleDbParameter
        Dim sqlString As String = "SELECT " & _
                                    "FTRP_ID, CLT_ID " & _
                                  " FROM FACTTRP, CLIENT "

        Dim strWhere As String = " FACTTRP.FTRP_CLT_ID = CLIENT.CLT_ID "
        Dim objCommand As OleDbCommand
        Dim objFCT As FactTRP
        Dim objRS As OleDbDataReader = Nothing
        Dim strId As String
        Dim objParam As OleDbParameter
        Dim nId As Long




        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection


        If pddeb <> DATE_DEFAUT Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & " FTRP_DATE >=  ?"
            objParam = objCommand.Parameters.AddWithValue("?", pddeb)

        End If
        If pdfin <> DATE_DEFAUT Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & " FTRP_DATE <=  ?"
            objParam = objCommand.Parameters.AddWithValue("?", pdfin)

        End If
        If pCodeClient <> "" Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "CLIENT.CLT_CODE LIKE ?"
            objParam = objCommand.Parameters.AddWithValue("?", pCodeClient)

        End If

        If pEtat <> vncEnums.vncEtatCommande.vncRien Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "FTRP_ETAT = ?"
            objParam = objCommand.Parameters.AddWithValue("?", pEtat)

        End If



        If strWhere <> "" Then
            sqlString = sqlString & "WHERE " & strWhere
        End If
        sqlString = sqlString & " ORDER BY CLIENT.CLT_CODE ASC"
        objCommand.CommandText = sqlString



        Try
            objRS = objCommand.ExecuteReader
            While objRS.Read()
                Try
                    strId = GetString(objRS, "FTRP_ID")
                    colTemp.Add(strId)
                Catch ex As InvalidCastException
                    colReturn = Nothing
                    Exit While
                End Try
            End While
            objRS.Close()
            objRS = Nothing
            For Each strId In colTemp
                nId = CLng(strId)
                objFCT = FactTRP.createandload(nId)
                objFCT.resetBooleans()

                If objFCT.id <> 0 Then
                    colReturn.Add(objFCT, objFCT.code)
                End If
            Next
            Return colReturn
        Catch ex As Exception
            setError("ListeFACTTRPEtat", ex.ToString() & sqlString)
            colReturn = Nothing
        End Try
        Debug.Assert(Not colReturn Is Nothing, "ListeFACTTRPEtat: colReturn is nothing" & FactTRP.getErreur())
        Return colReturn
    End Function 'ListeFactTRPEtat

    Protected Function insertFACTTRP() As Boolean
        '=======================================================================
        'Fonction : insertFACTTRP
        'Description : Insertion d'une Facture de Transport
        'Retour : Rend Vrai si l'INSERT s'est correctement effectué
        '=======================================================================
        Dim bReturn As Boolean
        Dim objFACT As FactTRP
        bReturn = False

        Debug.Assert(Me.GetType().Name.Equals("FactTRP"), "Objet de Type 'FactTRP' Requis")
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id = 0, "ID=0")
        objFACT = CType(Me, FactTRP)

        Dim sqlString As String = "INSERT INTO FACTTRP( " & _
                                    "FTRP_CODE," & _
                                    "FTRP_ETAT," & _
                                    "FTRP_CLT_ID," & _
                                    "FTRP_DATE," & _
                                    "FTRP_TOTAL_TAXES," & _
                                    "FTRP_TOTAL_HT," & _
                                    "FTRP_TOTAL_TTC," & _
                                    "FTRP_PERIODE," & _
                                    "FTRP_COM_FACT," & _
                                    "FTRP_MONTANT_REGLEMENT," & _
                                    "FTRP_DATE_REGLEMENT," & _
                                    "FTRP_REF_REGLEMENT," & _
                                    "FTRP_IDMODEREGLEMENT," & _
                                    "FTRP_DECHEANCE," & _
                                    "FTRP_BINTERNET" & _
                                  " ) VALUES ( " & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? " & _
                                    " )"
        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        '        Dim objParam As OleDbParameter


        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString
        m_dbconn.BeginTransaction()
        objCommand.Transaction = m_dbconn.transaction


        CreateParameterP_FTRP_CODE(objCommand)
        CreateParameterP_FTRP_ETAT(objCommand)
        CreateParameterP_FTRP_CLT_ID(objCommand)
        CreateParameterP_FTRP_DATE(objCommand)
        CreateParameterP_FTRP_TOTAL_TAXES(objCommand)
        CreateParameterP_FTRP_TOTAL_HT(objCommand)
        CreateParameterP_FTRP_TOTAL_TTC(objCommand)
        CreateParameterP_FTRP_PERIODE(objCommand)
        CreateParameterP_FTRP_COM_FACT(objCommand)
        CreateParameterP_FTRP_MONTANT_REGLEMENT(objCommand)
        CreateParameterP_FTRP_DATE_REGLEMENT(objCommand)
        CreateParameterP_FTRP_REF_REGLEMENT(objCommand)
        CreateParameterP_FTRP_IDMODEREGLEMENT(objCommand)
        CreateParameterP_FTRP_DECHEANCE(objCommand)
        CreateParameterP_FTRP_BINTERNET(objCommand)
        Try
            objCommand.ExecuteNonQuery()
            objCommand = New OleDbCommand("SELECT MAX(FTRP_ID) FROM FACTTRP", m_dbconn.Connection)
            objCommand.Transaction = m_dbconn.transaction
            objRS = objCommand.ExecuteReader
            If (objRS.Read) Then
                m_id = objRS.GetInt32(0)
                cleanErreur()
                bReturn = True
            Else
                bReturn = False
            End If
            objRS.Close()
            objRS = Nothing

        Catch ex As Exception
            setError("InsertFACTCOM", ex.ToString())
            bReturn = False
        End Try

        If Not objRS Is Nothing Then
            objRS.Close()
        End If
        If bReturn Then
            m_dbconn.transaction.Commit()
        Else
            m_dbconn.transaction.Rollback()
        End If
        '    Debug.Assert(m_id <> 0, "ID=0")
        Debug.Assert(bReturn, "InsertFACTTRP: " & getErreur())
        Return bReturn
    End Function 'insertFACTTRP

    Protected Function deleteFACTTRP() As Boolean
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "Id <> 0")
        Debug.Assert(Me.GetType().Name.Equals("FactTRP"))
        Dim bReturn As Boolean = False


        Dim sqlString As String = "DELETE FROM FACTTRP WHERE FTRP_ID=? "
        Dim objCommand As OleDbCommand
        Dim objFACT As FactTRP
        '        Dim objParam As OleDbParameter

        objFACT = CType(Me, FactTRP)
        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        CreateParameterP_ID(objCommand)
        Try
            objCommand.ExecuteNonQuery()
            m_id = 0
            objFACT.resetCode()
            bReturn = True

        Catch ex As Exception
            setError("DeleteFACTTRP", ex.ToString())
            bReturn = False
        End Try

        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'DeleteFACTTRP
    Protected Function deleteRefFACTTRP() As Boolean
        'Supprime la référence à la facture de transport dans la table commande
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "Id <> 0")
        Debug.Assert(Me.GetType().Name.Equals("FactTRP"))
        Dim bReturn As Boolean = False


        Dim sqlString As String = "UPDATE COMMANDE SET CMD_IDFACTTRP = 0 WHERE CMD_IDFACTTRP=? "
        Dim objCommand As OleDbCommand
        Dim objFACT As FactTRP
        '        Dim objParam As OleDbParameter

        objFACT = CType(Me, FactTRP)
        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        CreateParameterP_ID(objCommand)
        Try
            objCommand.ExecuteNonQuery()
            bReturn = True

        Catch ex As Exception
            setError("DeleteRefFACTTRP", ex.ToString())
            bReturn = False
        End Try

        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'DeleteRefFACTTRP
    Protected Function loadFACTTRP() As Boolean

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "L'id doit être renseigné")

        Dim sqlString As String = "SELECT " & _
                                    "FTRP_CODE," & _
                                    "FTRP_ETAT," & _
                                    "FTRP_CLT_ID," & _
                                    "FTRP_DATE," & _
                                    "FTRP_TOTAL_HT," & _
                                    "FTRP_TOTAL_TTC," & _
                                    "FTRP_TOTAL_TAXES," & _
                                    "FTRP_PERIODE," & _
                                    "FTRP_COM_FACT," & _
                                    "FTRP_MONTANT_REGLEMENT," & _
                                    "FTRP_DATE_REGLEMENT," & _
                                    "FTRP_REF_REGLEMENT," & _
                                    "FTRP_IDMODEREGLEMENT," & _
                                    "FTRP_DECHEANCE," & _
                                    "FTRP_BINTERNET" & _
                                    " FROM FACTTRP WHERE " & _
                                   " FTRP_ID = ?"

        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        Dim objFTRP As FactTRP
        '        Dim objParam As OleDbParameter
        Dim bReturn As Boolean

        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        'Si le paramètre existe on ne met à jour que sa valeur
        'Autrement on le crée
        CreateParameterP_ID(objCommand)

        Try
            objRS = objCommand.ExecuteReader
            If Not objRS.HasRows Then
                objRS.Close()
                m_id = 0
                Return False
            End If
            objRS.Read()
            objFTRP = CType(Me, FactTRP)
            Try
                objFTRP.code = GetString(objRS, "FTRP_CODE")
            Catch ex As InvalidCastException
                objFTRP.code = ""
            End Try
            Try
                objFTRP.setEtat(GetString(objRS, "FTRP_ETAT"))
            Catch ex As InvalidCastException
                objFTRP.setEtat(vncEnums.vncEtatCommande.vncFactTRPGeneree)
            End Try
            Try
                objFTRP.dateCommande = GetString(objRS, "FTRP_DATE")
            Catch ex As InvalidCastException
                objFTRP.dateCommande = DATE_DEFAUT
            End Try
            Try 'PERIODE
                objFTRP.periode = GetString(objRS, "FTRP_PERIODE")
            Catch ex As InvalidCastException
                objFTRP.periode = ""
            End Try
            Try 'Total Taxes
                objFTRP.montantTaxes = GetString(objRS, "FTRP_TOTAL_TAXES")
            Catch ex As InvalidCastException
                objFTRP.montantTaxes = 0
            End Try
            Try 'Total HT
                objFTRP.totalHT = GetString(objRS, "FTRP_TOTAL_HT")
            Catch ex As InvalidCastException
                objFTRP.totalHT = 0
            End Try
            Try 'Total TTC
                objFTRP.totalTTC = GetString(objRS, "FTRP_TOTAL_TTC")
            Catch ex As InvalidCastException
                objFTRP.totalTTC = 0
            End Try
            Try
                objFTRP.montantReglement = GetString(objRS, "FTRP_MONTANT_REGLEMENT")
            Catch ex As InvalidCastException
                objFTRP.montantReglement = 0
            End Try
            Try
                objFTRP.dateReglement = GetString(objRS, "FTRP_DATE_REGLEMENT")
            Catch ex As InvalidCastException
                objFTRP.dateReglement = DATE_DEFAUT
            End Try
            Try
                objFTRP.refReglement = GetString(objRS, "FTRP_REF_REGLEMENT")
            Catch ex As InvalidCastException
                objFTRP.refReglement = ""
            End Try
            Try
                objFTRP.idModeReglement = GetString(objRS, "FTRP_IDMODEREGLEMENT")
            Catch ex As InvalidCastException
                objFTRP.idModeReglement = 0
            End Try
            'Chargement de la date d'échéance après le mode de reglement pour éviter le recalcul
            Try
                objFTRP.dEcheance = GetString(objRS, "FTRP_DECHEANCE")
            Catch ex As InvalidCastException
                objFTRP.dEcheance = DATE_DEFAUT
            End Try
            Try
                objFTRP.bExportInternet = GetValue(objRS, "FTRP_BINTERNET")
            Catch ex As InvalidCastException
                objFTRP.bExportInternet = False
            End Try
            Try
                objFTRP.oTiers.setid(GetString(objRS, "FTRP_CLT_ID"))
            Catch ex As InvalidCastException
                objFTRP.oTiers.setid(0)
            End Try
            Try
                objFTRP.CommFacturation.comment = GetString(objRS, "FTRP_COM_FACT")
            Catch ex As InvalidCastException
                objFTRP.CommFacturation.comment = ""
            End Try

            cleanErreur()
            bReturn = True
        Catch ex As Exception
            setError("LoadFTRPCOM", ex.ToString())
            bReturn = False
        End Try
        If (Not objRS Is Nothing) Then
            objRS.Close()
            objRS = Nothing
        End If
        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'LoadFACTTRP
    Protected Function UpdateFACTTRP() As Boolean
        '=======================================================================
        'Fonction : UpdateFACTCOM
        'Description : Mise à jour  d'une Facture de commission
        'Retour : Rend Vrai si l'UPDATE s'est correctement effectué
        '=======================================================================
        Dim bReturn As Boolean
        Dim objFact As FactTRP
        bReturn = False

        Debug.Assert(Me.GetType().Name.Equals("FactTRP"), "Objet de Type FactTRP Requis")
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "ID<>0")
        objFact = CType(Me, FactTRP)

        Dim sqlString As String = "UPDATE FACTTRP SET " & _
                                    "FTRP_CODE = ? ," & _
                                    "FTRP_ETAT = ? ," & _
                                    "FTRP_CLT_ID = ? ," & _
                                    "FTRP_DATE = ? ," & _
                                    "FTRP_TOTAL_TAXES = ? ," & _
                                    "FTRP_TOTAL_HT = ? ," & _
                                    "FTRP_TOTAL_TTC = ? ," & _
                                    "FTRP_PERIODE = ? ," & _
                                    "FTRP_COM_FACT = ? ," & _
                                    "FTRP_MONTANT_REGLEMENT = ? ," & _
                                    "FTRP_DATE_REGLEMENT = ? ," & _
                                    "FTRP_REF_REGLEMENT = ? , " & _
                                    "FTRP_IDMODEREGLEMENT = ?," & _
                                    "FTRP_DECHEANCE = ?," & _
                                    "FTRP_BINTERNET = ?" & _
                                    " WHERE FTRP_ID = ?"

        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        '        Dim objParam As OleDbParameter



        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString
        m_dbconn.BeginTransaction()
        objCommand.Transaction = m_dbconn.transaction

        CreateParameterP_FTRP_CODE(objCommand)
        CreateParameterP_FTRP_ETAT(objCommand)
        CreateParameterP_FTRP_CLT_ID(objCommand)
        CreateParameterP_FTRP_DATE(objCommand)
        CreateParameterP_FTRP_TOTAL_TAXES(objCommand)
        CreateParameterP_FTRP_TOTAL_HT(objCommand)
        CreateParameterP_FTRP_TOTAL_TTC(objCommand)
        CreateParameterP_FTRP_PERIODE(objCommand)
        CreateParameterP_FTRP_COM_FACT(objCommand)
        CreateParameterP_FTRP_MONTANT_REGLEMENT(objCommand)
        CreateParameterP_FTRP_DATE_REGLEMENT(objCommand)
        CreateParameterP_FTRP_REF_REGLEMENT(objCommand)
        CreateParameterP_FTRP_IDMODEREGLEMENT(objCommand)
        CreateParameterP_FTRP_DECHEANCE(objCommand)
        CreateParameterP_FTRP_BINTERNET(objCommand)
        CreateParameterP_ID(objCommand)
        Try
            objCommand.ExecuteNonQuery()
            m_dbconn.transaction.Commit()
            bReturn = True

        Catch ex As Exception
            setError("UpdateFACTTRP", ex.ToString())
            m_dbconn.transaction.Rollback()
            bReturn = False
        End Try

        Debug.Assert(bReturn, "UpdateFACTTRP" & getErreur())
        Return bReturn
    End Function 'UpdateFACTTRP

    '==========================================================================
    'Methode : deletecolLGTRP
    'Description : Suppression des lignes d'une Facture de transport
    Protected Function deletecolLgTRP() As Boolean
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "Id <> 0")
        Debug.Assert(m_typedonnee = vncEnums.vncTypeDonnee.FACTTRP, "Objet de type FactTRP requis")


        Dim sqlString As String = "DELETE FROM LGFACTTRP "
        Dim strClauseWhere As String = " WHERE LGFACTTRP.LGTRP_FACTTRP_ID = ? "
        Dim objCommand As OleDbCommand
        '        Dim objParam As OleDbParameter
        Dim bReturn As Boolean

        'Clause where en Fonction du type de l'objet courant
        sqlString = sqlString & strClauseWhere

        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        CreateParameterP_ID(objCommand)
        Try
            objCommand.ExecuteNonQuery()

            bReturn = True

        Catch ex As Exception
            setError("deletecolLgTRP", ex.ToString())
            bReturn = False
        End Try
        Debug.Assert(bReturn, "deletecolLgTRP" & getErreur())
        Return bReturn
    End Function 'deletecolLgTRP

    Protected Function INSERTcolLGTRP() As Boolean
        Dim objFactTRP As Commande
        Dim bReturn As Boolean
        Dim objLgTRP As LgFactTRP

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "Le Client doit être Sauvegardé")
        Debug.Assert(m_typedonnee = vncEnums.vncTypeDonnee.FACTTRP, "Objet de type FactTRP requis")
        objFactTRP = Me
        Debug.Assert(Not objFactTRP.colLignes Is Nothing, "ColLignes is Nothing")


        Dim sqlString As String = "INSERT INTO LGFACTTRP (" & _
                                    "LGTRP_NUM," & _
                                    "LGTRP_FACTTRP_ID," & _
                                    "LGTRP_CMDCLT_ID," & _
                                    "LGTRP_TRPEUR," & _
                                    "LGTRP_DATE_LIV," & _
                                    "LGTRP_REF_LIV," & _
                                    "LGTRP_DATE_CMD," & _
                                    "LGTRP_REF_CMD," & _
                                    "LGTRP_QTE_COLIS," & _
                                    "LGTRP_QTE_PAL_PREP," & _
                                    "LGTRP_QTE_PAL_NONPREP," & _
                                    "LGTRP_PU_PAL_PREP," & _
                                    "LGTRP_PU_PAL_NONPREP," & _
                                    "LGTRP_POIDS," & _
                                    "LGTRP_MT_HT" & _
                                    ") VALUES ( " & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "?" & _
                                  " )"
        Dim sqlString2 As String = "UPDATE COMMANDE SET CMD_IDFACTTRP = ? WHERE CMD_ID = ?"
        Dim sqlString3 As String = "SELECT MAX(LGTRP_ID) FROM LGFACTTRP"
        Dim objCommand As OleDbCommand
        Dim objCommand2 As OleDbCommand
        Dim objCommandID As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        '        Dim objParam As OleDbParameter

        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString
        m_dbconn.BeginTransaction()
        objCommand.Transaction = m_dbconn.transaction


        objCommand2 = New OleDbCommand
        objCommand2.Connection = m_dbconn.Connection
        objCommand2.CommandText = sqlString2
        objCommand2.Transaction = m_dbconn.transaction

        objCommandID = New OleDbCommand
        objCommandID.Connection = m_dbconn.Connection
        objCommandID.CommandText = sqlString3
        objCommandID.Transaction = m_dbconn.transaction

        bReturn = True
        Try
            'Commande 1
            Dim oParam1 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.VarChar)
            Dim oParam2 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Integer)
            Dim oParam3 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Integer)
            Dim oParam4 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.VarChar)
            Dim oParam5 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.VarChar)
            Dim oParam6 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.VarChar)
            Dim oParam7 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.VarChar)
            Dim oParam8 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.VarChar)
            Dim oParam9 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Double)
            Dim oParam10 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Double)
            Dim oParam11 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Double)
            Dim oParam12 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Double)
            Dim oParam13 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Double)
            Dim oParam14 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Double)
            Dim oParam15 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Double)
            'Comande2
            Dim oParam20 As OleDbParameter = objCommand2.Parameters.Add("?", OleDbType.Integer)
            Dim oParam21 As OleDbParameter = objCommand2.Parameters.Add("?", OleDbType.Integer)

            For Each objLgTRP In objFactTRP.colLignes

                oParam1.Value = objLgTRP.num
                oParam2.Value = objFactTRP.id
                oParam3.Value = objLgTRP.idCmdCLT
                oParam4.Value = truncate(objLgTRP.nomTransporteur, 50)
                oParam5.Value = objLgTRP.dateLivraison
                oParam6.Value = truncate(objLgTRP.referenceLivraison, 50)
                oParam7.Value = objLgTRP.dateCommande
                oParam8.Value = objLgTRP.refCommande
                oParam9.Value = CDbl(objLgTRP.qteColis)
                oParam10.Value = CDbl(objLgTRP.qtePalettesPreparees)
                oParam11.Value = CDbl(objLgTRP.qtePalettesNonPreparees)
                oParam12.Value = CDbl(objLgTRP.puPalettesPreparees)
                oParam13.Value = CDbl(objLgTRP.puPalettesNonPreparees)
                oParam14.Value = CDbl(objLgTRP.poids)
                oParam15.Value = CDbl(objLgTRP.prixHT)
                objCommand.ExecuteNonQuery()
                objRS = objCommandID.ExecuteReader
                If objRS.Read Then
                    objLgTRP.setid(objRS.GetInt32(0))
                    cleanErreur()
                    bReturn = True
                Else
                    setError("InsertColLGTRP", "NO LGTRPID")
                    bReturn = False
                End If
                objRS.Close()
                objRS = Nothing

                'MAJ COMMANDE CLIENT (FLAG FactureTransport)
                If objLgTRP.idCmdCLT <> 0 Then
                    oParam20.Value = objFactTRP.id
                    oParam21.Value = objLgTRP.idCmdCLT
                    objCommand2.ExecuteNonQuery()
                End If
            Next
        Catch ex As Exception
            setError("InsertcolLGTRP", ex.ToString())
            bReturn = False
        End Try

        If bReturn Then
            m_dbconn.transaction.Commit()
        Else
            m_dbconn.transaction.Rollback()
        End If

        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function ' insertcolLGCMD

    Protected Function UPDATEcolLGTRP() As Boolean
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "La Commande.id  <> 0")
        Debug.Assert(m_typedonnee = vncEnums.vncTypeDonnee.FACTTRP, "Objet de type FACTTRP requis")


        Dim sqlString As String = "UPDATE LGFACTTRP SET " & _
                                    "LGTRP_NUM = ? ," & _
                                    "LGTRP_FACTTRP_ID = ? ," & _
                                    "LGTRP_CMDCLT_ID= ? , " & _
                                    "LGTRP_TRPEUR= ? , " & _
                                    "LGTRP_DATE_LIV= ? , " & _
                                    "LGTRP_REF_LIV= ? , " & _
                                    "LGTRP_DATE_CMD= ? , " & _
                                    "LGTRP_REF_CMD= ? , " & _
                                    "LGTRP_QTE_COLIS= ? , " & _
                                    "LGTRP_QTE_PAL_PREP= ? , " & _
                                    "LGTRP_QTE_PAL_NONPREP= ? , " & _
                                    "LGTRP_PU_PAL_PREP= ? , " & _
                                    "LGTRP_PU_PAL_NONPREP= ? , " & _
                                    "LGTRP_POIDS = ? , " & _
                                    "LGTRP_MT_HT = ? " & _
                                    " WHERE " & _
                                    " LGTRP_ID = ?"
        Dim objCommand As OleDbCommand
        Dim objFact As FactTRP
        '        Dim objParam As OleDbParameter
        Dim bReturn As Boolean
        Dim objLgTRP As LgFactTRP


        objFact = CType(Me, FactTRP)
        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString
        m_dbconn.BeginTransaction()
        objCommand.Transaction = m_dbconn.transaction



        bReturn = True
        For Each objLgTRP In objFact.colLignes
            Try
                objCommand.Parameters.Clear()
                objCommand.Parameters.AddWithValue("?", objLgTRP.num)
                objCommand.Parameters.AddWithValue("?", m_id)
                objCommand.Parameters.AddWithValue("?", objLgTRP.idCmdCLT)
                objCommand.Parameters.AddWithValue("?", truncate(objLgTRP.nomTransporteur, 50))
                objCommand.Parameters.AddWithValue("?", objLgTRP.dateLivraison)
                objCommand.Parameters.AddWithValue("?", truncate(objLgTRP.referenceLivraison, 50))
                objCommand.Parameters.AddWithValue("?", objLgTRP.dateCommande)
                objCommand.Parameters.AddWithValue("?", objLgTRP.refCommande)
                objCommand.Parameters.AddWithValue("?", objLgTRP.qteColis)
                objCommand.Parameters.AddWithValue("?", objLgTRP.qtePalettesPreparees)
                objCommand.Parameters.AddWithValue("?", objLgTRP.qtePalettesNonPreparees)
                objCommand.Parameters.AddWithValue("?", objLgTRP.puPalettesPreparees)
                objCommand.Parameters.AddWithValue("?", objLgTRP.puPalettesNonPreparees)
                objCommand.Parameters.AddWithValue("?", objLgTRP.poids)
                objCommand.Parameters.AddWithValue("?", objLgTRP.prixHT)
                objCommand.Parameters.AddWithValue("?", objLgTRP.id)
                objCommand.ExecuteNonQuery()
            Catch ex As Exception
                setError("UpdatecolLGTRP", ex.ToString())
                bReturn = False
                Exit For
            End Try
        Next objLgTRP

        If bReturn Then
            m_dbconn.transaction.Commit()
        Else
            m_dbconn.transaction.Rollback()
        End If

        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function ' UPDATEcolLGTRP
    Protected Function LoadcolLgFactTRP() As Boolean
        '================================================================================
        'Fonction : LoadcolLgFactTRP
        'Description : Chargement de la liste des lignes d'une factures de transport
        '================================================================================
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "Id <> 0")
        Debug.Assert(Me.GetType().Name.Equals("FactTRP"))

        Dim bReturn As Boolean
        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        Dim objFACTTRP As FactTRP
        Dim objLGTRP As LgFactTRP

        Dim sqlString As String = "SELECT " & _
                                    "LGTRP_ID," & _
                                    "LGTRP_NUM," & _
                                    "LGTRP_FACTTRP_ID," & _
                                    "LGTRP_CMDCLT_ID," & _
                                    "LGTRP_TRPEUR," & _
                                    "LGTRP_DATE_LIV," & _
                                    "LGTRP_REF_LIV," & _
                                    "LGTRP_DATE_CMD," & _
                                    "LGTRP_REF_CMD," & _
                                    "LGTRP_QTE_COLIS," & _
                                    "LGTRP_QTE_PAL_PREP," & _
                                    "LGTRP_QTE_PAL_NONPREP," & _
                                    "LGTRP_PU_PAL_PREP," & _
                                    "LGTRP_PU_PAL_NONPREP," & _
                                    "LGTRP_POIDS," & _
                                    "LGTRP_MT_HT" & _
                                  " FROM LGFACTTRP" & _
                                  " WHERE LGTRP_FACTTRP_ID = ? ORDER BY LGTRP_NUM"

        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        'Paramétre ClientID 
        CreateParameterP_ID(objCommand)

        Try
            objFACTTRP = CType(Me, FactTRP)
            objRS = objCommand.ExecuteReader
            While objRS.Read
                Try
                    objLGTRP = New LgFactTRP
                    Try
                        objLGTRP.setid(GetString(objRS, "LGTRP_ID"))
                    Catch ex As InvalidCastException
                        objLGTRP.setid(0)
                    End Try
                    Try
                        objLGTRP.num = GetString(objRS, "LGTRP_NUM")
                    Catch ex As InvalidCastException
                        objLGTRP.num = ""
                    End Try
                    Try
                        objLGTRP.idFactTRP = GetString(objRS, "LGTRP_FACTTRP_ID")
                    Catch ex As InvalidCastException
                        objLGTRP.idFactTRP = 0
                    End Try
                    Try
                        objLGTRP.idCmdCLT = GetString(objRS, "LGTRP_CMDCLT_ID")
                    Catch ex As InvalidCastException
                        objLGTRP.idCmdCLT = 0
                    End Try
                    Try
                        objLGTRP.nomTransporteur = GetString(objRS, "LGTRP_TRPEUR")
                    Catch ex As InvalidCastException
                        objLGTRP.nomTransporteur = ""
                    End Try
                    Try
                        objLGTRP.dateLivraison = GetString(objRS, "LGTRP_DATE_LIV")
                    Catch ex As InvalidCastException
                        objLGTRP.dateLivraison = DATE_DEFAUT
                    End Try
                    Try
                        objLGTRP.referenceLivraison = GetString(objRS, "LGTRP_REF_LIV")
                    Catch ex As InvalidCastException
                        objLGTRP.referenceLivraison = ""
                    End Try
                    Try
                        objLGTRP.dateCommande = GetString(objRS, "LGTRP_DATE_CMD")
                    Catch ex As InvalidCastException
                        objLGTRP.dateCommande = DATE_DEFAUT
                    End Try
                    Try
                        objLGTRP.refCommande = GetString(objRS, "LGTRP_REF_CMD")
                    Catch ex As InvalidCastException
                        objLGTRP.refCommande = ""
                    End Try
                    Try
                        objLGTRP.qteColis = GetString(objRS, "LGTRP_QTE_COLIS")
                    Catch ex As InvalidCastException
                        objLGTRP.qteColis = 0
                    End Try
                    Try
                        objLGTRP.qtePalettesPreparees = GetString(objRS, "LGTRP_QTE_PAL_PREP")
                    Catch ex As InvalidCastException
                        objLGTRP.qtePalettesPreparees = 0
                    End Try
                    Try
                        objLGTRP.qtePalettesNonPreparees = GetString(objRS, "LGTRP_QTE_PAL_NONPREP")
                    Catch ex As InvalidCastException
                        objLGTRP.qtePalettesNonPreparees = 0
                    End Try
                    Try
                        objLGTRP.puPalettesPreparees = GetString(objRS, "LGTRP_PU_PAL_PREP")
                    Catch ex As InvalidCastException
                        objLGTRP.puPalettesPreparees = 0
                    End Try
                    Try
                        objLGTRP.puPalettesNonPreparees = GetString(objRS, "LGTRP_PU_PAL_NONPREP")
                    Catch ex As InvalidCastException
                        objLGTRP.puPalettesNonPreparees = 0
                    End Try
                    Try
                        objLGTRP.poids = GetString(objRS, "LGTRP_POIDS")
                    Catch ex As InvalidCastException
                        objLGTRP.poids = 0
                    End Try
                    Try
                        objLGTRP.prixHT = GetString(objRS, "LGTRP_MT_HT")
                    Catch ex As InvalidCastException
                        objLGTRP.prixHT = 0
                    End Try
                    objFACTTRP.resetBooleans()
                    objFACTTRP.AjouteLigneFactTRP(objLGTRP, False)
                Catch ex As InvalidCastException
                    bReturn = False
                    Exit While
                End Try
            End While
            objRS.Close()
            objRS = Nothing
            bReturn = True
        Catch ex As Exception
            setError("LoadcolLgFactTRP", ex.ToString())
            bReturn = False
        End Try

        Debug.Assert(bReturn, "LoadcollgFactTRP" & getErreur())
        Return bReturn

    End Function ' LoadColLGFACTTRP

#End Region
#Region "Paramètres FACTTRP"
    Private Sub CreateParameterP_FTRP_CODE(ByVal objCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?" , objCMD.code)

    End Sub
    Private Sub CreateParameterP_FTRP_ETAT(ByVal objCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?" , objCMD.etat.codeEtat)
    End Sub
    Private Sub CreateParameterP_FTRP_CLT_ID(ByVal objCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?" , objCMD.oTiers.id)

    End Sub
    Private Sub CreateParameterP_FTRP_DATE(ByVal objCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?" , objCMD.dateCommande.ToShortDateString)

    End Sub
    Private Sub CreateParameterP_FTRP_TOTAL_TAXES(ByVal objCommand As OleDbCommand)
        Dim objCMD As FactTRP
        objCMD = Me
        objCommand.Parameters.AddWithValue("?" , objCMD.montantTaxes)
    End Sub
    Private Sub CreateParameterP_FTRP_TOTAL_HT(ByVal objCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?" , CDbl(objCMD.totalHT))
    End Sub

    Private Sub CreateParameterP_FTRP_TOTAL_TTC(ByVal objCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?" , CDbl(objCMD.totalTTC))
    End Sub
    Private Sub CreateParameterP_FTRP_COM_FACT(ByVal objCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCMD.CommFacturation.comment, 50))
    End Sub
    Private Sub CreateParameterP_FTRP_PERIODE(ByVal objCommand As OleDbCommand)
        Dim objCMD As FactTRP
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCMD.periode, 50))
    End Sub
    Private Sub CreateParameterP_FTRP_MONTANT_REGLEMENT(ByVal objCommand As OleDbCommand)
        Dim objCMD As FactTRP
        objCMD = Me
        objCommand.Parameters.AddWithValue("?" , CDbl(objCMD.montantReglement))

    End Sub
    Private Sub CreateParameterP_FTRP_DATE_REGLEMENT(ByVal objCommand As OleDbCommand)
        Dim objCMD As FactTRP
        objCMD = Me
        objCommand.Parameters.AddWithValue("?" , objCMD.dateReglement)

    End Sub
    Private Sub CreateParameterP_FTRP_REF_REGLEMENT(ByVal objCommand As OleDbCommand)
        Dim objCMD As FactTRP
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCMD.refReglement, 50))
    End Sub
    Private Sub CreateParameterP_FTRP_IDMODEREGLEMENT(ByVal objCommand As OleDbCommand)
        Dim objCMD As FactTRP
        objCMD = Me
        objCommand.Parameters.AddWithValue("?" , objCMD.idModeReglement)
    End Sub
    Private Sub CreateParameterP_FTRP_DECHEANCE(ByVal objCommand As OleDbCommand)
        Dim objCMD As FactTRP
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", objCMD.dEcheance)
    End Sub
    Private Sub CreateParameterP_FTRP_BINTERNET(ByVal objCommand As OleDbCommand)
        Dim objCMD As FactTRP
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", objCMD.bExportInternet)
    End Sub
#End Region
#Region "Facture HBV"
    Protected Function insertFACTHBV() As Boolean
        '=======================================================================
        'Fonction : insertFACTHBV
        'Description : Insertion d'une Facture de Transport
        'Retour : Rend Vrai si l'INSERT s'est correctement effectué
        '=======================================================================
        Dim bReturn As Boolean
        Dim objFACT As FactHBV
        bReturn = False

        Debug.Assert(Me.GetType().Name.Equals("FactHBV"), "Objet de Type 'FactHBV' Requis")
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id = 0, "ID=0")
        objFACT = CType(Me, FactHBV)

        Dim sqlString As String = "INSERT INTO FACTHBV( " & _
                                    "FHBV_CODE," & _
                                    "FHBV_ETAT," & _
                                    "FHBV_CLT_ID," & _
                                    "FHBV_DATE," & _
                                    "FHBV_TOTAL_HT," & _
                                    "FHBV_TOTAL_TTC," & _
                                    "FHBV_COM_FACT," & _
                                    "FHBV_IDMODEREGLEMENT," & _
                                    "FHBV_DECHEANCE," & _
                                    "FHBV_IDCOMMANDE" & _
                                  " ) VALUES ( " & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? " & _
                                    " )"
        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        '        Dim objParam As OleDbParameter


        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString
        m_dbconn.BeginTransaction()
        objCommand.Transaction = m_dbconn.transaction


        CreateParameterP_FHBV_CODE(objCommand)
        CreateParameterP_FHBV_ETAT(objCommand)
        CreateParameterP_FHBV_CLT_ID(objCommand)
        CreateParameterP_FHBV_DATE(objCommand)
        CreateParameterP_FHBV_TOTAL_HT(objCommand)
        CreateParameterP_FHBV_TOTAL_TTC(objCommand)
        CreateParameterP_FHBV_COM_FACT(objCommand)
        CreateParameterP_FHBV_IDMODEREGLEMENT(objCommand)
        CreateParameterP_FHBV_DECHEANCE(objCommand)
        CreateParameterP_FHBV_IDCOMMANDE(objCommand)
        Try
            objCommand.ExecuteNonQuery()
            objCommand = New OleDbCommand("SELECT MAX(FHBV_ID) FROM FACTHBV", m_dbconn.Connection)
            objCommand.Transaction = m_dbconn.transaction
            objRS = objCommand.ExecuteReader
            If (objRS.Read) Then
                m_id = objRS.GetInt32(0)
                cleanErreur()
                bReturn = True
            Else
                bReturn = False
            End If
            objRS.Close()
            objRS = Nothing

        Catch ex As Exception
            setError("InsertFACTHBV", ex.ToString())
            bReturn = False
        End Try

        If Not objRS Is Nothing Then
            objRS.Close()
        End If
        If bReturn Then
            m_dbconn.transaction.Commit()
        Else
            m_dbconn.transaction.Rollback()
        End If
        '    Debug.Assert(m_id <> 0, "ID=0")
        Debug.Assert(bReturn, "InsertFACTHBV: " & getErreur())
        Return bReturn
    End Function 'insertFACTHBV

    Protected Function deleteFACTHBV() As Boolean
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "Id <> 0")
        Debug.Assert(Me.GetType().Name.Equals("FactHBV"))
        Dim bReturn As Boolean = False


        Dim sqlString As String = "DELETE FROM FACTHBV WHERE FHBV_ID=? "
        Dim objCommand As OleDbCommand
        Dim objFACT As FactHBV
        '        Dim objParam As OleDbParameter

        objFACT = CType(Me, FactHBV)
        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        CreateParameterP_ID(objCommand)
        Try
            objCommand.ExecuteNonQuery()
            m_id = 0
            objFACT.resetCode()
            bReturn = True

        Catch ex As Exception
            setError("DeleteFACTHBV", ex.ToString())
            bReturn = False
        End Try

        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'DeleteFACTHBV
    Protected Function deleteRefFACTHBV() As Boolean
        'Supprime la référence à la facture de transport dans la table commande
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "Id <> 0")
        Debug.Assert(Me.GetType().Name.Equals("FactHBV"))
        Dim bReturn As Boolean = False


        Dim sqlString As String = "UPDATE COMMANDE SET CMD_IDFACTHBV = 0 WHERE CMD_IDFACTHBV=? "
        Dim objCommand As OleDbCommand
        Dim objFACT As FactHBV
        '        Dim objParam As OleDbParameter

        objFACT = CType(Me, FactHBV)
        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        CreateParameterP_ID(objCommand)
        Try
            objCommand.ExecuteNonQuery()
            bReturn = True

        Catch ex As Exception
            setError("DeleteRefFACTHBV", ex.ToString())
            bReturn = False
        End Try

        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'DeleteRefFACTHBV
    ''' <summary>
    ''' Chargement de la facture à partir de l'Id de la commande
    ''' </summary>
    ''' <param name="pIdCmd"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Function loadFACTHBVFromCmd(pIdCmd As Long) As Boolean
        Dim sqlString As String = "SELECT FHBV_ID " & _
                            "FROM FACTHBV WHERE FHBV_IDCOMMANDE = " & pIdCmd
        Dim bReturn As Boolean
        Dim objCommand As OleDbCommand
        Dim objFACT As FactHBV
        Dim objRS As OleDbDataReader = Nothing
        objCommand = m_dbconn.Connection.CreateCommand()
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString
        Try
            objRS = objCommand.ExecuteReader
            While (objRS.Read())
                Dim nId As Integer
                nId = getInteger(objRS, "FHBV_ID")
                m_id = nId
                loadFACTHBV()
            End While
            objRS.Close()
            objRS = Nothing
            bReturn = True
        Catch ex As Exception
            setError("ListeFACTHBVFromCmd", sqlString & vbCrLf & ex.ToString())
            bReturn = False
        End Try

        Debug.Assert(bReturn, getErreur())

        Return bReturn

    End Function 'loadFACTHBVFromCmd

    Protected Function loadFACTHBV() As Boolean

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "L'id doit être renseigné")

        Dim sqlString As String = "SELECT " & _
                                    "FHBV_CODE," & _
                                    "FHBV_ETAT," & _
                                    "FHBV_CLT_ID," & _
                                    "FHBV_DATE," & _
                                    "FHBV_TOTAL_HT," & _
                                    "FHBV_TOTAL_TTC," & _
                                    "FHBV_TOTAL_TAXES," & _
                                    "FHBV_COM_FACT," & _
                                    "FHBV_IDMODEREGLEMENT," & _
                                    "FHBV_DECHEANCE," & _
                                    "FHBV_IDCOMMANDE" & _
                                    " FROM FACTHBV WHERE " & _
                                   " FHBV_ID = ?"

        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        Dim objFHBV As FactHBV
        '        Dim objParam As OleDbParameter
        Dim bReturn As Boolean

        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        'Si le paramètre existe on ne met à jour que sa valeur
        'Autrement on le crée
        CreateParameterP_ID(objCommand)

        Try
            objRS = objCommand.ExecuteReader
            If Not objRS.HasRows Then
                objRS.Close()
                m_id = 0
                Return False
            End If
            objRS.Read()
            objFHBV = CType(Me, FactHBV)
            Try
                objFHBV.code = GetString(objRS, "FHBV_CODE")
            Catch ex As InvalidCastException
                objFHBV.code = ""
            End Try
            Try
                objFHBV.setEtat(GetString(objRS, "FHBV_ETAT"))
            Catch ex As InvalidCastException
                objFHBV.setEtat(vncEnums.vncEtatCommande.vncFactHBVGeneree)
            End Try
            Try
                objFHBV.dateCommande = GetString(objRS, "FHBV_DATE")
            Catch ex As InvalidCastException
                objFHBV.dateCommande = DATE_DEFAUT
            End Try
            Try 'Total HT
                objFHBV.totalHT = GetString(objRS, "FHBV_TOTAL_HT")
            Catch ex As InvalidCastException
                objFHBV.totalHT = 0
            End Try
            Try 'Total TTC
                objFHBV.totalTTC = GetString(objRS, "FHBV_TOTAL_TTC")
            Catch ex As InvalidCastException
                objFHBV.totalTTC = 0
            End Try
            Try
                objFHBV.idModeReglement = GetString(objRS, "FHBV_IDMODEREGLEMENT")
            Catch ex As InvalidCastException
                objFHBV.idModeReglement = 0
            End Try
            'Chargement de la date d'échéance après le mode de reglement pour éviter le recalcul
            Try
                objFHBV.dEcheance = GetString(objRS, "FHBV_DECHEANCE")
            Catch ex As InvalidCastException
                objFHBV.dEcheance = DATE_DEFAUT
            End Try
            Try
                objFHBV.oTiers.setid(GetString(objRS, "FHBV_CLT_ID"))
            Catch ex As InvalidCastException
                objFHBV.oTiers.setid(0)
            End Try
            Try
                objFHBV.CommFacturation.comment = GetString(objRS, "FHBV_COM_FACT")
            Catch ex As InvalidCastException
                objFHBV.CommFacturation.comment = ""
            End Try
            Try
                objFHBV.idCommande = getInteger(objRS, "FHBV_IDCOMMANDE")
            Catch ex As InvalidCastException
                objFHBV.idCommande = 0
            End Try

            cleanErreur()
            bReturn = True
        Catch ex As Exception
            setError("LoadFHBVCOM", ex.ToString())
            bReturn = False
        End Try
        If (Not objRS Is Nothing) Then
            objRS.Close()
            objRS = Nothing
        End If
        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'LoadFACTHBV
    Protected Function UpdateFACTHBV() As Boolean
        Dim bReturn As Boolean
        Dim objFact As FactHBV
        bReturn = False

        Debug.Assert(Me.GetType().Name.Equals("FactHBV"), "Objet de Type FactHBV Requis")
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "ID<>0")
        objFact = CType(Me, FactHBV)

        Dim sqlString As String = "UPDATE FACTHBV SET " & _
                                    "FHBV_CODE = ? ," & _
                                    "FHBV_ETAT = ? ," & _
                                    "FHBV_CLT_ID = ? ," & _
                                    "FHBV_DATE = ? ," & _
                                    "FHBV_TOTAL_HT = ? ," & _
                                    "FHBV_TOTAL_TTC = ? ," & _
                                    "FHBV_COM_FACT = ? ," & _
                                    "FHBV_IDMODEREGLEMENT = ?," & _
                                    "FHBV_DECHEANCE = ?," & _
                                    "FHBV_IDCOMMANDE = ?" & _
                                    " WHERE FHBV_ID = ?"

        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        '        Dim objParam As OleDbParameter



        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString
        m_dbconn.BeginTransaction()
        objCommand.Transaction = m_dbconn.transaction

        CreateParameterP_FHBV_CODE(objCommand)
        CreateParameterP_FHBV_ETAT(objCommand)
        CreateParameterP_FHBV_CLT_ID(objCommand)
        CreateParameterP_FHBV_DATE(objCommand)
        CreateParameterP_FHBV_TOTAL_HT(objCommand)
        CreateParameterP_FHBV_TOTAL_TTC(objCommand)
        CreateParameterP_FHBV_COM_FACT(objCommand)
        CreateParameterP_FHBV_IDMODEREGLEMENT(objCommand)
        CreateParameterP_FHBV_DECHEANCE(objCommand)
        CreateParameterP_FHBV_IDCOMMANDE(objCommand)
        CreateParameterP_ID(objCommand)
        Try
            objCommand.ExecuteNonQuery()
            m_dbconn.transaction.Commit()
            bReturn = True

        Catch ex As Exception
            setError("UpdateFACTHBV", ex.ToString())
            m_dbconn.transaction.Rollback()
            bReturn = False
        End Try

        Debug.Assert(bReturn, "UpdateFACTHBV" & getErreur())
        Return bReturn
    End Function 'UpdateFACTHBV
    ''' <summary>
    ''' Sauvegarde des Lignes des facture Hobivin
    ''' Chargement des ID des lignes existantes
    ''' pour chaque lignes de la collection
    ''' Si l'ID existe
    '''     Update de la ligne + Suppression de l'ID de la liste des Ids
    ''' Sinon
    '''     Création de la Ligne
    ''' 
    ''' Suppression des Lignes correspondant aux Ids restants
    ''' 
    ''' 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Function saveColLGFACTHBV() As Boolean
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "La Commande.id  <> 0")
        Debug.Assert(m_typedonnee = vncEnums.vncTypeDonnee.FACTHBV, "Objet de type FACTHBV requis")

        Dim objCommand As OleDbCommand
        Dim bReturn As Boolean
        Try

            'Marquage de toutes les lignes de factures
            Dim sqlStringMARK As String = "UPDATE LIGNE_FACTHBV SET " & _
                                        "LGFHBV_ETAT = 1 " & _
                                        "WHERE LGFHBV_FACT_ID =  " & Me.id


            objCommand = m_dbconn.Connection.CreateCommand()
            objCommand.CommandText = sqlStringMARK
            m_dbconn.BeginTransaction()
            objCommand.Transaction = m_dbconn.transaction
            objCommand.ExecuteNonQuery()


            'Parcours de Chaque ligne de la collection
            Dim oFact As FactHBV
            oFact = CType(Me, FactHBV)

            'Insert Or Update de chaque ligne
            bReturn = True
            For Each oLg As LgFactHBV In oFact.colLignes
                If Not oLg.bDeleted Then
                    If oLg.id = 0 Then
                        bReturn = insertLgFactHBV(oLg)
                    Else
                        bReturn = updateLgFactHBV(oLg)
                    End If
                End If
            Next oLg
            'Suppression des autres lignes
            Dim sqlStringDLT As String = "DELETE LIGNE_FACTHBV " & _
                                        "WHERE LGFHBV_ETAT=1 AND LGFHBV_FACT_ID =  " & Me.id


            objCommand = m_dbconn.Connection.CreateCommand()
            objCommand.CommandText = sqlStringDLT
            objCommand.Transaction = m_dbconn.transaction
            objCommand.ExecuteNonQuery()

            m_dbconn.transaction.Commit()
            bReturn = True
        Catch ex As Exception
            setError(ex.Message)
            m_dbconn.transaction.Rollback()
            bReturn = False
        End Try

        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function ' saveColLGFACTHBV
    'Insertion d'une ligne de Facture Hobivin
    Private Function insertLgFactHBV(pLg As LgFactHBV) As Boolean
        Dim bReturn As Boolean
        Dim oFact As FactHBV
        Try
            Dim oCmd As OleDbCommand
            oCmd = m_dbconn.Connection.CreateCommand()
            oFact = Me


            Dim sqlString As String = "INSERT INTO LIGNE_FACTHBV (" & _
                                        "LGFHBV_FACT_ID," & _
                                        "LGFHBV_PRD_ID," & _
                                        "LGFHBV_QTE_COMMANDE," & _
                                        "LGFHBV_QTE_LIV," & _
                                        "LGFHBV_QTE_FACT," & _
                                        "LGFHBV_PRIX_UNITAIRE," & _
                                        "LGFHBV_PRIX_HT," & _
                                        "LGFHBV_PRIX_TTC," & _
                                        "LGFHBV_BGRATUIT," & _
                                        "LGFHBV_ETAT" & _
                                        ") VALUES ( " & _
                                        "? ," & _
                                        "? ," & _
                                        "? ," & _
                                        "? ," & _
                                        "? ," & _
                                        "? ," & _
                                        "? ," & _
                                        "? ," & _
                                        "?, " & _
                                        "0" & _
                                      " )"
            Dim sqlString3 As String = "SELECT MAX(LGFHBV_ID) FROM LIGNE_FACTHBV"
            Dim objCommand As OleDbCommand
            Dim objCommandID As OleDbCommand
            Dim objRS As OleDbDataReader = Nothing

            objCommand = New OleDbCommand
            objCommand.Connection = m_dbconn.Connection
            objCommand.CommandText = sqlString
            objCommand.Transaction = m_dbconn.transaction

            objCommandID = New OleDbCommand
            objCommandID.Connection = m_dbconn.Connection
            objCommandID.CommandText = sqlString3
            objCommandID.Transaction = m_dbconn.transaction

            bReturn = True
            Try
                'Commande 1
                Dim oParam1 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Integer) 'LGFHBV_FACT_ID
                Dim oParam2 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Integer) 'LGFHBV_PRD_ID
                Dim oParam3 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Integer) 'LGFHBV_QTE_COMMANDE
                Dim oParam4 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Integer) 'LGFHBV_QTE_LIV
                Dim oParam5 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Integer) 'LGFHBV_QTE_FACT
                Dim oParam6 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Double) 'LGFHBV_PRIX_UNITAIRE
                Dim oParam7 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Double) 'LGFHBV_PRIX_HT
                Dim oParam8 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Double) 'LGFHBV_PRIX_TTC
                Dim oParam9 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Boolean) 'LGFHBV_BGRATUIT

                '"INSERT INTO LIGNE_FACTHBV (" & _
                ' "LGFHBV_FACT_ID," & _
                ' "LGFHBV_PRD_ID," & _
                ' "LGFHBV_QTE_COMMANDE," & _
                ' "LGFHBV_QTE_LIV," & _
                ' "LGFHBV_QTE_FACT," & _
                ' "LGFHBV_PRIX_UNITAIRE," & _
                ' "LGFHBV_PRIX_HT," & _
                ' "LGFHBV_PRIX_TTC," & _
                ' "LGFHBV_BGRATUIT," & _
                ' "LGFHBV_ETAT=0" & _

                oParam1.Value = oFact.id
                oParam2.Value = pLg.oProduit.id
                oParam3.Value = pLg.qteCommande
                oParam4.Value = pLg.qteLiv
                oParam5.Value = pLg.qteFact
                oParam6.Value = pLg.prixU
                oParam7.Value = pLg.prixHT
                oParam8.Value = pLg.prixTTC
                oParam9.Value = pLg.bGratuit
                objCommand.ExecuteNonQuery()

                objRS = objCommandID.ExecuteReader
                If objRS.Read Then
                    pLg.setid(objRS.GetInt32(0))
                    cleanErreur()
                    bReturn = True
                Else
                    setError("InsertLGFactHBV", "NO LGTRPID")
                    bReturn = False
                End If
                objRS.Close()
                objRS = Nothing

            Catch ex As Exception
                setError("InsertcolLGHBV", ex.ToString())
                bReturn = False
            End Try



        Catch ex As Exception
            bReturn = False
            setError(getErreur())
        End Try
        Debug.Assert(bReturn, "insertLgFactHBV.postCondition")
        Return bReturn
    End Function

    'Insertion d'une ligne de Facture Hobivin
    Private Function updateLgFactHBV(pLg As LgFactHBV) As Boolean
        Dim bReturn As Boolean
        Try
            Dim objCommand As OleDbCommand
            objCommand = m_dbconn.Connection.CreateCommand()
            objCommand.Transaction = m_dbconn.transaction

            Dim sqlString As String = "UPDATE LIGNE_FACTHBV SET " & _
                                        "LGFHBV_FACT_ID= ?," & _
                                        "LGFHBV_PRD_ID= ?," & _
                                        "LGFHBV_QTE_COMMANDE= ? " & " , " & _
                                        "LGFHBV_QTE_LIV= ? " & " , " & _
                                        "LGFHBV_QTE_FACT= ? " & "  ," & _
                                        "LGFHBV_PRIX_UNITAIRE= ?" & " , " & _
                                        "LGFHBV_PRIX_HT= ?" & " , " & _
                                        "LGFHBV_PRIX_TTC=?" & " , " & _
                                        "LGFHBV_BGRATUIT=?" & " , " & _
                                        "LGFHBV_ETAT=0 " & _
                                        "WHERE LGFHBV_ID = ?"
            objCommand.CommandText = sqlString
            objCommand.Parameters.AddWithValue("?", pLg.idFactHBV)
            objCommand.Parameters.AddWithValue("?", pLg.oProduit.id)
            objCommand.Parameters.AddWithValue("?", pLg.qteCommande)
            objCommand.Parameters.AddWithValue("?", pLg.qteLiv)
            objCommand.Parameters.AddWithValue("?", pLg.qteFact)
            objCommand.Parameters.AddWithValue("?", pLg.prixU)
            objCommand.Parameters.AddWithValue("?", pLg.prixHT)
            objCommand.Parameters.AddWithValue("?", pLg.prixTTC)
            objCommand.Parameters.AddWithValue("?", pLg.bGratuit)
            objCommand.Parameters.AddWithValue("?", pLg.id)

            bReturn = True
            Try
                objCommand.ExecuteNonQuery()

            Catch ex As Exception
                setError("updatecolLGHBV", ex.ToString())
                bReturn = False
            End Try



        Catch ex As Exception
            bReturn = False
            setError(getErreur())
        End Try
        Debug.Assert(bReturn, "updateLgFactHBV.postCondition")
        Return bReturn
    End Function

    Protected Function LoadcolLgFactHBV() As Boolean
        '================================================================================
        'Fonction : LoadcolLgFactHBV
        'Description : Chargement de la liste des lignes d'une factures de transport
        '================================================================================
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "Id <> 0")
        Debug.Assert(Me.GetType().Name.Equals("FactHBV"))

        Dim bReturn As Boolean
        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        Dim objFACTHBV As FactHBV
        Dim objLGHBV As LgFactHBV

        objFACTHBV = CType(Me, FactHBV)

        Dim sqlString As String = "SELECT " & _
                                    "LGFHBV_ID," & _
                                    "LGFHBV_FACT_ID," & _
                                    "LGFHBV_PRD_ID," & _
                                    "LGFHBV_QTE_COMMANDE," & _
                                    "LGFHBV_QTE_LIV," & _
                                    "LGFHBV_QTE_FACT," & _
                                    "LGFHBV_PRIX_UNITAIRE," & _
                                    "LGFHBV_PRIX_HT," & _
                                    "LGFHBV_PRIX_TTC," & _
                                    "LGFHBV_BGRATUIT" & _
                                  " FROM LIGNE_FACTHBV" & _
                                  " WHERE LGFHBV_FACT_ID =  " & objFACTHBV.id

        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString

        Try
            objRS = objCommand.ExecuteReader
            While objRS.Read
                Try
                    objLGHBV = New LgFactHBV
                    Try
                        objLGHBV.setid(getInteger(objRS, "LGFHBV_ID"))
                    Catch ex As InvalidCastException
                        objLGHBV.setid(0)
                    End Try
                    Try
                        objLGHBV.idFactHBV = getInteger(objRS, "LGFHBV_FACT_ID")
                    Catch ex As InvalidCastException
                        objLGHBV.idFactHBV = 0
                    End Try
                    Try
                        objLGHBV.oProduit = Produit.createandload(getInteger(objRS, "LGFHBV_PRD_ID"))
                    Catch ex As InvalidCastException
                        objLGHBV.oProduit = Nothing
                    End Try
                    Try
                        objLGHBV.qteCommande = getInteger(objRS, "LGFHBV_QTE_COMMANDE")
                    Catch ex As InvalidCastException
                        objLGHBV.qteCommande = 0
                    End Try
                    Try
                        objLGHBV.qteLiv = getInteger(objRS, "LGFHBV_QTE_LIV")
                    Catch ex As InvalidCastException
                        objLGHBV.qteLiv = 0
                    End Try
                    Try
                        objLGHBV.qteFact = getInteger(objRS, "LGFHBV_QTE_FACT")
                    Catch ex As InvalidCastException
                        objLGHBV.qteFact = 0
                    End Try
                    Try
                        objLGHBV.bGratuit = GetBoolean(objRS, "LGFHBV_BGRATUIT")
                    Catch ex As InvalidCastException
                        objLGHBV.bGratuit = 0
                    End Try
                    Try
                        objLGHBV.prixU = getDecimal(objRS, "LGFHBV_PRIX_UNITAIRE")
                    Catch ex As InvalidCastException
                        objLGHBV.prixU = 0
                    End Try
                    Try
                        objLGHBV.prixHT = getDecimal(objRS, "LGFHBV_PRIX_HT")
                    Catch ex As InvalidCastException
                        objLGHBV.prixHT = 0
                    End Try
                    Try
                        objLGHBV.prixTTC = getDecimal(objRS, "LGFHBV_PRIX_TTC")
                    Catch ex As InvalidCastException
                        objLGHBV.prixTTC = 0
                    End Try
                    objLGHBV.resetBooleans()
                    objFACTHBV.AjouteLigneFactHBV(objLGHBV, False)
                Catch ex As InvalidCastException
                    bReturn = False
                    Exit While
                End Try
            End While
            objRS.Close()
            objRS = Nothing
            bReturn = True
        Catch ex As Exception
            setError("LoadcolLgFactHBV", ex.ToString())
            bReturn = False
        End Try

        Debug.Assert(bReturn, "LoadcollgFactHBV" & getErreur())
        Return bReturn

    End Function ' LoadColLGFACTHBV
    ''' <summary>
    ''' Suppression des lignes d'une facture
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Function deletecolLgHBV() As Boolean
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "Id <> 0")
        Debug.Assert(m_typedonnee = vncEnums.vncTypeDonnee.FACTHBV, "Objet de type FactHBV requis")


        Dim sqlString As String = "DELETE FROM LIGNE_FACTHBV "
        Dim strClauseWhere As String = " WHERE LGFHBV_FACT_ID = ? "
        Dim objCommand As OleDbCommand
        '        Dim objParam As OleDbParameter
        Dim bReturn As Boolean

        'Clause where en Fonction du type de l'objet courant
        sqlString = sqlString & strClauseWhere

        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        CreateParameterP_ID(objCommand)
        Try
            objCommand.ExecuteNonQuery()

            bReturn = True

        Catch ex As Exception
            setError("deletecolLgHBV", ex.ToString())
            bReturn = False
        End Try
        Debug.Assert(bReturn, "deletecolLgHBV" & getErreur())
        Return bReturn
    End Function 'deletecolLgTRP
    ''' <summary>
    ''' Rend une collection de FactHBV
    ''' </summary>
    ''' <param name="strCode"></param>
    ''' <param name="strRSClient"></param>
    ''' <param name="pEtat"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Overloads Shared Function ListeFACTHBV(ByVal strCode As String, ByVal strRSClient As String, ByVal pEtat As vncEtatCommande) As Collection
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Dim colReturn As New Collection
        '        Dim objParam As OleDbParameter
        Dim sqlString As String = "SELECT FHBV_ID " & _
                                    "FROM FACTHBV, CLIENT "
        Dim strWhere As String = " FHBV_CLT_ID = CLT_ID"
        Dim objCommand As OleDbCommand
        Dim objFACT As FactHBV
        Dim objRS As OleDbDataReader = Nothing
        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        Dim bReturn As Boolean
        Dim objParam As OleDbParameter

        If strCode <> "" Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "FHBV_CODE LIKE ?"
            objParam = objCommand.Parameters.AddWithValue("?", strCode)

        End If

        If strRSClient <> "" Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "CLIENT.CLT_RS LIKE ?"
            objParam = objCommand.Parameters.AddWithValue("?", strRSClient)

        End If

        If pEtat <> vncEnums.vncEtatCommande.vncRien Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "FHBV_ETAT = ?"
            objParam = objCommand.Parameters.AddWithValue("?", pEtat)

        End If


        If strWhere <> "" Then
            sqlString = sqlString & " WHERE " & strWhere
        End If
        sqlString = sqlString & " ORDER BY FHBV_DATE DESC, CLT_CODE ASC"
        objCommand.CommandText = sqlString



        Try
            objRS = objCommand.ExecuteReader
            While (objRS.Read())
                Dim nId As Integer
                nId = getInteger(objRS, "FHBV_ID")

                objFACT = FactHBV.createandload(nId)
                If objFACT.id = nId Then
                    colReturn.Add(objFACT, CStr(objFACT.code))
                End If
            End While
            objRS.Close()
            objRS = Nothing
            bReturn = True
        Catch ex As Exception
            setError("ListeFACTHBV", sqlString & vbCrLf & ex.ToString())
            bReturn = False
            colReturn = Nothing
        End Try

        Debug.Assert(bReturn, getErreur())

        Return colReturn

    End Function 'ListeFACTHBV
    ''' <summary>
    ''' Rend une liste de factHBV en fonctino des date et de l'état
    ''' </summary>
    ''' <param name="pddeb"></param>
    ''' <param name="pdfin"></param>
    ''' <param name="pCodeClient"></param>
    ''' <param name="pEtat"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Shared Function ListeFACTHBVDate(Optional ByVal pddeb As Date = DATE_DEFAUT, Optional ByVal pdfin As Date = DATE_DEFAUT, Optional ByVal pCodeClient As String = "", Optional ByVal pEtat As vncEtatCommande = vncEnums.vncEtatCommande.vncRien) As Collection
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Dim colReturn As New Collection
        Dim colTemp As New Collection
        '        Dim objParam As OleDbParameter
        Dim sqlString As String = "SELECT " & _
                                    "FHBV_ID, CLT_ID " & _
                                  " FROM FACTHBV, CLIENT "

        Dim strWhere As String = " FACTHBV.FHBV_CLT_ID = CLIENT.CLT_ID "
        Dim objCommand As OleDbCommand
        Dim objFCT As FactHBV
        Dim objRS As OleDbDataReader = Nothing
        Dim strId As String
        Dim objParam As OleDbParameter
        Dim nId As Long




        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection


        If pddeb <> DATE_DEFAUT Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & " FHBV_DATE >=  ?"
            objParam = objCommand.Parameters.AddWithValue("?", pddeb)

        End If
        If pdfin <> DATE_DEFAUT Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & " FHBV_DATE <=  ?"
            objParam = objCommand.Parameters.AddWithValue("?", pdfin)

        End If
        If pCodeClient <> "" Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "CLIENT.CLT_CODE LIKE ?"
            objParam = objCommand.Parameters.AddWithValue("?", pCodeClient)

        End If

        If pEtat <> vncEnums.vncEtatCommande.vncRien Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "FHBV_ETAT = ?"
            objParam = objCommand.Parameters.AddWithValue("?", pEtat)

        End If



        If strWhere <> "" Then
            sqlString = sqlString & "WHERE " & strWhere
        End If
        sqlString = sqlString & " ORDER BY CLIENT.CLT_CODE ASC"
        objCommand.CommandText = sqlString



        Try
            objRS = objCommand.ExecuteReader
            While objRS.Read()
                Try
                    strId = GetString(objRS, "FHBV_ID")
                    colTemp.Add(strId)
                Catch ex As InvalidCastException
                    colReturn = Nothing
                    Exit While
                End Try
            End While
            objRS.Close()
            objRS = Nothing
            For Each strId In colTemp
                nId = CLng(strId)
                objFCT = FactHBV.createandload(nId)
                objFCT.resetBooleans()

                If objFCT.id = nId Then
                    colReturn.Add(objFCT, objFCT.code)
                End If
            Next
            Return colReturn
        Catch ex As Exception
            setError("ListeFACTHBVEtat", ex.ToString() & sqlString)
            colReturn = Nothing
        End Try
        Debug.Assert(Not colReturn Is Nothing, "ListeFACTHBVEtat: colReturn is nothing" & FactHBV.getErreur())
        Return colReturn
    End Function 'ListeFactTRPEtat
#Region "Paramètres FACTHBV"
    Private Sub CreateParameterP_FHBV_CODE(ByVal objCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", objCMD.code)

    End Sub
    Private Sub CreateParameterP_FHBV_ETAT(ByVal objCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", objCMD.etat.codeEtat)
    End Sub
    Private Sub CreateParameterP_FHBV_CLT_ID(ByVal objCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", objCMD.oTiers.id)

    End Sub
    Private Sub CreateParameterP_FHBV_DATE(ByVal objCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", objCMD.dateCommande.ToShortDateString)

    End Sub
     Private Sub CreateParameterP_FHBV_TOTAL_HT(ByVal objCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", CDbl(objCMD.totalHT))
    End Sub

    Private Sub CreateParameterP_FHBV_TOTAL_TTC(ByVal objCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", CDbl(objCMD.totalTTC))
    End Sub
    Private Sub CreateParameterP_FHBV_COM_FACT(ByVal objCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCMD.CommFacturation.comment, 50))
    End Sub
     Private Sub CreateParameterP_FHBV_IDMODEREGLEMENT(ByVal objCommand As OleDbCommand)
        Dim objCMD As FactHBV
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", objCMD.idModeReglement)
    End Sub
    Private Sub CreateParameterP_FHBV_DECHEANCE(ByVal objCommand As OleDbCommand)
        Dim objCMD As FactHBV
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", objCMD.dEcheance)
    End Sub
    Private Sub CreateParameterP_FHBV_IDCOMMANDE(ByVal objCommand As OleDbCommand)
        Dim objCMD As FactHBV
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", objCMD.idCommande)
    End Sub
#End Region

#End Region
#Region "FactColisage"
    Protected Function insertFactColisage() As Boolean
        '=======================================================================
        'Fonction : insertFactColisage
        'Description : Insertion d'une Facture de Transport
        'Retour : Rend Vrai si l'INSERT s'est correctement effectué
        '=======================================================================
        Dim bReturn As Boolean
        Dim objFACT As FactColisage
        bReturn = False

        Debug.Assert(Me.GetType().Name.Equals("FactColisage"), "Objet de Type 'FactColisage' Requis")
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id = 0, "ID=0")
        objFACT = CType(Me, FactColisage)

        Dim sqlString As String = "INSERT INTO FACTCOLISAGE( " & _
                                    "FCOL_CODE," & _
                                    "FCOL_ETAT," & _
                                    "FCOL_FRN_ID," & _
                                    "FCOL_DATE," & _
                                    "FCOL_TOTAL_TAXES," & _
                                    "FCOL_TOTAL_HT," & _
                                    "FCOL_TOTAL_TTC," & _
                                    "FCOL_PERIODE," & _
                                    "FCOL_COM_FACT," & _
                                    "FCOL_MONTANT_REGLEMENT," & _
                                    "FCOL_DATE_REGLEMENT," & _
                                    "FCOL_REF_REGLEMENT," & _
                                    "FCOL_IDMODEREGLEMENT," & _
                                    "FCOL_DECHEANCE," & _
                                    "FCOL_BINTERNET" & _
                                  " ) VALUES ( " & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? " & _
                                    " )"
        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        '        Dim objParam As OleDbParameter


        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString
        m_dbconn.BeginTransaction()
        objCommand.Transaction = m_dbconn.transaction


        CreateParameterP_FCOL_CODE(objCommand)
        CreateParameterP_FCOL_ETAT(objCommand)
        CreateParameterP_FCOL_FRN_ID(objCommand)
        CreateParameterP_FCOL_DATE(objCommand)
        CreateParameterP_FCOL_TOTAL_TAXES(objCommand)
        CreateParameterP_FCOL_TOTAL_HT(objCommand)
        CreateParameterP_FCOL_TOTAL_TTC(objCommand)
        CreateParameterP_FCOL_PERIODE(objCommand)
        CreateParameterP_FCOL_COM_FACT(objCommand)
        CreateParameterP_FCOL_MONTANT_REGLEMENT(objCommand)
        CreateParameterP_FCOL_DATE_REGLEMENT(objCommand)
        CreateParameterP_FCOL_REF_REGLEMENT(objCommand)
        CreateParameterP_FCOL_IDMODEREGLEMENT(objCommand)
        CreateParameterP_FCOL_DECHEANCE(objCommand)
        CreateParameterP_FCOL_BINTERNET(objCommand)
        Try
            objCommand.ExecuteNonQuery()
            objCommand = New OleDbCommand("SELECT MAX(FCOL_ID) FROM FACTCOLISAGE", m_dbconn.Connection)
            objCommand.Transaction = m_dbconn.transaction
            objRS = objCommand.ExecuteReader
            If (objRS.Read) Then
                m_id = objRS.GetInt32(0)
                cleanErreur()
                bReturn = True
            Else
                bReturn = False
            End If
            objRS.Close()
            objRS = Nothing

        Catch ex As Exception
            setError("InsertFACTCOM", ex.ToString())
            bReturn = False
        End Try

        If Not objRS Is Nothing Then
            objRS.Close()
        End If
        If bReturn Then
            m_dbconn.transaction.Commit()
        Else
            m_dbconn.transaction.Rollback()
        End If
        '    Debug.Assert(m_id <> 0, "ID=0")
        Debug.Assert(bReturn, "InsertFactColisage: " & getErreur())
        Return bReturn
    End Function 'insertFactColisage

    Protected Function deleteFactColisage() As Boolean
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "Id <> 0")
        Debug.Assert(Me.GetType().Name.Equals("FactColisage"))
        Dim bReturn As Boolean = False


        Dim sqlString As String = "DELETE FROM FactColisage WHERE FCOL_ID=? "
        Dim objCommand As OleDbCommand
        Dim objFACT As FactColisage
        '        Dim objParam As OleDbParameter

        objFACT = CType(Me, FactColisage)
        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        CreateParameterP_ID(objCommand)
        Try
            objCommand.ExecuteNonQuery()
            m_id = 0
            objFACT.resetCode()
            bReturn = True

        Catch ex As Exception
            setError("DeleteFactColisage", ex.ToString())
            bReturn = False
        End Try

        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'DeleteFactColisage
    Protected Function loadFactColisage() As Boolean

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "L'id doit être renseigné")

        Dim sqlString As String = "SELECT " & _
                                    "FCOL_CODE," & _
                                    "FCOL_ETAT," & _
                                    "FCOL_FRN_ID," & _
                                    "FCOL_DATE," & _
                                    "FCOL_TOTAL_HT," & _
                                    "FCOL_TOTAL_TTC," & _
                                    "FCOL_TOTAL_TAXES," & _
                                    "FCOL_PERIODE," & _
                                    "FCOL_COM_FACT," & _
                                    "FCOL_MONTANT_REGLEMENT," & _
                                    "FCOL_DATE_REGLEMENT," & _
                                    "FCOL_REF_REGLEMENT," & _
                                    "FCOL_IDMODEREGLEMENT," & _
                                    "FCOL_DECHEANCE," & _
                                    "FCOL_BINTERNET" & _
                                    " FROM FactColisage WHERE " & _
                                   " FCOL_ID = ?"

        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        Dim objFCOL As FactColisage
        '        Dim objParam As OleDbParameter
        Dim bReturn As Boolean

        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        'Si le paramètre existe on ne met à jour que sa valeur
        'Autrement on le crée
        CreateParameterP_ID(objCommand)

        Try
            objRS = objCommand.ExecuteReader
            If Not objRS.HasRows Then
                objRS.Close()
                m_id = 0
                Return False
            End If
            objRS.Read()
            objFCOL = CType(Me, FactColisage)
            Try
                objFCOL.code = GetString(objRS, "FCOL_CODE")
            Catch ex As InvalidCastException
                objFCOL.code = ""
            End Try
            Try
                objFCOL.setEtat(GetString(objRS, "FCOL_ETAT"))
            Catch ex As InvalidCastException
                objFCOL.setEtat(vncEnums.vncEtatCommande.vncFactCOLGeneree)
            End Try
            Try
                objFCOL.dateCommande = GetString(objRS, "FCOL_DATE")
            Catch ex As InvalidCastException
                objFCOL.dateCommande = DATE_DEFAUT
            End Try
            Try 'PERIODE
                objFCOL.periode = GetString(objRS, "FCOL_PERIODE")
            Catch ex As InvalidCastException
                objFCOL.periode = ""
            End Try
            Try 'Total Taxes
                objFCOL.montantTaxes = GetString(objRS, "FCOL_TOTAL_TAXES")
            Catch ex As InvalidCastException
                objFCOL.montantTaxes = 0
            End Try
            Try 'Total HT
                objFCOL.totalHT = GetString(objRS, "FCOL_TOTAL_HT")
            Catch ex As InvalidCastException
                objFCOL.totalHT = 0
            End Try
            Try 'Total TTC
                objFCOL.totalTTC = GetString(objRS, "FCOL_TOTAL_TTC")
            Catch ex As InvalidCastException
                objFCOL.totalTTC = 0
            End Try
            Try
                objFCOL.montantReglement = GetString(objRS, "FCOL_MONTANT_REGLEMENT")
            Catch ex As InvalidCastException
                objFCOL.montantReglement = 0
            End Try
            Try
                objFCOL.dateReglement = GetString(objRS, "FCOL_DATE_REGLEMENT")
            Catch ex As InvalidCastException
                objFCOL.dateReglement = DATE_DEFAUT
            End Try
            Try
                objFCOL.refReglement = GetString(objRS, "FCOL_REF_REGLEMENT")
            Catch ex As InvalidCastException
                objFCOL.refReglement = ""
            End Try
            Try
                objFCOL.idModeReglement = GetString(objRS, "FCOL_IDMODEREGLEMENT")
            Catch ex As InvalidCastException
                objFCOL.idModeReglement = 0
            End Try
            Try
                objFCOL.bExportInternet = GetValue(objRS, "FCOL_BINTERNET")
            Catch ex As InvalidCastException
                objFCOL.bExportInternet = False
            End Try
            'Chargement de la date d'échéance après le mode de reglement pour éviter le recalcul
            Try
                objFCOL.dEcheance = GetString(objRS, "FCOL_DECHEANCE")
            Catch ex As InvalidCastException
                objFCOL.dEcheance = DATE_DEFAUT
            End Try
            Try
                objFCOL.oTiers.setid(GetString(objRS, "FCOL_FRN_ID"))
            Catch ex As InvalidCastException
                objFCOL.oTiers.setid(0)
            End Try
            Try
                objFCOL.CommFacturation.comment = GetString(objRS, "FCOL_COM_FACT")
            Catch ex As InvalidCastException
                objFCOL.CommFacturation.comment = ""
            End Try

            cleanErreur()
            bReturn = True
        Catch ex As Exception
            setError("LoadFactColisage", ex.ToString())
            bReturn = False
        End Try
        If (Not objRS Is Nothing) Then
            objRS.Close()
            objRS = Nothing
        End If
        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'LoadFactColisage
    Protected Function UpdateFactColisage() As Boolean
        '=======================================================================
        'Fonction : UpdateFACTCOlisage
        'Description : Mise à jour  d'une Facture de colisage
        'Retour : Rend Vrai si l'UPDATE s'est correctement effectué
        '=======================================================================
        Dim bReturn As Boolean
        Dim objFact As FactColisage
        bReturn = False

        Debug.Assert(Me.GetType().Name.Equals("FactColisage"), "Objet de Type FactColisage Requis")
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "ID<>0")
        objFact = CType(Me, FactColisage)

        Dim sqlString As String = "UPDATE FactColisage SET " & _
                                    "FCOL_CODE = ? ," & _
                                    "FCOL_ETAT = ? ," & _
                                    "FCOL_FRN_ID = ? ," & _
                                    "FCOL_DATE = ? ," & _
                                    "FCOL_TOTAL_TAXES = ? ," & _
                                    "FCOL_TOTAL_HT = ? ," & _
                                    "FCOL_TOTAL_TTC = ? ," & _
                                    "FCOL_PERIODE = ? ," & _
                                    "FCOL_COM_FACT = ? ," & _
                                    "FCOL_MONTANT_REGLEMENT = ? ," & _
                                    "FCOL_DATE_REGLEMENT = ? ," & _
                                    "FCOL_REF_REGLEMENT = ? , " & _
                                    "FCOL_IDMODEREGLEMENT = ?, " & _
                                    "FCOL_DECHEANCE = ?, " & _
                                    "FCOL_BINTERNET = ? " & _
                                    " WHERE FCOL_ID = ?"

        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        '        Dim objParam As OleDbParameter



        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString
        m_dbconn.BeginTransaction()
        objCommand.Transaction = m_dbconn.transaction

        CreateParameterP_FCOL_CODE(objCommand)
        CreateParameterP_FCOL_ETAT(objCommand)
        CreateParameterP_FCOL_FRN_ID(objCommand)
        CreateParameterP_FCOL_DATE(objCommand)
        CreateParameterP_FCOL_TOTAL_TAXES(objCommand)
        CreateParameterP_FCOL_TOTAL_HT(objCommand)
        CreateParameterP_FCOL_TOTAL_TTC(objCommand)
        CreateParameterP_FCOL_PERIODE(objCommand)
        CreateParameterP_FCOL_COM_FACT(objCommand)
        CreateParameterP_FCOL_MONTANT_REGLEMENT(objCommand)
        CreateParameterP_FCOL_DATE_REGLEMENT(objCommand)
        CreateParameterP_FCOL_REF_REGLEMENT(objCommand)
        CreateParameterP_FCOL_IDMODEREGLEMENT(objCommand)
        CreateParameterP_FCOL_DECHEANCE(objCommand)
        CreateParameterP_FCOL_BINTERNET(objCommand)
        CreateParameterP_ID(objCommand)
        Try
            objCommand.ExecuteNonQuery()
            m_dbconn.transaction.Commit()
            bReturn = True

        Catch ex As Exception
            setError("UpdateFactColisage", ex.ToString())
            m_dbconn.transaction.Rollback()
            bReturn = False
        End Try

        Debug.Assert(bReturn, "UpdateFactColisage" & getErreur())
        Return bReturn
    End Function 'UpdateFactColisage

    '==========================================================================
    'Methode : deletecolLGCOL
    'Description : Suppression des lignes d'une Facture de transport
    Protected Function deletecolLgCOL() As Boolean
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "Id <> 0")
        Debug.Assert(m_typedonnee = vncEnums.vncTypeDonnee.FACTCOL, "Objet de type FactColisage requis")


        Dim sqlString As String = "DELETE FROM LGFactColisage "
        Dim strClauseWhere As String = " WHERE LGFactColisage.LGCOL_FACTCOL_ID = ? "
        Dim objCommand As OleDbCommand
        '        Dim objParam As OleDbParameter
        Dim bReturn As Boolean

        'Clause where en Fonction du type de l'objet courant
        sqlString = sqlString & strClauseWhere

        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        CreateParameterP_ID(objCommand)
        Try
            objCommand.ExecuteNonQuery()

            bReturn = True

        Catch ex As Exception
            setError("deletecolLgCOL", ex.ToString())
            bReturn = False
        End Try
        Debug.Assert(bReturn, "deletecolLgCOL" & getErreur())
        Return bReturn
    End Function 'deletecolLgCOL

    Protected Function INSERTcolLGCOL() As Boolean
        Dim objFactColisage As Commande
        Dim bReturn As Boolean
        Dim objLgCOL As LgFactColisage

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "Le Client doit être Sauvegardé")
        Debug.Assert(m_typedonnee = vncEnums.vncTypeDonnee.FACTCOL, "Objet de type FactColisage requis")
        objFactColisage = Me
        Debug.Assert(Not objFactColisage.colLignes Is Nothing, "ColLignes is Nothing")


        Dim sqlString As String = "INSERT INTO LGFACTCOLISAGE (" & _
                                    "LGCOL_NUM," & _
                                    "LGCOL_FACTCOL_ID," & _
                                    "LGCOL_DDEB," & _
                                    "LGCOL_DFIN," & _
                                    "LGCOL_STK_INI," & _
                                    "LGCOL_STK_IN," & _
                                    "LGCOL_STK_OUT," & _
                                    "LGCOL_STK_FINAL," & _
                                    "LGCOL_QTE," & _
                                    "LGCOL_PU," & _
                                    "LGCOL_MT_HT" & _
                                    ") VALUES ( " & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? " & _
                                  " )"
        Dim sqlString3 As String = "SELECT MAX(LGCOL_ID) FROM LGFactColisage"
        Dim objCommand As OleDbCommand
        Dim objCommandID As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        '        Dim objParam As OleDbParameter

        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString
        m_dbconn.BeginTransaction()
        objCommand.Transaction = m_dbconn.transaction



        objCommandID = New OleDbCommand
        objCommandID.Connection = m_dbconn.Connection
        objCommandID.CommandText = sqlString3
        objCommandID.Transaction = m_dbconn.transaction

        bReturn = True
        Try
            'Commande 1
            Dim oParam1 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.VarChar)
            Dim oParam2 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Integer)
            Dim oParam3 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.VarChar)
            Dim oParam4 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.VarChar)
            Dim oParam5 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Integer)
            Dim oParam6 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Integer)
            Dim oParam7 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Integer)
            Dim oParam8 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Integer)
            Dim oParam9 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Integer)
            Dim oParam10 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Double)
            Dim oParam11 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Double)

            For Each objLgCOL In objFactColisage.colLignes

                oParam1.Value = objLgCOL.num
                oParam2.Value = objFactColisage.id
                oParam3.Value = objLgCOL.dDeb
                oParam4.Value = objLgCOL.dFin
                oParam5.Value = objLgCOL.StockInitial
                oParam6.Value = objLgCOL.Entrees
                oParam7.Value = objLgCOL.Sorties
                oParam8.Value = objLgCOL.StockFinal
                oParam9.Value = objLgCOL.qte
                oParam10.Value = CDbl(objLgCOL.pu)
                oParam11.Value = CDbl(objLgCOL.MontantHT)
                objCommand.ExecuteNonQuery()
                objRS = objCommandID.ExecuteReader
                If objRS.Read Then
                    objLgCOL.setid(objRS.GetInt32(0))
                    cleanErreur()
                    bReturn = True
                Else
                    setError("InsertColLGCOL", "NO LGCOLID")
                    bReturn = False
                End If
                objRS.Close()
                objRS = Nothing

            Next
        Catch ex As Exception
            setError("InsertcolLGCOL", ex.ToString())
            bReturn = False
        End Try

        If bReturn Then
            m_dbconn.transaction.Commit()
        Else
            m_dbconn.transaction.Rollback()
        End If

        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function ' insertcolLGCMD

    Protected Function UPDATEcolLGCOL() As Boolean
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "La Commande.id  <> 0")
        Debug.Assert(m_typedonnee = vncEnums.vncTypeDonnee.FACTCOL, "Objet de type FactColisage requis")


        Dim sqlString As String = "UPDATE LGFactColisage SET " & _
                                    "LGCOL_NUM = ? ," & _
                                    "LGCOL_FACTCOL_ID = ? ," & _
                                    "LGCOL_DDEB= ? , " & _
                                    "LGCOL_DFIN= ? , " & _
                                    "LGCOL_STK_INI= ? , " & _
                                    "LGCOL_STK_IN= ? , " & _
                                    "LGCOL_STK_OUT= ? , " & _
                                    "LGCOL_STK_FINAL= ? , " & _
                                    "LGCOL_QTE= ? , " & _
                                    "LGCOL_PU= ? , " & _
                                    "LGCOL_MT_HT = ? " & _
                                    " WHERE " & _
                                    " LGCOL_ID = ?"
        Dim objCommand As OleDbCommand
        Dim objFact As FactColisage
        '        Dim objParam As OleDbParameter
        Dim bReturn As Boolean
        Dim objLgCOL As LgFactColisage


        objFact = CType(Me, FactColisage)
        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString
        m_dbconn.BeginTransaction()
        objCommand.Transaction = m_dbconn.transaction



        bReturn = True
        For Each objLgCOL In objFact.colLignes
            Try
                objCommand.Parameters.Clear()
                objCommand.Parameters.AddWithValue("?", objLgCOL.num)
                objCommand.Parameters.AddWithValue("?", m_id)
                objCommand.Parameters.AddWithValue("?", objLgCOL.dDeb)
                objCommand.Parameters.AddWithValue("?", objLgCOL.dFin)
                objCommand.Parameters.AddWithValue("?", objLgCOL.StockInitial)
                objCommand.Parameters.AddWithValue("?", objLgCOL.Entrees)
                objCommand.Parameters.AddWithValue("?", objLgCOL.Sorties)
                objCommand.Parameters.AddWithValue("?", objLgCOL.StockFinal)
                objCommand.Parameters.AddWithValue("?", objLgCOL.qte)
                objCommand.Parameters.AddWithValue("?", objLgCOL.pu)
                objCommand.Parameters.AddWithValue("?", objLgCOL.MontantHT)
                objCommand.Parameters.AddWithValue("?", objLgCOL.id)
                objCommand.ExecuteNonQuery()
            Catch ex As Exception
                setError("UpdatecolLGCOL", ex.ToString())
                bReturn = False
                Exit For
            End Try
        Next objLgCOL

        If bReturn Then
            m_dbconn.transaction.Commit()
        Else
            m_dbconn.transaction.Rollback()
        End If

        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function ' UPDATEcolLGCOL
    ''' <summary>
    ''' Chargement de la liste des lignes d'une facture de Colisage
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Function LoadcolLgFactColisage() As Boolean
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "Id <> 0")
        Debug.Assert(Me.GetType().Name.Equals("FactColisage"))

        Dim bReturn As Boolean
        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        Dim objFactColisage As FactColisage
        Dim objLGCOL As LgFactColisage

        Dim sqlString As String = "SELECT " & _
                                    "LGCOL_ID," & _
                                    "LGCOL_NUM," & _
                                    "LGCOL_FACTCOL_ID," & _
                                    "LGCOL_DDEB," & _
                                    "LGCOL_DFIN," & _
                                    "LGCOL_STK_INI," & _
                                    "LGCOL_STK_IN," & _
                                    "LGCOL_STK_OUT," & _
                                    "LGCOL_STK_FINAL," & _
                                    "LGCOL_QTE," & _
                                    "LGCOL_PU," & _
                                    "LGCOL_MT_HT " & _
                                  " FROM LGFactColisage" & _
                                  " WHERE LGCOL_FACTCOL_ID = ? ORDER BY LGCOL_NUM"

        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        'Paramétre ClientID 
        CreateParameterP_ID(objCommand)

        Try
            objFactColisage = CType(Me, FactColisage)
            objRS = objCommand.ExecuteReader
            While objRS.Read
                Try
                    objLGCOL = New LgFactColisage
                    Try
                        objLGCOL.setid(GetString(objRS, "LGCOL_ID"))
                    Catch ex As InvalidCastException
                        objLGCOL.setid(0)
                    End Try
                    Try
                        objLGCOL.num = GetString(objRS, "LGCOL_NUM")
                    Catch ex As InvalidCastException
                        objLGCOL.num = ""
                    End Try
                    Try
                        objLGCOL.idFactColisage = GetString(objRS, "LGCOL_FACTCOL_ID")
                    Catch ex As InvalidCastException
                        objLGCOL.idFactColisage = 0
                    End Try
                    Try
                        objLGCOL.dDeb = GetString(objRS, "LGCOL_DDEB")
                    Catch ex As InvalidCastException
                        objLGCOL.dDeb = DATE_DEFAUT
                    End Try
                    Try
                        objLGCOL.dFin = GetString(objRS, "LGCOL_DFIN")
                    Catch ex As InvalidCastException
                        objLGCOL.dFin = DATE_DEFAUT
                    End Try
                    Try
                        objLGCOL.StockInitial = GetString(objRS, "LGCOL_STK_INI")
                    Catch ex As InvalidCastException
                        objLGCOL.StockInitial = 0
                    End Try
                    Try
                        objLGCOL.StockFinal = GetString(objRS, "LGCOL_STK_FINAL")
                    Catch ex As InvalidCastException
                        objLGCOL.StockFinal = 0
                    End Try
                    Try
                        objLGCOL.Entrees = GetString(objRS, "LGCOL_STK_IN")
                    Catch ex As InvalidCastException
                        objLGCOL.Entrees = 0
                    End Try
                    Try
                        objLGCOL.Sorties = GetString(objRS, "LGCOL_STK_OUT")
                    Catch ex As InvalidCastException
                        objLGCOL.Sorties = 0
                    End Try
                    Try
                        objLGCOL.qte = GetString(objRS, "LGCOL_QTE")
                    Catch ex As InvalidCastException
                        objLGCOL.qte = 0
                    End Try
                    Try
                        objLGCOL.pu = GetString(objRS, "LGCOL_PU")
                    Catch ex As InvalidCastException
                        objLGCOL.pu = ""
                    End Try
                    Try
                        objLGCOL.MontantHT = GetString(objRS, "LGCOL_MT_HT")
                    Catch ex As InvalidCastException
                        objLGCOL.MontantHT = 0
                    End Try
                    objFactColisage.resetBooleans()
                    objFactColisage.AjouteLigneFactColisage(objLGCOL, False)
                Catch ex As InvalidCastException
                    bReturn = False
                    Exit While
                End Try
            End While
            objRS.Close()
            objRS = Nothing
            bReturn = True
        Catch ex As Exception
            setError("LoadcolLgFactColisage", ex.ToString())
            bReturn = False
        End Try

        Debug.Assert(bReturn, "LoadcollgFactColisage" & getErreur())
        Return bReturn

    End Function ' LoadColLGFactColisage
    ''' <summary>
    ''' REnd une collection de facture de colisage correspondant au critère
    ''' </summary>
    ''' <param name="pddeb">date de Facture (inclus)</param>
    ''' <param name="pdfin">Date de Facture (inclus)</param>
    ''' <param name="pCodeFournisseur">Code Fournisseur (Optionel)</param>
    ''' <param name="pEtat">Etat de la Facture (Par defaut tous)</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Shared Function ListeFACTColisage(Optional ByVal pddeb As Date = DATE_DEFAUT, Optional ByVal pdfin As Date = DATE_DEFAUT, Optional ByVal pCodeFournisseur As String = "", Optional ByVal pEtat As vncEtatCommande = vncEnums.vncEtatCommande.vncRien) As Collection
        '============================================================================
        'Function : ListeFACTTRPEtat
        'Description : Rend une liste de Facture de transport en fonction leur état
        '============================================================================
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Dim colReturn As New Collection
        Dim colTemp As New Collection
        '        Dim objParam As OleDbParameter
        Dim sqlString As String = "SELECT " & _
                                    "FCOL_ID, FRN_ID " & _
                                  " FROM FACTCOLISAGE INNER JOIN FOURNISSEUR ON FACTCOLISAGE.FCOL_FRN_ID = FOURNISSEUR.FRN_ID "

        Dim strWhere As String = ""
        Dim objCommand As OleDbCommand
        Dim objFCT As FactColisage
        Dim objRS As OleDbDataReader = Nothing
        Dim strId As String
        Dim objParam As OleDbParameter
        Dim nId As Long



        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection


        If pddeb <> DATE_DEFAUT Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & " FCOL_DATE >=  ?"
            objParam = objCommand.Parameters.AddWithValue("?", pddeb)

        End If
        If pdfin <> DATE_DEFAUT Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & " FCOL_DATE <=  ?"
            objParam = objCommand.Parameters.AddWithValue("?", pdfin)

        End If
        If pCodeFournisseur <> "" Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "FOURNISSEUR.FRN_CODE LIKE ?"
            objParam = objCommand.Parameters.AddWithValue("?", pCodeFournisseur)

        End If

        If pEtat <> vncEnums.vncEtatCommande.vncRien Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "FCOL_ETAT = ?"
            objParam = objCommand.Parameters.AddWithValue("?", pEtat)

        End If



        If strWhere <> "" Then
            sqlString = sqlString & "WHERE " & strWhere
        End If
        sqlString = sqlString & " ORDER BY FOURNISSEUR.FRN_CODE ASC"
        objCommand.CommandText = sqlString



        Try
            objRS = objCommand.ExecuteReader
            While objRS.Read()
                Try
                    strId = GetString(objRS, "FCOL_ID")
                    colTemp.Add(strId)
                Catch ex As InvalidCastException
                    colReturn = Nothing
                    Exit While
                End Try
            End While
            objRS.Close()
            objRS = Nothing
            For Each strId In colTemp
                nId = CLng(strId)
                objFCT = FactColisage.createandload(nId)
                objFCT.resetBooleans()

                If objFCT.id <> 0 Then
                    colReturn.Add(objFCT, objFCT.code)
                End If
            Next
            Return colReturn
        Catch ex As Exception
            setError("ListeFACTColisage", ex.ToString() & sqlString)
            colReturn = Nothing
        End Try
        Debug.Assert(Not colReturn Is Nothing, "ListeFACTColisage: colReturn is nothing" & FactTRP.getErreur())
        Return colReturn
    End Function 'ListeFactColisage
    ''' <summary>
    ''' REnd une collection de facture de colisage correspondant au critère
    ''' </summary>
    ''' <param name="strCode">code de la facture</param>
    ''' <param name="strRSFournisseur">Raison Sociale du fournisseur</param>
    ''' <param name="pEtat">Etat de la Facture (Par defaut tous)</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Shared Function ListeFACTColisage(ByVal strCode As String, ByVal strRSFournisseur As String, Optional ByVal pEtat As vncEtatCommande = vncEtatCommande.vncRien) As Collection
        '============================================================================
        'Function : ListeFACTTRPEtat
        'Description : Rend une liste de Facture de transport en fonction leur état
        '============================================================================
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Dim colReturn As New Collection
        Dim colTemp As New Collection
        '        Dim objParam As OleDbParameter
        Dim sqlString As String = "SELECT " & _
                                    "FCOL_ID, FRN_ID " & _
                                  " FROM FACTCOLISAGE INNER JOIN FOURNISSEUR ON FACTCOLISAGE.FCOL_FRN_ID = FOURNISSEUR.FRN_ID "

        Dim strWhere As String = ""
        Dim objCommand As OleDbCommand
        Dim objFCT As FactColisage
        Dim objRS As OleDbDataReader = Nothing
        Dim strId As String
        Dim nId As Long
        Dim objParam As OleDbParameter



        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection


        If Not String.IsNullOrEmpty(strCode) Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & " FCOL_CODE =  ?"
            objParam = objCommand.Parameters.AddWithValue("?", strCode)

        End If
        If Not String.IsNullOrEmpty(strRSFournisseur) Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & " FRN_RS LIKE  ?"
            objParam = objCommand.Parameters.AddWithValue("?", strRSFournisseur)

        End If

        If pEtat <> vncEnums.vncEtatCommande.vncRien Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "FCOL_ETAT = ?"
            objParam = objCommand.Parameters.AddWithValue("?", pEtat)

        End If



        If strWhere <> "" Then
            sqlString = sqlString & "WHERE " & strWhere
        End If
        sqlString = sqlString & " ORDER BY FCOL_DATE DESC"
        objCommand.CommandText = sqlString



        Try
            objRS = objCommand.ExecuteReader
            While objRS.Read()
                Try
                    strId = GetString(objRS, "FCOL_ID")
                    colTemp.Add(strId)
                Catch ex As InvalidCastException
                    colReturn = Nothing
                    Exit While
                End Try
            End While
            objRS.Close()
            objRS = Nothing
            For Each strId In colTemp
                nId = CType(strId, Long)
                objFCT = FactColisage.createandload(nId)
                objFCT.resetBooleans()

                If objFCT.id <> 0 Then
                    colReturn.Add(objFCT, objFCT.code)
                End If
            Next
            Return colReturn
        Catch ex As Exception
            setError("ListeFACTColisage", ex.ToString() & sqlString)
            colReturn = Nothing
        End Try
        Debug.Assert(Not colReturn Is Nothing, "ListeFACTColisage: colReturn is nothing" & FactTRP.getErreur())
        Return colReturn
    End Function 'ListeFactColisage
    ''' <summary>
    ''' Liste des facture de colisage non reglee
    ''' </summary>
    ''' <param name="strCode"></param>
    ''' <param name="strRSFournisseur"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Overloads Shared Function ListeFACTColisageNonReglee(ByVal strCode As String, ByVal strRSFournisseur As String) As Collection
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Dim colReturn As New Collection
        '        Dim objParam As OleDbParameter
        Dim sqlString As String = "SELECT FCOL_ID, FCOL_CODE, FCOL_DATE, FRN_ID , FRN_CODE, FRN_RS, FCOL_TOTAL_HT , FCOL_TOTAL_TTC, FCOL_BINTERNET " & _
                                    "FROM RQ_FACTURES, FACTCOLISAGE , FOURNISSEUR "
        Dim strWhere As String = " FACTCOLISAGE.FCOL_FRN_ID = FOURNISSEUR.FRN_ID AND RQ_FACTURES.FACT_ID = FCOL_ID AND RQ_FACTURES.FACT_TYPEFACT = " & vncEnums.vncTypeDonnee.FACTCOL & _
        " AND (RQ_FACTURES.FACT_TOTAL_REGLEMENT IS NULL OR RQ_FACTURES.FACT_TOTAL_REGLEMENT < RQ_FACTURES.FACT_TOTAL_TTC)"
        Dim objCommand As OleDbCommand
        Dim objFACT As FactColisage
        Dim objRS As OleDbDataReader = Nothing
        Dim objParam As OleDbParameter
        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection

        If strCode <> "" Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "FCOL_CODE LIKE ?"
            objParam = objCommand.Parameters.AddWithValue("?", strCode)

        End If

        If strRSFournisseur <> "" Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & "FOURNISSEUR.FRN_RS LIKE ?"
            objParam = objCommand.Parameters.AddWithValue("?", strRSFournisseur)

        End If


        If strWhere <> "" Then
            sqlString = sqlString & " WHERE " & strWhere
        End If
        sqlString = sqlString & " ORDER BY FCOL_DATE DESC, FRN_CODE ASC"
        'sqlString = sqlString & " ORDER BY FRN_CODE ASC"
        objCommand.CommandText = sqlString



        Try
            objRS = objCommand.ExecuteReader
            While (objRS.Read())
                objFACT = New FactColisage
                objFACT.m_bResume = True ' C'est un objet Résumé
                Try
                    objFACT.m_id = GetString(objRS, "FCOL_ID")
                Catch ex As InvalidCastException
                    objFACT.m_id = 0
                End Try
                Try
                    objFACT.code = GetString(objRS, "FCOL_CODE")
                Catch ex As InvalidCastException
                    objFACT.code = ""
                End Try
                Try
                    objFACT.oTiers.setid(GetString(objRS, "FRN_ID"))
                Catch ex As InvalidCastException
                    objFACT.oTiers.setid(0)
                End Try
                Try
                    objFACT.oTiers.code = GetString(objRS, "FRN_CODE")
                Catch ex As InvalidCastException
                    objFACT.oTiers.code = ""
                End Try
                Try
                    objFACT.oTiers.rs = GetString(objRS, "FRN_RS")
                Catch ex As InvalidCastException
                    objFACT.oTiers.nom = ""
                End Try

                Try
                    objFACT.dateCommande = GetString(objRS, "FCOL_DATE")
                Catch ex As InvalidCastException
                    objFACT.dateCommande = Now()
                End Try

                Try
                    objFACT.totalHT = GetString(objRS, "FCOL_TOTAL_HT")
                Catch ex As InvalidCastException
                    objFACT.totalHT = 0
                End Try
                Try
                    objFACT.totalTTC = GetString(objRS, "FCOL_TOTAL_TTC")
                Catch ex As InvalidCastException
                    objFACT.totalTTC = 0
                End Try
                Try
                    objFACT.bExportInternet = GetValue(objRS, "FCOL_BINTERNET")
                Catch ex As InvalidCastException
                    objFACT.bExportInternet = False
                End Try
                objFACT.resetBooleans()
                colReturn.Add(objFACT, CStr(objFACT.code))
            End While
            objRS.Close()
            Return colReturn
        Catch ex As Exception
            setError("ListeFACTCOMNonReglee", ex.ToString())
            Return Nothing
        End Try
    End Function 'ListeFACTCOMNonReglee

#End Region
#Region "Paramètres FACTColisage"
    Private Sub CreateParameterP_FCOL_CODE(ByVal objCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", objCMD.code)

    End Sub
    Private Sub CreateParameterP_FCOL_ETAT(ByVal objCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", objCMD.etat.codeEtat)
    End Sub
    Private Sub CreateParameterP_FCOL_FRN_ID(ByVal objCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", objCMD.oTiers.id)

    End Sub
    Private Sub CreateParameterP_FCOL_DATE(ByVal objCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", objCMD.dateCommande.ToShortDateString)

    End Sub
    Private Sub CreateParameterP_FCOL_TOTAL_TAXES(ByVal objCommand As OleDbCommand)
        Dim objCMD As FactColisage
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", objCMD.montantTaxes)
    End Sub
    Private Sub CreateParameterP_FCOL_TOTAL_HT(ByVal objCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", CDbl(objCMD.totalHT))
    End Sub

    Private Sub CreateParameterP_FCOL_TOTAL_TTC(ByVal objCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", CDbl(objCMD.totalTTC))
    End Sub
    Private Sub CreateParameterP_FCOL_COM_FACT(ByVal objCommand As OleDbCommand)
        Dim objCMD As Commande
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", objCMD.CommFacturation.comment)
    End Sub
    Private Sub CreateParameterP_FCOL_PERIODE(ByVal objCommand As OleDbCommand)
        Dim objCMD As FactColisage
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCMD.periode, 50))
    End Sub
    Private Sub CreateParameterP_FCOL_MONTANT_REGLEMENT(ByVal objCommand As OleDbCommand)
        Dim objCMD As FactColisage
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", CDbl(objCMD.montantReglement))

    End Sub
    Private Sub CreateParameterP_FCOL_DATE_REGLEMENT(ByVal objCommand As OleDbCommand)

        Dim objCMD As FactColisage
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", objCMD.dateReglement)

    End Sub
    Private Sub CreateParameterP_FCOL_REF_REGLEMENT(ByVal objCommand As OleDbCommand)
        Dim objCMD As FactColisage
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", truncate(objCMD.refReglement, 50))
    End Sub
    Private Sub CreateParameterP_FCOL_IDMODEREGLEMENT(ByVal objCommand As OleDbCommand)
        Dim objCMD As FactColisage
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", objCMD.idModeReglement)
    End Sub
    Private Sub CreateParameterP_FCOL_DECHEANCE(ByVal objCommand As OleDbCommand)
        Dim objCMD As FactColisage
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", objCMD.dEcheance)
    End Sub
    Private Sub CreateParameterP_FCOL_BINTERNET(ByVal objCommand As OleDbCommand)
        Dim objCMD As FactColisage
        objCMD = Me
        objCommand.Parameters.AddWithValue("?", objCMD.bExportInternet)
    End Sub
#End Region

#Region "Users et Right"
    Shared Function listeUSERS() As Collection
        '=================================================================
        ' Function : loadcolUsers 
        'Description : Changement d'une collection de users
        '=================================================================

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")

        Dim sqlString As String = " SELECT USR_ID,USR_CODE, USR_PASSWORD, USR_ROLE " & _
                                  "  FROM USERS "

        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        Dim objUSer As aut_user
        '        Dim objParam As OleDbParameter
        Dim nid As Integer
        Dim colReturn As New Collection

        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString

        Try
            objRS = objCommand.ExecuteReader
            While objRS.Read
                nid = getInteger(objRS, "USR_ID")
                objUSer = New aut_user
                objUSer.setid(nid)
                Try
                    objUSer.code = GetString(objRS, "USR_CODE")
                Catch ex As InvalidCastException
                    objUSer.code = ""
                End Try
                Try
                    objUSer.password = GetString(objRS, "USR_PASSWORD")
                Catch ex As InvalidCastException
                    objUSer.password = ""
                End Try

                Try
                    objUSer.setRole(GetString(objRS, "USR_ROLE"))
                Catch ex As InvalidCastException
                    objUSer.setRole(vncEnums.userRole.INVITE)
                End Try

                If objUSer.id <> 0 Then
                    objUSer.resetBooleans()
                    colReturn.Add(objUSer, objUSer.code)
                End If
            End While
            objRS.Close()
            'Chargement des droits pour chaque user
            For Each objUSer In colReturn
                objUSer.loadcolUSERSRIGHTS()
            Next objUSer

        Catch ex As Exception
            setError("LoadcolUSERs", ex.ToString())
            colReturn = Nothing
        End Try
        Debug.Assert(Not colReturn Is Nothing, getErreur())
        Return colReturn
    End Function 'LoadcolUSERS
    Protected Function loadUSER() As Boolean

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "L'id doit être renseigné")

        Dim sqlString As String = " SELECT USR_CODE, USR_PASSWORD, USR_ROLE " & _
                                  "  FROM USERS " & _
                                  " WHERE USR_ID = ?"

        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        Dim objUSer As aut_user
        '        Dim objParam As OleDbParameter
        Dim bReturn As Boolean

        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString
        'Si le paramètre existe on ne met à jour que sa valeur
        'Autrement on le crée
        CreateParameterP_ID(objCommand)

        Try
            objRS = objCommand.ExecuteReader
            objRS.Read()
            If Not objRS.HasRows Then
                objRS.Close()
                Return False
            End If
            objUSer = CType(Me, aut_user)
            Try
                objUSer.code = GetString(objRS, "USR_CODE")
            Catch ex As InvalidCastException
                objUSer.code = ""
            End Try
            Try
                objUSer.password = GetString(objRS, "USR_PASSWORD")
            Catch ex As InvalidCastException
                objUSer.password = ""
            End Try

            Try
                objUSer.setRole(GetString(objRS, "USR_ROLE"))
            Catch ex As InvalidCastException
                objUSer.setRole(vncEnums.userRole.INVITE)
            End Try

            bReturn = True

        Catch ex As Exception
            setError("LoadUSER", ex.ToString())
            bReturn = False
        End Try
        If Not objRS Is Nothing Then
            objRS.Close()
        End If
        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'LoadUSER
    Protected Function loadRIGHT() As Boolean

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "L'id doit être renseigné")

        Dim sqlString As String = " SELECT RGH_TAG, RGH_ROLE, RGH_DROIT, RGH_TEXT" & _
                                  "  FROM USERSRIGHTS " & _
                                  " WHERE RGH_ID = ?"

        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        Dim objRight As aut_right
        '        Dim objParam As OleDbParameter
        Dim bReturn As Boolean

        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        'Si le paramètre existe on ne met à jour que sa valeur
        'Autrement on le crée
        CreateParameterP_ID(objCommand)

        Try
            objRS = objCommand.ExecuteReader
            If Not objRS.HasRows Then
                objRS.Close()
                Return False
            End If
            objRight = CType(Me, aut_right)
            Try
                objRight.tag = GetString(objRS, "RGH_TAG")
            Catch ex As InvalidCastException
                objRight.tag = ""
            End Try
            Try
                objRight.role = GetString(objRS, "RGH_ROLE")
            Catch ex As InvalidCastException
                objRight.role = vncEnums.userRole.INVITE
            End Try

            Try
                objRight.droit = GetString(objRS, "RGH_DROIT")
            Catch ex As InvalidCastException
                objRight.droit = False
            End Try

            Try
                objRight.text = GetString(objRS, "RGH_TEXT")
            Catch ex As InvalidCastException
                objRight.text = ""
            End Try

            bReturn = True

        Catch ex As Exception
            setError("LoadRIGHT", ex.ToString())
            bReturn = False
        End Try
        objRS.Close()
        objCommand.Connection.Close()
        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'LoadRight
    ''' <summary>
    ''' Chargement des Lignes de menu INTERDIT pour un user
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Function loadcolUSERSRIGHTS() As Boolean

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "L'id doit être renseigné")

        'On ne charge que les lignes interdites par role
        Dim sqlString As String = " SELECT RGH_TAG, RGH_DROIT, RGH_TEXT" & _
                                  "  FROM USERSRIGHTS WHERE RGH_DROIT = 0 and " & _
                                  "  RGH_ROLE = ?"

        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        Dim objUser As aut_user
        '        Dim objParam As OleDbParameter
        Dim bReturn As Boolean
        Dim tag As String
        Dim bDroit As Boolean
        Dim text As String

        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        'Si le paramètre existe on ne met à jour que sa valeur
        'Autrement on le crée
        CreateParameterP_ROLE(objCommand)

        Try
            objRS = objCommand.ExecuteReader
            objUser = CType(Me, aut_user)
            While objRS.Read()
                Try
                    tag = GetString(objRS, "RGH_TAG")
                Catch ex As InvalidCastException
                    tag = ""
                End Try

                Try
                    bDroit = GetString(objRS, "RGH_DROIT")
                Catch ex As InvalidCastException
                    bDroit = False
                End Try

                Try
                    text = GetString(objRS, "RGH_TEXT")
                Catch ex As InvalidCastException
                    text = False
                End Try

                If tag <> "" Then
                    objUser.ajouteDroit(tag, bDroit, text)
                Else
                    'Debug.Assert(False, "rght_tag = ''")
                End If
            End While
            objRS.Close()


            bReturn = True

        Catch ex As Exception
            setError("LoadRIGHT", ex.ToString())
            bReturn = False
        End Try
        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'LoadcolUSERSRIGHTS

    Protected Function insertUSER() As Boolean
        Dim obj As aut_user
        Dim bReturn As Boolean

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        obj = CType(Me, aut_user)

        Dim sqlString As String = " INSERT INTO USERS (" & _
                                    " USR_CODE , " & _
                                    " USR_PASSWORD , " & _
                                    " USR_ROLE  " & _
                                    " ) VALUES (" & _
                                    " ? , " & _
                                    " ? , " & _
                                    " ? " & _
                                    " )"
        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        '        Dim objParam As OleDbParameter




        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        CreateParameterP_CODE(objCommand)
        CreateParameterP_PASSWORD(objCommand)
        CreateParameterP_ROLE(objCommand)
        m_dbconn.BeginTransaction()
        Try
            objCommand.ExecuteNonQuery()
            objCommand.CommandText = "SELECT MAX(USER_ID) FROM USERS"
            objRS = objCommand.ExecuteReader
            m_id = objRS.GetInt32(0)
            cleanErreur()
            objRS.Close()
            bReturn = True
            m_dbconn.transaction.Commit()

        Catch ex As Exception
            setError("InsertUSER", ex.ToString())
            bReturn = False
            m_dbconn.transaction.Rollback()
        End Try

        '    Debug.Assert(m_id <> 0, "ID=0")
        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'InsertUser

    Protected Function deletecolUSERSRIGHTS() As Boolean
        '==========================================================
        'Function :deletecolUSERSRIGHTS
        ' Description : suppression des droits pour un role
        '==========================================================

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "L'id doit être renseigné")

        Dim sqlString As String = " DELETE FROM USERSRIGHTS" & _
                                  " WHERE RGH_ROLE = ?"

        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        '        Dim objParam As OleDbParameter
        Dim bReturn As Boolean

        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        'Si le paramètre existe on ne met à jour que sa valeur
        'Autrement on le crée
        CreateParameterP_ROLE(objCommand)

        Try
            objCommand.ExecuteNonQuery()
            bReturn = True

        Catch ex As Exception
            setError("deletecolUSERSRIGHTS", ex.ToString())
            bReturn = False
        End Try
        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'deletecolUSERSRIGHTS
    Protected Function insertcolUSERSRIGHTS() As Boolean
        Dim obj As aut_user
        Dim bReturn As Boolean
        Dim objRight As aut_right


        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "Le Produit doit être Sauvegardé")

        obj = CType(Me, aut_user) 'Appelle depuis un objet aut_user

        bReturn = True
        m_dbconn.BeginTransaction()
        For Each objRight In obj.colRigths
            objRight.setid(0)
            objRight.role = obj.role
            bReturn = bReturn And objRight.insert
        Next
        If bReturn Then
            m_dbconn.transaction.Commit()
        Else
            m_dbconn.transaction.Rollback()
        End If

        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'insertcolUSERSRIGHTS
    Protected Function insertRIGHT() As Boolean
        Dim obj As aut_right
        Dim bReturn As Boolean

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        obj = CType(Me, aut_right)

        Dim sqlString As String = " INSERT INTO USERSRIGHTS (" & _
                                    " RGH_TAG , " & _
                                    " RGH_ROLE , " & _
                                    " RGH_DROIT,  " & _
                                    " RGH_TEXT  " & _
                                    " ) VALUES (" & _
                                    " ? , " & _
                                    " ? , " & _
                                    " ? , " & _
                                    " ? " & _
                                    " )"
        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        '        Dim objParam As OleDbParameter



        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        CreateParameterP_RGH_TAG(objCommand)
        CreateParameterP_RGH_ROLE(objCommand)
        CreateParameterP_RGH_DROIT(objCommand)
        CreateParameterP_RGH_TEXT(objCommand)
        m_dbconn.BeginTransaction()
        Try
            objCommand.ExecuteNonQuery()
            objCommand.CommandText = "SELECT MAX(RGH_ID) FROM USERSRIGHTS"
            objRS = objCommand.ExecuteReader
            m_id = objRS.GetInt32(0)
            cleanErreur()
            objRS.Close()
            bReturn = True
            m_dbconn.transaction.Commit()

        Catch ex As Exception
            setError("InsertRIGHT", ex.ToString())
            bReturn = False
            m_dbconn.transaction.Rollback()
        End Try

        '    Debug.Assert(m_id <> 0, "ID=0")
        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'InsertRIGHT

    Protected Function updateUSER() As Boolean
        Dim obj As aut_user
        Dim bReturn As Boolean

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        obj = CType(Me, aut_user)

        Dim sqlString As String = " UPDATE USERS SET" & _
                                    " USR_CODE  = ? , " & _
                                    " USR_PASSWORD = ? , " & _
                                    " USR_ROLE  =? " & _
                                    " WHERE USR_ID = ? "
        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        '        Dim objParam As OleDbParameter



        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        CreateParameterP_CODE(objCommand)
        CreateParameterP_PASSWORD(objCommand)
        CreateParameterP_ROLE(objCommand)
        CreateParameterP_ID(objCommand)
        Try
            m_dbconn.BeginTransaction()
            objCommand.ExecuteNonQuery()
            bReturn = True
            m_dbconn.transaction.Commit()

        Catch ex As Exception
            setError("UpdateUSER", ex.ToString())
            bReturn = False
            m_dbconn.transaction.Rollback()
        End Try

        '    Debug.Assert(m_id <> 0, "ID=0")
        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'UpdateUser
    Private Sub CreateParameterP_ROLE(ByVal objCommand As OleDbCommand)
        Dim obj As aut_user
        obj = Me
        objCommand.Parameters.AddWithValue("?", obj.role)
    End Sub
    Private Sub CreateParameterP_CODE(ByVal objCommand As OleDbCommand)
        Dim obj As aut_user
        obj = Me
        objCommand.Parameters.AddWithValue("?", obj.code)
    End Sub
    Private Sub CreateParameterP_PASSWORD(ByVal objCommand As OleDbCommand)
        Dim obj As aut_user
        obj = Me
        If obj.password Is Nothing Then
            objCommand.Parameters.AddWithValue("?", "")
        Else
            objCommand.Parameters.AddWithValue("?", obj.password)
        End If

    End Sub
    Private Sub CreateParameterP_RGH_TAG(ByVal objCommand As OleDbCommand)
        Dim obj As aut_right
        obj = Me
        objCommand.Parameters.AddWithValue("?", obj.tag)
    End Sub
    Private Sub CreateParameterP_RGH_ROLE(ByVal objCommand As OleDbCommand)
        Dim obj As aut_right
        obj = Me
        objCommand.Parameters.AddWithValue("?", obj.role)
    End Sub
    Private Sub CreateParameterP_RGH_DROIT(ByVal objCommand As OleDbCommand)
        Dim obj As aut_right
        obj = Me
        objCommand.Parameters.AddWithValue("?", obj.droit)
    End Sub
    Private Sub CreateParameterP_RGH_TEXT(ByVal objCommand As OleDbCommand)
        Dim obj As aut_right
        obj = Me
        objCommand.Parameters.AddWithValue("?", obj.text)
    End Sub

#End Region

#Region "TestStock"

    Public Shared Function shared_testStock(Optional ByVal pddeb As Date = DATE_DEFAUT, Optional ByVal pdfin As Date = DATE_DEFAUT) As Collection
        '============================================================================
        'Function : ListeFACTCOMEtat
        'Description : Rend une liste de Facture de commission en fonction leur état
        '============================================================================
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Dim colReturn As New Collection
        '        Dim objParam As OleDbParameter
        Dim sqlString As String = "SELECT " & _
                                    "CMD_ID " & _
                                  " FROM COMMANDE "

        Dim strWhere As String = " COMMANDE.CMD_ETAT > 2 "
        Dim objCommand As OleDbCommand
        Dim objCommandSTK As OleDbCommand
        Dim objCMD As CommandeClient
        Dim objRS As OleDbDataReader = Nothing
        Dim objRSSTK As OleDbDataReader = Nothing
        Dim nId As Integer
        Dim objLgCmd As LgCommande
        Dim sqlStringSTK As String
        Dim nQte As Decimal
        Dim strErreur As String = ""
        Dim nEnr As Integer
        Dim objParam As OleDbParameter


        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection

        objCommandSTK = New OleDbCommand
        objCommandSTK.Connection = m_dbconn.Connection
        sqlStringSTK = "SELECT STK_ID, STK_QTE FROM MVT_STOCK WHERE STK_TYPE='2' AND STK_REF_ID= ? AND STK_PRD_ID = ?"
        objCommandSTK.CommandText = sqlStringSTK



        If pddeb <> DATE_DEFAUT Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & " CMD_DATE >=  ?"
            objParam = objCommand.Parameters.AddWithValue("?", pddeb)

        End If
        If pdfin <> DATE_DEFAUT Then
            If strWhere <> "" Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & " CMD_DATE <=  ?"
            objParam = objCommand.Parameters.AddWithValue("?", pdfin)

        End If

        If strWhere <> "" Then
            sqlString = sqlString & " WHERE " & strWhere
        End If
        objCommand.CommandText = sqlString



        'Parcours des commandes dont l'état > Validée
        Try
            objRS = objCommand.ExecuteReader
            While objRS.HasRows
                'Chargement de chaque commande
                nId = GetString(objRS, "CMD_ID")
                objCMD = CommandeClient.createandload(nId)
                If objCMD.id <> 0 Then
                    'Chargement des lignes de chaque commande
                    For Each objLgCmd In objCMD.colLignes
                        objCommandSTK.Parameters.AddWithValue("?", objCMD.id)
                        objCommandSTK.Parameters.AddWithValue("P_PRD_ID", objLgCmd.oProduit.id)
                        'Lecture de la table MVTSTK
                        strErreur = "Commande = " & objCMD.code & "(" & objCMD.id & "), PRODUIT = " & objLgCmd.oProduit.code & "(" & objLgCmd.oProduit.id & ")"
                        Try
                            objRSSTK = objCommandSTK.ExecuteReader()
                            nEnr = 0
                            'S'il y a plus d'un enr Erreur
                            While objRSSTK.HasRows()
                                nEnr = nEnr + 1
                                nQte = GetString(objRSSTK, "STK_QTE")
                                'Si la Qte n'est pas bonne => Erreur
                                If (nQte * -1) <> objLgCmd.qteLiv Then
                                    strErreur = strErreur & "Qte dans Commande = " & objLgCmd.qteLiv & " dans MVTSTOCK= " & nQte
                                    colReturn.Add(strErreur)
                                End If
                                objRSSTK.NextResult()
                            End While
                            If nEnr <> 1 Then
                                strErreur = strErreur & "Nbre lignes dans MVTSTOCK = " & nEnr
                                colReturn.Add(strErreur)
                            End If

                        Catch ex As Exception
                            strErreur = strErreur & "Erreur en lecture de MVTSTK " & sqlStringSTK
                            colReturn.Add(strErreur)
                        End Try

                    Next 'objLgCmd
                    objRSSTK.Close()
                End If 'objCmd.id <> 0
                objRS.NextResult()
            End While 'Boucle sur RSCommande
            objRS.Close()
            Return colReturn
        Catch ex As Exception
            strErreur = strErreur & "Erreur en lecture de COMMANDE " & sqlString
            colReturn.Add(strErreur)
        End Try

        Return colReturn
    End Function ' testStock

    Protected Function loadParam() As Boolean
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "L'id doit être renseigné")

        Dim colReturn As New Collection
        Dim objCommand As OleDbCommand
        Dim objParam As Param
        Dim objRS As OleDbDataReader = Nothing
        Dim strWhere As String = ""
        Dim bReturn As Boolean

        Try
            objParam = Me
            Dim sqlString As String = "SELECT PAR_ID, PAR_CODE, PAR_TYPE, PAR_VALUE, PAR_VALUE2, PAR_DEFAUT, PAR_DDEB_ECHEANCE FROM PARAMETRE WHERE PAR_ID = ?"
            objCommand = New OleDbCommand
            objCommand.Connection = m_dbconn.Connection

            CreateParameterP_ID(objCommand)

            objCommand.CommandText = sqlString
            objRS = objCommand.ExecuteReader

            While (objRS.Read)
                Me.fromRS(objRS)

                colReturn.Add(Me, objParam.code)
                objRS.NextResult()
            End While
            bReturn = True
            objRS.Close()
            objRS = Nothing
        Catch ex As Exception
            setError("loadParam", ex.ToString())
            bReturn = False
        End Try

        Debug.Assert(bReturn, "Persist.loadParam:" & getErreur())
        Return bReturn
    End Function 'LoadParam
#End Region
#Region "Transporteur"
    '=======================================================================
    'Fonction : LoadTRP
    'Description : Chargement de l'objet en base
    'Retour : Rend Vrai si le chargement s'est correctement effectué
    '=======================================================================
    Protected Function loadTRP() As Boolean

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "L'id doit être renseigné")

        Dim sqlString As String = "SELECT TRP_ID,  TRP_CODE, TRP_NOM," & _
                                    " TRP_LIV_RUE1, TRP_LIV_RUE2, TRP_LIV_CP, TRP_LIV_VILLE, " & _
                                    " TRP_LIV_TEL, TRP_LIV_FAX, TRP_LIV_PORT, TRP_LIV_EMAIL, TRP_DEFAUT " & _
                                  " FROM TRANSPORTEUR" & _
                                  " WHERE TRANSPORTEUR.TRP_ID=? "
        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        Dim objTRP As Transporteur
        '        Dim objParam As OleDbParameter
        Dim bReturn As Boolean

        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        'Si le paramètre existe on ne met à jour que sa valeur
        'Autrement on le crée
        CreateParameterP_ID(objCommand)

        Try
            'objRS = objcommand.executeReader
            objRS = objCommand.ExecuteReader
            If Not objRS.HasRows Then
                objRS.Close()
                Return False
            End If
            objRS.Read()
            'm_id = CType(GetString(objRS,"CLT_ID",), Integer)
            objTRP = CType(Me, Transporteur)

            Try
                objTRP.code = GetString(objRS, "TRP_CODE")
            Catch ex As InvalidCastException
                objTRP.code = ""
            End Try
            Try
                objTRP.nom = GetString(objRS, "TRP_NOM")
                objTRP.AdresseLivraison.nom = GetString(objRS, "TRP_NOM")
            Catch ex As InvalidCastException
                objTRP.nom = ""
                objTRP.AdresseLivraison.nom = ""
            End Try


            Try
                objTRP.AdresseLivraison.rue1 = GetString(objRS, "TRP_LIV_RUE1")
            Catch ex As InvalidCastException
                objTRP.AdresseLivraison.rue1 = ""
            End Try

            Try
                objTRP.AdresseLivraison.rue2 = GetString(objRS, "TRP_LIV_RUE2")
            Catch ex As InvalidCastException
                objTRP.AdresseLivraison.rue2 = ""
            End Try

            Try
                objTRP.AdresseLivraison.cp = GetString(objRS, "TRP_LIV_CP")
            Catch ex As InvalidCastException
                objTRP.AdresseLivraison.cp = ""
            End Try

            Try
                objTRP.AdresseLivraison.ville = GetString(objRS, "TRP_LIV_VILLE")
            Catch ex As InvalidCastException
                objTRP.AdresseLivraison.ville = ""
            End Try

            Try
                objTRP.AdresseLivraison.tel = GetString(objRS, "TRP_LIV_TEL")
            Catch ex As InvalidCastException
                objTRP.AdresseLivraison.tel = ""
            End Try

            Try
                objTRP.AdresseLivraison.fax = GetString(objRS, "TRP_LIV_FAX")
            Catch ex As InvalidCastException
                objTRP.AdresseLivraison.fax = ""
            End Try
            Try
                objTRP.AdresseLivraison.port = GetString(objRS, "TRP_LIV_PORT")
            Catch ex As InvalidCastException
                objTRP.AdresseLivraison.port = ""
            End Try

            Try
                objTRP.AdresseLivraison.Email = GetString(objRS, "TRP_LIV_EMAIL")
            Catch ex As InvalidCastException
                objTRP.AdresseLivraison.Email = ""
            End Try
            Try
                objTRP.bDefaut = GetString(objRS, "TRP_DEFAUT")
            Catch ex As InvalidCastException
                objTRP.bDefaut = ""
            End Try

            objRS.Close()
            objRS = Nothing
            bReturn = True
            cleanErreur()

        Catch ex As Exception
            setError("LoadTRP", ex.ToString())
            bReturn = False
        End Try
        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'LoadTRP
    '=======================================================================
    'Fonction : InsertTRP
    'Description : Insertion d'un Transporteur
    'Retour : Rend Vrai si l'INSERT s'est correctement effectué
    '=======================================================================
    Protected Function insertTRP() As Boolean
        Dim objTRP As Transporteur
        Dim bReturn As Boolean

        Debug.Assert(Me.GetType().Name.Equals("Transporteur"), "Objet de Type Transporteur Requis")
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id = 0, "ID=0")
        objTRP = CType(Me, Transporteur)

        Dim sqlString As String = "INSERT INTO TRANSPORTEUR( " & _
                                    "TRP_CODE, TRP_NOM, TRP_LIV_RUE1,  TRP_LIV_RUE2,TRP_LIV_CP,TRP_LIV_VILLE,TRP_LIV_TEL,TRP_LIV_FAX,TRP_LIV_PORT,TRP_LIV_EMAIL, TRP_DEFAUT " & _
                                    ") VALUES (?,?,?,?,?,?,?,?,?,?,?)"
        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        '        Dim objParam As OleDbParameter


        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString
        m_dbconn.BeginTransaction()
        objCommand.Transaction = m_dbconn.transaction


        CreateParamP_TRP_CODE(objCommand)
        CreateParamP_TRP_NOM(objCommand)
        CreateParamP_TRP_LIV_RUE1(objCommand)
        CreateParamP_TRP_LIV_RUE2(objCommand)
        CreateParamP_TRP_LIV_CP(objCommand)
        CreateParamP_TRP_LIV_VILLE(objCommand)
        CreateParamP_TRP_LIV_TEL(objCommand)
        CreateParamP_TRP_LIV_FAX(objCommand)
        CreateParamP_TRP_LIV_PORT(objCommand)
        CreateParamP_TRP_LIV_EMAIL(objCommand)
        CreateParamP_TRP_DEFAUT(objCommand)
        Try
            objCommand.ExecuteNonQuery()
            objCommand = New OleDbCommand("SELECT MAX(TRP_ID) FROM TRANSPORTEUR", m_dbconn.Connection)
            objCommand.Transaction = m_dbconn.transaction
            objRS = objCommand.ExecuteReader
            bReturn = True
            If (objRS.Read()) Then
                m_id = objRS.GetInt32(0)
                cleanErreur()
            Else
                setError("InsertTRP", "No TRP ID")
                bReturn = False
            End If
            objRS.Close()
            objRS = Nothing

        Catch ex As Exception
            setError("InsertTRP", ex.ToString())
            bReturn = False
        End Try

        If bReturn Then
            m_dbconn.transaction.Commit()
        Else
            m_dbconn.transaction.Rollback()
        End If
        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'InsertTRP
    '=======================================================================
    'Fonction : UpdateTRP
    'Description : Insertion d'un Transporteur
    'Retour : Rend Vrai si l'UPDATE s'est correctement effectué
    '=======================================================================
    Protected Function updateTRP() As Boolean
        Dim objTRP As Transporteur
        Dim bReturn As Boolean

        Debug.Assert(Me.GetType.Name.Equals("Transporteur"), "Objet de type Client Requis")
        objTRP = CType(Me, Transporteur)
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "ID<>0")

        Dim sqlString As String = "UPDATE TRANSPORTEUR SET TRP_NOM = ?, " & _
                                        "TRP_CODE = ? , " & _
                                        "TRP_LIV_RUE1 = ? , " & _
                                        "TRP_LIV_RUE2 = ? , " & _
                                        "TRP_LIV_CP = ? , " & _
                                        "TRP_LIV_VILLE = ? , " & _
                                        "TRP_LIV_TEL = ? , " & _
                                        "TRP_LIV_FAX = ? , " & _
                                        "TRP_LIV_PORT = ? , " & _
                                        "TRP_LIV_EMAIL = ? , " & _
                                        "TRP_DEFAUT = ? , " & _
                                  " WHERE TRP_ID = ?"
        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        '        Dim objParam As OleDbParameter

        objTRP = CType(Me, Transporteur)

        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        CreateParamP_TRP_NOM(objCommand)
        CreateParamP_TRP_CODE(objCommand)
        CreateParamP_TRP_LIV_RUE1(objCommand)
        CreateParamP_TRP_LIV_RUE2(objCommand)
        CreateParamP_TRP_LIV_CP(objCommand)
        CreateParamP_TRP_LIV_VILLE(objCommand)
        CreateParamP_TRP_LIV_TEL(objCommand)
        CreateParamP_TRP_LIV_FAX(objCommand)
        CreateParamP_TRP_LIV_PORT(objCommand)
        CreateParamP_TRP_LIV_EMAIL(objCommand)
        CreateParamP_TRP_DEFAUT(objCommand)
        CreateParameterP_ID(objCommand)


        Try
            objCommand.ExecuteNonQuery()
            cleanErreur()
            bReturn = True

        Catch ex As Exception
            setError("UpdateTRP", ex.ToString())
            bReturn = False
        End Try

        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'UpdateTRP

    '=======================================================================
    'Fonction : DeleteTRP
    'Description : Suppression d'un client
    'Retour : Rend Vrai si le DELETE s'est correctement effectué
    '=======================================================================
    Protected Function deleteTRP() As Boolean
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "Id <> 0")
        Debug.Assert(Me.GetType().Name.Equals("Client"))


        Dim sqlString As String = "DELETE FROM TRANSPORTEUR WHERE TRP_ID=? "
        Dim objCommand As OleDbCommand
        Dim objTRP As Transporteur
        '        Dim objParam As OleDbParameter

        objTRP = CType(Me, Transporteur)
        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        CreateParameterP_ID(objCommand)
        Try
            objCommand.ExecuteNonQuery()
            m_id = 0
            objTRP.code = ""
            objTRP.nom = ""
            Return True

        Catch ex As Exception
            setError("DeleteTRP", ex.ToString())
            Return False
        End Try
    End Function 'DeleteTRP
    Protected Shared Function listeTRANSPORTEURS() As Collection
        '=================================================================
        ' Function : listeTRANSPORTEURS
        'Description : Changement d'une collection de Transporteurs
        '=================================================================

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")

        Dim sqlString As String = "SELECT TRP_ID,  TRP_CODE, TRP_NOM," & _
                                    " TRP_LIV_RUE1, TRP_LIV_RUE2, TRP_LIV_CP, TRP_LIV_VILLE, " & _
                                    " TRP_LIV_TEL, TRP_LIV_FAX, TRP_LIV_PORT, TRP_LIV_EMAIL, TRP_DEFAUT " & _
                                  " FROM TRANSPORTEUR"

        Dim objCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        Dim objTRP As Transporteur
        '        Dim objParam As OleDbParameter
        Dim nid As Integer
        Dim colReturn As New Collection

        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString




        Try
            objRS = objCommand.ExecuteReader
            While objRS.Read
                nid = GetString(objRS, "TRP_ID")
                objTRP = New Transporteur
                objTRP.setid(nid)
                Try
                    objTRP.code = GetString(objRS, "TRP_CODE")
                Catch ex As InvalidCastException
                    objTRP.code = ""
                End Try
                Try
                    objTRP.nom = GetString(objRS, "TRP_NOM")
                    objTRP.AdresseLivraison.nom = GetString(objRS, "TRP_NOM")
                Catch ex As InvalidCastException
                    objTRP.nom = ""
                    objTRP.AdresseLivraison.nom = ""
                End Try


                Try
                    objTRP.AdresseLivraison.rue1 = GetString(objRS, "TRP_LIV_RUE1")
                Catch ex As InvalidCastException
                    objTRP.AdresseLivraison.rue1 = ""
                End Try

                Try
                    objTRP.AdresseLivraison.rue2 = GetString(objRS, "TRP_LIV_RUE2")
                Catch ex As InvalidCastException
                    objTRP.AdresseLivraison.rue2 = ""
                End Try

                Try
                    objTRP.AdresseLivraison.cp = GetString(objRS, "TRP_LIV_CP")
                Catch ex As InvalidCastException
                    objTRP.AdresseLivraison.cp = ""
                End Try

                Try
                    objTRP.AdresseLivraison.ville = GetString(objRS, "TRP_LIV_VILLE")
                Catch ex As InvalidCastException
                    objTRP.AdresseLivraison.ville = ""
                End Try

                Try
                    objTRP.AdresseLivraison.tel = GetString(objRS, "TRP_LIV_TEL")
                Catch ex As InvalidCastException
                    objTRP.AdresseLivraison.tel = ""
                End Try

                Try
                    objTRP.AdresseLivraison.fax = GetString(objRS, "TRP_LIV_FAX")
                Catch ex As InvalidCastException
                    objTRP.AdresseLivraison.fax = ""
                End Try
                Try
                    objTRP.AdresseLivraison.port = GetString(objRS, "TRP_LIV_PORT")
                Catch ex As InvalidCastException
                    objTRP.AdresseLivraison.port = ""
                End Try

                Try
                    objTRP.AdresseLivraison.Email = GetString(objRS, "TRP_LIV_EMAIL")
                Catch ex As InvalidCastException
                    objTRP.AdresseLivraison.Email = ""
                End Try
                Try
                    objTRP.bDefaut = GetString(objRS, "TRP_DEFAUT")
                Catch ex As InvalidCastException
                    objTRP.bDefaut = ""
                End Try
                If objTRP.id <> 0 Then
                    objTRP.resetBooleans()
                    colReturn.Add(objTRP, objTRP.code)
                End If
            End While
            objRS.Close()

        Catch ex As Exception
            setError("ListeTransporteurs", ex.ToString())
            colReturn = Nothing
        End Try
        Debug.Assert(Not colReturn Is Nothing, getErreur())
        Return colReturn
    End Function 'listeTRANSPORTEURS



    Private Sub CreateParamP_TRP_NOM(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objTRP As Transporteur
        objTRP = Me
        objCommand.Parameters.AddWithValue("?", truncate(objTRP.nom, 50))
    End Sub

    Private Sub CreateParamP_TRP_CODE(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objTRP As Transporteur
        objTRP = Me
        objCommand.Parameters.AddWithValue("?", objTRP.code)
    End Sub
    Private Sub CreateParamP_TRP_LIV_RUE1(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objTRP As Transporteur
        objTRP = Me
        objCommand.Parameters.AddWithValue("?", truncate(objTRP.AdresseLivraison.rue1, 50))
    End Sub
    Private Sub CreateParamP_TRP_LIV_RUE2(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objTRP As Transporteur
        objTRP = Me
        objCommand.Parameters.AddWithValue("?", truncate(objTRP.AdresseLivraison.rue2, 50))
    End Sub

    Private Sub CreateParamP_TRP_LIV_CP(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objTRP As Transporteur
        objTRP = Me
        objCommand.Parameters.AddWithValue("?", truncate(objTRP.AdresseLivraison.cp, 50))
    End Sub
    Private Sub CreateParamP_TRP_LIV_VILLE(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objTRP As Transporteur
        objTRP = Me
        objCommand.Parameters.AddWithValue("?", truncate(objTRP.AdresseLivraison.ville, 50))
    End Sub
    Private Sub CreateParamP_TRP_LIV_TEL(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objTRP As Transporteur
        objTRP = Me
        objCommand.Parameters.AddWithValue("?", truncate(objTRP.AdresseLivraison.tel, 50))
    End Sub
    Private Sub CreateParamP_TRP_LIV_FAX(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objTRP As Transporteur
        objTRP = Me
        objCommand.Parameters.AddWithValue("?", truncate(objTRP.AdresseLivraison.fax, 50))
    End Sub
    Private Sub CreateParamP_TRP_LIV_PORT(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objTRP As Transporteur
        objTRP = Me
        objCommand.Parameters.AddWithValue("?", truncate(objTRP.AdresseLivraison.port, 50))
    End Sub
    Private Sub CreateParamP_TRP_LIV_EMAIL(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objTRP As Transporteur
        objTRP = Me
        objCommand.Parameters.AddWithValue("?", truncate(objTRP.AdresseLivraison.Email, 50))
    End Sub

    Private Sub CreateParamP_TRP_DEFAUT(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objTRP As Transporteur
        objTRP = Me
        objCommand.Parameters.AddWithValue("?", objTRP.bDefaut)
    End Sub


#End Region
#Region "Reglement"
    '=======================================================================
    'Fonction : LoadReglement
    'Description : Chargement de l'objet en base
    'Retour : Rend Vrai si le chargement s'est correctement effectué
    '=======================================================================
    Protected Function loadReglement() As Boolean

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "L'id doit être renseigné")

        Dim oTA As dsVinicomTableAdapters.REGLEMENTTableAdapter = New dsVinicomTableAdapters.REGLEMENTTableAdapter()
        Dim oDT As dsVinicom.REGLEMENTDataTable
        Dim objReglement As Reglement
        Dim bReturn As Boolean
        Dim oRow As dsVinicom.REGLEMENTRow


        Try
            oTA.Connection = m_dbconn.Connection
            oDT = oTA.GetDataBy_ID(m_id)
            If oDT.Rows.Count = 1 Then
                objReglement = CType(Me, Reglement)
                oRow = oDT(0)
                objReglement.getData(oRow)
            End If

            bReturn = True
            cleanErreur()

        Catch ex As Exception
            setError("LoadREGLEMENT", ex.ToString())
            bReturn = False
        End Try
        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'LoadREGLEMENT
    '=======================================================================
    'Fonction : InsertREGLEMENT
    'Description : Insertion d'un Reglement
    'Retour : Rend Vrai si l'INSERT s'est correctement effectué
    '=======================================================================
    Protected Function insertREGLEMENT() As Boolean
        Debug.Assert(Me.GetType().Name.Equals("Reglement"), "Objet de Type Transporteur Requis")
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id = 0, "ID=0")

        Dim objReglement As Reglement
        Dim bReturn As Boolean
        Dim oTA As New dsVinicomTableAdapters.REGLEMENTTableAdapter

        oTA.Connection = m_dbconn.Connection
        objReglement = CType(Me, Reglement)




        Try
            oTA.Insert(objReglement.IdFact, objReglement.Montant, objReglement.DateReglement.ToShortDateString(), objReglement.Reference, objReglement.Etat.codeEtat, objReglement.TypeFact, objReglement.Commentaire)
            setid(oTA.getMaxID())
            bReturn = True

        Catch ex As Exception
            setError("InsertTRP", ex.ToString())
            bReturn = False
        End Try

        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'InsertReglement
    '=======================================================================
    'Fonction : UpdateReglement
    'Description : Insertion d'un Reglement
    'Retour : Rend Vrai si l'UPDATE s'est correctement effectué
    '=======================================================================
    Protected Function updateReglement() As Boolean
        Dim objReglement As Reglement
        Dim bReturn As Boolean

        Debug.Assert(Me.GetType.Name.Equals("Reglement"), "Objet de type Reglement Requis")
        objReglement = CType(Me, Reglement)
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "ID<>0")

        Dim oTA As dsVinicomTableAdapters.REGLEMENTTableAdapter

        oTA = New dsVinicomTableAdapters.REGLEMENTTableAdapter()
        oTA.Connection = m_dbconn.Connection

        objReglement = CType(Me, Reglement)

        Try
            oTA.Update(objReglement.IdFact, objReglement.Montant, objReglement.DateReglement.ToShortDateString(), objReglement.Reference, objReglement.Etat.codeEtat, objReglement.TypeFact, objReglement.Commentaire, objReglement.id)
            cleanErreur()
            bReturn = True

        Catch ex As Exception
            setError("UpdateReglement", ex.ToString())
            bReturn = False
        End Try

        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'UpdateReglement

    '=======================================================================
    'Fonction : DeleteReglement
    'Description : Suppression d'un Reglement
    'Retour : Rend Vrai si le DELETE s'est correctement effectué
    '=======================================================================
    Protected Function deleteReglement() As Boolean
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")


        Dim oTa As dsVinicomTableAdapters.REGLEMENTTableAdapter
        oTa = New dsVinicomTableAdapters.REGLEMENTTableAdapter()
        oTa.Connection = m_dbconn.Connection
        Try
            oTa.Delete(Me.id)
            m_id = 0
            Return True

        Catch ex As Exception
            setError("Persist.DeleteReglement", ex.ToString())
            Return False
        End Try
    End Function 'DeleteReglement
    Public Function deleteReglements() As Boolean
        Dim bReturn As Boolean
        Dim oTA As dsVinicomTableAdapters.REGLEMENTTableAdapter
        Try
            oTA = New dsVinicomTableAdapters.REGLEMENTTableAdapter()
            oTA.Connection = m_dbconn.Connection
            oTA.DeleteBy_TYPFACT_IDFACT(typeDonnee, id)
            bReturn = True
        Catch ex As Exception

            setError("Persist.deleteReglements", ex.Message)
            bReturn = False
        End Try
        Return bReturn
    End Function

    Public Shared Function getListeReglement(ByVal pDateDeb As Date, ByVal pdateFin As Date) As Collection
        Dim colReturn As Collection
        Dim oTA As dsVinicomTableAdapters.REGLEMENTTableAdapter
        Dim dateDeb As Date
        Dim dateFin As Date
        Dim oDT As dsVinicom.REGLEMENTDataTable
        Dim oRow As dsVinicom.REGLEMENTRow
        Dim obj As Reglement


        Try
            colReturn = New Collection()
            oTA = New dsVinicomTableAdapters.REGLEMENTTableAdapter()
            oTA.Connection = m_dbconn.Connection
            dateDeb = CDate("01/01/1901")
            dateFin = CDate("31/12/2999")
            If Not pDateDeb.Equals(DATE_DEFAUT) Then
                dateDeb = pDateDeb
            End If
            If Not pdateFin.Equals(DATE_DEFAUT) Then
                dateFin = pdateFin
            End If
            oDT = oTA.GetDataBy_DATE(dateDeb, dateFin)

            For Each oRow In oDT
                obj = New Reglement()
                obj.getData(oRow)
                colReturn.Add(obj)
            Next


        Catch ex As Exception
            setError(System.Environment.StackTrace, ex.Message)
            colReturn = New Collection()
        End Try

        Return colReturn
    End Function

    Public Shared Function getListeReglement(ByVal pDateDeb As Date, ByVal pdateFin As Date, ByVal pidFact As Integer, ByVal pTypeFact As Integer) As Collection
        Dim colReturn As Collection
        Dim oTA As dsVinicomTableAdapters.REGLEMENTTableAdapter
        Dim dateDeb As Date
        Dim dateFin As Date
        Dim oDT As dsVinicom.REGLEMENTDataTable
        Dim oRow As dsVinicom.REGLEMENTRow
        Dim obj As Reglement


        Try
            colReturn = New Collection()
            oTA = New dsVinicomTableAdapters.REGLEMENTTableAdapter()
            oTA.Connection = Persist.oleDBConnection
            dateDeb = CDate("01/01/1901")
            dateFin = CDate("31/12/2999")
            If Not pDateDeb.Equals(DATE_DEFAUT) Then
                dateDeb = pDateDeb
            End If
            If Not pdateFin.Equals(DATE_DEFAUT) Then
                dateFin = pdateFin
            End If
            oDT = oTA.GetDataBy_DATE_IDFACT_TYPEFACT(dateDeb, dateFin, pidFact, pTypeFact)

            For Each oRow In oDT
                obj = New Reglement()
                obj.getData(oRow)
                colReturn.Add(obj)
            Next


        Catch ex As Exception
            setError(System.Environment.StackTrace, ex.Message)
            colReturn = New Collection()
        End Try

        Return colReturn
    End Function

    Public Shared Function getListeReglement(ByVal pDateDeb As Date, ByVal pdateFin As Date, ByVal pType As vncEnums.vncTypeDonnee, ByVal pEtat As vncEtatReglement) As Collection
        Dim colReturn As Collection
        Dim oTA As dsVinicomTableAdapters.REGLEMENTTableAdapter
        Dim dateDeb As Date
        Dim dateFin As Date
        Dim oDT As dsVinicom.REGLEMENTDataTable
        Dim oRow As dsVinicom.REGLEMENTRow
        Dim obj As Reglement


        Try
            colReturn = New Collection()
            oTA = New dsVinicomTableAdapters.REGLEMENTTableAdapter()
            oTA.Connection = Persist.oleDBConnection
            dateDeb = CDate("01/01/1901")
            dateFin = CDate("31/12/2999")
            If Not pDateDeb.Equals(DATE_DEFAUT) Then
                dateDeb = pDateDeb
            End If
            If Not pdateFin.Equals(DATE_DEFAUT) Then
                dateFin = pdateFin
            End If
            oDT = oTA.GetDataBy_DATE_TYPEFACT(dateDeb, dateFin, pType)

            For Each oRow In oDT
                obj = New Reglement()
                obj.getData(oRow)
                If (pEtat = vncEtatReglement.vncRGLMT_Tous Or obj.Etat.codeEtat = pEtat) Then
                    colReturn.Add(obj)
                End If
            Next


        Catch ex As Exception
            setError(System.Environment.StackTrace, ex.Message)
            colReturn = New Collection()
        End Try

        Return colReturn
    End Function

#End Region
#Region "TauxComm"
    '=======================================================================
    'Fonction : LoadTauxComm
    'Description : Chargement de l'objet en base
    'Retour : Rend Vrai si le chargement s'est correctement effectué
    '=======================================================================
    Protected Function loadTauxComm() As Boolean

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "L'id doit être renseigné")

        Dim oTA As dsVinicomTableAdapters.FRN_COMMTableAdapter = New dsVinicomTableAdapters.FRN_COMMTableAdapter()
        Dim oDT As dsVinicom.FRN_COMMDataTable
        Dim objTauxComm As TauxComm
        Dim bReturn As Boolean
        Dim oRow As dsVinicom.FRN_COMMRow


        Try
            oTA.Connection = m_dbconn.Connection
            oDT = oTA.GetDataByID(m_id)
            If oDT.Rows.Count = 1 Then
                objTauxComm = CType(Me, TauxComm)
                oRow = oDT(0)
                objTauxComm.getData(oRow)
            End If

            bReturn = True
            cleanErreur()

        Catch ex As Exception
            setError("LoadTauxComm", ex.ToString())
            bReturn = False
        End Try
        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'LoadTauxComm
    '=======================================================================
    'Fonction : InsertTauxComm
    'Description : Insertion d'un TauxComm
    'Retour : Rend Vrai si l'INSERT s'est correctement effectué
    '=======================================================================
    Protected Function insertTauxComm() As Boolean
        Debug.Assert(Me.GetType().Name.Equals("TauxComm"), "Objet de Type Transporteur Requis")
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")

        Dim objTauxComm As TauxComm
        Dim bReturn As Boolean
        Dim oTA As New dsVinicomTableAdapters.FRN_COMMTableAdapter

        oTA.Connection = m_dbconn.Connection
        objTauxComm = CType(Me, TauxComm)

        Try
            oTA.Insert(objTauxComm.IdFRN, objTauxComm.TypeClt, objTauxComm.TauxComm)
            setid(oTA.getMaxID())
            bReturn = True

        Catch ex As Exception
            setError("InsertTauxComm", ex.ToString())
            bReturn = False
        End Try

        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'InsertTauxComm
    '=======================================================================
    'Fonction : UpdateTauxComm
    'Description : Insertion d'un TauxComm
    'Retour : Rend Vrai si l'UPDATE s'est correctement effectué
    '=======================================================================
    Protected Function updateTauxComm() As Boolean
        Dim objTauxComm As TauxComm
        Dim bReturn As Boolean

        Debug.Assert(Me.GetType.Name.Equals("TauxComm"), "Objet de type TauxComm Requis")
        objTauxComm = CType(Me, TauxComm)
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "ID<>0")

        Dim oTA As dsVinicomTableAdapters.FRN_COMMTableAdapter

        oTA = New dsVinicomTableAdapters.FRN_COMMTableAdapter()
        oTA.Connection = m_dbconn.Connection

        objTauxComm = CType(Me, TauxComm)

        Try
            oTA.Update(objTauxComm.IdFRN, objTauxComm.TypeClt, objTauxComm.TauxComm, objTauxComm.id)
            cleanErreur()
            bReturn = True

        Catch ex As Exception
            setError("UpdateTauxComm", ex.ToString())
            bReturn = False
        End Try

        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'UpdateTauxComm

    '=======================================================================
    'Fonction : DeleteTauxComm
    'Description : Suppression d'un TauxComm
    'Retour : Rend Vrai si le DELETE s'est correctement effectué
    '=======================================================================
    Protected Function deleteTauxComm() As Boolean
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")


        Dim oTa As dsVinicomTableAdapters.FRN_COMMTableAdapter
        oTa = New dsVinicomTableAdapters.FRN_COMMTableAdapter()
        oTa.Connection = m_dbconn.Connection
        Try
            oTa.Delete(Me.id)
            m_id = 0
            Return True

        Catch ex As Exception
            setError("Persist.DeleteTauxComm", ex.ToString())
            Return False
        End Try
    End Function 'DeleteTauxComm
    ''' <summary>
    ''' Supprime les Taux de Coms du Fournisseur
    ''' </summary>
    ''' <param name="pFRNId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function deleteTauxComms(ByVal pFRNId As Integer) As Boolean
        Dim bReturn As Boolean
        Dim oTA As dsVinicomTableAdapters.FRN_COMMTableAdapter
        Try
            oTA = New dsVinicomTableAdapters.FRN_COMMTableAdapter()
            oTA.Connection = m_dbconn.Connection
            oTA.DeleteFromFRN_ID(pFRNId)
            bReturn = True
        Catch ex As Exception

            setError("Persist.deleteTauxComms", ex.Message)
            bReturn = False
        End Try
        Return bReturn
    End Function
    ''' <summary>
    ''' Rend la liste des taux de commision d'un fournisseur
    ''' </summary>
    ''' <param name="pidFRN">id Fournisseur</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Shared Function getListeTauxComm(ByVal pidFRN As Integer) As Collection
        Dim colReturn As Collection
        Dim oTA As dsVinicomTableAdapters.FRN_COMMTableAdapter
        Dim oDT As dsVinicom.FRN_COMMDataTable
        Dim oRow As dsVinicom.FRN_COMMRow
        Dim obj As TauxComm


        Try
            colReturn = New Collection()
            oTA = New dsVinicomTableAdapters.FRN_COMMTableAdapter()
            oTA.Connection = m_dbconn.Connection
            oDT = oTA.GetDataByFRNID(pidFRN)

            For Each oRow In oDT
                obj = New TauxComm(oRow)
                colReturn.Add(obj)
            Next


        Catch ex As Exception
            setError(System.Environment.StackTrace, ex.Message)
            colReturn = New Collection()
        End Try

        Return colReturn
    End Function
    ''' <summary>
    ''' Rend le taux de commissons pour un Fournisseur et un type de client 
    ''' </summary>
    ''' <param name="pidFRN">id Fournisseur</param>
    ''' <param name="pTypeClt">Code Type de client</param>
    ''' <returns>Un Taux de commission</returns>
    ''' <remarks>Si Pas de Taux Rend Nothing</remarks>
    Protected Shared Function getTauxComm(ByVal pidFRN As Integer, ByVal pTypeClt As String) As TauxComm
        Dim oReturn As TauxComm
        Dim oTA As dsVinicomTableAdapters.FRN_COMMTableAdapter
        Dim oDT As dsVinicom.FRN_COMMDataTable


        Try
            oReturn = Nothing
            oTA = New dsVinicomTableAdapters.FRN_COMMTableAdapter()
            oTA.Connection = m_dbconn.Connection
            oDT = oTA.GetDataByFRNID_TYPECLT(pidFRN, pTypeClt)
            If oDT.Rows.Count > 0 Then
                oReturn = New TauxComm(oDT.Rows(0))
            End If


        Catch ex As Exception
            setError(System.Environment.StackTrace, ex.Message)
            oReturn = Nothing
        End Try

        Return oReturn
    End Function
#End Region

    Private Sub Persist_Updated() Handles MyBase.Updated
        m_bUpdated = True
    End Sub
    Protected Sub setUpdatedFalse()
        m_bUpdated = False
    End Sub
    Protected Sub majBooleenAlaFinDuNew()
        m_bUpdated = False
        m_bNew = True
        m_bDeleted = False
    End Sub

    Public Overrides ReadOnly Property shortResume() As String
        Get
            Return String.Empty
        End Get
    End Property

    Public Shared Function GetValue(ByVal objRS As OleDbDataReader, ByVal strColumnName As String) As Object
        Debug.Assert(Not objRS Is Nothing)
        Debug.Assert(Not objRS.IsClosed())
        Debug.Assert(objRS.HasRows)

        Dim nOrdinal As Integer
        Dim oReturn As Object
        nOrdinal = objRS.GetOrdinal(strColumnName)
        If Not objRS.IsDBNull(nOrdinal) Then
            Try
                oReturn = objRS.GetValue(nOrdinal)
            Catch e As InvalidCastException
                oReturn = Nothing
            End Try
        Else
            oReturn = Nothing
        End If

        Return oReturn
    End Function 'getValue

    Public Shared Function GetString(ByVal objRS As OleDbDataReader, ByVal strColumnName As String) As String
        Debug.Assert(Not objRS Is Nothing)
        Debug.Assert(Not objRS.IsClosed())
        Debug.Assert(objRS.HasRows)

        Dim nOrdinal As Integer
        Dim strReturn As String
        nOrdinal = objRS.GetOrdinal(strColumnName)
        If Not objRS.IsDBNull(nOrdinal) Then
            Try
                strReturn = objRS.GetValue(nOrdinal).ToString()
            Catch e As InvalidCastException
                strReturn = ""
            End Try
        Else
            strReturn = ""
        End If

        Return strReturn
    End Function 'getString
    Public Shared Function getLong(ByVal objDataReader As OleDbDataReader, ByVal strColumnName As String) As String
        Debug.Assert(Not objDataReader Is Nothing)
        Debug.Assert(Not objDataReader.IsClosed())
        Debug.Assert(objDataReader.HasRows)

        Dim nOrdinal As Integer
        Dim iReturn As Integer
        nOrdinal = objDataReader.GetOrdinal(strColumnName)
        If Not objDataReader.IsDBNull(nOrdinal) Then
            iReturn = objDataReader.GetInt32(nOrdinal)
        Else
            iReturn = -1
        End If
        Return iReturn
    End Function 'getInteger

    Public Shared Function getInteger(ByVal objDataReader As OleDbDataReader, ByVal strColumnName As String) As String
        Debug.Assert(Not objDataReader Is Nothing)
        Debug.Assert(Not objDataReader.IsClosed())
        Debug.Assert(objDataReader.HasRows)

        Dim nOrdinal As Integer
        Dim iReturn As Integer
        nOrdinal = objDataReader.GetOrdinal(strColumnName)
        If Not objDataReader.IsDBNull(nOrdinal) Then
            iReturn = objDataReader.GetInt32(nOrdinal)
        Else
            iReturn = -1
        End If
        Return iReturn
    End Function 'getInteger
    Public Shared Function getDecimal(ByVal objDataReader As OleDbDataReader, ByVal strColumnName As String) As String
        Debug.Assert(Not objDataReader Is Nothing)
        Debug.Assert(Not objDataReader.IsClosed())
        Debug.Assert(objDataReader.HasRows)

        Dim nOrdinal As Integer
        Dim iReturn As Decimal
        nOrdinal = objDataReader.GetOrdinal(strColumnName)
        If Not objDataReader.IsDBNull(nOrdinal) Then
            iReturn = objDataReader.GetDecimal(nOrdinal)
        Else
            iReturn = 0D
        End If
        Return iReturn
    End Function 'getDecimal
    Public Shared Function GetBoolean(ByVal objRS As OleDbDataReader, ByVal strColumnName As String) As Boolean
        Debug.Assert(Not objRS Is Nothing)
        Debug.Assert(Not objRS.IsClosed())
        Debug.Assert(objRS.HasRows)

        Dim nOrdinal As Integer
        Dim strReturn As Boolean
        nOrdinal = objRS.GetOrdinal(strColumnName)
        If Not objRS.IsDBNull(nOrdinal) Then
            Try
                strReturn = objRS.GetBoolean(nOrdinal)
            Catch e As InvalidCastException
                strReturn = False
            End Try
        Else
            strReturn = False
        End If

        Return strReturn
    End Function 'getBoolean
    Public Shared Function executeSQLNonQuery(ByVal strSQL As String) As Boolean
        Debug.Assert(shared_isConnected())

        Dim objCommand As OleDbCommand
        Dim bReturn As Boolean
        Try
            objCommand = New OleDbCommand(strSQL, m_dbconn.Connection)
            objCommand.ExecuteNonQuery()
            bReturn = True
        Catch e As Exception
            bReturn = False
        End Try

        Return bReturn

    End Function
    ''' <summary>
    ''' Execute a Query and returns the First Col of the First Line in the result
    ''' </summary>
    ''' <param name="strSQL"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function executeSQLQuery(ByVal strSQL As String) As String

        Dim objCommand As OleDbCommand
        Dim sReturn As String
        Dim oRs As OleDb.OleDbDataReader
        Try
            shared_connect()
            objCommand = New OleDbCommand(strSQL, m_dbconn.Connection)
            oRs = objCommand.ExecuteReader()
            If oRs.HasRows Then
                oRs.Read()
                sReturn = oRs.GetValue(0).ToString()
                oRs.Close()
            Else
                sReturn = String.Empty
            End If

        Catch e As Exception
            sReturn = String.Empty
        End Try

        Return sReturn

    End Function
    Friend Sub resetBooleans()
        m_bNew = False
        setUpdatedFalse()
        m_bDeleted = False
    End Sub

    Private Shared Sub m_dbconn_Connected() Handles m_dbconn.Connected
        RaiseEvent Connected()
    End Sub

    Private Shared Sub m_dbconn_Disconnected() Handles m_dbconn.Disconnected
        RaiseEvent Disconnected()
    End Sub
    Protected Overrides Sub RaiseUpdated()
        m_bUpdated = True
        MyBase.RaiseUpdated()
    End Sub
    Public Shared Sub setReportConnection(ByVal objReport As ReportDocument)

        Dim myConnectionInfo As ConnectionInfo = New ConnectionInfo()
        myConnectionInfo.ServerName = Persist.oleDBConnection.DataSource
        myConnectionInfo.DatabaseName = Persist.oleDBConnection.Database
        myConnectionInfo.UserID = m_ReportCnxUser
        myConnectionInfo.Password = m_ReportCnxPassword
        Trace.WriteLine("setReportconnection,ServerName=" + myConnectionInfo.ServerName + ",DBNAme=" + myConnectionInfo.DatabaseName + "User=" + myConnectionInfo.UserID + ",Password=" + myConnectionInfo.Password)


        Dim mySections As Sections = objReport.ReportDefinition.Sections
        For Each mySection As Section In mySections
            Dim myReportObjects As ReportObjects = mySection.ReportObjects
            For Each myReportObject As ReportObject In myReportObjects
                If myReportObject.Kind = ReportObjectKind.SubreportObject Then
                    Dim mySubreportObject As SubreportObject = CType(myReportObject, SubreportObject)
                    Dim subReportDocument As ReportDocument = mySubreportObject.OpenSubreport(mySubreportObject.SubreportName)
                    setReportConnection(subReportDocument)
                End If
            Next
        Next
        'Applique la connection sur les tables du rapport
        Dim myTables As Tables = objReport.Database.Tables
        For Each myTable As CrystalDecisions.CrystalReports.Engine.Table In myTables
            Dim myTableLogonInfo As TableLogOnInfo = myTable.LogOnInfo
            myTableLogonInfo.ConnectionInfo = myConnectionInfo
            myTable.ApplyLogOnInfo(myTableLogonInfo)
        Next
    End Sub
    ''' <summary>
    ''' Returns the String truncate to the given number
    ''' </summary>
    ''' <param name="pstr"></param>
    ''' <param name="pLength"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function truncate(ByVal pstr As String, ByVal pLength As Integer) As String
        Dim strReturn As String
        strReturn = Trim(pstr.PadRight(pLength).Substring(0, pLength))
        Return strReturn
    End Function

    '=======================================================================
    'Fonction : DeleteParam
    'Description : Suppression d'un PRODUIT
    'Retour : Rend Vrai si le DELETE s'est correctement effectué
    '=======================================================================
    Protected Function deleteParam() As Boolean
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "Id <> 0")


        Dim sqlString As String = "DELETE FROM PARAMETRE WHERE PAR_ID=? "
        Dim objOLeDBCommand As OleDbCommand
        Dim objParam As Param

        objParam = CType(Me, Param)
        objOLeDBCommand = New OleDbCommand
        objOLeDBCommand.Connection = m_dbconn.Connection
        objOLeDBCommand.CommandText = sqlString



        CreateParameterP_ID(objOLeDBCommand)
        Try
            objOLeDBCommand.ExecuteNonQuery()
            m_id = 0
            objParam.code = ""
            objParam.valeur = ""
            Return True

        Catch ex As Exception
            setError("deletePARAM", ex.ToString())

            Return False
        End Try
    End Function 'DeleteParam
    '=======================================================================
    'Fonction : InsertParam
    'Description : Insertion d'un Paramètre
    'Retour : Rend Vrai si l'INSERT s'est correctement effectué
    '=======================================================================
    Protected Function insertParam() As Boolean
        Dim objParam As Param
        Dim bReturn As Boolean


        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id = 0, "ID=0")
        objParam = CType(Me, Param)

        Dim sqlString As String = "INSERT INTO PARAMETRE( " & _
                                    "PAR_CODE," & _
                                    "PAR_TYPE," & _
                                    "PAR_VALUE, " & _
                                    "PAR_DEFAUT, " & _
                                    "PAR_VALUE2, " & _
                                    "PAR_DDEB_ECHEANCE " & _
                                    " ) VALUES (" & _
                                    "?,?,?,?,?,?" & _
                                    " )"
        Dim objOLeDBCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing



        objOLeDBCommand = New OleDbCommand
        objOLeDBCommand.Connection = m_dbconn.Connection
        objOLeDBCommand.CommandText = sqlString
        m_dbconn.BeginTransaction()
        objOLeDBCommand.Transaction = m_dbconn.transaction


        CreateParamP_PAR_CODE(objOLeDBCommand)
        CreateParamP_PAR_TYPE(objOLeDBCommand)
        CreateParamP_PAR_VALUE(objOLeDBCommand)
        CreateParamP_PAR_DEFAUT(objOLeDBCommand)
        CreateParamP_PAR_VALUE2(objOLeDBCommand)
        CreateParamP_PAR_DDEB_ECHEANCE(objOLeDBCommand)
        Try
            objOLeDBCommand.ExecuteNonQuery()
            objOLeDBCommand = New OleDbCommand("SELECT MAX(PAR_ID) FROM PARAMETRE", m_dbconn.Connection)
            objOLeDBCommand.Transaction = m_dbconn.transaction
            objRS = objOLeDBCommand.ExecuteReader()
            If objRS.Read() Then
                m_id = objRS.GetInt32(0)
                cleanErreur()
                bReturn = True
            Else
                bReturn = False
            End If

        Catch ex As Exception
            setError("InsertParam", ex.ToString())
            bReturn = False
        End Try
        If Not objRS Is Nothing Then
            objRS.Close()
            objRS = Nothing
        End If
        If bReturn Then
            m_dbconn.transaction.Commit()
        Else
            m_dbconn.transaction.Rollback()
        End If

        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'InsertParam


    '=======================================================================
    'Fonction : UpdateParam
    'Description : Insertion d'un PARAMETRE
    'Retour : Rend Vrai si l'UPDATE s'est correctement effectué
    '=======================================================================
    Protected Function updateParam() As Boolean
        Dim objParam As Param
        Dim bReturn As Boolean

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "ID<>0")
        objParam = CType(Me, Param)

        Dim sqlString As String = "UPDATE PARAMETRE  SET " & _
                                    "PAR_CODE= ? , " & _
                                    "PAR_TYPE = ? , " & _
                                    "PAR_VALUE = ? , " & _
                                    "PAR_DEFAUT = ? , " & _
                                    "PAR_VALUE2 = ? , " & _
                                    "PAR_DDEB_ECHEANCE = ?  " & _
                                  " WHERE PAR_ID = ?"
        Dim objOLeDBCommand As OleDbCommand


        objOLeDBCommand = New OleDbCommand
        objOLeDBCommand.Connection = m_dbconn.Connection
        objOLeDBCommand.CommandText = sqlString



        CreateParamP_PAR_CODE(objOLeDBCommand)
        CreateParamP_PAR_TYPE(objOLeDBCommand)
        CreateParamP_PAR_VALUE(objOLeDBCommand)
        CreateParamP_PAR_DEFAUT(objOLeDBCommand)
        CreateParamP_PAR_VALUE2(objOLeDBCommand)
        CreateParamP_PAR_DDEB_ECHEANCE(objOLeDBCommand)
        CreateParameterP_ID(objOLeDBCommand)


        Try
            objOLeDBCommand.ExecuteNonQuery()
            cleanErreur()
            bReturn = True

        Catch ex As Exception

            setError("UpdateParam", ex.ToString())
            bReturn = False
        End Try
        Debug.Assert(bReturn, getErreur())
        Return bReturn

    End Function 'UpdateParam
#Region "CreateParameter Param"
    '=====================================================================================================================
    'TABLE PARAMETRE '
    '===============
    Private Sub CreateParamP_PAR_CODE(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objPAR As Param
        objPAR = Me
        objCommand.Parameters.AddWithValue("?", truncate(objPAR.code, 50))
    End Sub
    Private Sub CreateParamP_PAR_TYPE(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objPAR As Param
        objPAR = Me

        objCommand.Parameters.AddWithValue("?", objPAR.type)
    End Sub
    Private Sub CreateParamP_PAR_VALUE(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objPAR As Param
        objPAR = Me
        objCommand.Parameters.AddWithValue("?", objPAR.valeur)
    End Sub
    Private Sub CreateParamP_PAR_DEFAUT(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objPAR As Param
        objPAR = Me
        objCommand.Parameters.AddWithValue("?", objPAR.defaut)
    End Sub
    Private Sub CreateParamP_PAR_VALUE2(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objPAR As Param
        objPAR = Me
        objCommand.Parameters.AddWithValue("?", objPAR.valeur2)
    End Sub
    Private Sub CreateParamP_PAR_DDEB_ECHEANCE(ByVal objCommand As OleDbCommand)
        '        Dim objParam As OleDbParameter
        Dim objPAR As ParamModeReglement
        Try
            objPAR = Me
            objCommand.Parameters.AddWithValue("?", objPAR.dDebutEcheance)
        Catch
            objCommand.Parameters.AddWithValue("?", DBNull.Value)
        End Try

    End Sub
#End Region
#Region "Contenant"
    Protected Function loadContenant() As Boolean
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "L'id doit être renseigné")

        Dim colReturn As New Collection
        Dim objCommand As OleDbCommand
        Dim objParam As contenant
        Dim objRS As OleDbDataReader = Nothing
        Dim strWhere As String = ""
        Dim bReturn As Boolean

        Try
            objParam = Me
            Dim sqlString As String = "SELECT CONT_ID, CONT_CODE, CONT_LIBELLE, CONT_CENT, CONT_BOUT , CONT_DEFAUT, CONT_POIDS FROM CONTENANT WHERE CONT_ID = ?"
            objCommand = New OleDbCommand
            objCommand.Connection = m_dbconn.Connection

            CreateParameterP_ID(objCommand)

            objCommand.CommandText = sqlString
            objRS = objCommand.ExecuteReader

            While (objRS.Read)
                Me.fromRS(objRS)
                Me.resetBooleans()

                colReturn.Add(Me, objParam.code)
                objRS.NextResult()
            End While
            bReturn = True
            objRS.Close()
            objRS = Nothing
        Catch ex As Exception
            setError("loadContenant", ex.ToString())
            bReturn = False
        End Try

        Debug.Assert(bReturn, "Persist.loadContenant:" & getErreur())
        Return bReturn
    End Function 'LoadContenant
    '=======================================================================
    'Fonction : InsertContenant
    'Description : Insertion d'un Paramètre contenant
    'Retour : Rend Vrai si l'INSERT s'est correctement effectué
    '=======================================================================
    Protected Function insertContenant() As Boolean
        Dim objParam As contenant
        Dim bReturn As Boolean


        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id = 0, "ID=0")
        objParam = CType(Me, contenant)

        Dim sqlString As String = "INSERT INTO CONTENANT( " & _
                                    "CONT_CODE," & _
                                    "CONT_LIBELLE," & _
                                    "CONT_CENT, " & _
                                    "CONT_BOUT, " & _
                                    "CONT_DEFAUT, " & _
                                    "CONT_POIDS" & _
                                    " ) VALUES (" & _
                                    "?,?,?,?,?,?" & _
                                    " )"
        Dim objOLeDBCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing



        objOLeDBCommand = New OleDbCommand
        objOLeDBCommand.Connection = m_dbconn.Connection
        objOLeDBCommand.CommandText = sqlString
        m_dbconn.BeginTransaction()
        objOLeDBCommand.Transaction = m_dbconn.transaction


        CreateParamP_CONT_CODE(objOLeDBCommand)
        createParamP_CONT_LIBELLE(objOLeDBCommand)
        createParamP_CONT_CENT(objOLeDBCommand)
        createParamP_CONT_BOUT(objOLeDBCommand)
        createParamP_CONT_DEFAUT(objOLeDBCommand)
        createParamP_CONT_POIDS(objOLeDBCommand)
        Try
            objOLeDBCommand.ExecuteNonQuery()
            objOLeDBCommand = New OleDbCommand("SELECT MAX(CONT_ID) FROM CONTENANT", m_dbconn.Connection)
            objOLeDBCommand.Transaction = m_dbconn.transaction
            objRS = objOLeDBCommand.ExecuteReader()
            If objRS.Read() Then
                m_id = objRS.GetInt32(0)
                cleanErreur()
                bReturn = True
            Else
                bReturn = False
            End If

        Catch ex As Exception
            setError("InsertContenant", ex.ToString())
            bReturn = False
        End Try
        If Not objRS Is Nothing Then
            objRS.Close()
            objRS = Nothing
        End If
        If bReturn Then
            m_dbconn.transaction.Commit()
        Else
            m_dbconn.transaction.Rollback()
        End If

        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'InsertContenant
#Region "CreateParameter CONTENANT"
    '=====================================================================================================================
    'TABLE CONTENANT '
    '===============
    Private Sub CreateParamP_CONT_CODE(ByVal objCommand As OleDbCommand)
        Dim objPAR As contenant
        objPAR = Me
        objCommand.Parameters.AddWithValue("?", truncate(objPAR.code, 50))
    End Sub
    Private Sub createParamP_CONT_LIBELLE(ByVal objCommand As OleDbCommand)
        Dim objCONT As contenant
        objCONT = Me

        objCommand.Parameters.AddWithValue("?", objCONT.libelle)
    End Sub
    Private Sub createParamP_CONT_CENT(ByVal objCommand As OleDbCommand)
        Dim objCONT As contenant
        objCONT = Me
        objCommand.Parameters.AddWithValue("?", objCONT.cent)
    End Sub
    Private Sub createParamP_CONT_DEFAUT(ByVal objCommand As OleDbCommand)
        '        Dim objContenant As OleDbContenanteter
        Dim objCONT As contenant
        objCONT = Me
        objCommand.Parameters.AddWithValue("?", objCONT.defaut)
    End Sub
    Private Sub createParamP_CONT_BOUT(ByVal objCommand As OleDbCommand)
        Dim objCONT As contenant
        objCONT = Me
        objCommand.Parameters.AddWithValue("?", objCONT.bout)
    End Sub
    Private Sub createParamP_CONT_POIDS(ByVal objCommand As OleDbCommand)
        Dim objCONT As contenant
        objCONT = Me
        objCommand.Parameters.AddWithValue("?", objCONT.poids)
    End Sub
#End Region
    '=======================================================================
    'Fonction : DeleteContentant
    'Description : Suppression d'un Contenant
    'Retour : Rend Vrai si le DELETE s'est correctement effectué
    '=======================================================================
    Protected Function deleteContenant() As Boolean
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "Id <> 0")


        Dim sqlString As String = "DELETE FROM CONTENANT WHERE CONT_ID=? "
        Dim objOLeDBCommand As OleDbCommand
        Dim objParam As contenant

        objParam = CType(Me, contenant)
        objOLeDBCommand = New OleDbCommand
        objOLeDBCommand.Connection = m_dbconn.Connection
        objOLeDBCommand.CommandText = sqlString



        CreateParameterP_ID(objOLeDBCommand)
        Try
            objOLeDBCommand.ExecuteNonQuery()
            m_id = 0
            objParam.code = ""
            objParam.libelle = ""
            Return True

        Catch ex As Exception
            setError("deleteCont", ex.ToString())

            Return False
        End Try
    End Function 'DeleteCont
    '=======================================================================
    'Fonction : UpdateContenant
    'Description : Insertion d'un ContenantETRE
    'Retour : Rend Vrai si l'UPDATE s'est correctement effectué
    '=======================================================================
    Protected Function updateContenant() As Boolean
        Dim objContenant As contenant
        Dim bReturn As Boolean

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "ID<>0")
        objContenant = CType(Me, contenant)

        Dim sqlString As String = "UPDATE CONTENANT  SET " & _
                                    "CONT_CODE= ? , " & _
                                    "CONT_LIBELLE = ? , " & _
                                    "CONT_CENT = ? , " & _
                                    "CONT_DEFAUT = ? , " & _
                                    "CONT_BOUT = ? , " & _
                                    "CONT_POIDS= ?  " & _
                                  " WHERE CONT_ID = ?"
        Dim objOLeDBCommand As OleDbCommand


        objOLeDBCommand = New OleDbCommand
        objOLeDBCommand.Connection = m_dbconn.Connection
        objOLeDBCommand.CommandText = sqlString



        CreateParamP_CONT_CODE(objOLeDBCommand)
        createParamP_CONT_LIBELLE(objOLeDBCommand)
        createParamP_CONT_CENT(objOLeDBCommand)
        createParamP_CONT_DEFAUT(objOLeDBCommand)
        createParamP_CONT_BOUT(objOLeDBCommand)
        createParamP_CONT_POIDS(objOLeDBCommand)
        CreateParameterP_ID(objOLeDBCommand)


        Try
            objOLeDBCommand.ExecuteNonQuery()
            cleanErreur()
            bReturn = True

        Catch ex As Exception

            setError("UpdateContenant", ex.ToString())
            bReturn = False
        End Try
        Debug.Assert(bReturn, getErreur())
        Return bReturn

    End Function 'UpdateContenant

#End Region
#Region "ExportParam"
    Protected Shared Function ListeExportParam(ByVal pExpType As String) As Collection
        Dim colReturn As New Collection
        Try
            Dim oTa As dsVinicomTableAdapters.EXPORTPARAMTableAdapter
            oTa = New dsVinicomTableAdapters.EXPORTPARAMTableAdapter
            oTa.Connection = oleDBConnection
            Dim oDatatable As New dsVinicom.EXPORTPARAMDataTable()
            Dim oRow As dsVinicom.EXPORTPARAMRow
            oTa.FillByExp_Type(oDatatable, pExpType)
            For Each oRow In oDatatable
                Dim oExpParam As ExportParam
                oExpParam = New ExportParam(oRow.EXP_TYPE, oRow.EXP_NLIGNE, oRow.EXP_NCHAMPS, oRow.EXP_TYPECHAMPS, oRow.EXP_VALEUR, oRow.EXP_LONGUEUR, oRow.EXP_DESCRIPTION)
                colReturn.Add(oExpParam)
            Next

        Catch ex As Exception
            setError(ex.Message)
            colReturn = Nothing
        End Try
        Return colReturn
    End Function 'ListeExportParam
#End Region

    ''' <summary>
    ''' Retourne l'ID de sous commande le plus petit
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetSCMDMinID() As Integer
        Dim nReturn As Integer
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")

        Dim objOLeDBCommand As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        objOLeDBCommand = New OleDbCommand
        objOLeDBCommand.Connection = m_dbconn.Connection
        objOLeDBCommand.CommandText = "SELECT MIN(SCMD_ID) FROM SOUSCOMMANDE"
        objRS = objOLeDBCommand.ExecuteReader()
        If objRS.Read() Then
            nReturn = objRS.GetInt32(0)
            cleanErreur()
        End If
        Return nReturn
    End Function 'GetSCMDMinID
#Region "FicheTechniqueFourn"
    '=======================================================================
    'Fonction : InsertFicheTechniqueFourn
    'Description : Insertion d'une fiche technique founisseur
    'Retour : Rend Vrai si l'INSERT s'est correctement effectué
    '=======================================================================
    Protected Function insertFicheTechniqueFourn() As Boolean
        Dim bReturn As Boolean

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(Me.GetType.Name.Equals("FicheTechniqueFourn"), "Object de type FicheTechniqueFourn requis")
        Try
            'On utilise la techno dataset pour éviter de coder en dur les requêtes
            Dim obj As FicheTechniqueFourn
            obj = CType(Me, FicheTechniqueFourn)
            Dim oTa As dsVinicomTableAdapters.FICHETECHNIQUE_FOURNISSEURTableAdapter
            oTa = New dsVinicomTableAdapters.FICHETECHNIQUE_FOURNISSEURTableAdapter
            oTa.Connection = oleDBConnection
            Dim oDatatable As New dsVinicom.FICHETECHNIQUE_FOURNISSEURDataTable()
            Dim oRow As dsVinicom.FICHETECHNIQUE_FOURNISSEURRow
            oRow = oDatatable.NewFICHETECHNIQUE_FOURNISSEURRow()
            oRow.FTFRN_ID = obj.id
            oRow.FTFRN_TXTDOMAINE = obj.txtPresDomaine
            oRow.FTFRN_TXTFOURNISSEUR = obj.txtPresFournisseur
            oRow.FTFRN_TXTTERROIR = obj.txtTerroir
            oDatatable.AddFICHETECHNIQUE_FOURNISSEURRow(oRow)
            oTa.Update(oDatatable)
            setid(oRow.FTFRN_ID)
            bReturn = True

        Catch ex As Exception
            setError(ex.Message)
            bReturn = Nothing
        End Try
        Return bReturn
    End Function 'insertFicheTechniqueFourn
    '=======================================================================
    'Fonction : updateFicheTechniqueFourn
    'Description : Insertion d'une fiche technique  founisseur
    'Retour : Rend Vrai si l'INSERT s'est correctement effectué
    '=======================================================================
    Protected Function updateFicheTechniqueFourn() As Boolean
        Dim bReturn As Boolean

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(Me.GetType.Name.Equals("FicheTechniqueFourn"), "Object de type FicheTechniqueFourn requis")
        Debug.Assert(m_id <> 0, "L'ID doit être renseigné")
        Try
            Dim obj As FicheTechniqueFourn
            obj = CType(Me, FicheTechniqueFourn)
            Dim oTa As dsVinicomTableAdapters.FICHETECHNIQUE_FOURNISSEURTableAdapter
            oTa = New dsVinicomTableAdapters.FICHETECHNIQUE_FOURNISSEURTableAdapter
            oTa.Connection = oleDBConnection
            oTa.Update(obj.txtPresDomaine, obj.txtPresFournisseur, obj.txtTerroir, obj.specialite, id)
            bReturn = True

        Catch ex As Exception
            setError(ex.Message)
            bReturn = Nothing
        End Try
        Return bReturn
    End Function
    '=======================================================================
    'Fonction : deleteFicheTechniqueFourn
    'Description : Suppression d'une fiche technique founisseur
    'Retour : Rend Vrai si le DELETE s'est correctement effectué
    '=======================================================================
    Protected Function deleteFicheTechniqueFourn() As Boolean
        Dim bReturn As Boolean

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "L'ID doit être renseigné")
        Debug.Assert(Me.GetType.Name.Equals("FicheTechniqueFourn"), "Object de type FicheTechniqueFourn requis")
        Try
            Dim obj As FicheTechniqueFourn
            obj = CType(Me, FicheTechniqueFourn)
            Dim oTa As dsVinicomTableAdapters.FICHETECHNIQUE_FOURNISSEURTableAdapter
            oTa = New dsVinicomTableAdapters.FICHETECHNIQUE_FOURNISSEURTableAdapter
            oTa.Connection = oleDBConnection
            oTa.Delete(id)
            bReturn = True

        Catch ex As Exception
            setError(ex.Message)
            bReturn = Nothing
        End Try
        Return bReturn
    End Function 'deleteFicheTechniqueFourn
    Protected Function loadFicheTechniqueFourn() As Boolean

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "L'id doit être renseigné")

        Dim oTA As dsVinicomTableAdapters.FICHETECHNIQUE_FOURNISSEURTableAdapter = New dsVinicomTableAdapters.FICHETECHNIQUE_FOURNISSEURTableAdapter()
        Dim oDT As dsVinicom.FICHETECHNIQUE_FOURNISSEURDataTable
        Dim obj As FicheTechniqueFourn
        Dim bReturn As Boolean
        Dim oRow As dsVinicom.FICHETECHNIQUE_FOURNISSEURRow


        Try
            oTA.Connection = m_dbconn.Connection
            oDT = oTA.GetDataByID(m_id)
            If oDT.Rows.Count = 1 Then
                obj = CType(Me, FicheTechniqueFourn)
                oRow = oDT(0)
                obj.getData(oRow)
            Else
                m_id = 0
            End If

            bReturn = True
            cleanErreur()

        Catch ex As Exception
            setError("LoadFichetechniqueFourn", ex.ToString())
            bReturn = False
        End Try
        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'LoadFicheTechniqueFrn
    Protected Function loadFicheTechniqueFournImages() As Boolean

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "L'id doit être renseigné")

        Dim oTA As dsVinicomTableAdapters.IMAGESTableAdapter = New dsVinicomTableAdapters.IMAGESTableAdapter()
        Dim oDT As dsVinicom.IMAGESDataTable
        Dim obj As FicheTechniqueFourn
        Dim bReturn As Boolean
        Dim oRow As dsVinicom.IMAGESRow


        bReturn = False
        Try
            obj = CType(Me, FicheTechniqueFourn)
            oTA.Connection = m_dbconn.Connection
            oDT = oTA.GetDataByObjectID_Type_Num(m_id, "FTFRN_DOM", 1)
            If oDT.Rows.Count = 1 Then
                oRow = oDT(0)
                bReturn = obj.IMGDomaine.getData(oRow)
            End If
            If (bReturn) Then
                oDT = oTA.GetDataByObjectID_Type_Num(m_id, "FTFRN_FRN", 1)
                If oDT.Rows.Count = 1 Then
                    oRow = oDT(0)
                    bReturn = obj.IMGFournisseur.getData(oRow)
                End If
            End If

            cleanErreur()

        Catch ex As Exception
            setError("LoadFichetechniqueFournImages", ex.ToString())
            bReturn = False
        End Try
        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'LoadFicheTechniqueFrnImages
#End Region
#Region "vini_Image"
    '=======================================================================
    'Fonction : InsertImage
    'Description : Insertion d'une Image 
    'Retour : Rend Vrai si l'INSERT s'est correctement effectué
    '=======================================================================
    Protected Function insertImage() As Boolean
        Dim bReturn As Boolean

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(Me.GetType.Name.Equals("vini_Image"), "Object de type vini_Image requis")
        Try
            'On utilise la techno dataset pour éviter de coder en dur les requêtes
            Dim obj As vini_Image
            Dim oTa As New dsVinicomTableAdapters.IMAGESTableAdapter()
            Dim nId As Integer
            obj = CType(Me, vini_Image)
            oTa.Connection = oleDBConnection
            oTa.Insert(obj.IdObject, obj.Type, obj.Num, obj.Image, obj.Desc)
            nId = oTa.GETNEWID()
            setid(oTa.GETNEWID())

            bReturn = True

        Catch ex As Exception
            setError(ex.Message)
            bReturn = Nothing
        End Try
        Return bReturn
    End Function 'insertImage
    '=======================================================================
    'Fonction : updateImage
    'Description : Mise à jour d'une image en base de donnée
    'Retour : Rend Vrai si l'UPDATE s'est correctement effectué
    '=======================================================================
    Protected Function updateImage() As Boolean
        Dim bReturn As Boolean

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(Me.GetType.Name.Equals("vini_Image"), "Object de type vini_Image requis")
        Debug.Assert(m_id <> 0, "L'ID doit être renseigné")
        Try
            Dim obj As vini_Image
            obj = CType(Me, vini_Image)
            Dim oTa As New dsVinicomTableAdapters.IMAGESTableAdapter
            oTa.Connection = oleDBConnection
            oTa.Update(obj.IdObject, obj.Type, obj.Num, obj.Image, obj.Desc, id)
            bReturn = True

        Catch ex As Exception
            setError(ex.Message)
            bReturn = Nothing
        End Try
        Return bReturn
    End Function
    '=======================================================================
    'Fonction : deleteImage
    'Description : Suppression d'une Image
    'Retour : Rend Vrai si le DELETE s'est correctement effectué
    '=======================================================================
    Protected Function deleteImage() As Boolean
        Dim bReturn As Boolean

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "L'ID doit être renseigné")
        Debug.Assert(Me.GetType.Name.Equals("vini_Image"), "Object de type vini_Image requis")
        Try
            Dim obj As vini_Image
            obj = CType(Me, vini_Image)
            Dim oTa As New dsVinicomTableAdapters.IMAGESTableAdapter()
            oTa.Connection = oleDBConnection
            oTa.Delete(id)
            bReturn = True

        Catch ex As Exception
            setError(ex.Message)
            bReturn = Nothing
        End Try
        Return bReturn
    End Function 'deleteImage
    Protected Function loadImage() As Boolean

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "L'id doit être renseigné")

        Dim oTA As dsVinicomTableAdapters.IMAGESTableAdapter = New dsVinicomTableAdapters.IMAGESTableAdapter()
        Dim oDT As dsVinicom.IMAGESDataTable
        Dim obj As vini_Image
        Dim bReturn As Boolean
        Dim oRow As dsVinicom.IMAGESRow


        Try
            oTA.Connection = m_dbconn.Connection
            oDT = oTA.GetDataByID(m_id)
            If oDT.Rows.Count = 1 Then
                obj = CType(Me, vini_Image)
                oRow = oDT(0)
                obj.getData(oRow)
            Else
                m_id = 0
            End If

            bReturn = True
            cleanErreur()

        Catch ex As Exception
            setError("Persist.loadImage", ex.ToString())
            bReturn = False
        End Try
        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'LoadImage
    Protected Function loadImage(pObjectID As Integer, pType As String, pNum As Integer) As Boolean

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")

        Dim oTA As dsVinicomTableAdapters.IMAGESTableAdapter = New dsVinicomTableAdapters.IMAGESTableAdapter()
        Dim oDT As dsVinicom.IMAGESDataTable
        Dim obj As vini_Image
        Dim bReturn As Boolean
        Dim oRow As dsVinicom.IMAGESRow


        Try
            oTA.Connection = m_dbconn.Connection
            oDT = oTA.GetDataByObjectID_Type_Num(pObjectID, pType, pNum)
            If oDT.Rows.Count = 1 Then
                obj = CType(Me, vini_Image)
                oRow = oDT(0)
                obj.getData(oRow)
            Else
                m_id = 0
            End If

            bReturn = True
            cleanErreur()

        Catch ex As Exception
            setError("Persist.loadImage", ex.ToString())
            bReturn = False
        End Try
        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'LoadImage
#End Region
End Class
