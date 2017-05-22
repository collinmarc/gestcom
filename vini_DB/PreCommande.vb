Imports System.IO
'===================================================================================================================================
'Projet : Vinicom
'Auteur : Marc Collin 
'Création : 23/06/04
'===================================================================================================================================
' Classe : PreCommande Client   
' Description : PréCommandeClient = Liste des produits commandé par le client
'               Hérite de Commande
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
Public Class preCommande
    Inherits Commande

#Region "Accesseurs"
    Public Sub New(ByVal poClient As Client)
        MyBase.New(poClient, New EtatPrecommande())
        m_typedonnee = vncEnums.vncTypeDonnee.PRECOMMANDE
        majBooleenAlaFinDuNew()
    End Sub

    Public Function setClientId(ByVal pClient As Client) As Boolean
        Dim bReturn As Boolean
        Try
            setid(pClient.id)
            bReturn = True
        Catch ex As Exception
            setError("Precommande.setclientId", ex.Message)
            bReturn = False
        End Try
        Return bReturn
    End Function

    Public Shared Function createandload(ByVal pid As Integer) As preCommande
        '=======================================================================
        ' Contructeur pour chargement
        '=======================================================================
        Dim objPreCommande As preCommande
        Dim bReturn As Boolean
        objPreCommande = New preCommande(New Client("", ""))
        Try
            If pid <> 0 Then
                bReturn = objPreCommande.load(pid)
                If Not bReturn Then
                    setError("PreCommande.createAndLoad", getErreur())
                End If

            End If
        Catch ex As Exception
            setError("PreCommande.createAndLoad", ex.ToString)
        End Try
        Debug.Assert(objPreCommande.id = pid, "Commande Client " & pid & " non chargée")
        Return objPreCommande
    End Function 'createandload
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

        Debug.Assert(id <> 0, "idPrecommande<> 0")
        bReturn = LoadCLTPRECOMMANDE()
        Return bReturn
    End Function 'DBLoad
    Public Overrides Function save() As Boolean
        Dim bReturn As Boolean
        shared_connect()
        If bcolLignesUpated Then
            'On Supprime la precommande avant de la recréer
            bReturn = deletePRECOMMANDE()
            If bReturn Then
                bReturn = insertPRECOMMANDE(colLignes)
            End If
        Else
            bReturn = UPDATECLTPRECOMMANDE()
        End If
        shared_disconnect()
        Return bReturn
    End Function ' Save
    Friend Overrides Function delete() As Boolean
        '=======================================================================
        'Fonction : delete()
        'Description : Suppression de l'objet dans la base de l'objet
        'Détails    :  
        'Retour : Vrai si l'opération s'est bien déroulée
        '=======================================================================
        Debug.Assert(id <> 0, "idCommande <> 0")
        Dim bReturn As Boolean
        bReturn = deletecolLgCMD()
        Return bReturn

    End Function ' delete

    '=======================================================================
    'Fonction : insert()
    'Description : insertion de l'objet dans la base de l'objet
    'Détails    :  
    'Retour : Vrai di l'opération s'est bien déroulée
    '=======================================================================
    Friend Overrides Function insert() As Boolean
        Debug.Assert(Not oTiers Is Nothing, "Le Client n'est pas Renseigné")
        Debug.Assert(code.Equals(""), "Le Code doit être nul")
        Debug.Assert(oTiers.id <> 0, "Le Client n'est pas sauvegardé")
        Debug.Assert(id = 0, "idCommande = 0")

        Dim bReturn As Boolean
        If setNewcode() Then
            bReturn = insertPRECOMMANDE(m_colLignes)
        End If
        Return bReturn
    End Function 'insert
    '=======================================================================
    'Fonction : Update()
    'Description : Mise à jour de l'objet
    'Détails    :  
    'Retour : Vrai di l'opération s'est bien déroulée
    '=======================================================================
    Friend Overrides Function update() As Boolean
        Debug.Assert(id <> 0, "id <> 0")
        Dim bReturn As Boolean
        bReturn = UPDATECLTPRECOMMANDE()
        Return bReturn

    End Function 'Update

    ''' <summary>
    ''' Purge des lignes de précommandes
    ''' </summary>
    ''' <param name="pdtFin">date Max</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Purge(pdtFin As Date) As Boolean
        Return Persist.PurgeLgPrecommande(pdtFin)
    End Function
#End Region

#Region "Methodes overrides"
    Public Overrides Function checkForDelete() As Boolean

    End Function

    Public Overrides Function ControleMvtStock() As Microsoft.VisualBasic.Collection
        Return Nothing
    End Function

    Public Function toCSV() As String
        Debug.Assert(bcolLignesLoaded, "Les Lignes doivent être chargées au préalable")
        Dim strResult1Line As String
        Dim strResult As String
        Dim objLgPrecom As lgPrecomm
        Dim objProduit As Produit
        Dim objClient As Client

        objClient = Client.createandload(m_id)

        strResult = ""
        For Each objLgPrecom In colLignes
            strResult1Line = ""
            objProduit = Produit.createandload(objLgPrecom.idProduit) 'Chargement du produit
            strResult1Line = strResult1Line & Trim(objClient.code) & ";"
            strResult1Line = strResult1Line & Trim(objProduit.code) & ";"
            strResult1Line = strResult1Line & Trim(objLgPrecom.qteDern) & ";"
            strResult1Line = strResult1Line & Trim(objLgPrecom.prixU) & ";"
            strResult1Line = strResult1Line & Trim(Format(objLgPrecom.dateDerniereCommande, "ddMMyyyy")) & ";"
            strResult1Line = strResult1Line & Trim(objLgPrecom.refDerniereCommande)
            strResult = strResult & strResult1Line & vbCrLf
        Next objLgPrecom
        Return strResult

    End Function 'ToCSV

    Public Overrides Sub Exporter(ByVal pstrFileName As String)
        ExporterPreCommande(pstrFileName, "PRC")
    End Sub

    ''' <summary>
    ''' Export Quadra
    ''' </summary>
    ''' <param name="pstFileName"></param>
    ''' <param name="pExportType"></param>
    ''' <remarks></remarks>
    Protected Sub ExporterPreCommande(ByVal pstFileName As String, ByVal pExportType As String)
        Dim strLine As String
        Dim oTA As vini_DB.dsVinicomTableAdapters.EXPORTPARAMTableAdapter
        Dim oTAConstantes As vini_DB.dsVinicomTableAdapters.CONSTANTESTableAdapter
        Dim nbLignes As Integer
        Dim nLigne As Integer
        Dim nbChamps As Integer
        Dim nChamps As Integer
        Dim oDT As vini_DB.dsVinicom.EXPORTPARAMDataTable
        Dim oRow As vini_DB.dsVinicom.EXPORTPARAMRow
        Dim oDTConstantes As vini_DB.dsVinicom.CONSTANTESDataTable
        Dim strValeur As String
        Dim objLgPrecom As lgPrecomm

        Try
            'Chargement des informatios du tiers
            oTiers.load()
            strLine = String.Empty
            oTA = New vini_DB.dsVinicomTableAdapters.EXPORTPARAMTableAdapter()
            oTA.Connection = Persist.oleDBConnection
            oTAConstantes = New vini_DB.dsVinicomTableAdapters.CONSTANTESTableAdapter()
            oTAConstantes.Connection = Persist.oleDBConnection

            oDTConstantes = oTAConstantes.GetData()
            'Recupération du nombre de lignes de paramètre
            nbLignes = oTA.getNbLignes(pExportType)

            For Each objLgPrecom In colLignes
                For nLigne = 1 To nbLignes
                    ' Pour chaque ligne
                    'Lecture des paramètres
                    oDT = oTA.GetDataBy_TYPE_NLIGNE(pExportType, nLigne)
                    'Recupéation du nombre de champs
                    nbChamps = oTA.getNbChamps(pExportType, nLigne)
                    For nChamps = 1 To nbChamps
                        'Pour chaque Champs
                        oRow = oDT.FindByEXP_TYPEEXP_NLIGNEEXP_NCHAMPS(pExportType, nLigne, nChamps)
                        If oRow IsNot Nothing Then
                            If Trim(oRow.EXP_TYPECHAMPS).Equals("C") Then
                                'Exportation d'une Constante
                                strLine = strLine + Left(oRow.EXP_VALEUR + Space(oRow.EXP_LONGUEUR), oRow.EXP_LONGUEUR)
                            End If
                            If Trim(oRow.EXP_TYPECHAMPS).Equals("V") Then
                                'Exportation d'une Valeur
                                strValeur = getAttributeValue(oRow.EXP_VALEUR, oDTConstantes, objLgPrecom)
                                strLine = strLine + Left(strValeur + Space(oRow.EXP_LONGUEUR), oRow.EXP_LONGUEUR)
                            End If
                        End If
                    Next 'champs
                    strLine = strLine + vbCrLf

                Next 'Lignes
            Next

            File.AppendAllText(pstFileName, strLine)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Function getAttributeValue(ByVal pstrAttributeName As String, ByVal pDTConstantes As dsVinicom.CONSTANTESDataTable, ByVal pLgPrecom As lgPrecomm) As String
        Dim strReturn As String
        strReturn = String.Empty

        Select Case pstrAttributeName
            Case "IDCLIENT"
                strReturn = m_id
            Case "CODECLIENT"
                strReturn = Client.createandload(m_id).code
            Case "CODEPRODUIT"
                strReturn = getTXTString(pLgPrecom.codeProduit)
            Case "QTEHAB"
                strReturn = pLgPrecom.qteHab.ToString()
            Case "QTEDERN"
                strReturn = pLgPrecom.qteDern.ToString()
            Case "PRIXU"
                strReturn = pLgPrecom.prixU.ToString()
            Case "DATEDERNCOMMANDE"
                strReturn = Format(pLgPrecom.dateDerniereCommande, "ddMMyy")
            Case "REFDERNCOMMANDE"
                strReturn = pLgPrecom.refDerniereCommande
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


    Public Overrides Function genereMvtStock() As Boolean

    End Function

    Friend Overrides Function setNewcode() As Boolean

    End Function

    Public Overrides Function supprimeMvtStock() As Boolean

    End Function

#End Region

    '=======================================================================
    'Fonction : AjouteLgPrecomm
    'Description : Ajoute une ligne LgPrecom au client parametre = Produit
    'Retour : Boolean
    '=======================================================================
    Public Function ajouteLgPrecom(ByVal objProduit As Produit, Optional ByVal qteHab As Double = 0, Optional ByVal qteDern As Double = 0, Optional ByVal pPrixU As Double = 0, Optional ByVal pDateDernCom As Date = DATE_DEFAUT, Optional ByVal prefDernCommande As String = "") As lgPrecomm
        Dim oLgPrecom As lgPrecomm
        oLgPrecom = ajouteLgPrecom(objProduit.id, objProduit.code, objProduit.nom, qteHab, qteDern, pPrixU, pDateDernCom, prefDernCommande)
        oLgPrecom.millesime = objProduit.millesime
        oLgPrecom.libConditionnement = objProduit.libConditionnement
        Return oLgPrecom
    End Function 'AjouteLgPrecomm
    '=======================================================================
    'Fonction : AjouteLgPrecomm
    'Description : Ajoute une ligne LgPrecom au client parametres (idProduit, CodeProduit,LibelleProduit)
    'Retour : LgPrecomm
    '=======================================================================
    Public Function ajouteLgPrecom(ByVal idProduit As Integer, ByVal codeProduit As String, ByVal libProduit As String, Optional ByVal qteHab As Decimal = 0, Optional ByVal qteDern As Decimal = 0, Optional ByVal pPrixU As Double = 0, Optional ByVal pDateDernCom As Date = DATE_DEFAUT, Optional ByVal prefDernCommande As String = "") As lgPrecomm
        Debug.Assert(idProduit <> 0, "ID Produit <> 0")

        Dim objLgPrecom As lgPrecomm
        objLgPrecom = Nothing
        If Not (lgPrecomExists(idProduit)) Then
            objLgPrecom = New lgPrecomm(m_id, idProduit, codeProduit, libProduit, qteHab, qteDern, pPrixU, pDateDernCom, prefDernCommande)
            AjouteLigne(objLgPrecom, False)
        End If

        Return objLgPrecom
    End Function 'AjouteLgPrecomm
    Public Function ajouteLgPrecom() As lgPrecomm

        Dim objLgPrecom As lgPrecomm
        objLgPrecom = Nothing
        objLgPrecom = New lgPrecomm(m_id)
        m_colLignes.Add(objLgPrecom, "")
        Return objLgPrecom
    End Function 'AjouteLgPrecomm
    '=======================================================================
    'Fonction : SupprimeLgPrecomm
    'Description : Supprime une ligne de precommande au client
    'Retour : Boolean
    '=======================================================================
    Public Function supprimeLgPrecom(ByVal nIndex As Integer) As Boolean
        Debug.Assert(nIndex <= m_colLignes.Count And nIndex > 0, "nIndex ")

        Try
            m_colLignes.Remove(nIndex)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function 'supprimeLgPrecomm

    '=======================================================================
    'Fonction : SupprimePrecomm
    'Description : Supprime la collection de precommande du client
    'Retour : Boolean
    '=======================================================================
    Public Function supprimePrecom() As Boolean
        Try
            m_colLignes.clear()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function 'supprimePrecomm

    '=======================================================================
    'Fonction : getLgPrecom
    'Description : Renvoie la ligne de precommande
    'Retour : Boolean
    '=======================================================================
    Public Function getLgPrecom(ByVal pidProduit As Integer) As lgPrecomm
        Dim objLgPrecomm As lgPrecomm

        Try
            objLgPrecomm = m_colLignes(CStr(pidProduit))
        Catch ex As Exception
            objLgPrecomm = Nothing
        End Try
        Return objLgPrecomm
    End Function 'getLgPrecomm
    '=======================================================================
    'Fonction : lgPrecomExists
    'Description : Renvoie Vrai si la ligne de precommande existe
    'Retour : Boolean
    '=======================================================================
    Public Function lgPrecomExists(ByVal pIdProduit As Integer) As Boolean
        Dim bReturn As Boolean

        Try
            bReturn = m_colLignes.keyExists(pIdProduit)
        Catch ex As Exception
            bReturn = False
        End Try
        Return bReturn
    End Function 'lgPrecommExists
    '=======================================================================
    'Fonction : getlgPrecomCount
    'Description : Renvoie Le nombre de ligne de precommande
    'Retour : integer
    '=======================================================================
    Public Function getlgPrecomCount() As Integer
        Dim nReturn As Integer

        Try
            nReturn = m_colLignes.Count
        Catch ex As Exception
            nReturn = 0
        End Try
        Return nReturn
    End Function 'lgPrecommCount

End Class ' preCommande
