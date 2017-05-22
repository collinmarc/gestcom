'Test de la classe dbConnection
Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports vini_DB
Imports System.Globalization



<TestClass()> Public Class test_Base

    Dim m_idCLT As Integer
    Dim m_idFRN As Integer
    Dim m_idPRD As Integer
    Dim m_idCMD As Integer
    Dim m_idBA As Integer
    Dim m_idFactCOL As Integer
    Dim m_idFactCOM As Integer
    Dim m_idFactTRP As Integer
    Dim m_idFRNComm As Integer
    Dim m_idLGFactCOL As Integer
    Dim m_idLGFactTRP As Integer
    Dim m_idLGCMD As Integer
    Dim m_idMVTStock As Integer
    Dim m_idPRECOMMANDE As Integer
    Dim m_idREGLEMENT As Integer
    Dim m_idSCMD As Integer
    Dim m_idTRP As Integer
    Dim m_idParam As Integer
    Dim m_idContenant As Integer
    Dim m_idFicheTechniqueFourn As Integer
    Dim m_idImage As Integer


    <TestInitialize()> Public Overridable Sub TestInitialize()
        If System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator <> "." Then
            'Le séparateur instancié dans le panneau de configuration est la virgule : ","
            Dim forceDotCulture As CultureInfo
            'Code un peu louche il faut avouer, mais il faut faire avec car le framework pose problème
            'ici; en effet, il faut cloner la culture pour pouvoir modifier les paramètres de l'application
            'car sinon la culture de base est en lecture seule.
            forceDotCulture = CultureInfo.CurrentCulture.Clone()
            'On affecte le point : "." comme paramètre de séparateur décimal
            forceDotCulture.NumberFormat.NumberDecimalSeparator = "."
            'Là, on affecte l'application cloné à celle où l'on travaille 
            'C'est un passage flou car en fait, l'appli est en mode readonly et l'on ne peut pas
            'la modifier directement, d'où cette affectation
            System.Threading.Thread.CurrentThread.CurrentCulture = forceDotCulture
        End If
        initConstantes()
        Persist.shared_connect()
        Param.LoadcolParams()
        Assert.IsTrue(getIdsReference())
        Persist.shared_disconnect()
    End Sub
    <TestCleanup()> Public Overridable Sub TestCleanup()
        Persist.shared_connect()
        Assert.IsTrue(cleanDataTest())
        Persist.shared_disconnect()
    End Sub

    Private Function getIdsReference() As Boolean

        Dim bReturn As Boolean

        Try

            m_idCLT = CInt(Persist.executeSQLQuery("SELECT MAX(CLT_ID) FROM CLIENT"))
            m_idFRN = CInt(Persist.executeSQLQuery("SELECT MAX(FRN_ID) FROM FOURNISSEUR"))
            m_idPRD = CInt(Persist.executeSQLQuery("SELECT MAX(PRD_ID) FROM PRODUIT"))
            m_idCMD = CInt(Persist.executeSQLQuery("SELECT MAX(CMD_ID) FROM COMMANDE"))
            m_idBA = CInt(Persist.executeSQLQuery("SELECT MAX(CMD_ID) FROM BONAPPRO"))
            m_idFactCOL = CInt(Persist.executeSQLQuery("SELECT MAX(FCOL_ID) FROM FACTCOLISAGE"))
            m_idFactCOM = CInt(Persist.executeSQLQuery("SELECT MAX(FCT_ID) FROM FACTCOM"))
            m_idFactTRP = CInt(Persist.executeSQLQuery("SELECT MAX(FTRP_ID) FROM FACTTRP"))
            m_idFRNComm = CInt(Persist.executeSQLQuery("SELECT MAX(FRNC_ID) FROM FRN_COMM"))
            m_idLGFactCOL = CInt(Persist.executeSQLQuery("SELECT MAX(LGCOL_ID) FROM LGFACTCOLISAGE"))
            m_idLGFactTRP = CInt(Persist.executeSQLQuery("SELECT MAX(LGTRP_ID) FROM LGFACTTRP"))
            m_idLGCMD = CInt(Persist.executeSQLQuery("SELECT MAX(LGCM_ID) FROM LIGNE_COMMANDE"))
            m_idMVTStock = CInt(Persist.executeSQLQuery("SELECT MAX(STK_ID) FROM MVT_STOCK"))
            m_idPRECOMMANDE = CInt(Persist.executeSQLQuery("SELECT MAX(PCMD_ID) FROM PRECOMMANDE"))
            m_idREGLEMENT = CInt(Persist.executeSQLQuery("SELECT MAX(RGL_ID) FROM REGLEMENT"))
            m_idSCMD = CInt(Persist.executeSQLQuery("SELECT MAX(SCMD_ID) FROM SOUSCOMMANDE"))
            m_idTRP = CInt(Persist.executeSQLQuery("SELECT MAX(TRP_ID) FROM TRANSPORTEUR"))
            m_idParam = CInt(Persist.executeSQLQuery("SELECT MAX(PAR_ID) FROM PARAMETRE"))
            m_idContenant = CInt(Persist.executeSQLQuery("SELECT MAX(CONT_ID) FROM CONTENANT"))
            Try
                m_idFicheTechniqueFourn = CInt(Persist.executeSQLQuery("SELECT MAX(FTFRN_ID) FROM FICHETECHNIQUE_FOURNISSEUR"))
            Catch
                m_idFicheTechniqueFourn = 0
            End Try
            Try
                m_idImage = CInt(Persist.executeSQLQuery("SELECT MAX(IMG_ID) FROM IMAGES"))
            Catch
                m_idImage = 0
            End Try
            bReturn = True
        Catch ex As Exception
            bReturn = False

        End Try

            Return bReturn
    End Function

    Private Function cleanDataTest() As Boolean

        Dim bReturn As Boolean

        Try

            Persist.executeSQLNonQuery("DELETE FROM FICHETECHNIQUE_FOURNISSEUR WHERE FTFRN_ID > " + m_idFicheTechniqueFourn.ToString())
            Persist.executeSQLNonQuery("DELETE FROM IMAGE WHERE IMG_ID > " + m_idImage.ToString())
            Persist.executeSQLNonQuery("DELETE FROM REGLEMENT WHERE RGL_ID > " + m_idREGLEMENT.ToString())
            Persist.executeSQLNonQuery("DELETE FROM LGFACTCOLISAGE WHERE LGCOL_ID > " + m_idLGFactCOL.ToString())
            Persist.executeSQLNonQuery("DELETE FROM FACTCOLISAGE WHERE FCOL_ID > " + m_idFactCOL.ToString())
            Persist.executeSQLNonQuery("DELETE FROM LGFACTTRP WHERE LGTRP_ID > " + m_idLGFactTRP.ToString())
            Persist.executeSQLNonQuery("DELETE FROM FACTTRP WHERE FTRP_ID > " + m_idFactTRP.ToString())
            Persist.executeSQLNonQuery("DELETE FROM FACTCOM WHERE FCT_ID > " + m_idFactCOM.ToString())
            Persist.executeSQLNonQuery("DELETE FROM LIGNE_COMMANDE WHERE LGCM_ID > " + m_idLGCMD.ToString())
            Persist.executeSQLNonQuery("DELETE FROM SOUSCOMMANDE WHERE SCMD_ID > " + m_idSCMD.ToString())
            Persist.executeSQLNonQuery("DELETE FROM COMMANDE where CMD_ID > " + m_idCMD.ToString())
            Persist.executeSQLNonQuery("DELETE FROM BONAPPRO WHERE CMD_ID > " + m_idBA.ToString())
            Persist.executeSQLNonQuery("DELETE FROM MVT_STOCK WHERE STK_ID > " + m_idMVTStock.ToString())
            Persist.executeSQLNonQuery("DELETE FROM PRECOMMANDE WHERE PCMD_ID > " + m_idPRECOMMANDE.ToString())

            Persist.executeSQLNonQuery("DELETE FROM PRODUIT WHERE PRD_ID > " + m_idPRD.ToString())
            Persist.executeSQLNonQuery("DELETE FROM CLIENT WHERE CLT_ID> " + m_idCLT.ToString())
            Persist.executeSQLNonQuery("DELETE FROM FOURNISSEUR WHERE FRN_ID > " + m_idFRN.ToString())
            Persist.executeSQLNonQuery("DELETE FROM FRN_COMM WHERE FRNC_ID > " + m_idFRNComm.ToString())
            Persist.executeSQLNonQuery("DELETE FROM TRANSPORTEUR WHERE TRP_ID > " + m_idTRP.ToString())
            Persist.executeSQLNonQuery("DELETE FROM PARAMETRE WHERE PAR_ID > " + m_idParam.ToString())
            Persist.executeSQLNonQuery("DELETE FROM CONTENANT WHERE CONT_ID > " + m_idContenant.ToString())

            bReturn = True
        Catch ex As Exception
            bReturn = False

        End Try

        Return bReturn
    End Function

    <TestMethod()> Sub T00_Base()
        Dim obj As Fournisseur
        obj = New Fournisseur("TST", "nom")
        Assert.IsTrue(obj.Save)
    End Sub
End Class






