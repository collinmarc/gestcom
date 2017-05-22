'Test de la classe parametre
Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports vini_DB



<TestClass()> Public Class test_param
    Inherits test_Base
    Private nMaxId As Integer
    <TestInitialize()> Public Overrides Sub TestInitialize()
        MyBase.TestInitialize()
        Persist.shared_connect()
        nMaxId = Persist.executeSQLQuery("SELECT MAX(PAR_ID) FROM PARAMETRE")
    End Sub
    <TestCleanup()> Public Overrides Sub TestCleanup()
        MyBase.TestCleanup()
        Persist.executeSQLQuery("DELETE PARAMETRE WHERE PAR_ID >" & nMaxId)
        Persist.shared_disconnect()
    End Sub
    <TestMethod()> Public Sub T10_Object()
        Dim objPAR As Param
        Dim objPAR2 As Param

        objPAR = New Param
        objPAR.code = "CODE"
        objPAR.type = "R"
        objPAR.valeur = "COMPTANT"

        Assert.AreEqual(objPAR.code, "CODE")
        Assert.AreEqual(objPAR.type, "R")
        Assert.AreEqual(objPAR.valeur, "COMPTANT")



        Assert.IsTrue(objPAR.Equals(objPAR), "Egal à Lui même")
        objPAR2 = New Param
        objPAR2.code = "CODE"
        objPAR2.type = "R"
        objPAR2.valeur = "COMPTANT"
        Assert.IsTrue(objPAR.Equals(objPAR2), "Egal à un semblable")
        objPAR2.valeur = "XXX"
        Assert.IsFalse(objPAR.Equals(objPAR2), "Egal à un Différent")
        Dim obj As Object
        Assert.IsFalse(objPAR.Equals(obj), "Egal autrecjhose")


    End Sub
    <TestMethod()> Public Sub T20_ColParam()
        Dim objcol As Collection
        Dim objPar As Param

        'Console.Out.WriteLine("Conditionnement")
        objcol = Param.colConditionnement()
        Assert.IsTrue(objcol.Count > 0, "ColConditionnement.Count > 0")
        For Each objPar In objcol
            Assert.IsTrue(objPar.type = "D")
            'Console.Out.WriteLine(objPar.toString())
        Next
        Assert.IsTrue(Param.conditionnementdefaut.defaut, "Pas de Conditionnement par defaut")

        objcol = Param.colModeReglement()
        Assert.IsTrue(objcol.Count > 0, "colModeReglement.Count > 0")
        For Each objPar In objcol
            Assert.AreEqual(objPar.type, "R")
        Next
        Assert.IsTrue(Param.ModeReglementdefaut.defaut, "Pas de Mode de reglement par defaut")
        objcol = Param.colCouleur()
        Assert.IsTrue(objcol.Count > 0, "colCouleur.Count > 0")
        For Each objPar In objcol
            Assert.AreEqual(objPar.type, "O")
        Next
        Assert.IsTrue(Param.couleurdefaut.defaut, "Pas de Couleur par defaut")

        objcol = Param.colRegion()
        Assert.IsTrue(objcol.Count > 0, "colRegion.Count > 0")
        For Each objPar In objcol
            Assert.AreEqual(objPar.type, "G")
        Next
        Assert.IsTrue(Param.regiondefaut.defaut, "Pas de Region par defaut")

        objcol = Param.colTVA()
        Assert.IsTrue(objcol.Count > 0, "colRegion.Count > 0")
        For Each objPar In objcol
            Assert.AreEqual(objPar.type, "V")
        Next
        Assert.IsTrue(Param.TVAdefaut.defaut, "Pas de TVA par defaut")

        'Assert.IsFalse(Param.getConstante("CST_NUMFAXPLATEFORME").Equals(""))
        Assert.IsTrue(Param.getConstante("CST_RIEN").Equals(""))
        '            Assert.IsTrue(Not Persist.shared_isConnected(), "La Base est encore connectée")
    End Sub
    <TestMethod()> Public Sub T30_ObjectContenant()
        Dim objPAR As contenant

        Dim objPAR2 As contenant

        objPAR = New contenant
        objPAR.code = "CONTCODE"
        objPAR.libelle = "CONTLIB"
        objPAR.cent = 37.5
        objPAR.bout = 3

        Assert.AreEqual(objPAR.id, 0)
        Assert.AreEqual(objPAR.code, "CONTCODE")
        Assert.AreEqual(objPAR.libelle, "CONTLIB")
        Assert.AreEqual(objPAR.cent, 37.5)
        Assert.AreEqual(objPAR.bout, 3)



        Assert.IsTrue(objPAR.Equals(objPAR), "Egal à Lui même")
        objPAR2 = New contenant
        objPAR2.code = "CONTCODE"
        objPAR2.libelle = "CONTLIB"
        objPAR2.cent = 37.5
        objPAR2.bout = 3
        Assert.IsTrue(objPAR.Equals(objPAR2), "Egal à un semblable")
        objPAR2.cent = 150
        Assert.IsFalse(objPAR.Equals(objPAR2), "Egal à un Différent")
        Dim obj As Object
        Assert.IsFalse(objPAR.Equals(obj), "Egal autrecjhose")


    End Sub

    <TestMethod()> Public Sub T40_colContenant()
        Dim objcol As Collection
        Dim objPar As contenant

        'Console.Out.WriteLine("Contenant")
        objcol = contenant.colContenant()
        'Console.Out.WriteLine(Persist.getErreur)
        Assert.IsTrue(objcol.Count > 0, "Colcontenant.Count > 0")
        For Each objPar In objcol
            'Console.Out.WriteLine(objPar.toString())
        Next
        Assert.IsTrue(contenant.contenantDefaut.defaut, "Pas de Contenant par defaut")
    End Sub
    <TestMethod()> Public Sub T50_Constantes()
        Dim strVersion As String = Param.getConstante("CST_VERSION_BD")
        Assert.IsTrue(Len(strVersion) > 0)

    End Sub

    <TestMethod()> Public Sub T60_DB()
        Dim objPAR As Param
        Dim nId As Integer

        objPAR = New Param
        objPAR.code = "CODE"
        objPAR.type = "R"
        objPAR.valeur = "COMPTANT"

        Assert.IsTrue(objPAR.Save())
        Assert.AreNotEqual(0, objPAR.id)

        nId = objPAR.id

        objPAR = New Param
        objPAR.load(nId)
        Assert.AreEqual("CODE", objPAR.code)
        Assert.AreEqual("R", objPAR.type)
        Assert.AreEqual("COMPTANT", objPAR.valeur)

        objPAR.valeur = "VALEUR2"

        Assert.IsTrue(objPAR.Save())

        objPAR = New Param
        objPAR.load(nId)
        Assert.AreEqual("CODE", objPAR.code)
        Assert.AreEqual("R", objPAR.type)
        Assert.AreEqual("VALEUR2", objPAR.valeur)

        objPAR.bDeleted = True
        Assert.IsTrue(objPAR.Save())
        objPAR = New Param
        objPAR.load(nId)
        Assert.AreEqual(0, objPAR.code.Length)

    End Sub
    <TestMethod()> Public Sub T60_DB_ModeReglement()
        Dim objPAR As ParamModeReglement
        Dim nId As Integer

        objPAR = New ParamModeReglement
        objPAR.code = "CODE"
        objPAR.type = "R"
        objPAR.valeur = "COMPTANT"
        objPAR.valeur2 = "30"
        objPAR.dDebutEcheance = "FDM"

        Assert.IsTrue(objPAR.Save())
        Assert.AreNotEqual(0, objPAR.id)

        nId = objPAR.id

        objPAR = New ParamModeReglement
        objPAR.load(nId)
        Assert.AreEqual("CODE", objPAR.code)
        Assert.AreEqual("R", objPAR.type)
        Assert.AreEqual("COMPTANT", objPAR.valeur)
        Assert.AreEqual("30", objPAR.valeur2)
        Assert.AreEqual("FDM", objPAR.dDebutEcheance)

        objPAR.valeur = "VALEUR2"
        objPAR.dDebutEcheance = "FACT"

        Assert.IsTrue(objPAR.Save())

        objPAR = New ParamModeReglement
        objPAR.load(nId)
        Assert.AreEqual("CODE", objPAR.code)
        Assert.AreEqual("R", objPAR.type)
        Assert.AreEqual("VALEUR2", objPAR.valeur)
        Assert.AreEqual("30", objPAR.valeur2)
        Assert.AreEqual("FACT", objPAR.dDebutEcheance)

        objPAR.bDeleted = True
        Assert.IsTrue(objPAR.Save())
        objPAR = New ParamModeReglement
        objPAR.load(nId)
        Assert.AreEqual(0, objPAR.code.Length)

    End Sub
    <TestMethod()> Public Sub T60_DB_Contenant()
        Dim objPAR As contenant
        Dim nId As Integer

        objPAR = New contenant
        objPAR.code = "CODE"
        objPAR.libelle = "R"
        objPAR.bout = 1.5
        objPAR.cent = 1.25
        objPAR.poids = 0.5

        Assert.IsTrue(objPAR.Save())
        Assert.AreNotEqual(0, objPAR.id)

        nId = objPAR.id

        objPAR = New contenant
        objPAR.load(nId)
        Assert.AreEqual("CODE", objPAR.code)
        Assert.AreEqual("R", objPAR.libelle)
        Assert.AreEqual(1.5, objPAR.bout)
        Assert.AreEqual(1.25, objPAR.cent)
        Assert.AreEqual(0.5, objPAR.poids)

        objPAR.poids = 2

        Assert.IsTrue(objPAR.Save())

        objPAR = New contenant
        objPAR.load(nId)
        Assert.AreEqual("CODE", objPAR.code)
        Assert.AreEqual("R", objPAR.libelle)
        Assert.AreEqual(1.5, objPAR.bout)
        Assert.AreEqual(1.25, objPAR.cent)
        Assert.AreEqual(2, objPAR.poids)

        objPAR.bDeleted = True
        Assert.IsTrue(objPAR.Save())
        objPAR = New contenant
        objPAR.load(nId)
        Assert.AreEqual(0, objPAR.code.Length)

    End Sub
    <TestMethod()> Public Sub T100_CalcDateEcheance()
        Dim objParam As ParamModeReglement
        Dim ddeb As Date

        objParam = New ParamModeReglement()
        ddeb = "05/02/1964"

        objParam.dDebutEcheance = "FACT"
        objParam.valeur2 = "30"
        Assert.AreEqual(CDate("05/03/1964"), objParam.calDateEcheance(ddeb))

        objParam.dDebutEcheance = "FACT"
        objParam.valeur2 = "60"
        Assert.AreEqual(CDate("05/04/1964"), objParam.calDateEcheance(ddeb))

        objParam.dDebutEcheance = "FACT"
        objParam.valeur2 = "90"
        Assert.AreEqual(CDate("05/05/1964"), objParam.calDateEcheance(ddeb))

        objParam.dDebutEcheance = "FACT"
        objParam.valeur2 = "20"
        Assert.AreEqual(CDate("25/02/1964"), objParam.calDateEcheance(ddeb))

        objParam.dDebutEcheance = "FDM"
        objParam.valeur2 = "30"
        Assert.AreEqual(CDate("31/03/1964"), objParam.calDateEcheance(ddeb))

        objParam.dDebutEcheance = "FDM"
        objParam.valeur2 = "60"
        Assert.AreEqual(CDate("30/04/1964"), objParam.calDateEcheance(ddeb))

        objParam.dDebutEcheance = "FDM"
        objParam.valeur2 = "90"
        Assert.AreEqual(CDate("31/05/1964"), objParam.calDateEcheance(ddeb))

        objParam.dDebutEcheance = "FDM"
        objParam.valeur2 = "20"
        Assert.AreEqual(CDate("20/03/1964"), objParam.calDateEcheance(ddeb))
    End Sub
End Class


