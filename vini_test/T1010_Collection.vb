'Test de la classe dbConnection
Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports vini_DB



<TestClass()> Public Class T1010_Collection
    Inherits test_Base


    <TestMethod()> Public Sub T01_SortedList()
        Dim objcol As SortedList
        Dim obj As DictionaryEntry
        Dim objmvt As mvtStock
        objcol = New SortedList
        objmvt = New mvtStock(Now(), 0, vncEnums.vncTypeMvt.vncmvtBonAppro, 10, "Premier mvt")
        objcol.Add(objmvt.key, objmvt)

        objmvt = New mvtStock(Now(), 0, vncEnums.vncTypeMvt.vncmvtBonAppro, 10, "Second mvt")
        If objcol.ContainsKey(objmvt.key) Then
            objcol.Add(objmvt.key & "2", objmvt)
        End If
        For Each obj In objcol
            Console.WriteLine(obj.Value.ToString)
        Next

    End Sub
    <TestMethod()> Public Sub T01_ColEventSorted()
        Dim objcol As ColEventSorted
        Dim objmvt As mvtStock
        objcol = New ColEventSorted

        objmvt = New mvtStock("01/06/2004", 0, vncEnums.vncTypeMvt.vncmvtBonAppro, 10, "Premier mvt 01/06/2004")
        objcol.Add(objmvt, objmvt.key)

        objmvt = New mvtStock("30/05/2004", 0, vncEnums.vncTypeMvt.vncmvtBonAppro, 10, "Second mvt 30/05/2004")
        objcol.Add(objmvt, objmvt.key)

        objmvt = New mvtStock("30/05/2003", 0, vncEnums.vncTypeMvt.vncmvtBonAppro, 10, "Troiseime mvt 30/05/2003")
        objcol.Add(objmvt, objmvt.key)

        objmvt = New mvtStock("30/05/2005", 0, vncEnums.vncTypeMvt.vncmvtBonAppro, 10, "4eme mvt 30/05/2005")
        objcol.Add(objmvt, objmvt.key)

        objmvt = New mvtStock("01/01/2000", 0, vncEnums.vncTypeMvt.vncmvtBonAppro, 10, "5eme mvt 01/01/2000")
        objcol.Add(objmvt, objmvt.key)

        objmvt = New mvtStock("30/05/2005", 0, vncEnums.vncTypeMvt.vncMvtInventaire, 10, "6eme mvt Inventaire 30/05/2005")
        objcol.Add(objmvt, objmvt.key)
        objmvt = objcol(5)
        Assert.IsTrue(objmvt.datemvt.ToShortDateString.Equals("01/01/2000"))
        objmvt = objcol(4)
        Assert.IsTrue(objmvt.datemvt.ToShortDateString.Equals("30/05/2003"))
        objmvt = objcol(3)
        Assert.IsTrue(objmvt.datemvt.ToShortDateString.Equals("30/05/2004"))
        objmvt = objcol(2)
        Assert.IsTrue(objmvt.datemvt.ToShortDateString.Equals("01/06/2004"))
        objmvt = objcol(1)
        Assert.IsTrue(objmvt.datemvt.ToShortDateString.Equals("30/05/2005"))
        Assert.IsTrue(objmvt.typeMvt = vncEnums.vncTypeMvt.vncMvtInventaire)
        objmvt = objcol(0)
        Assert.IsTrue(objmvt.datemvt.ToShortDateString.Equals("30/05/2005"))
        Assert.IsTrue(objmvt.typeMvt = vncEnums.vncTypeMvt.vncmvtBonAppro)
        For Each objmvt In objcol
            Console.WriteLine(objmvt.toString)
        Next

    End Sub

    <TestMethod()> Public Sub T10_DupplicateKey()
        Dim objcol As ColEvent
        Dim objLgCmd As LgCommande
        Dim bReturn As Boolean

        objLgCmd = New LgCommande(0)
        objLgCmd.num = "10"
        objcol = New ColEvent
        Try
            objcol.Add(objLgCmd, CStr(objLgCmd.num))
            objcol.Add(objLgCmd, CStr(objLgCmd.num))
            objcol.Add(objLgCmd, CStr(objLgCmd.num))
            bReturn = False 'Erreur car normalement on ne peut pas ajouter de clés duppiquées
        Catch ex As Exception
            bReturn = True
        End Try
        Assert.IsTrue(bReturn, "Ajout de clé duppliquée possible")

    End Sub
End Class



