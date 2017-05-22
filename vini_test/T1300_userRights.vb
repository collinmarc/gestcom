'Test de la classe dbConnection
Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports vini_DB



<TestClass()> Public Class T1300_UserRights

    <TestInitialize()> Public Sub TestInitialize()
    End Sub
    <TestCleanup()> Public Sub TestCleanup()
    End Sub
    'Ignore("a tester en manuel")
    <TestMethod(), Ignore()> Public Sub T10_Object()

        Dim objuser As aut_user
        Dim objright As aut_right

        objuser = New aut_user("Moi" & Now().DayOfYear & Now().Year, "password", vncEnums.userRole.ADMIN)
        objuser.ajouteDroit("0-1", True, "Menu 01")
        objuser.ajouteDroit("0-2", False, "Menu 02")
        Assert.IsTrue(objuser.Save(), "objuser.save()")
        Assert.IsTrue(objuser.id <> 0, "Id <> 0")

        objuser.code = "Toi" & Now().DayOfYear & Now().Year
        objright = objuser.colRigths(1)
        objright.text = objright.text & "updated"
        objright.droit = Not objright.droit
        Assert.IsTrue(objuser.Save)


    End Sub

    <TestMethod()> Public Sub T20_LISTE()

        Dim objuser As aut_user
        Dim oCol As Collection

        Persist.shared_connect()
        oCol = objuser.listeUSERS()
        For Each objuser In oCol
            Assert.IsTrue(objuser.colRigths.Count > 0, "Y PAS DE DROITS")
            Console.WriteLine(objuser.shortResume)
        Next
        Persist.shared_disconnect()
    End Sub
End Class



