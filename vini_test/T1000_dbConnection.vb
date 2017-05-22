'Test de la classe dbConnection
Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports vini_DB



<TestClass()> Public Class T1000_dbConnection
    Inherits test_Base
    <TestMethod()> Public Sub T1Connect_Disconnect()
        Assert.IsTrue(Fournisseur.shared_connect(), Fournisseur.getErreur())
        Assert.IsFalse(Fournisseur.bErreur)
        Assert.IsTrue(Fournisseur.shared_isConnected())
        Assert.IsTrue(Fournisseur.shared_disconnect())
        Assert.IsFalse(Fournisseur.bErreur)
        Assert.IsFalse(Fournisseur.shared_isConnected())
    End Sub
    '<TestMethod()> Public Sub T2TestConnectNOK()
    '    If Persist.shared_isConnected Then
    '        Persist.shared_disconnect()
    '    End If
    '    Assert.IsFalse(Persist.shared_connect("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\MesDocuments\NewCo\vinicom\V2\vinicom.mdb2"))
    '    Assert.IsFalse(Fournisseur.shared_isConnected())
    '    Assert.IsTrue(Len(Fournisseur.getErreur()) > 0)
    'End Sub
    <TestMethod(), Ignore()> Public Sub T3TestnConnect()
        Assert.AreEqual(Persist.nConnection, 0, "NConnextion")
        Assert.IsTrue(Persist.shared_connect(), Persist.getErreur())
        Assert.AreEqual(Persist.nConnection, 1, "nConnection =1")
        Persist.shared_connect()
        Assert.AreEqual(Persist.nConnection, 2, "nConnection =2")
        Persist.shared_disconnect()
        Assert.AreEqual(Persist.nConnection, 1, "nConnection =1")
        Assert.IsTrue(Persist.shared_isConnected(), "On reste connecté")
        Persist.shared_disconnect()
        Assert.AreEqual(Persist.nConnection, 0, "nConnection =0")
        Assert.IsFalse(Persist.shared_isConnected(), "On est deconnecté")

    End Sub

End Class


