'Imports System
'imports Microsoft.VisualStudio.TestTools.UnitTesting
'Imports vini_DB
'
'    <TestClass()> Public Class test_Fax
'        Inherits test_Base

'        <TestInitialize()> Public Overrides Sub TestInitialize()
'            MyBase.TestInitialize()
'        End Sub
'        <TestCleanup()> Public Overrides Sub TestCleanup()
'            MyBase.TestCleanup()
'        End Sub
'        <TestMethod()> Public Sub T10_Object()
'            Dim objFax As clsFax
'            Dim objClient As Client
'            Dim nJob As Integer

'            System.IO.File.CreateText("adel.txt")

'            objClient = New Client("TFX01", "TextFax")
'            objClient.rs = "TextFax RS"

'            objClient.AdresseLivraison.fax = "0299555277"
'            Assert.IsTrue(objClient.save(), "Creation du client")

'            objFax = New clsFax
'            nJob = objFax.sendFax("Moi", "0123456789", "TestAvecPagedeGarde", "Ceci est une note", True, "TOTO.DOC", "02 99 55 52 77", objClient, True)
'            Assert.IsTrue(nJob <> 0)
'            Assert.IsTrue(MsgBox("Vérifiez que le dernier message dans la file d'attente fax soit correct et contient une page de garde", MsgBoxStyle.YesNo) = MsgBoxResult.Yes)

'            nJob = objFax.sendFax("Moi", "0123456789", "TestAvecPagedeGarde", "Ceci est une note", False, "TOTO.DOC", "02 99 55 52 77", objClient, True)
'            Assert.IsTrue(nJob <> 0)
'            Assert.IsTrue(MsgBox("Vérifiez que le dernier message dans la file d'attente fax soit correct et ne contient pas de page de garde", MsgBoxStyle.YesNo) = MsgBoxResult.Yes)

'            System.IO.File.Delete("adel.txt")
'            objClient.bDeleted = True
'            objClient.save()
'        End Sub

'    End Class
'

