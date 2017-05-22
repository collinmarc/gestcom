
'Test de la classe dbConnection
Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports vini_DB



<TestClass()> Public Class test_V192
    Inherits test_Base
    '"Tester en manuel"
    <TestMethod(), Ignore()> Public Sub T10_Fax()
        Dim objCmd As CommandeClient
        Dim i As Integer
        Dim objSCMD As SousCommande
        Dim strFileName As String

        '            Persist.shared_connect()
        For i = 1 To 20
            'Console.Out.WriteLine("Boucle" & i)
            objCmd = CommandeClient.createandload(938)
            objCmd.loadcolLignes()
            objCmd.LoadColSousCommande()
            objCmd.save()
            For Each objSCMD In objCmd.colSousCommandes
                'Console.Out.WriteLine("Scmd id = " & objSCMD.id & "code =" & objSCMD.code)
                'Console.Out.Flush()
                strFileName = "D:\TEMP\SCMD\" & objSCMD.code & ".doc"
                If objSCMD.faxer("D:\MesDocuments\NewCo\vinicom\V3\gestcom\vini_app\", "Moi", "0299555555", "Bon à Facturer", "", "False", strFileName, "0299556339") Then
                    objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDFaxer)
                Else
                    Assert.IsTrue(False, objCmd.getErreur())
                End If
                objCmd.save()
            Next objSCMD
        Next i
        '            Persist.shared_disconnect()

    End Sub


End Class



