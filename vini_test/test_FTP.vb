'Test de la classe ftpwininet
Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports vini_DB

<TestClass(), Ignore()> Public Class test_FTP
    Inherits test_Base
    Private m_FTP As clsFTPVinicom
    <TestInitialize()> Public Overrides Sub TestInitialize()
        MyBase.TestInitialize()
        m_FTP = New clsFTPVinicom("ftp.phpnet.org", "marccollin_vnctest", "vinicom")
        m_FTP.setWaitParam(1, 5)
    End Sub
    <TestMethod(), Ignore()>
    Public Sub T10_LOCK()



        Debug.Print(Now() + "Begin of Test")
        Assert.IsTrue(m_FTP.ftpConnection.FtpFileExists("semaphore"))

        Debug.Print(Now() + "End of Test")
    End Sub 'T10_LOCK
    <TestMethod(), Ignore()> Public Sub T40_FTPVINICOMUPLOAD()
        Dim nFile As Integer

        Assert.AreEqual(m_FTP.ftpConnection.Hostname, "ftp://ftp.phpnet.org", "HostName")
        Assert.AreEqual(m_FTP.ftpConnection.Username, "marccollin_vnctest", "Username")
        Assert.AreEqual(m_FTP.ftpConnection.Password, "vinicom", "Password")
        Assert.AreEqual(m_FTP.remoteDir, Param.getConstante("FTP_REMOTEDIR").Replace(".", ""), "RemoteDir")
        Assert.AreEqual(m_FTP.strLockFromFileName, Param.getConstante("FTP_LOCKFROMFILENAME"), "FTP_LOCKFROMFILENAME")
        Assert.AreEqual(m_FTP.strLockToFileName, Param.getConstante("FTP_LOCKTOFILENAME"), "FTP_LOCKTOFILENAME")
        'm_FTP.connect()
        'Creation d'un repertoire
        My.Computer.FileSystem.CreateDirectory("./T40")
        nFile = FreeFile()
        FileOpen(nFile, "./T40/adel1.txt", OpenMode.Output, OpenAccess.Write)
        WriteLine(nFile, "test")
        FileClose(nFile)
        nFile = FreeFile()
        FileOpen(nFile, "./T40/adel2.txt", OpenMode.Output, OpenAccess.Write)
        WriteLine(nFile, "test")
        FileClose(nFile)
        nFile = FreeFile()
        FileOpen(nFile, "./T40/adel3.txt", OpenMode.Output, OpenAccess.Write)
        WriteLine(nFile, "test")
        FileClose(nFile)


        Debug.Print(Now() + "Upload Files")
        Assert.IsTrue(m_FTP.uploadFromDir("./T40"), "Upload From Dir")
        Debug.Print(Now() + "Test Files")
        Assert.IsTrue(m_FTP.ftpConnection.FtpFileExists("adel1.txt"), "Remote File 1 exists")
        Assert.IsTrue(m_FTP.ftpConnection.FtpFileExists("adel2.txt"), "Remote File 3 exists")
        Assert.IsTrue(m_FTP.ftpConnection.FtpFileExists("adel3.txt"), "Remote File 2 exists")

        My.Computer.FileSystem.DeleteDirectory("./T40", FileIO.DeleteDirectoryOption.DeleteAllContents)
        Debug.Print(Now() + "End of Test")
    End Sub 'T50_FTPVINICOMUPLOAD
    <TestMethod(), Ignore()> Public Sub T50_FTPVINICOMDOWNLOAD()
        Dim nFile As Integer

        'm_FTP.connect()
        'Creation d'un repertoire
        My.Computer.FileSystem.CreateDirectory("./T50")

        nFile = FreeFile()
        FileOpen(nFile, "T50/toVinicom.csv", OpenMode.Output, OpenAccess.Write)
        WriteLine(nFile, "test")
        FileClose(nFile)

        Assert.IsTrue(m_FTP.uploadFromDir("T50"), "Upload From Dir")
        Assert.IsTrue(m_FTP.ftpConnection.FtpFileExists("toVinicom.csv"), "Remote File toVinicom.csv exists")

        'Suppression des fichiers en local
        If System.IO.File.Exists("T50/toVinicom.csv") Then
            System.IO.File.Delete("T50/toVinicom.csv")
        End If

        Assert.IsTrue(m_FTP.downloadToDir("T50"), "DownloadToDir")
        Assert.IsTrue(System.IO.File.Exists("T50/toVinicom.csv"), "Local File toVinicom.csv exists")
        My.Computer.FileSystem.DeleteDirectory("./T50", FileIO.DeleteDirectoryOption.DeleteAllContents)

    End Sub 'T50_FTPVINICOMDOWNLOAD
    <TestMethod(), Ignore()> Public Sub T60_DeteRemoteFile()
        Dim nFile As Integer

        My.Computer.FileSystem.CreateDirectory("./T60")
        nFile = FreeFile()
        FileOpen(nFile, "./T60/adel.txt", OpenMode.Output, OpenAccess.Write)
        WriteLine(nFile, "TEST")
        FileClose(nFile)
        Assert.IsTrue(m_FTP.ftpConnection.Upload("./T60/adel.txt", "adel.txt"), "Upoad File")

        '        Assert.IsFalse(m_FTP.bDeleteFile("adel1.txt"), "Delete unknown file")
        Assert.IsTrue(m_FTP.ftpConnection.FtpDelete("adel.txt"), "Delete unknown file")
        Assert.IsFalse(m_FTP.ftpConnection.FtpFileExists("adel.txt"), "Remote file doesn't exist")

        My.Computer.FileSystem.DeleteDirectory("./T60", FileIO.DeleteDirectoryOption.DeleteAllContents)

        'Console.Out.Write(m_FTP.sListDir)
    End Sub
End Class

