Imports Microsoft.VisualStudio.TestTools.UnitTesting

Imports vini_DB



'''<summary>
'''Classe de test pour FicheTechniqueFournTest, destinée à contenir tous
'''les tests unitaires FicheTechniqueFournTest
'''</summary>
<TestClass()> _
Public Class T0100_vini_base
    Inherits test_Base

#Region "TestContext"
    Private testContextInstance As TestContext

    '''<summary>
    '''Obtient ou définit le contexte de test qui fournit
    '''des informations sur la série de tests active ainsi que ses fonctionnalités.
    '''</summary>
    Public Property TestContext() As TestContext
        Get
            Return testContextInstance
        End Get
        Set(value As TestContext)
            testContextInstance = value
        End Set
    End Property
#End Region

#Region "Attributs de tests supplémentaires"
    '
    'Vous pouvez utiliser les attributs supplémentaires suivants lorsque vous écrivez vos tests :
    '
    'Utilisez ClassInitialize pour exécuter du code avant d'exécuter le premier test dans la classe
    <ClassInitialize()> _
    Public Shared Sub MyClassInitialize(ByVal testContext As TestContext)
    End Sub
    '
    'Utilisez ClassCleanup pour exécuter du code après que tous les tests ont été exécutés dans une classe
    <ClassCleanup()> _
    Public Shared Sub MyClassCleanup()
    End Sub
    '
    'Utilisez TestInitialize pour exécuter du code avant d'exécuter chaque test
    <TestInitialize()> _
    Public Sub MyTestInitialize()
        MyBase.TestInitialize()
    End Sub
    '
    'Utilisez TestCleanup pour exécuter du code après que chaque test a été exécuté
    <TestCleanup()> _
    Public Sub MyTestCleanup()
        MyBase.TestCleanup()
    End Sub

#End Region



    '''<summary>
    '''Test pour Constructeur vini_Image
    '''</summary>
    <TestMethod()> _
    Public Sub T01_Constructeur()
        Dim pfrnId As Integer = 10 ' TODO: initialisez à une valeur appropriée
        Dim oImage As vini_image = New vini_image()
        Dim obj As vini_image
        Dim nId As Integer

        oImage.IdObject = 10
        oImage.Type = "IMG"
        oImage.Num = 5
        Assert.AreEqual(10, oImage.IdObject)
        Assert.AreEqual("IMG", oImage.Type)
        Assert.AreEqual(5, oImage.Num)

        'Sauvegarde de l'objet (Insert)
        Assert.IsTrue(oImage.Save())
        Assert.AreNotEqual(0, oImage.id, "Id ne doit pas être 0 après un insert")
        Assert.AreNotEqual(-1, oImage.id, "Id ne doit pas être -1 après un insert")
        nId = oImage.id

        obj = New vini_image()
        Assert.IsTrue(obj.load(nId), "Load")
        Assert.IsTrue(oImage.Equals(obj), "Egalité après rechargement")

        oImage.Desc = "descrption"
        'Sauvegarde de l'objet (UPDATE)
        Assert.IsTrue(oImage.Save())
        Assert.AreEqual(nId, oImage.id, "Id ne doit pas être modifié après un update")

        obj = New vini_image()
        Assert.IsTrue(obj.load(nId), "Load")
        Assert.IsTrue(oImage.Equals(obj), "Egalité après rechargement UPDATE")
        Assert.AreEqual(oImage.Desc, obj.Desc)

        'Suppression de l'objet
        oImage.bDeleted = True
        Assert.IsTrue(oImage.Save())

        obj = New vini_image()
        Assert.IsTrue(obj.load(nId), "Load")
        Assert.AreEqual(0, obj.id, "ID à 0 après delete")


    End Sub

    '''<summary>
    '''Test pour LoadImage et Save Image
    '''</summary>
    <TestMethod()> _
    Public Sub T10_LoadSaveImage()
        Dim oImage As vini_Image = New vini_Image()
        Dim obj As vini_Image
        Dim nId As Integer

        oImage.IdObject = 10
        oImage.Type = "IMG"
        oImage.Num = 1
        'chargement de l'image
        Assert.IsTrue(oImage.LoadFromFile(TestContext.DeploymentDirectory & "/Image2.jpg"), "Load file")
        Assert.AreNotEqual(0, oImage.Image.Length, "Image Chargée")
        'Sauvegarde de l'image
        Assert.IsTrue(oImage.SaveToFile(TestContext.DeploymentDirectory & "/Image3.jpg"), "SaveToFile")
        Assert.IsTrue(My.Computer.FileSystem.FileExists(TestContext.DeploymentDirectory & "/Image3.jpg"), "File Image3.jpg n'existe pas")

        Dim nLen3 As Integer
        Dim nLen2 As Integer
        nLen3 = My.Computer.FileSystem.ReadAllBytes(TestContext.DeploymentDirectory & "/Image3.jpg").Length
        nLen2 = My.Computer.FileSystem.ReadAllBytes(TestContext.DeploymentDirectory & "/Image2.jpg").Length

        Assert.AreEqual(nLen2, nLen3, "Comparaison des tailles de fichier")

    End Sub
    '''<summary>
    '''Test pour Save abd Load Image en baseDeDonnées
    '''</summary>
    <TestMethod()> _
    Public Sub T20_LoadDBSave()
        Dim oImage As vini_Image = New vini_Image()
        Dim obj As vini_Image
        Dim nId As Integer

        oImage.IdObject = 10
        oImage.Type = "IMG"
        oImage.Num = 1
        'chargement de l'image
        Assert.IsTrue(oImage.LoadFromFile(TestContext.DeploymentDirectory & "/Image2.jpg"), "Load file")
        'Sauvegarde de l'Objet en base de donnée
        Assert.IsTrue(oImage.Save(), "Sauvegarde en base")
        nId = oImage.id

        'Rechargement de l'image
        obj = New vini_Image()
        Assert.IsTrue(obj.load(nId), "Rechargement de l'objet")
        Assert.AreNotEqual(0, obj.Image.Length, "Image chargée")
        obj.SaveToFile(TestContext.DeploymentDirectory & "/Image3.jpg")

        Dim nLen3 As Integer
        Dim nLen2 As Integer
        nLen2 = My.Computer.FileSystem.ReadAllBytes(TestContext.DeploymentDirectory & "/Image2.jpg").Length
        nLen3 = My.Computer.FileSystem.ReadAllBytes(TestContext.DeploymentDirectory & "/Image3.jpg").Length

        Assert.AreEqual(nLen2, nLen3, "Comparaison des tailles de fichier")

    End Sub
End Class
