Imports Microsoft.VisualStudio.TestTools.UnitTesting

Imports vini_DB



'''<summary>
'''Classe de test pour FicheTechniqueFournTest, destinée à contenir tous
'''les tests unitaires FicheTechniqueFournTest
'''</summary>
<TestClass()> _
Public Class T1050_FicheTechniqueFourn
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
    '''Test pour Equals
    '''</summary>
    <TestMethod()> _
    Public Sub EqualsTest()
        Dim pfrnId As Long = 15879 ' TODO: initialisez à une valeur appropriée
        Dim target As FicheTechniqueFourn = New FicheTechniqueFourn(pfrnId) ' TODO: initialisez à une valeur appropriée
        Dim obj As Object = Nothing ' TODO: initialisez à une valeur appropriée
        Assert.AreEqual(target, target)
    End Sub

    '''<summary>
    '''Test pour Constructeur FicheTechniqueFourn
    '''</summary>
    <TestMethod()> _
    Public Sub FicheTechniqueFournConstructorTest()
        Dim pfrnId As Integer = 10 ' TODO: initialisez à une valeur appropriée
        Dim target As FicheTechniqueFourn = New FicheTechniqueFourn(pfrnId)
        Dim obj As FicheTechniqueFourn
        Assert.AreEqual(pfrnId, target.id, "Constructeur")
        target.txtPresDomaine = "TxtPresDomaine"
        Assert.AreEqual("TxtPresDomaine", target.txtPresDomaine, "TxtPresDomaine")
        target.txtPresFournisseur = "txtPresFournisseur"
        Assert.AreEqual("txtPresFournisseur", target.txtPresFournisseur, "txtPresFournisseur")
        target.txtTerroir = "txtTerroir"
        Assert.AreEqual("txtTerroir", target.txtTerroir, "txtTerroir")

        'Sauvegarde de l'objet (Insert)
        Assert.IsTrue(target.Save())

        obj = New FicheTechniqueFourn(pfrnId)
        Assert.IsTrue(obj.load(), "Load")
        Assert.IsTrue(target.Equals(obj), "Egalité après rechargement")

        target.specialite = "1"
        'Sauvegarde de l'objet (UPDATE)
        Assert.IsTrue(target.Save())

        obj = New FicheTechniqueFourn(pfrnId)
        Assert.IsTrue(obj.load(), "Load")
        Assert.IsTrue(target.Equals(obj), "Egalité après rechargement UPDATE")

        'Suppression de l'objet
        target.bDeleted = True
        Assert.IsTrue(target.Save())

        obj = New FicheTechniqueFourn(pfrnId)
        Assert.IsTrue(obj.load(), "Load")
        Assert.AreEqual(0, obj.id, "ID à 0 après delete")


    End Sub
    '''<summary>
    '''Test pour Constructeur FicheTechniqueFourn
    '''</summary>
    <TestMethod()> _
    Public Sub LoadImages()

        Dim pfrnId As Integer = 10 ' TODO: initialisez à une valeur appropriée
        Dim target As FicheTechniqueFourn = New FicheTechniqueFourn(pfrnId)
        Dim obj As FicheTechniqueFourn
        Dim nOri As Integer
        Dim nTarg As Integer
        target.txtPresDomaine = "Chateau Trimoulet"
        target.txtPresFournisseur = "Michel jean"
        target.txtTerroir = "Nous considérons que le vin est un produit vivant qui tire ses caractéristiques du sol sur lequel les vignes qui le produisent sont enracinées."
        target.loadIMGDomaine(TestContext.DeploymentDirectory & "/imgRaisin.jpg")
        target.loadIMGFournisseur(TestContext.DeploymentDirectory & "/Image2.jpg")

        'Sauvegarde de l'objet (Insert)
        Assert.IsTrue(target.Save())

        obj = New FicheTechniqueFourn(pfrnId)
        Assert.IsTrue(obj.load(), "Load")
        Assert.AreNotEqual(0, obj.IMGDomaine.Image.Length, "Image domaine non chargée")
        Assert.AreNotEqual(0, obj.IMGFournisseur.Image.Length, "Image Fournisseur non chargée")

        'Vérification du bon chargement
        obj.IMGDomaine.SaveToFile(TestContext.DeploymentDirectory & "/IMGdomaine.jpg")
        nOri = My.Computer.FileSystem.ReadAllBytes(TestContext.DeploymentDirectory & "/imgRaisin.jpg").Length
        nTarg = My.Computer.FileSystem.ReadAllBytes(TestContext.DeploymentDirectory & "/IMGdomaine.jpg").Length
        Assert.AreEqual(nOri, nTarg, "Les fichiers ImgRaisin et ImgDomaine ne sont pas égaux")
        obj.IMGFournisseur.SaveToFile(TestContext.DeploymentDirectory & "/IMGFourn.jpg")
        nOri = My.Computer.FileSystem.ReadAllBytes(TestContext.DeploymentDirectory & "/Image2.jpg").Length
        nTarg = My.Computer.FileSystem.ReadAllBytes(TestContext.DeploymentDirectory & "/IMGFourn.jpg").Length
        Assert.AreEqual(nOri, nTarg, "Les fichiers Image2 et ImgFourn ne sont pas égaux")

        'Mise à jour d'une Image
        target.loadIMGDomaine(TestContext.TestDeploymentDir & "/IMGvignes.jpg")
        'Sauvegarde de l'objet (Update)
        Assert.IsTrue(target.Save())
        'Rechargement de l'objet
        obj = New FicheTechniqueFourn(pfrnId)
        Assert.IsTrue(obj.load(), "Load")
        Assert.AreNotEqual(0, obj.IMGDomaine.Image.Length, "Image domaine non chargée")
        Assert.AreNotEqual(0, obj.IMGFournisseur.Image.Length, "Image Fournisseur non chargée")
        'Vérification du bon chargement
        obj.IMGDomaine.SaveToFile(TestContext.DeploymentDirectory & "/IMGdomaine2.jpg")
        nOri = My.Computer.FileSystem.ReadAllBytes(TestContext.DeploymentDirectory & "/IMGVignes.jpg").Length
        nTarg = My.Computer.FileSystem.ReadAllBytes(TestContext.DeploymentDirectory & "/IMGdomaine2.jpg").Length
        Assert.AreEqual(nOri, nTarg, "Les fichiers ImgVignes et ImgDomaine2 ne sont pas égaux")



        'Suppression de l'objet
        target.bDeleted = True
        Assert.IsTrue(target.Save())

    End Sub




End Class
