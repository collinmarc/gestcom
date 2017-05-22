Imports Microsoft.VisualStudio.TestTools.UnitTesting


<TestClass()> _
Public Class First
    <TestMethod()> Public Sub TestOK()
    End Sub
    <TestMethod()> Public Sub TestNOK()
        Assert.AreEqual(1, 2)
    End Sub
    <TestMethod(), Ignore()> Public Sub TestIgnore()
        Assert.AreEqual(1, 2)
    End Sub
End Class

