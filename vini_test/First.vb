Imports Microsoft.VisualStudio.TestTools.UnitTesting


<TestClass()> _
Public Class First
    <TestMethod()> Public Sub TestOK()
    End Sub
    <TestMethod(), Ignore()> Public Sub TestNOK()
        Assert.AreEqual(1, 2)
    End Sub
    <TestMethod(), Ignore()> Public Sub TestIgnore()
        Assert.AreEqual(1, 2)
    End Sub

    <TestMethod(), Ignore()> Public Sub TestDate()
        Dim dDeb As Date
        Assert.AreEqual("octobre 2018", Now.ToString("MMMM yyyy"))
        dDeb = CDate(Now.ToString("MMMM yyyy"))
        Assert.AreEqual(#10/1/2018#, dDeb)
    End Sub

End Class

