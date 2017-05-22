
Public Class textBoxCurrency
    Inherits textBoxNumeric
    Public Overrides Property Text() As String
        Get
            If IsNumeric(MyBase.Text) Then
                Return CDec(MyBase.Text)
            Else
                Return 0.0
            End If
        End Get
        Set(ByVal Value As String)
            Dim ValueDec As Decimal
            Try
                ValueDec = Value
                MyBase.Text = ValueDec.ToString("C")
            Catch ex As Exception
                MyBase.Text = Value
            End Try
        End Set
    End Property
End Class
