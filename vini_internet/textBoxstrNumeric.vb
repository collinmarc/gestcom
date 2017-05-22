' Test Box contenant du texte devant être numéric mais la mise en forme ne change pas

Public Class textBoxStrNumeric
    Inherits TextBox

    Public Overrides Property Text() As String
        Get
            If IsNumeric(MyBase.Text) Then
                Return MyBase.Text
            Else
                Return 0.0
            End If
        End Get
        Set(ByVal Value As String)
            Try
                If IsNumeric(Value) Then
                    MyBase.Text = Value
                Else
                    If Value.Equals("") Then
                        MyBase.Text = "0"
                    Else
                        Throw New InvalidCastException
                    End If
                End If
            Catch ex As Exception
            End Try
        End Set
    End Property

    Private Sub textBoxNumeric_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Validating
        Dim ofrmVinicom As FrmVinicom
        If Len(Me.Text) = 0 Then
            Me.Text = 0
        End If
        If Not IsNumeric(Me.Text) And Len(Me.Text) > 0 Then
            e.Cancel = True
            Me.SelectAll()
            Try
                ofrmVinicom = Me.FindForm()
                ofrmVinicom.DisplayError(Me.Name, "Champs Numérique Attendu")
            Catch ex As Exception

            End Try
        Else
            Try
                ofrmVinicom = Me.FindForm()
                ofrmVinicom.DisplayError("", "")
            Catch ex As Exception

            End Try

        End If
    End Sub

    Public Function getTextlong() As Long
        Dim nReturn As Long
        Try
            nReturn = CLng(Text)
        Catch ex As Exception
            nReturn = 0
        End Try
        Return nReturn
    End Function
    Public Function getTextDouble() As Double
        Dim nReturn As Double
        Try
            nReturn = CDbl(Text)
        Catch ex As Exception
            nReturn = 0
        End Try
        Return nReturn
    End Function
End Class
