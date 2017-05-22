Partial Class dsVinicom
    Partial Class CONSTANTESDataTable

        Private Sub CONSTANTESDataTable_ColumnChanging(ByVal sender As System.Object, ByVal e As System.Data.DataColumnChangeEventArgs) Handles Me.ColumnChanging
            If (e.Column.ColumnName = Me.CST_SOC_ADRESSE_RUE2Column.ColumnName) Then
                'Ajoutez le code utilisateur ici
            End If

        End Sub

        Private Sub CONSTANTESDataTable_CONSTANTESRowChanging(ByVal sender As System.Object, ByVal e As CONSTANTESRowChangeEvent) Handles Me.CONSTANTESRowChanging

        End Sub

    End Class

End Class
