

Namespace dsSaw8TableAdapters
    
    Partial Public Class DataTable1TableAdapter
    End Class
End Namespace

Namespace dsSaw8TableAdapters
    
    Partial Public Class QuoteRealLUTableAdapter
    End Class
End Namespace

Namespace dsSaw8TableAdapters
    
    Partial Public Class qutlinepriceTableAdapter
    End Class
End Namespace

Partial Class dsSaw8
    Partial Class SpecRegFollowUpDataTable

        Private Sub SpecRegFollowUpDataTable_ColumnChanging(sender As System.Object, e As System.Data.DataColumnChangeEventArgs) Handles Me.ColumnChanging
            If (e.Column.ColumnName = Me.QtyColumn.ColumnName) Then
                'Add user code here
            End If

        End Sub

    End Class

    Partial Class projectcustDataTable

    End Class

    Partial Class quoteDataTable

    End Class

    Partial Class quotelinesDataTable

    End Class

End Class

Namespace dsSaw8TableAdapters
    
    Partial Public Class quoteTableAdapter
    End Class
End Namespace
